Attribute VB_Name = "extractCCtoWordWithTables"
' ==========================================================
' vbaWord: Content Controls to Word (With Table Support)
' Updated: Removed "【题目】:" and "--- 来源: " prefixes
' ==========================================================

Sub SummarizeCCWithTablesToWord()
    Dim wdApp As Object, wdDoc As Object, targetDoc As Object
    Dim fd As FileDialog, fileItem As Variant
    Dim cc As Object, targetRng As Object
    Dim currentTag As String, displayName As String, docPwd As String
    Dim tagFound As Boolean, pType As Long, isUnprotected As Boolean
    Dim priorityMode As String
    
    ' 获取配置
    docPwd = modConfig.GetDocPassword()
    priorityMode = modConfig.GetPriorityMode()
    
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    Set targetDoc = wdApp.Documents.Add
    wdApp.Visible = True
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "请选择内容控件文档 (含表格汇总)"
        .Filters.Add "Word Documents", "*.docx; *.docm", 1
        
        If .Show = -1 Then
            For Each fileItem In .SelectedItems
                Set wdDoc = wdApp.Documents.Open(Filename:=fileItem, ReadOnly:=False, Visible:=False)
                displayName = Dir(fileItem)
                
                pType = wdDoc.ProtectionType: isUnprotected = False
                If pType <> -1 Then
                    On Error Resume Next
                    wdDoc.Unprotect Password:=docPwd
                    If Err.Number = 0 Then isUnprotected = True
                    On Error GoTo 0
                End If
                
                For Each cc In wdDoc.ContentControls
                    ' --- 使用动态优先级识别题目 ---
                    currentTag = GetEffectiveTag(cc, priorityMode)
                    
                    Set targetRng = targetDoc.Content: tagFound = False
                    With targetRng.Find
                        .Text = currentTag
                        .Forward = True: .Wrap = 1
                        If .Execute Then tagFound = True: targetRng.Collapse 0
                    End With
                    
                    If Not tagFound Then
                        Set targetRng = targetDoc.Content: targetRng.Collapse 0
                        targetRng.InsertAfter vbCrLf & currentTag & vbCrLf
                        targetRng.Font.Bold = True: targetRng.Collapse 0
                    End If
                    
                    ' 仅保留来源文件名，移除前缀
                    targetRng.InsertAfter vbCrLf & displayName & vbCrLf
                    targetRng.Font.Bold = False: targetRng.Collapse 0
                    
                    cc.Range.Copy
                    targetRng.Paste
                    targetRng.Collapse 0: targetRng.InsertAfter vbCrLf
                Next cc
                
                If pType <> -1 And isUnprotected Then
                    On Error Resume Next
                    wdDoc.Protect Type:=pType, NoReset:=True, Password:=docPwd
                    On Error GoTo 0
                End If
                
                wdDoc.Close SaveChanges:=False
            Next fileItem
            MsgBox "含表格的汇总报告已生成！", vbInformation
        End If
    End With
    Set targetDoc = Nothing: Set wdDoc = Nothing: Set wdApp = Nothing
End Sub

Private Function GetEffectiveTag(cc As Object, mode As String) As String
    Dim result As String
    If mode = "TITLE" Then
        result = cc.Title: If result = "" Then result = cc.Tag
    Else
        result = cc.Tag: If result = "" Then result = cc.Title
    End If
    If result = "" Then result = "未命名题目"
    GetEffectiveTag = result
End Function
