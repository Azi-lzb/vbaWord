Attribute VB_Name = "extractCCtoWord"
' ==========================================================
' vbaWord: Content Controls to Word Summarizer
' Updated: Use modConfig for settings
' ==========================================================

Sub SummarizeCCToWord()
    Dim wdApp As Object, wdDoc As Object, targetDoc As Object
    Dim fd As FileDialog, fileItem As Variant, cc As Object, dict As Object
    Dim key As Variant, currentTag As String, docPwd As String, priorityMode As String
    Dim pType As Long, isUnprotected As Boolean
    
    docPwd = modConfig.GetDocPassword()
    priorityMode = modConfig.GetPriorityMode()
    
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If wdApp Is Nothing Then
        MsgBox "无法启动 Word。", vbCritical
        Exit Sub
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "请选择内容控件文档 (汇总至 Word)"
        .Filters.Add "Word Documents", "*.docx; *.docm", 1
        
        If .Show = -1 Then
            For Each fileItem In .SelectedItems
                Set wdDoc = wdApp.Documents.Open(Filename:=fileItem, ReadOnly:=False, Visible:=False)
                
                pType = wdDoc.ProtectionType: isUnprotected = False
                If pType <> -1 Then
                    On Error Resume Next
                    wdDoc.Unprotect Password:=docPwd
                    If Err.Number = 0 Then isUnprotected = True
                    On Error GoTo 0
                End If
                
                For Each cc In wdDoc.ContentControls
                    currentTag = GetEffectiveTag(cc, priorityMode)
                    
                    If Not dict.Exists(currentTag) Then
                        dict.Add currentTag, "【" & Dir(fileItem) & "】: " & cc.Range.Text & "; "
                    Else
                        dict(currentTag) = dict(currentTag) & "【" & Dir(fileItem) & "】: " & cc.Range.Text & "; "
                    End If
                Next cc
                
                If pType <> -1 And isUnprotected Then
                    On Error Resume Next
                    wdDoc.Protect Type:=pType, NoReset:=True, Password:=docPwd
                    On Error GoTo 0
                End If
                
                wdDoc.Close SaveChanges:=False
            Next fileItem
            
            Set targetDoc = wdApp.Documents.Add
            wdApp.Visible = True
            With targetDoc.Range
                For Each key In dict.Keys
                    .InsertAfter key & vbCrLf & dict(key) & vbCrLf & vbCrLf
                Next key
            End With
            MsgBox "Word 汇总报告已生成！", vbInformation
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
