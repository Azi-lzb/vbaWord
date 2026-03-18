Attribute VB_Name = "extractToWord"
' ==========================================================
' vbaWord: Word Form Batch Summarizer (Output to Word)
' Fixed Error 6124 (Protection Issue)
' ==========================================================

Sub SummarizeToNewWordDoc()
    Dim wdApp As Object, wdDoc As Object, targetDoc As Object
    Dim fd As FileDialog, fileItem As Variant, ff As Object, dict As Object 
    Dim key As Variant, rawFileName As String, displayName As String
    Dim qText As String, i As Integer, prevEnd As Long
    Dim pType As Long ' 记录原始保护类型
    
    On Error Resume Next
    Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If wdApp Is Nothing Then
        MsgBox "无法启动 Word。", vbCritical
        Exit Sub
    End If
    
    Set dict = CreateObject("Scripting.Dictionary")
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "请选择旧式窗体问卷汇总至 Word"
        .Filters.Add "Word Documents", "*.doc; *.docx; *.docm", 1
        
        If .Show = -1 Then
            For Each fileItem In .SelectedItems
                rawFileName = Dir(fileItem)
                displayName = GetMappedName(rawFileName)
                Set wdDoc = wdApp.Documents.Open(Filename:=fileItem, ReadOnly:=True, Visible:=False)
                
                ' --- 修复 6124 错误的关键：临时解除保护 ---
                pType = wdDoc.ProtectionType
                If pType <> -1 Then ' -1 表示 wdNoProtection
                    On Error Resume Next
                    wdDoc.Unprotect Password:="" ' 假设无密码，若有密码需修改此处
                    On Error GoTo 0
                End If
                
                prevEnd = 0
                For i = 1 To wdDoc.FormFields.Count
                    Set ff = wdDoc.FormFields(i)
                    ' 现在可以安全读取题目文本了
                    qText = wdDoc.Range(prevEnd, ff.Range.Start).Text
                    qText = Replace(qText, vbCr, ""): qText = Replace(qText, vbLf, "")
                    qText = Replace(qText, ":", ""): qText = Replace(qText, "：", "")
                    qText = Trim(qText)
                    
                    If qText = "" Then qText = ff.Name
                    
                    If Not dict.Exists(qText) Then
                        dict.Add qText, "【" & displayName & "】: " & ff.Result & "; "
                    Else
                        dict(qText) = dict(qText) & "【" & displayName & "】: " & ff.Result & "; "
                    End If
                    prevEnd = ff.Range.End
                Next i
                
                ' 恢复原始保护状态
                If pType <> -1 Then wdDoc.Protect Type:=pType, NoReset:=True
                
                wdDoc.Close SaveChanges:=False
            Next fileItem
            
            ' 生成汇总文档
            Set targetDoc = wdApp.Documents.Add
            wdApp.Visible = True
            With targetDoc.Range
                For Each key In dict.Keys
                    .InsertAfter key & vbCrLf & dict(key) & vbCrLf & vbCrLf
                Next key
            End With
            MsgBox "汇总完成！", vbInformation
        End If
    End With
    Set targetDoc = Nothing: Set wdDoc = Nothing: Set wdApp = Nothing
End Sub

' 映射函数 (保持不变)
Private Function GetMappedName(originalName As String) As String
    Dim mapWs As Worksheet: Dim i As Long
    GetMappedName = originalName
    On Error Resume Next
    Set mapWs = ThisWorkbook.Sheets("mapping")
    On Error GoTo 0
    If Not mapWs Is Nothing Then
        For i = 2 To mapWs.Cells(mapWs.Rows.Count, "A").End(xlUp).Row
            If InStr(1, originalName, mapWs.Cells(i, 1).Value, vbTextCompare) > 0 Then
                GetMappedName = mapWs.Cells(i, 2).Value
                Exit Function
            End If
        Next i
    End If
End Function
