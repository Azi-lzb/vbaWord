Attribute VB_Name = "extractToWord"
' ==========================================================
' vbaWord: Word Form Batch Summarizer (Output to Word)
' Updated: Use modConfig for settings
' ==========================================================

Sub SummarizeToNewWordDoc()
    Dim wdApp As Object, wdDoc As Object, targetDoc As Object
    Dim fd As FileDialog, fileItem As Variant, ff As Object, dict As Object 
    Dim key As Variant, rawFileName As String, displayName As String
    Dim qText As String, i As Integer, prevEnd As Long
    Dim pType As Long, isUnprotected As Boolean, docPwd As String
    
    docPwd = modConfig.GetDocPassword()
    
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
        .Title = "请选择旧式窗体问卷汇总至 Word"
        .Filters.Clear: .Filters.Add "Word Documents", "*.doc; *.docx; *.docm", 1
        
        If .Show = -1 Then
            For Each fileItem In .SelectedItems
                rawFileName = Dir(fileItem)
                displayName = GetMappedName(rawFileName)
                
                Set wdDoc = wdApp.Documents.Open(Filename:=fileItem, ReadOnly:=False, Visible:=False)
                
                pType = wdDoc.ProtectionType: isUnprotected = False
                If pType <> -1 Then
                    On Error Resume Next
                    wdDoc.Unprotect Password:=docPwd
                    If Err.Number = 0 Then isUnprotected = True
                    On Error GoTo 0
                End If
                
                prevEnd = 0
                On Error Resume Next
                For i = 1 To wdDoc.FormFields.Count
                    Set ff = wdDoc.FormFields(i)
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
                On Error GoTo 0
                
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
            MsgBox "汇总完成！", vbInformation
        End If
    End With
End Sub

Private Function GetMappedName(originalName As String) As String
    Dim mapWs As Worksheet: Dim i As Long: GetMappedName = originalName
    On Error Resume Next: Set mapWs = ThisWorkbook.Sheets("mapping"): On Error GoTo 0
    If Not mapWs Is Nothing Then
        For i = 2 To mapWs.Cells(mapWs.Rows.Count, "A").End(xlUp).Row
            If InStr(1, originalName, mapWs.Cells(i, 1).Value, vbTextCompare) > 0 Then
                GetMappedName = mapWs.Cells(i, 2).Value: Exit Function
            End If
        Next i
    End If
End Function
