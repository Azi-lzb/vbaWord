Attribute VB_Name = "extractWord"
' ==========================================================
' vbaWord: Word Form to Excel Summarizer
' Updated: Use modConfig for settings
' ==========================================================

Sub BatchSummarizeWordForms()
    Dim wdApp As Object, wdDoc As Object, fd As FileDialog
    Dim fileItem As Variant, iRow As Long, iCol As Integer
    Dim ff As Object, ws As Worksheet, rawFileName As String
    Dim displayName As String, docPwd As String
    Dim pType As Long, isUnprotected As Boolean
    
    ' 从新配置模块读取
    docPwd = modConfig.GetDocPassword()
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("output")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "output"
    Else
        ws.Cells.Clear
    End If
    
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If wdApp Is Nothing Then
        MsgBox "无法启动 Word。", vbCritical
        Exit Sub
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "请选择旧式窗体问卷 (汇总至 Excel)"
        .Filters.Add "Word Documents", "*.doc; *.docx; *.docm", 1
        
        If .Show = -1 Then
            iRow = 2
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
                
                If iRow = 2 Then
                    ws.Cells(1, 1).Value = "源文件名 (映射后)"
                    iCol = 2
                    For Each ff In wdDoc.FormFields
                        ws.Cells(1, iCol).Value = ff.Name
                        iCol = iCol + 1
                    Next ff
                    ws.Range("A1:Z1").Font.Bold = True
                End If
                
                ws.Cells(iRow, 1).Value = displayName
                iCol = 2
                For Each ff In wdDoc.FormFields
                    ws.Cells(iRow, iCol).Value = ff.Result
                    iCol = iCol + 1
                Next ff
                
                If pType <> -1 And isUnprotected Then
                    On Error Resume Next
                    wdDoc.Protect Type:=pType, NoReset:=True, Password:=docPwd
                    On Error GoTo 0
                End If
                
                wdDoc.Close SaveChanges:=False
                iRow = iRow + 1
            Next fileItem
            
            ws.Columns.AutoFit
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
