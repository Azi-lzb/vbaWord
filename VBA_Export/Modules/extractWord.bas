Attribute VB_Name = "extractWord"
' ==========================================================
' vbaWord: Word Form to Excel Summarizer
' Author: Azi-lzb
' Description: Extracts data from legacy form fields into "output" sheet.
' ==========================================================

Sub BatchSummarizeWordForms()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim iRow As Long, iCol As Integer
    Dim ff As Object
    Dim ws As Worksheet
    Dim rawFileName As String
    Dim displayName As String
    
    ' --- 查找或创建 output 工作表 ---
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("output")
    On Error GoTo 0
    
    If ws Is Nothing Then
        ' 如果工作表不存在则新建一个
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "output"
    Else
        ' 如果工作表已存在则清空内容
        ws.Cells.Clear
    End If
    
    ' 初始化 Word
    On Error Resume Next
    Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If wdApp Is Nothing Then
        MsgBox "无法启动 Word。", vbCritical
        Exit Sub
    End If
    
    ' 选择文件
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "请选择要汇总的 Word 问卷文档"
        .Filters.Add "Word Documents", "*.doc; *.docx; *.docm", 1
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            iRow = 2
            
            For Each fileItem In .SelectedItems
                rawFileName = Dir(fileItem)
                displayName = GetMappedName(rawFileName)
                
                Set wdDoc = wdApp.Documents.Open(fileName:=fileItem, ReadOnly:=True, Visible:=False)
                
                ' 设置表头
                If iRow = 2 Then
                    ws.Cells(1, 1).Value = "源文件名 (映射后)"
                    iCol = 2
                    For Each ff In wdDoc.FormFields
                        ws.Cells(1, iCol).Value = ff.Name
                        iCol = iCol + 1
                    Next ff
                    ws.Range("A1:Z1").Font.Bold = True
                End If
                
                ' 填入数据
                ws.Cells(iRow, 1).Value = displayName
                iCol = 2
                For Each ff In wdDoc.FormFields
                    ws.Cells(iRow, iCol).Value = ff.result
                    iCol = iCol + 1
                Next ff
                
                wdDoc.Close SaveChanges:=False
                iRow = iRow + 1
            Next fileItem
            
            ws.Columns.AutoFit
            ws.Activate ' 汇总完成后自动跳转到 output 表
            MsgBox "汇总完成！结果已保存至 'output' 工作表。", vbInformation
        End If
    End With
    
    wdApp.Quit
    Set wdDoc = Nothing
    Set wdApp = Nothing
End Sub

' 映射逻辑函数
Private Function GetMappedName(originalName As String) As String
    Dim mapWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim result As String
    
    result = originalName
    On Error Resume Next
    Set mapWs = ThisWorkbook.Sheets("mapping")
    On Error GoTo 0
    
    If Not mapWs Is Nothing Then
        lastRow = mapWs.Cells(mapWs.Rows.Count, "A").End(xlUp).Row
        For i = 2 To lastRow
            If InStr(1, originalName, mapWs.Cells(i, 1).Value, vbTextCompare) > 0 Then
                result = mapWs.Cells(i, 2).Value
                Exit For
            End If
        Next i
    End If
    
    GetMappedName = result
End Function
