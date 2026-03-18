Attribute VB_Name = "extractCCtoExcel"
' ==========================================================
' vbaWord: Content Controls to Excel Summarizer
' Author: Azi-lzb
' Description: Extracts data from Content Controls into "output_cc" sheet.
' ==========================================================

Sub SummarizeContentControlsToExcel()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim iRow As Long, iCol As Integer
    Dim cc As Object
    Dim ws As Worksheet
    Dim tagList As Object
    Dim currentTag As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("output_cc")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "output_cc"
    Else
        ws.Cells.Clear
    End If
    
    Set tagList = CreateObject("Scripting.Dictionary")
    
    On Error Resume Next
    Set wdApp = CreateObject("Word.Application")
    On Error GoTo 0
    
    If wdApp Is Nothing Then
        MsgBox "无法启动 Word。", vbCritical
        Exit Sub
    End If
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "请选择包含内容控件的 Word 文档"
        .Filters.Add "Word Documents", "*.docx; *.docm", 1
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            ws.Cells(1, 1).Value = "源文件名"
            iRow = 2
            iCol = 2
            
            For Each fileItem In .SelectedItems
                Set wdDoc = wdApp.Documents.Open(fileName:=fileItem, ReadOnly:=True, Visible:=False)
                ws.Cells(iRow, 1).Value = Dir(fileItem)
                
                For Each cc In wdDoc.ContentControls
                    currentTag = cc.Tag
                    If currentTag = "" Then currentTag = cc.Title
                    If currentTag = "" Then currentTag = "未命名控件"
                    
                    If Not tagList.Exists(currentTag) Then
                        tagList.Add currentTag, iCol
                        ws.Cells(1, iCol).Value = currentTag
                        iCol = iCol + 1
                    End If
                    
                    ws.Cells(iRow, tagList(currentTag)).Value = cc.Range.Text
                Next cc
                
                wdDoc.Close SaveChanges:=False
                iRow = iRow + 1
            Next fileItem
            
            ws.Columns.AutoFit
            ws.Activate
            MsgBox "内容控件汇总完成！", vbInformation
        End If
    End With
    
    wdApp.Quit
    Set wdDoc = Nothing: Set wdApp = Nothing
End Sub
