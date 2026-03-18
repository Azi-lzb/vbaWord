Attribute VB_Name = "extractCCtoWord"
' ==========================================================
' vbaWord: Content Controls to Word Summarizer
' Author: Azi-lzb
' Description: Extracts data from Content Controls and summarizes into a new Word doc.
' ==========================================================

Sub SummarizeCCToWord()
    Dim wdApp As Object
    Dim wdDoc As Object
    Dim targetDoc As Object
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim cc As Object
    Dim dict As Object
    Dim key As Variant
    Dim currentTag As String
    
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
        .Title = "汇总内容控件至 Word"
        .Filters.Add "Word Documents", "*.docx; *.docm", 1
        .AllowMultiSelect = True
        
        If .Show = -1 Then
            For Each fileItem In .SelectedItems
                Set wdDoc = wdApp.Documents.Open(fileName:=fileItem, ReadOnly:=True, Visible:=False)
                
                For Each cc In wdDoc.ContentControls
                    currentTag = cc.Tag
                    If currentTag = "" Then currentTag = cc.Title
                    If currentTag = "" Then currentTag = "未命名"
                    
                    If Not dict.Exists(currentTag) Then
                        dict.Add currentTag, "【" & Dir(fileItem) & "】: " & cc.Range.Text & "; "
                    Else
                        dict(currentTag) = dict(currentTag) & "【" & Dir(fileItem) & "】: " & cc.Range.Text & "; "
                    End If
                Next cc
                wdDoc.Close SaveChanges:=False
            Next fileItem
            
            Set targetDoc = wdApp.Documents.Add
            wdApp.Visible = True
            With targetDoc.Range
                For Each key In dict.Keys
                    .InsertAfter key & vbCrLf & dict(key) & vbCrLf & vbCrLf
                Next key
            End With
            MsgBox "Word 汇总报告生成完成！", vbInformation
        End If
    End With
    
    Set targetDoc = Nothing: Set wdDoc = Nothing: Set wdApp = Nothing
End Sub
