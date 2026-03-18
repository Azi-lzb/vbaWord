Attribute VB_Name = "GenerateCCTestData"
' ==========================================================
' Content Control Test Data Generator
' Author: Azi-lzb
' Optimized: Save before Protect; Selection Clean-up
' ==========================================================

Sub CreateCCSampleDocs()
    Dim i As Integer
    Dim doc As Object, cc As Object, rng As Object, wdApp As Object
    Dim folderPath As String
    
    ' 1. 将路径设置为当前 Excel 文件所在的目录
    folderPath = ThisWorkbook.Path & "\"
    
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    If wdApp Is Nothing Then
        MsgBox "无法启动 Word。", vbCritical
        Exit Sub
    End If
    
    wdApp.Visible = True

    ' 2. 循环生成
    For i = 1 To 3
        Set doc = wdApp.Documents.Add

        ' A. 插入内容 (内容控件)
        Set rng = doc.Content
        rng.Collapse 0
        rng.InsertAfter "姓名: "
        rng.Collapse 0
        Set cc = doc.ContentControls.Add(1, rng) ' wdContentControlText
        cc.Title = "姓名"
        cc.Tag = "UserName"
        cc.Range.Text = "新式员工_" & i

        Set rng = doc.Content
        rng.Collapse 0
        rng.InsertAfter vbCrLf & "部门: "
        rng.Collapse 0
        Set cc = doc.ContentControls.Add(3, rng) ' wdContentControlDropdownList
        cc.Title = "部门"
        cc.DropdownListEntries.Add "财务部", "Fin"
        cc.DropdownListEntries.Add "技术部", "Tech"
        cc.Range.Text = "技术部"

        ' B. 先保存一次 (解决 6124 错误的关键)
        On Error Resume Next
        doc.SaveAs2 Filename:=folderPath & "CC_Sample_" & i & ".docx"
        On Error GoTo 0
        
        ' C. 统一限制编辑 (仅限填写窗体)
        ' 给 Word 一个极短的缓冲时间
        DoEvents
        
        ' 尝试清除文档中的光标位置以确保保护不会因 Selection 锁定而失败
        wdApp.Selection.WholeStory
        wdApp.Selection.Collapse 1 ' wdCollapseStart

        On Error Resume Next
        ' 3 = wdAllowOnlyFormFields
        doc.Protect Type:=3, NoReset:=True, Password:=""
        If Err.Number <> 0 Then Debug.Print "Error protecting CC doc: " & Err.Description
        On Error GoTo 0

        ' D. 再次保存并关闭
        doc.Save
        doc.Close
    Next i

    MsgBox "已在 Excel 目录下生成 3 份内容控件 (新式窗体) 测试文件。", vbInformation
End Sub
