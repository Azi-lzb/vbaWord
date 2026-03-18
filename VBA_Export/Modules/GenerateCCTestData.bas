Attribute VB_Name = "GenerateCCTestData"
' ==========================================================
' Content Control Test Data Generator
' Author: Azi-lzb
' Logic: Create -> Unprotect -> Fill Content -> Protect -> Save
' ==========================================================

Sub CreateCCSampleDocs()
    Dim i As Integer
    Dim doc As Object ' Word.Document
    Dim folderPath As String
    Dim cc As Object ' Word.ContentControl
    Dim rng As Object ' Word.Range
    Dim wdApp As Object
    
    ' 设置输出目录路径
    folderPath = "C:\Users\AZI\Desktop\VibeCode项目\vbaWord\"
    
    ' 1. 初始化 Word 实例
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    ' 2. 检查目录有效性
    If Dir(folderPath, vbDirectory) = "" Then
        On Error Resume Next
        folderPath = ThisWorkbook.Path & "\"
        On Error GoTo 0
    End If

    ' 3. 循环生成文档
    For i = 1 To 3
        Set doc = wdApp.Documents.Add

        ' --- 确保文档未受保护以进行编辑 ---
        On Error Resume Next
        If doc.ProtectionType <> -1 Then ' -1 = wdNoProtection
            doc.Unprotect Password:=""
        End If
        On Error GoTo 0

        ' A. 插入姓名 (文本控件)
        Set rng = doc.Content
        rng.Collapse 0 ' wdCollapseEnd
        rng.InsertAfter "姓名: "
        rng.Collapse 0
        Set cc = doc.ContentControls.Add(1, rng) ' 1 = wdContentControlText
        cc.Title = "姓名"
        cc.Tag = "UserName"
        cc.Range.Text = "员工_" & i

        ' B. 插入日期 (日期控件)
        Set rng = doc.Content
        rng.Collapse 0
        rng.InsertAfter vbCrLf & "填写日期: "
        rng.Collapse 0
        Set cc = doc.ContentControls.Add(4, rng) ' 4 = wdContentControlDate
        cc.Title = "日期"
        cc.Tag = "FillDate"
        cc.Range.Text = CStr(DateAdd("d", i, Date))

        ' C. 插入部门 (下拉列表)
        Set rng = doc.Content
        rng.Collapse 0
        rng.InsertAfter vbCrLf & "所属部门: "
        rng.Collapse 0
        Set cc = doc.ContentControls.Add(3, rng) ' 3 = wdContentControlDropdownList
        cc.Title = "部门"
        cc.Tag = "Dept"
        cc.DropdownListEntries.Add "财务部", "Fin"
        cc.DropdownListEntries.Add "技术部", "Tech"
        cc.DropdownListEntries.Add "市场部", "Mkt"
        cc.Range.Text = "技术部"

        ' D. 底部提示
        Set rng = doc.Content
        rng.Collapse 0
        rng.InsertAfter vbCrLf & vbCrLf & "---------------------------------------" & vbCrLf
        rng.InsertAfter "提示：该问卷文档已被设置为【仅限填写窗体】模式。"

        ' --- 重新应用保护，防止直接编辑非控件区域 ---
        ' 3 = wdAllowOnlyFormFields (旧式窗体保护也适用于内容控件填写)
        On Error Resume Next
        doc.Protect Type:=3, NoReset:=True, Password:=""
        On Error GoTo 0

        ' 4. 保存文档
        On Error Resume Next
        doc.SaveAs2 fileName:=folderPath & "CC_Sample_" & i & ".docx"
        If Err.Number <> 0 Then
            MsgBox "保存失败，请检查路径权限：" & vbCrLf & folderPath, vbExclamation
            doc.Close False
            Exit Sub
        End If
        On Error GoTo 0

        doc.Close
    Next i

    wdApp.Visible = True
    MsgBox "内容控件测试数据生成完成！" & vbCrLf & _
           "逻辑：生成 -> 解除保护 -> 填充 -> 重新保护。" & vbCrLf & _
           "路径：" & folderPath, vbInformation
End Sub
