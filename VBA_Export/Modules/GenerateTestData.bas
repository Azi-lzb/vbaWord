Attribute VB_Name = "GenerateTestData"
' ==========================================================
' Test Data Generator for vbaWord (Legacy Form Fields)
' Optimized Flow: Create -> Save -> Protect -> Final Save
' ==========================================================

Sub CreateSampleDocs()
    Dim i As Integer
    Dim doc As Object
    Dim wdApp As Object
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
        
        ' --- 步骤 A: 插入内容 ---
        doc.Range.InsertAfter "姓名: "
        doc.FormFields.Add Range:=doc.Range(doc.Range.End - 1, doc.Range.End - 1), Type:=70
        doc.FormFields(doc.FormFields.Count).Name = "UserName"
        doc.FormFields(doc.FormFields.Count).Result = "用户_" & i
        
        doc.Range.InsertAfter vbCrLf & "反馈: "
        doc.FormFields.Add Range:=doc.Range(doc.Range.End - 1, doc.Range.End - 1), Type:=70
        doc.FormFields(doc.FormFields.Count).Result = "测试反馈 " & i
        
        ' --- 步骤 B: 先保存一次 (解决 6124 错误的关键) ---
        On Error Resume Next
        doc.SaveAs2 Filename:=folderPath & "Sample_Data_" & i & ".docx"
        On Error GoTo 0
        
        ' --- 步骤 C: 统一限制编辑 ---
        ' 3 = wdAllowOnlyFormFields
        On Error Resume Next
        doc.Protect Type:=3, NoReset:=True, Password:=""
        If Err.Number <> 0 Then Debug.Print "Protection Error: " & Err.Description
        On Error GoTo 0
        
        ' --- 步骤 D: 再次保存并关闭 ---
        doc.Save
        doc.Close
    Next i
    
    MsgBox "已在 Excel 目录下生成 3 份旧式窗体测试文件。", vbInformation
End Sub
