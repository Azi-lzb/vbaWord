Attribute VB_Name = "GenerateTestData"
' ==========================================================
' Test Data Generator for vbaWord (Legacy Form Fields)
' Logic: Create Word App -> Add Doc -> Insert Fields -> Protect -> Save
' ==========================================================

Sub CreateSampleDocs()
    Dim i As Integer
    Dim doc As Object
    Dim wdApp As Object
    Dim folderPath As String
    
    ' 设置项目根目录路径
    folderPath = "C:\Users\AZI\Desktop\VibeCode项目\vbaWord\"
    
    ' 1. 初始化 Word 应用程序实例 (解决 Error 424 关键)
    On Error Resume Next
    Set wdApp = GetObject(, "Word.Application")
    If wdApp Is Nothing Then
        Set wdApp = CreateObject("Word.Application")
    End If
    On Error GoTo 0
    
    If wdApp Is Nothing Then
        MsgBox "无法启动 Word，请确保已安装 Word。", vbCritical
        Exit Sub
    End If
    
    wdApp.Visible = True
    
    ' 2. 创建 3 份示例问卷
    For i = 1 To 3
        ' 使用 wdApp.Documents 而不是直接用 Documents
        Set doc = wdApp.Documents.Add
        
        ' 插入问卷内容
        doc.Range.InsertAfter "姓名: "
        ' 插入窗体域 (文本输入) - 使用常量 70 (wdFieldFormTextInput)
        doc.FormFields.Add Range:=doc.Range(doc.Range.End - 1, doc.Range.End - 1), Type:=70
        doc.FormFields(doc.FormFields.Count).Name = "UserName"
        doc.FormFields(doc.FormFields.Count).Result = "用户_" & i
        
        doc.Range.InsertAfter vbCrLf & "年龄: "
        doc.FormFields.Add Range:=doc.Range(doc.Range.End - 1, doc.Range.End - 1), Type:=70
        doc.FormFields(doc.FormFields.Count).Name = "UserAge"
        doc.FormFields(doc.FormFields.Count).Result = 20 + i
        
        doc.Range.InsertAfter vbCrLf & "反馈意见: "
        doc.FormFields.Add Range:=doc.Range(doc.Range.End - 1, doc.Range.End - 1), Type:=70
        doc.FormFields(doc.FormFields.Count).Name = "Comments"
        doc.FormFields(doc.FormFields.Count).Result = "来自用户 " & i & " 的测试反馈内容。"
        
        ' 保护文档 (仅限填写窗体) - 3 = wdAllowOnlyFormFields
        doc.Protect Type:=3, NoReset:=True, Password:=""
        
        ' 保存文件
        On Error Resume Next
        doc.SaveAs2 Filename:=folderPath & "Sample_Data_" & i & ".docx"
        If Err.Number <> 0 Then
            MsgBox "保存失败，请检查路径权限：" & vbCrLf & folderPath, vbExclamation
            doc.Close False
            Exit Sub
        End If
        On Error GoTo 0
        
        doc.Close
    Next i
    
    MsgBox "已在以下路径生成 3 份测试文件：" & vbCrLf & folderPath, vbInformation
End Sub
