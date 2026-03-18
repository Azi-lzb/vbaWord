Attribute VB_Name = "modConfig"
' ==========================================================
' vbaWord: Configuration Management Module
' Description: Manages the 'config' sheet and provides settings to other modules.
' ==========================================================

' --- 1. 初始化配置表 ---
Sub InitConfigSheet()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("config")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = "config"
    End If
    
    ws.Cells.Clear
    
    ' 设置表头
    ws.Range("A1:C1").Value = Array("配置项说明", "配置值", "填写指导")
    ws.Range("A1:C1").Font.Bold = True
    ws.Range("A1:C1").Interior.Color = RGB(220, 230, 241)
    
    ' 填充配置行
    ' 行 2: 密码
    ws.Cells(2, 1).Value = "文档保护密码"
    ws.Cells(2, 2).Value = "" ' 默认空
    ws.Cells(2, 3).Value = "若 Word 文档设置了限制编辑，请在此填写密码；若无密码请留空。"
    
    ' 行 3: 优先级
    ws.Cells(3, 1).Value = "题目识别优先级"
    ws.Cells(3, 2).Value = "TAG"
    ws.Cells(3, 3).Value = "可选：TAG (标记优先) 或 TITLE (标题优先)。仅适用于新式窗体。"
    
    ' 美化
    ws.Columns("A:C").AutoFit
    ws.Columns("B").ColumnWidth = 20
    ws.Range("B2:B3").Interior.Color = RGB(255, 255, 204) ' 黄色背景提示填写
    
    MsgBox "config 配置表初始化完成！", vbInformation
End Sub

' --- 2. 读取密码 ---
Public Function GetDocPassword() As String
    GetDocPassword = GetConfigValue("文档保护密码")
End Function

' --- 3. 读取优先级模式 ---
Public Function GetPriorityMode() As String
    Dim val As String
    val = UCase(GetConfigValue("题目识别优先级"))
    If val = "" Then val = "TAG"
    GetPriorityMode = val
End Function

' --- 4. 通用读取私有函数 ---
Private Function GetConfigValue(itemName As String) As String
    Dim ws As Worksheet
    Dim cell As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("config")
    On Error GoTo 0
    
    If ws Is Nothing Then
        GetConfigValue = ""
        Exit Function
    End If
    
    ' 在 A 列搜索名称
    Set cell = ws.Columns("A").Find(What:=itemName, LookIn:=xlValues, LookAt:=xlWhole)
    If Not cell Is Nothing Then
        GetConfigValue = Trim(CStr(cell.Offset(0, 1).Value))
    Else
        GetConfigValue = ""
    End If
End Function
