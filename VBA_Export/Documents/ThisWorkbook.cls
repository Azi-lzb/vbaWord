VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  ' True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' ==========================================================
' ThisWorkbook: vbaWord Main Dispatcher
' Updated: Added Config Init Menu
' ==========================================================

Private Sub Workbook_Open()
    Call CreateCustomMenu
End Sub

Private Sub Workbook_Activate()
    Call CreateCustomMenu
End Sub

Private Sub Workbook_Deactivate()
    On Error Resume Next
    Application.CommandBars("Worksheet menu bar").Controls("vbaWord工具(&S)").Delete
    On Error GoTo 0
End Sub

Public Sub CreateCustomMenu()
    Dim ToolMenu As Object, NewMenu As Object
    
    With Application.CommandBars("Worksheet menu bar")
        On Error Resume Next
        .Controls("vbaWord工具(&S)").Delete
        On Error GoTo 0
        
        Set ToolMenu = .FindControl(ID:=30007)
        If ToolMenu Is Nothing Then
            Set NewMenu = .Controls.Add(Type:=10)
        Else
            Set NewMenu = .Controls.Add(Type:=10, before:=ToolMenu.Index)
        End If
        
        With NewMenu
            .Caption = "vbaWord工具(&S)"
            
            ' --- 1. 旧式窗体汇总 ---
            With .Controls.Add(Type:=10)
                .Caption = "1. 旧式窗体汇总"
                With .Controls.Add(Type:=1)
                    .Caption = "1.1 汇总至 Excel (output表)"
                    .OnAction = "extractWord.BatchSummarizeWordForms"
                End With
                With .Controls.Add(Type:=1)
                    .Caption = "1.2 汇总至 Word (新文档)"
                    .OnAction = "extractToWord.SummarizeToNewWordDoc"
                End With
            End With
            
            ' --- 2. 内容控件汇总 ---
            With .Controls.Add(Type:=10)
                .Caption = "2. 内容控件汇总"
                With .Controls.Add(Type:=1)
                    .Caption = "2.1 汇总至 Excel (output_cc表)"
                    .OnAction = "extractCCtoExcel.SummarizeContentControlsToExcel"
                End With
                With .Controls.Add(Type:=1)
                    .Caption = "2.2 汇总至 Word (文本模式)"
                    .OnAction = "extractCCtoWord.SummarizeCCToWord"
                End With
                With .Controls.Add(Type:=1)
                    .Caption = "2.3 汇总至 Word (含表格模式)"
                    .OnAction = "extractCCtoWordWithTables.SummarizeCCWithTablesToWord"
                End With
            End With
            
            ' --- 3. 配置与工具 ---
            With .Controls.Add(Type:=10)
                .Caption = "3. 配置与工具"
                With .Controls.Add(Type:=1)
                    .Caption = "3.1 初始化 config 配置表"
                    .OnAction = "modConfig.InitConfigSheet"
                End With
                With .Controls.Add(Type:=1)
                    .Caption = "3.2 生成旧式窗体示例"
                    .OnAction = "GenerateTestData.CreateSampleDocs"
                End With
                With .Controls.Add(Type:=1)
                    .Caption = "3.3 生成内容控件示例"
                    .OnAction = "GenerateCCTestData.CreateCCSampleDocs"
                End With
            End With
        End With
    End With
End Sub
