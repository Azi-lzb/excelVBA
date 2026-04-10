Option Explicit

Private Sub Workbook_Deactivate()
    Application.CommandBars("Worksheet menu bar").Reset
End Sub

Private Sub Workbook_Open()
    Call BuildReportMenu
End Sub

Private Sub Workbook_Activate()
    Call BuildReportMenu
End Sub

Private Sub BuildReportMenu()
    Dim toolMenu As Object
    Dim rootMenu As Object
    Dim popupMenu As Object

    With Application.CommandBars("Worksheet menu bar")
        .Reset

        On Error Resume Next
        Set toolMenu = .FindControl(ID:=.Controls("工具(&T)").ID)
        On Error GoTo 0

        If toolMenu Is Nothing Then
            Set rootMenu = .Controls.Add(Type:=msoControlPopup)
        Else
            Set rootMenu = .Controls.Add(Type:=msoControlPopup, before:=toolMenu.Index)
        End If
    End With

    rootMenu.Caption = "报表工具(&S)"

    Set popupMenu = AddPopup(rootMenu, "1. 文件选择")
    AddButton popupMenu, "1.1 选择模板文件", "选择模板文件"
    AddButton popupMenu, "1.2 选择外部文件", "选择外部文件"
    AddButton popupMenu, "1.3 选择源文件", "选择源文件"
    AddButton popupMenu, "1.4 批量选择源文件", "批量选择源文件"
    AddButton popupMenu, "1.5 批量选择源文件(含子文件夹)", "批量选择源文件含子文件夹"
    AddButton popupMenu, "1.6 模板与源数据表格校验", "模板与源数据表格校验"
    AddButton popupMenu, "1.7 注入校验区域", "注入校验区域"
    AddButton popupMenu, "1.8 汇总校验结果", "汇总校验结果"
    AddButton popupMenu, "1.9 清空校验区域", "清空校验区域"
    AddButton popupMenu, "1.10 一键校验流程（注入→汇总→清空）", "一键校验流程"

    Set popupMenu = AddPopup(rootMenu, "2. 一二批功能")
    AddButton popupMenu, "2.1 拆分村镇银行（三币种、批发用）", "批量填入村镇银行计算结果"
    AddButton popupMenu, "2.2 将机构名称标准化（三币种、批发用）", "批量修改分机构文件"
    AddButton popupMenu, "2.3 删除分机构表的外资行（批发用）", "删除外资行_批量"
    AddButton popupMenu, "2.4 调整外汇表页眉（批发用）", "批量调整外汇Sheet页眉_优化版"
    AddButton popupMenu, "2.5 核对地区总分校验（商业银行共用一批）", "核对地区汇总数据_完整版"
    AddButton popupMenu, "2.6 提取工作表数据（专项二批次）", "批量提取工作表数据"
    AddButton popupMenu, "2.7 修正涉农贷款比上月（专项二批次）", "修正涉农贷款比上月"

    Set popupMenu = AddPopup(rootMenu, "3. 工作簿操作")
    AddButton popupMenu, "3.1 表格比对", "执行表格比对"
    AddSummaryPopup popupMenu
    AddButton popupMenu, "3.3 批量重命名文件", "批量重命名文件"
    AddButton popupMenu, "3.4 批量Excel格式转换", "批量Excel格式转换"
    AddButton popupMenu, "3.5 追加列（按批注）", "追加列按批注"
    AddButton popupMenu, "3.6 批量Word格式转换", "批量Word格式转换"
    AddButton popupMenu, "3.7 批量修改Sheet名（按config_rename J→K）", "批量修改Sheet名"
    AddButton popupMenu, "3.8 批量转时序数据", "批量转时序数据"
    AddButton popupMenu, "3.9 指标关系校验", "执行指标关系校验"

    Set popupMenu = AddPopup(rootMenu, "4. 工作表操作")
    AddButton popupMenu, "4.1 取消合并并填充左上角值", "取消合并并填充左上角值"
    AddButton popupMenu, "4.2 合并相邻同值单元格", "合并相邻同值单元格"
    AddButton popupMenu, "4.3 分列（数字/符号/中文）", "功能3dot10_分列数字符号中文"

    Set popupMenu = AddPopup(rootMenu, "5. VBA")
    AddButton popupMenu, "5.1 导入VBA模块", "Sync_ImportAll"
    AddButton popupMenu, "5.2 初始化config", "初始化config"
    AddButton popupMenu, "5.3 初始化执行面板", "初始化执行面板"
    AddButton popupMenu, "5.4 注入VBA到所选源文件", "注入VBA到所选源文件"
    AddButton popupMenu, "5.5 清除目标工作簿VBA", "清除目标工作簿VBA"

    Set popupMenu = AddPopup(rootMenu, "6. 初始化")
    AddButton popupMenu, "6.1 初始化config", "初始化config"
    AddButton popupMenu, "6.2 初始化执行面板", "初始化执行面板"
    AddButton popupMenu, "6.3 初始化运行日志", "初始化运行日志"
    AddButton popupMenu, "6.4 初始化机构映射表", "初始化机构映射表"
    AddButton popupMenu, "6.5 初始化表格比对", "初始化表格比对"
    AddButton popupMenu, "6.6 初始化工作表提取", "初始化工作表提取"
    AddButton popupMenu, "6.7 初始化config_rename", "初始化config_rename"
    AddButton popupMenu, "6.8 初始化统计工具", "初始化统计工具"
    AddButton popupMenu, "6.9 导出初始化数据", "导出初始化数据"

    Set popupMenu = Nothing
    Set toolMenu = Nothing
    Set rootMenu = Nothing
End Sub

Private Function AddPopup(ByVal parentMenu As Object, ByVal captionText As String) As Object
    Set AddPopup = parentMenu.Controls.Add(Type:=msoControlPopup)
    AddPopup.Caption = captionText
End Function

Private Sub AddButton(ByVal parentMenu As Object, ByVal captionText As String, ByVal actionName As String)
    With parentMenu.Controls.Add(Type:=msoControlButton)
        .Caption = captionText
        .OnAction = actionName
    End With
End Sub

Private Sub AddSummaryPopup(ByVal parentMenu As Object)
    Dim summaryMenu As Object

    Set summaryMenu = AddPopup(parentMenu, "3.2 汇总")
    AddButton summaryMenu, "3.2.1 按使用区域汇总", "SummarizeByUsedRange"
    AddButton summaryMenu, "3.2.2 按批注汇总", "SummarizeBySheetThenSplit"
End Sub
