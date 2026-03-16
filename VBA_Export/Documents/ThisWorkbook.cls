Option Explicit

Private Sub Workbook_Deactivate()
    Application.CommandBars("Worksheet menu bar").Reset
End Sub

Private Sub Workbook_Open()
    Call 报表工具
End Sub

Private Sub 报表工具()
    Dim ToolMenu As Object
    Dim NewMenu As Object
    With Application.CommandBars("Worksheet menu bar")
        .Reset
        '利用ID属性查找帮助菜单
        Set ToolMenu = .FindControl(ID:=.Controls("工具(&T)").ID)
        If ToolMenu Is Nothing Then
            '如果该菜单不存在,则将新菜单添加到末尾
            Set NewMenu = .Controls.Add(Type:=msoControlPopup)
        Else
            '将新菜单添加到帮助菜单之前
            Set NewMenu = .Controls.Add(Type:=msoControlPopup, _
                before:=ToolMenu.Index)
        End If
        With NewMenu
            .Caption = "报表工具(&S)"
            ' 1. 文件选择（执行面板用）
            With .Controls.Add(Type:=msoControlPopup)
                .Caption = "1. 文件选择"
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "1.1 选择模板文件"
                    .OnAction = "选择模板文件"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "1.2 选择外部文件"
                    .OnAction = "选择外部文件"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "1.3 选择源文件"
                    .OnAction = "选择源文件"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "1.4 批量选择源文件"
                    .OnAction = "批量选择源文件"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "1.5 批量选择源文件(含子文件夹)"
                    .OnAction = "批量选择源文件含子文件夹"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "1.6 模板与源数据表格校验"
                    .OnAction = "模板与源数据表格校验"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "1.7 注入校验区域"
                    .OnAction = "注入校验区域"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "1.8 汇总校验结果"
                    .OnAction = "汇总校验结果"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "1.9 清除校验区域"
                    .OnAction = "清除校验区域"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "1.10 输出校验结果（清除→注入→汇总）"
                    .OnAction = "输出校验结果"
                End With
            End With
            ' 2. 一二批功能
            With .Controls.Add(Type:=msoControlPopup)
                .Caption = "2. 一二批功能"
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "2.1 拆分村镇银行（三币种、外发用）"
                    .OnAction = "批量填入村镇银行计算结果"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "2.2 将机构名称标准化（三币种、外发用）"
                    .OnAction = "批量修改分机构文件"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "2.3 删除分机构表的外资行（外发用）"
                    .OnAction = "删除外资行_批量"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "2.4 修改外汇表的页眉（外发用）"
                    .OnAction = "批量调整外汇Sheet页眉_优化版"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "2.5 核对地区总分校验（商业银行共享一批）"
                    .OnAction = "核对地区汇总数据_完整版"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "2.6 提取工作表（本外币、专项二批次）"
                    .OnAction = "批量提取工作表数据"
                End With

            End With
            ' 3. 工作簿操作
            With .Controls.Add(Type:=msoControlPopup)
                .Caption = "3. 工作簿操作"
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "3.1 工作表比较"
                    .OnAction = "执行表格比对"
                End With
                With .Controls.Add(Type:=msoControlPopup)
                    .Caption = "3.2 汇总"
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "3.2.1 按使用区域汇总"
                        .OnAction = "SummarizeByUsedRange"
                    End With
                    With .Controls.Add(Type:=msoControlButton)
                        .Caption = "3.2.2 按批注汇总"
                        .OnAction = "SummarizeBySheetThenSplit"
                    End With
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "3.3 批量重命名文件"
                    .OnAction = "批量重命名文件"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "3.4 批量Excel格式转换"
                    .OnAction = "批量Excel格式转换"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "3.5 追加列（按批注）"
                    .OnAction = "追加列按批注"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "3.6 批量Word格式转换"
                    .OnAction = "批量Word格式转换"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "3.7 批量修改Sheet名（按config_rename J→K）"
                    .OnAction = "批量修改Sheet名"
                End With
            End With
            ' 4. 工作表操作
            With .Controls.Add(Type:=msoControlPopup)
                .Caption = "4. 工作表操作"
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "4.1 取消合并并填充左上角值"
                    .OnAction = "取消合并并填充左上角值"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "4.2 合并相邻同值单元格"
                    .OnAction = "合并相邻同值单元格"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "4.3 分列（数字/符号/中文）"
                    .OnAction = "功能3dot10_分列数字符号中文"
                End With
            End With
            ' 5. VBA
            With .Controls.Add(Type:=msoControlPopup)
                .Caption = "5. VBA"
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "5.1 导入 VBA 模块"
                    .OnAction = "Sync_ImportAll"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "5.2 初始化配置"
                    .OnAction = "初始化config"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "5.3 初始化执行面板"
                    .OnAction = "初始化执行面板"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "5.4 注入VBA到所选源文件"
                    .OnAction = "注入VBA到所选源文件"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "5.5 清除目标工作簿VBA"
                    .OnAction = "清除目标工作簿VBA"
                End With
            End With
            ' 6. 初始化（机构映射表、表格比对、config_rename 等整表初始化与导出）
            With .Controls.Add(Type:=msoControlPopup)
                .Caption = "6. 初始化"
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "6.1 初始化config"
                    .OnAction = "初始化config"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "6.2 初始化执行面板"
                    .OnAction = "初始化执行面板"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "6.3 初始化运行日志"
                    .OnAction = "初始化运行日志"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "6.4 初始化机构映射表"
                    .OnAction = "初始化机构映射表"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "6.5 初始化表格比对"
                    .OnAction = "初始化表格比对"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "6.6 初始化工作表提取"
                    .OnAction = "初始化工作表提取"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "6.7 初始化config_rename"
                    .OnAction = "初始化config_rename"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "6.8 初始化统计工具"
                    .OnAction = "初始化统计工具"
                End With
                With .Controls.Add(Type:=msoControlButton)
                    .Caption = "6.9 导出初始化数据"
                    .OnAction = "导出初始化数据"
                End With
            End With

        End With
    End With
    Set ToolMenu = Nothing
    Set NewMenu = Nothing
End Sub

Private Sub Workbook_Activate()
    Call 报表工具
End Sub








