Attribute VB_Name = "功能1dot5地区数值核对"
Sub 核对地区汇总数据_完整版()
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim checkedFiles As Integer
    Dim checkedSheets As Integer
    Dim totalErrors As Long
    Dim t0 As Double

    t0 = Timer
    RunLog_WriteRow "1.5 核对地区总分校验", "开始", "", "", "", "", "开始", ""
    
    ' 初始化
    checkedFiles = 0
    checkedSheets = 0
    totalErrors = 0
    
    ' 创建文件选择对话框
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "请选择要核对的Excel文件（可多选）"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xls; *.xlsx; *.xlsm"
        
        If .Show = -1 Then
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
            
            ' 创建结果工作簿
            Dim resultWb As Workbook
            Dim resultWs As Worksheet
            Set resultWb = Workbooks.Add
            Set resultWs = resultWb.Worksheets(1)
            resultWs.Name = "核对结果"
            
            ' 设置结果表头
            设置结果表头 resultWs
            
            Dim resultRow As Long
            resultRow = 2
            
            ' 遍历所有选择的文件
            For Each fileItem In .SelectedItems
                On Error Resume Next ' 添加错误处理
                Set wb = Workbooks.Open(CStr(fileItem), ReadOnly:=True)
                
                If Err.Number = 0 Then
                    checkedFiles = checkedFiles + 1
                    
                    ' 遍历工作簿中的所有工作表
                    For Each ws In wb.Worksheets
                        ' 检查工作表名称是否包含"分地区"
                        If InStr(1, ws.Name, "分地区", vbTextCompare) > 0 Then
                            ' 执行核对并获取错误信息
                            Dim errorCount As Long
                            Dim errorMsgs As String
                            
                            errorCount = 核对地区数据(ws, errorMsgs)
                            
                            ' 记录结果
                            记录结果 resultWs, resultRow, _
                                       wb.Name, ws.Name, _
                                       errorCount, errorMsgs
                            
                            totalErrors = totalErrors + errorCount
                            resultRow = resultRow + 1
                            checkedSheets = checkedSheets + 1
                            
                            RunLog_WriteRow "1.5 核对地区总分校验", "核对表", wb.Name & "|" & ws.Name, "", "", IIf(errorCount > 0, "有误", "通过"), IIf(errorCount > 0, errorCount & " 个错误", "OK"), ""
                            Debug.Print "已核对: " & wb.Name & " - " & ws.Name & _
                                        " (" & errorCount & " 个错误)"
                        End If
                    Next ws
                    
                    wb.Close SaveChanges:=False
                Else
                    RunLog_WriteRow "1.5 核对地区总分校验", "打开文件", CStr(fileItem), "", "", "失败", Err.Number & " " & Err.Description, ""
                    Debug.Print "无法打开文件: " & fileItem
                End If
                On Error GoTo 0
            Next fileItem
            
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
            
            If resultRow >= 2 Then
                On Error Resume Next
                resultWs.Cells(resultRow, 1) = "=== 统计汇总 ==="
                resultWs.Cells(resultRow, 2) = "已检查文件数: " & checkedFiles
                resultWs.Cells(resultRow, 3) = "已检查工作表数: " & checkedSheets
                resultWs.Cells(resultRow, 4) = "发现错误总数: " & totalErrors
                On Error GoTo 0
            End If
            
            ' 自动调整列宽
            resultWs.Columns("A:F").AutoFit
            
            ' 高亮显示错误行
            格式化结果表 resultWs
            
            RunLog_WriteRow "1.5 核对地区总分校验", "完成", "", "", "", "", "Done", CStr(Round(Timer - t0, 2))
            ' 显示结果
            Dim msg As String
            msg = "核对完成！" & vbCrLf & vbCrLf
            If checkedFiles > 0 Then
                msg = msg & "已检查文件: " & checkedFiles & " 个" & vbCrLf
                msg = msg & "已检查工作表: " & checkedSheets & " 个" & vbCrLf
                msg = msg & "发现错误: " & totalErrors & " 个"
            Else
                msg = msg & "没有找到可处理的工作表。"
            End If
            
            MsgBox msg, vbInformation, "核对完成"
        Else
            RunLog_WriteRow "1.5 核对地区总分校验", "完成", "", "", "", "", "已取消", CStr(Round(Timer - t0, 2))
            MsgBox "操作已取消", vbInformation
        End If
    End With
    
    Set fd = Nothing
End Sub

' 核心函数：核对地区数据
Function 核对地区数据(ws As Worksheet, ByRef 错误信息 As String) As Long
    Dim 最后行 As Long, 最后列 As Long
    Dim 起始行 As Long, 起始列 As Long
    Dim i As Long, j As Long, 列 As Long
    Dim 惠州行 As Long
    Dim 地区行号(0 To 6) As Long
    Dim 地区名称(6) As String
    Dim 错误数 As Long
    Dim 错误详情 As String
    
    On Error GoTo ErrorHandler ' 添加错误处理
    
    ' 定义需要核对的县区
    地区名称(0) = "惠州市"
    地区名称(1) = "惠阳区"
    地区名称(2) = "博罗县"
    地区名称(3) = "惠东县"
    地区名称(4) = "龙门县"
    地区名称(5) = "惠州市本部"
    地区名称(6) = "大亚湾"
    
    ' 初始化数组存储行号
    For i = 0 To 6
        地区行号(i) = 0
    Next i
    
    错误数 = 0
    错误详情 = ""
    
    ' 获取实际数据范围
    最后行 = 获取最后有数据的行(ws)
    最后列 = 获取最后有数据的列(ws)
    
    ' 检查是否有数据
    If 最后行 = 0 Or 最后列 = 0 Then
        错误信息 = "工作表为空或没有数据"
        核对地区数据 = 1
        Exit Function
    End If
    
    ' 确定起始位置
    If ws.Cells(2, 2).value <> "" Then
        起始行 = 3  ' 数据从第3行开始
        起始列 = 3  ' 数据从C列开始
    Else
        ' 寻找数据起始行
        起始行 = 查找数据起始行(ws)
        If 起始行 = 0 Then
            起始行 = 3
        End If
        起始列 = 3
    End If
    
    ' 如果没有找到有效数据，退出
    If 起始行 = 0 Or 最后列 < 起始列 Or 最后行 < 起始行 Then
        错误信息 = "未找到有效数据区域"
        核对地区数据 = 1
        Exit Function
    End If
    
    ' 查找各县区和惠州市的行号
    For i = 起始行 To 最后行
        Dim 单元格内容 As String
        单元格内容 = Trim(CStr(ws.Cells(i, 2).value))  ' B列是地区名称
        
        If 单元格内容 <> "" Then
            ' 检查是否是我们关心的地区
            For j = 0 To 6
                If 单元格内容 = 地区名称(j) Then
                    地区行号(j) = i
                    Exit For
                End If
            Next j
        End If
    Next i
    
    ' 检查是否找到惠州市
    If 地区行号(0) = 0 Then
        错误信息 = "未找到'惠州市'行"
        核对地区数据 = 1
        Exit Function
    End If
    
    ' 从C列开始遍历每一列进行核对
    For 列 = 起始列 To 最后列
        Dim 惠州值 As Double
        Dim 县区合计 As Double
        Dim 有有效数据 As Boolean
        
        ' 获取惠州市的值（索引0是惠州市）
        惠州行 = 地区行号(0)
        If IsNumeric(ws.Cells(惠州行, 列).value) Then
            惠州值 = CDbl(ws.Cells(惠州行, 列).value)
            有有效数据 = True
        ElseIf ws.Cells(惠州行, 列).value = "" Or ws.Cells(惠州行, 列).value = "-" Then
            惠州值 = 0
            有有效数据 = True
        Else
            ' 非数字值，跳过本列
            有有效数据 = False
        End If
        
        ' 计算县区合计（索引1-6是县区）
        县区合计 = 0
        For j = 1 To 6
            If 地区行号(j) > 0 Then  ' 如果找到了该县区
                Dim 县区行 As Long
                县区行 = 地区行号(j)
                
                If IsNumeric(ws.Cells(县区行, 列).value) Then
                    县区合计 = 县区合计 + CDbl(ws.Cells(县区行, 列).value)
                    有有效数据 = True
                ElseIf ws.Cells(县区行, 列).value = "" Or ws.Cells(县区行, 列).value = "-" Then
                    ' 空值或横杠视为0
                    ' 不改变有有效数据状态
                End If
            End If
        Next j
        
        ' 如果本列有有效数据，进行核对
        If 有有效数据 Then
            Dim 差异 As Double
            Dim 差异率 As Double
            
            差异 = 惠州值 - 县区合计
            
            ' 检查差异是否显著
            Dim 是否显著差异 As Boolean
            是否显著差异 = False
            
            If Abs(差异) > 0.01 Then  ' 绝对差异大于0.01
                是否显著差异 = True
            ElseIf 惠州值 <> 0 Then
                差异率 = Abs(差异 / 惠州值)
                If 差异率 > 0.001 Then  ' 相对差异大于0.1%
                    是否显著差异 = True
                End If
            ElseIf 县区合计 <> 0 Then  ' 惠州市为0，县区合计不为0
                是否显著差异 = True
            End If
            
            If 是否显著差异 Then
                错误数 = 错误数 + 1
                
                Dim 错误消息 As String
                ' 获取表头名称（第2行）
                Dim 表头名称 As String
                On Error Resume Next
                表头名称 = Trim(CStr(ws.Cells(2, 列).value))
                On Error GoTo 0
                
                If 表头名称 = "" Then
                    表头名称 = "列" & 列号转字母(列)
                End If
                
                错误消息 = 表头名称 & ": " & _
                           Format(惠州值, "0.00") & " ≠ " & _
                           Format(县区合计, "0.00") & " (差" & _
                           Format(差异, "0.00") & ")"
                
                ' 累积错误信息（最多显示5条）
                If 错误数 > 0 Then
                    If 错误详情 <> "" Then 错误详情 = 错误详情 & "; "
                    错误详情 = 错误详情 & 错误消息
                End If
                
                ' 调试输出
                Debug.Print "  错误: " & 错误消息
            End If
        End If
    Next 列
    
    ' 检查是否有未找到的县区
    Dim 未找到的地区 As String
    未找到的地区 = ""
    For j = 1 To 6
        If 地区行号(j) = 0 Then
            If 未找到的地区 <> "" Then 未找到的地区 = 未找到的地区 & ", "
            未找到的地区 = 未找到的地区 & 地区名称(j)
        End If
    Next j
    
    If 未找到的地区 <> "" Then
        错误详情 = 错误详情 & IIf(错误详情 <> "", "; ", "") & _
                  "未找到: " & 未找到的地区
    End If
    
    ' 返回结果
    错误信息 = 错误详情
    核对地区数据 = 错误数
    
    Exit Function
    
ErrorHandler:
    错误信息 = "处理过程中发生错误: " & Err.Description
    核对地区数据 = 1
End Function

' 获取最后有数据的行 - 修复版
Function 获取最后有数据的行(ws As Worksheet) As Long
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = ws.Cells.Find("*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If Not rng Is Nothing Then
        获取最后有数据的行 = rng.row
    Else
        获取最后有数据的行 = 0
    End If
    
    Exit Function
    
ErrorHandler:
    获取最后有数据的行 = 0
End Function

' 获取最后有数据的列 - 修复版
Function 获取最后有数据的列(ws As Worksheet) As Long
    On Error GoTo ErrorHandler
    
    Dim rng As Range
    Set rng = ws.Cells.Find("*", LookIn:=xlValues, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    If Not rng Is Nothing Then
        获取最后有数据的列 = rng.Column
    Else
        获取最后有数据的列 = 0
    End If
    
    Exit Function
    
ErrorHandler:
    获取最后有数据的列 = 0
End Function

' 列号转字母 - 完全重写，避免使用Cells函数
Function 列号转字母(列号 As Long) As String
    If 列号 <= 0 Then
        列号转字母 = "A"
        Exit Function
    End If
    
    Dim 字母 As String
    Dim 余数 As Long
    
    字母 = ""
    
    While 列号 > 0
        余数 = (列号 - 1) Mod 26
        字母 = Chr(65 + 余数) & 字母
        列号 = (列号 - 1) \ 26
    Wend
    
    列号转字母 = 字母
End Function

' 其他辅助函数保持不变，但需要修复一些细节

' 设置结果表头（重命名函数，避免英文名）
Sub 设置结果表头(ws As Worksheet)
    With ws
        .Cells(1, 1) = "序号"
        .Cells(1, 2) = "文件名"
        .Cells(1, 3) = "工作表名"
        .Cells(1, 4) = "核对状态"
        .Cells(1, 5) = "错误数量"
        .Cells(1, 6) = "错误信息"
        
        ' 设置表头格式
        With .Rows(1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .Interior.Color = RGB(200, 220, 240)
        End With
    End With
End Sub

' 查找数据起始行
Function 查找数据起始行(ws As Worksheet) As Long
    Dim 行 As Long
    Dim 最大检查行数 As Long
    
    最大检查行数 = 100  ' 最多检查100行
    
    For 行 = 1 To 最大检查行数
        ' 检查B列是否有我们关心的地区名称
        Dim 单元格内容 As String
        单元格内容 = Trim(CStr(ws.Cells(行, 2).value))
        
        ' 检查是否包含关键词
        If 单元格内容 = "惠州市" Or _
           单元格内容 = "惠阳区" Or _
           单元格内容 = "博罗县" Or _
           单元格内容 = "惠东县" Or _
           单元格内容 = "龙门县" Or _
           单元格内容 = "惠州市本部" Or _
           单元格内容 = "大亚湾" Then
            查找数据起始行 = 行
            Exit Function
        End If
    Next 行
    
    查找数据起始行 = 0  ' 未找到
End Function

' 记录结果
Sub 记录结果(结果表 As Worksheet, 行号 As Long, _
           文件名 As String, 工作表名 As String, _
           错误数量 As Long, 错误信息 As String)
    With 结果表
        .Cells(行号, 1) = 行号 - 1                    ' 序号
        .Cells(行号, 2) = 文件名                      ' 文件名
        .Cells(行号, 3) = 工作表名                    ' 工作表名
        
        If 错误数量 > 0 Then
            .Cells(行号, 4) = "错误"                  ' 状态
            .Cells(行号, 5) = 错误数量                ' 错误数量
            .Cells(行号, 6) = 错误信息       ' 错误信息（截断）
        Else
            .Cells(行号, 4) = "正确"                  ' 状态
            .Cells(行号, 5) = 0                       ' 错误数量
            .Cells(行号, 6) = "核对通过"              ' 备注
        End If
    End With
End Sub

' 格式化结果表
Sub 格式化结果表(ws As Worksheet)
    Dim 最后行 As Long
    最后行 = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    ' 确保至少有一行数据
    If 最后行 < 2 Then Exit Sub
    
    Dim 行 As Long
    For 行 = 2 To 最后行
        Dim 状态 As String
        状态 = CStr(ws.Cells(行, 4).value)
        
        If 状态 = "错误" Then
            ' 错误行 - 红色背景
            With ws.Rows(行)
                .Interior.Color = RGB(255, 200, 200)
                .Font.Color = RGB(192, 0, 0)
            End With
        ElseIf 状态 = "正确" Then
            ' 正确行 - 绿色背景
            With ws.Rows(行)
                .Interior.Color = RGB(200, 255, 200)
                .Font.Color = RGB(0, 100, 0)
            End With
        End If
    Next 行
End Sub

' ==================== 以下是使用示例和测试函数 ====================

' 测试单个文件 - 修复版
Sub 测试核对功能()
    Dim fd As FileDialog
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "请选择要测试的Excel文件"
    fd.Filters.Add "Excel文件", "*.xls; *.xlsx; *.xlsm"
    
    If fd.Show = -1 Then
        On Error Resume Next
        Set wb = Workbooks.Open(fd.SelectedItems(1), ReadOnly:=True)
        On Error GoTo 0
        
        If wb Is Nothing Then
            MsgBox "无法打开文件！", vbExclamation
            Exit Sub
        End If
        
        Dim 找到工作表 As Boolean
        找到工作表 = False
        
        ' 查找并核对每个包含"分地区"的工作表
        For Each ws In wb.Worksheets
            If InStr(1, ws.Name, "分地区", vbTextCompare) > 0 Then
                找到工作表 = True
                
                Dim 错误信息 As String
                Dim 错误数量 As Long
                
                错误数量 = 核对地区数据(ws, 错误信息)
                
                If 错误数量 > 0 Then
                    MsgBox "工作表 '" & ws.Name & "' 核对结果:" & vbCrLf & _
                           "发现 " & 错误数量 & " 个问题" & vbCrLf & _
                           "具体问题: " & 错误信息, vbExclamation, "核对结果"
                Else
                    MsgBox "工作表 '" & ws.Name & "' 核对结果:" & vbCrLf & _
                           "? 所有数据核对正确！", vbInformation, "核对结果"
                End If
            End If
        Next ws
        
        If Not 找到工作表 Then
            MsgBox "未找到包含'分地区'的工作表！", vbExclamation
        End If
        
        wb.Close SaveChanges:=False
    End If
    
    Set fd = Nothing
End Sub

' 创建一个示例数据工作表（用于测试）- 修复版
Sub 创建示例数据()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "分地区数据示例"
    
    ' 创建表头
    ws.Cells(2, 2) = "地区"  ' B2是地区表头
    ws.Cells(2, 3) = "指标1"
    ws.Cells(2, 4) = "指标2"
    ws.Cells(2, 5) = "指标3"
    ws.Cells(2, 6) = "指标4"
    
    ' 创建数据行（正确的数据）
    Dim 数据(1 To 7, 1 To 4) As Double
    数据(1, 1) = 1000  ' 惠州市 - 指标1
    数据(1, 2) = 2000  ' 惠州市 - 指标2
    数据(1, 3) = 3000  ' 惠州市 - 指标3
    数据(1, 4) = 4000  ' 惠州市 - 指标4
    
    数据(2, 1) = 200   ' 惠阳区 - 指标1
    数据(2, 2) = 400   ' 惠阳区 - 指标2
    数据(2, 3) = 600   ' 惠阳区 - 指标3
    数据(2, 4) = 800   ' 惠阳区 - 指标4
    
    数据(3, 1) = 300   ' 博罗县 - 指标1
    数据(3, 2) = 500   ' 博罗县 - 指标2
    数据(3, 3) = 700   ' 博罗县 - 指标3
    数据(3, 4) = 900   ' 博罗县 - 指标4
    
    数据(4, 1) = 200   ' 惠东县 - 指标1
    数据(4, 2) = 400   ' 惠东县 - 指标2
    数据(4, 3) = 600   ' 惠东县 - 指标3
    数据(4, 4) = 800   ' 惠东县 - 指标4
    
    数据(5, 1) = 100   ' 龙门县 - 指标1
    数据(5, 2) = 200   ' 龙门县 - 指标2
    数据(5, 3) = 300   ' 龙门县 - 指标3
    数据(5, 4) = 400   ' 龙门县 - 指标4
    
    数据(6, 1) = 150   ' 惠州市本部 - 指标1
    数据(6, 2) = 250   ' 惠州市本部 - 指标2
    数据(6, 3) = 350   ' 惠州市本部 - 指标3
    数据(6, 4) = 450   ' 惠州市本部 - 指标4
    
    数据(7, 1) = 50    ' 大亚湾 - 指标1
    数据(7, 2) = 50    ' 大亚湾 - 指标2
    数据(7, 3) = 50    ' 大亚湾 - 指标3
    数据(7, 4) = 50    ' 大亚湾 - 指标4
    
    ' 填写数据
    Dim 地区名称(1 To 7) As String
    地区名称(1) = "惠州市"
    地区名称(2) = "惠阳区"
    地区名称(3) = "博罗县"
    地区名称(4) = "惠东县"
    地区名称(5) = "龙门县"
    地区名称(6) = "惠州市本部"
    地区名称(7) = "大亚湾"
    
    Dim i As Long, j As Long
    For i = 1 To 7
        ws.Cells(i + 2, 2) = 地区名称(i)  ' 从B3开始
        For j = 1 To 4
            ws.Cells(i + 2, j + 2) = 数据(i, j)  ' 从C3开始
        Next j
    Next i
    
    ' 格式设置
    ws.Columns("B:F").AutoFit
    With ws.Rows(2)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(220, 230, 241)
    End With
    
    MsgBox "示例数据已创建！" & vbCrLf & _
           "表头在第二行，数据从第三行开始。" & vbCrLf & _
           "请使用'测试核对功能'进行测试。", vbInformation
End Sub

