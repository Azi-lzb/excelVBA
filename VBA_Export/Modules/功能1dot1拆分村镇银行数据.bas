Attribute VB_Name = "功能1dot1拆分村镇银行数据"
Function 获取村镇银行所在行(ws As Worksheet) As String
    '功能：在A列中查找"村镇银行"，删除该行，并插入两行新数据
    '参数：ws - 要处理的工作表
    '返回值：字符串，包含新插入两行的行号信息
    
    Dim lastRow As Long
    Dim i As Long
    Dim targetRow As Long
    Dim newRow1 As Long, newRow2 As Long
    Dim resultMsg As String
    
    '初始化
    targetRow = 0
    resultMsg = ""
    
    '检查工作表是否有效
    If ws Is Nothing Then
        获取村镇银行所在行 = "错误：工作表对象无效"
        Exit Function
    End If
    
    '获取最后一行
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    
    '在A列中查找"村镇银行"
    For i = 1 To lastRow
        If ws.Cells(i, 1).value = "村镇银行" Then
            targetRow = i
            Exit For
        End If
    Next i
    
    '如果没找到，返回提示
    If targetRow = 0 Then
        获取村镇银行所在行 = "未找到包含'村镇银行'的行"
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    
    '删除找到的行
    ws.Rows(targetRow).Delete Shift:=xlUp
    
    '在删除行的位置插入两行
    ws.Rows(targetRow & ":" & targetRow + 1).Insert Shift:=xlDown
    
    '设置新插入行的数据
    ws.Cells(targetRow, 1).value = "博罗长江村镇银行"
    ws.Cells(targetRow + 1, 1).value = "惠东惠民村镇银行"

    
    Application.ScreenUpdating = True
    获取村镇银行所在行 = targetRow
End Function

Function 计算村镇银行() As Variant
    '功能：计算本期与去年同期村镇银行三类指标的存款和贷款数据同比变化
    '返回值：6行8列的数组，每行对应一个指标，前4列存款数据，后4列贷款数据
    '格式：
    '行1: 村镇银行本外币 [存款C5D5E5+同比, 贷款CI5CJ5CK5+同比]
    '行2: 村镇银行人民币 [存款C5D5E5+同比, 贷款CI5CJ5CK5+同比]
    '行3: 村镇银行外汇   [存款C5D5E5+同比, 贷款CI5CJ5CK5+同比]
    '行4: 村镇银行本外币 [存款C6D6E6+同比, 贷款CI6CJ6CK6+同比]
    '行5: 村镇银行人民币 [存款C6D6E6+同比, 贷款CI6CJ6CK6+同比]
    '行6: 村镇银行外汇   [存款C6D6E6+同比, 贷款CI6CJ6CK6+同比]
    
    Dim fd As FileDialog
    Dim 本期文件 As Workbook, 去年同期文件 As Workbook
    Dim ws1 As Worksheet, ws2 As Worksheet
    Dim 结果数组(1 To 6, 1 To 8) As Variant
    Dim 指标列表 As Variant
    Dim 当前行 As Long
    Dim 本期值 As Double, 去年值 As Double
    Dim 同比数据 As Variant
    Dim 本期文件路径 As String, 去年同期文件路径 As String
    
    '初始化结果数组
    For i = 1 To 6
        For j = 1 To 8
            结果数组(i, j) = 0
        Next j
    Next i
    
    '定义要处理的三个指标
    指标列表 = Array("村镇银行本外币", "村镇银行人民币", "村镇银行外汇")
    
    '选择本期文件
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "选择本期文件（包含村镇银行数据）"
    fd.AllowMultiSelect = False
    fd.Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"
    
    If Not fd.Show Then
        计算村镇银行 = 结果数组
        Exit Function
    End If
    本期文件路径 = fd.SelectedItems(1)
    
    '选择去年同期文件
    fd.Title = "选择去年同期文件（包含村镇银行数据）"
    If Not fd.Show Then
        计算村镇银行 = 结果数组
        Exit Function
    End If
    去年同期文件路径 = fd.SelectedItems(1)
    
    '隐式打开文件（不显示窗口）
    On Error Resume Next
    Set 本期文件 = GetObject(本期文件路径)
    Set 去年同期文件 = GetObject(去年同期文件路径)
    On Error GoTo 0
    
    '检查文件是否成功打开
    If 本期文件 Is Nothing Or 去年同期文件 Is Nothing Then
        MsgBox "无法打开文件，请检查文件路径和格式", vbExclamation
        计算村镇银行 = 结果数组
        Exit Function
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    '遍历三个指标
    当前行 = 1
    For Each 指标 In 指标列表
        '查找包含当前指标的工作表
        For Each ws1 In 本期文件.Worksheets
            If InStr(1, ws1.Name, 指标, vbTextCompare) > 0 Then
                '在去年同期文件中查找同名工作表
                On Error Resume Next
                Set ws2 = 去年同期文件.Sheets(ws1.Name)
                On Error GoTo 0
                
                If Not ws2 Is Nothing Then
                    '=== 处理C5行数据 ===
                    '存款数据（C、D、E列）- 前4列
                    本期值 = 0
                    去年值 = 0
                    
                    If IsNumeric(ws1.Range("C5").value) Then
                        本期值 = CDbl(ws1.Range("C5").value)
                    End If
                    If IsNumeric(ws2.Range("C5").value) Then
                        去年值 = CDbl(ws2.Range("C5").value)
                    End If
                    
                    同比数据 = 安全同比计算(本期值, 去年值)
                    结果数组(当前行, 1) = 本期值
                    结果数组(当前行, 2) = ws1.Range("D5").value
                    结果数组(当前行, 3) = ws1.Range("E5").value
                    结果数组(当前行, 4) = 同比数据
                    
                    '贷款数据（CI、CJ、CK列）- 后4列
                    本期值 = 0
                    去年值 = 0
                    
                    If IsNumeric(ws1.Range("CI5").value) Then
                        本期值 = CDbl(ws1.Range("CI5").value)
                    End If
                    If IsNumeric(ws2.Range("CI5").value) Then
                        去年值 = CDbl(ws2.Range("CI5").value)
                    End If
                    
                    同比数据 = 安全同比计算(本期值, 去年值)
                    结果数组(当前行, 5) = 本期值
                    结果数组(当前行, 6) = ws1.Range("CJ5").value
                    结果数组(当前行, 7) = ws1.Range("CK5").value
                    结果数组(当前行, 8) = 同比数据
                    
                    '=== 处理C6行数据（当前行+3）===
                    If 当前行 + 3 <= 6 Then '防止数组越界
                        '存款数据（C、D、E列）- 前4列
                        本期值 = 0
                        去年值 = 0
                        
                        If IsNumeric(ws1.Range("C6").value) Then
                            本期值 = CDbl(ws1.Range("C6").value)
                        End If
                        If IsNumeric(ws2.Range("C6").value) Then
                            去年值 = CDbl(ws2.Range("C6").value)
                        End If
                        
                        同比数据 = 安全同比计算(本期值, 去年值)
                        结果数组(当前行 + 3, 1) = 本期值
                        结果数组(当前行 + 3, 2) = ws1.Range("D6").value
                        结果数组(当前行 + 3, 3) = ws1.Range("E6").value
                        结果数组(当前行 + 3, 4) = 同比数据
                        
                        '贷款数据（CI、CJ、CK列）- 后4列
                        本期值 = 0
                        去年值 = 0
                        
                        If IsNumeric(ws1.Range("CI6").value) Then
                            本期值 = CDbl(ws1.Range("CI6").value)
                        End If
                        If IsNumeric(ws2.Range("CI6").value) Then
                            去年值 = CDbl(ws2.Range("CI6").value)
                        End If
                        
                        同比数据 = 安全同比计算(本期值, 去年值)
                        结果数组(当前行 + 3, 5) = 本期值
                        结果数组(当前行 + 3, 6) = ws1.Range("CJ6").value
                        结果数组(当前行 + 3, 7) = ws1.Range("CK6").value
                        结果数组(当前行 + 3, 8) = 同比数据
                    End If
                    
                    Exit For '找到第一个匹配工作表就退出
                End If
            End If
        Next ws1
        
        当前行 = 当前行 + 1 '移动到下一个指标
    Next 指标
    
    '关闭文件（不保存更改）
    On Error Resume Next
    本期文件.Close SaveChanges:=False
    去年同期文件.Close SaveChanges:=False
    On Error GoTo 0
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    '返回结果数组
    计算村镇银行 = 结果数组
End Function

'安全同比计算函数
Function 安全同比计算(本期值 As Double, 去年值 As Double) As Variant
    '专门用于同比计算的安全函数
    '返回值：Double数值或字符串标记
    
    If 去年值 = 0 Then
        If 本期值 = 0 Then
            安全同比计算 = 0 '都为0，同比为0
        Else
            安全同比计算 = "∞" '去年为0，本期不为0，视为无穷大增长
        End If
    Else
        安全同比计算 = Round(((本期值 / 去年值) - 1) * 100, 2)
    End If
End Function

Sub 批量填入村镇银行计算结果()
    Dim fd As FileDialog
    Dim selectedFile As Variant
    Dim targetWb As Workbook
    Dim targetWs As Worksheet
    Dim 结果数组 As Variant
    Dim 村镇银行行号 As Long
    Dim fileCount As Integer
    Dim 分机构工作表数 As Integer
    Dim 行号结果 As String
    Dim t0 As Double
    Dim 当前文件 As String
    Dim 本文件分机构数 As Integer

    t0 = Timer
    RunLog_WriteRow "1.1 拆分村镇银行", "开始", "", "", "", "", "开始", ""

    '先计算村镇银行数据
    结果数组 = 计算村镇银行()
    
    '检查计算结果是否有效
    If 结果数组(1, 1) = 0 And 结果数组(1, 5) = 0 Then
        RunLog_WriteRow "1.1 拆分村镇银行", "完成", "", "", "", "失败", "计算结果为空", CStr(Round(Timer - t0, 2))
        MsgBox "村镇银行计算结果为空，请先确保计算成功", vbExclamation
        Exit Sub
    End If
    
    '选择包含"分机构"工作表的Excel文件
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "选择包含'分机构'工作表的Excel文件（注：本外币、人民币、外币、外发用）"
    fd.AllowMultiSelect = True
    fd.Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"
    
    If fd.Show Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        
        fileCount = 0
        分机构工作表数 = 0
        
        For Each selectedFile In fd.SelectedItems
            本文件分机构数 = 0
            '使用常规方式打开文件，添加更多参数确保文件不被锁定
            On Error Resume Next
            Set targetWb = Workbooks.Open( _
                fileName:=selectedFile, _
                ReadOnly:=False, _
                UpdateLinks:=0, _
                IgnoreReadOnlyRecommended:=True, _
                Notify:=False _
            )
            If Err.Number <> 0 Then
                当前文件 = Dir(selectedFile): If 当前文件 = "" Then 当前文件 = CStr(selectedFile)
                RunLog_WriteRow "1.1 拆分村镇银行", "处理文件", 当前文件, "", "", "失败", Err.Number & " " & Err.Description, ""
                Debug.Print "无法打开文件: " & selectedFile & " - " & Err.Description
                On Error GoTo 0
                GoTo NextFile
            End If
            On Error GoTo 0
            
            fileCount = fileCount + 1
            
            '遍历工作簿中的工作表，查找包含"分机构"的工作表
            For Each targetWs In targetWb.Worksheets
                If InStr(1, targetWs.Name, "分机构", vbTextCompare) > 0 Then
                    分机构工作表数 = 分机构工作表数 + 1
                    本文件分机构数 = 本文件分机构数 + 1
                    
                    '调用你现有的函数获取村镇银行行号
                    行号结果 = 获取村镇银行所在行(targetWs)
                    
                    '解析返回的行号
                    If IsNumeric(行号结果) Then
                        村镇银行行号 = CLng(行号结果)
                        
                        '根据工作表名称判断类型并填入数据
                        Call 填入计算结果(targetWs, 村镇银行行号, 结果数组)
                    Else
                        '如果返回的是错误信息，显示提示
                        Debug.Print "工作表 " & targetWs.Name & ": " & 行号结果
                    End If
                End If
            Next targetWs
            
            RunLog_WriteRow "1.1 拆分村镇银行", "处理文件", targetWb.Name, "", "", "成功", "分机构表 " & 本文件分机构数 & " 个", ""

            '重要：确保文件完全关闭和释放
            On Error Resume Next
            targetWb.Save
            targetWb.Close SaveChanges:=True
            Set targetWb = Nothing
            '添加短暂延迟，确保系统释放文件句柄
            Application.Wait (Now + TimeValue("0:00:01"))
            On Error GoTo 0
            
NextFile:
        Next selectedFile
        
        '恢复Excel设置
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.CutCopyMode = False
        
        RunLog_WriteRow "1.1 拆分村镇银行", "完成", "", "", "", "", "Done", CStr(Round(Timer - t0, 2))
        '显示处理结果
        MsgBox "批量填入完成！" & vbCrLf & _
               "处理文件数: " & fileCount & vbCrLf & _
               "处理分机构工作表数: " & 分机构工作表数 & vbCrLf & _
               "村镇银行计算结果已填入对应行", vbInformation
    Else
        RunLog_WriteRow "1.1 拆分村镇银行", "完成", "", "", "", "", "未选择文件", CStr(Round(Timer - t0, 2))
        MsgBox "未选择任何文件", vbInformation
    End If
End Sub

Sub 填入计算结果(ws As Worksheet, 村镇银行行号 As Long, 结果数组 As Variant)
    '功能：根据工作表类型将计算结果填入指定行
    '参数：
    '   ws: 目标工作表
    '   村镇银行行号: 包含"村镇银行"的行号
    '   结果数组: 6×8的计算结果数组
    
    Dim 工作表类型 As String
    
    '根据工作表名称判断类型
    If InStr(1, ws.Name, "本外币", vbTextCompare) > 0 Then
        工作表类型 = "本外币"
    ElseIf InStr(1, ws.Name, "人民币", vbTextCompare) > 0 Then
        工作表类型 = "人民币"
    ElseIf InStr(1, ws.Name, "外汇", vbTextCompare) > 0 Then
        工作表类型 = "外汇"
    Else
        '无法识别类型，跳过
        Exit Sub
    End If
    
    '根据工作表类型填入数据
    Select Case 工作表类型
        Case "本外币"
            '本外币分机构：结果数组(1,i)填入K行，结果数组(4,i)填入K+1行
            For i = 1 To 8
                ws.Cells(村镇银行行号, i + 1) = 结果数组(1, i) 'K行
                ws.Cells(村镇银行行号 + 1, i + 1) = 结果数组(4, i) 'K+1行
            Next i
            
        Case "人民币"
            '人民币分机构：结果数组(2,i)填入K行，结果数组(5,i)填入K+1行
            For i = 1 To 8
                ws.Cells(村镇银行行号, i + 1) = 结果数组(2, i) 'K行
                ws.Cells(村镇银行行号 + 1, i + 1) = 结果数组(5, i) 'K+1行
            Next i
            
        Case "外汇"
            '外汇分机构：结果数组(3,i)填入K行，结果数组(6,i)填入K+1行
            For i = 1 To 8
                ws.Cells(村镇银行行号, i + 1) = 结果数组(3, i) 'K行
                ws.Cells(村镇银行行号 + 1, i + 1) = 结果数组(6, i) 'K+1行
            Next i
    End Select
    
End Sub

