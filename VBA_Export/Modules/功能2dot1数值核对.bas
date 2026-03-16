Attribute VB_Name = "功能2dot1数值核对"
Sub 执行表格比对()
    Dim configWs As Worksheet
    Dim configData As Variant
    Dim fd As FileDialog
    Dim wbA As Workbook, wbB As Workbook
    Dim selectedFileA As Variant, selectedFileB As Variant
    Dim i As Long
    Dim resultWb As Workbook
    Dim resultWs As Worksheet
    Dim resultRow As Long
    Dim totalChecks As Long, passedChecks As Long, failedChecks As Long
    Dim t0 As Double

    t0 = Timer
    RunLog_WriteRow "2.1 表格对比", "开始", "", "", "", "", "开始", ""
    
    ' 获取配置表
    On Error Resume Next
    Set configWs = ThisWorkbook.Sheets("表格比对")
    If Err.Number <> 0 Then
        MsgBox "未找到名为'表格比对'的工作表，请检查！", vbExclamation
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 读取配置数据
    configData = 读取配置表(configWs)
    If configData(1, 1) = "" Then
        RunLog_WriteRow "2.1 表格对比", "完成", "", "", "", "失败", "配置表为空", CStr(Round(Timer - t0, 2))
        MsgBox "配置表为空，请检查！", vbExclamation
        Exit Sub
    End If
    
    ' 选择工作簿A
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "选择工作簿A"
    fd.AllowMultiSelect = False
    fd.Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"
    
    If Not fd.Show Then RunLog_WriteRow "2.1 表格对比", "完成", "", "", "", "", "已取消(未选表A)", CStr(Round(Timer - t0, 2)): Exit Sub
    selectedFileA = fd.SelectedItems(1)
    
    ' 选择工作簿B
    fd.Title = "选择工作簿B"
    If Not fd.Show Then RunLog_WriteRow "2.1 表格对比", "完成", "", "", "", "", "已取消(未选表B)", CStr(Round(Timer - t0, 2)): Exit Sub
    selectedFileB = fd.SelectedItems(1)
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    
    ' 打开工作簿
    On Error Resume Next
    Set wbA = Workbooks.Open(selectedFileA, ReadOnly:=True)
    Set wbB = Workbooks.Open(selectedFileB, ReadOnly:=True)
    If Err.Number <> 0 Then
        RunLog_WriteRow "2.1 表格对比", "完成", "", "", "", "失败", "无法打开工作簿 " & Err.Number, CStr(Round(Timer - t0, 2))
        MsgBox "无法打开工作簿，请检查文件路径和格式！", vbExclamation
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.EnableEvents = True
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 创建结果工作簿
    Set resultWb = Workbooks.Add
    Set resultWs = resultWb.Sheets(1)
    resultWs.Name = "比对结果"
    
    ' 设置结果表头
    resultWs.Cells(1, 1) = "序号"
    resultWs.Cells(1, 2) = "表A工作表"
    resultWs.Cells(1, 3) = "表B工作表"
    resultWs.Cells(1, 4) = "比对类型"
    resultWs.Cells(1, 5) = "单元格位置"
    resultWs.Cells(1, 6) = "表A值"
    resultWs.Cells(1, 7) = "表B值"
    resultWs.Cells(1, 8) = "比对结果"
    resultWs.Cells(1, 9) = "差异说明"
    resultWs.Range("A1:I1").Font.Bold = True
    resultWs.Range("A1:I1").Interior.Color = RGB(200, 200, 200)
    
    resultRow = 2
    totalChecks = 0
    passedChecks = 0
    failedChecks = 0
    
    ' 遍历配置表执行比对
    For i = 2 To UBound(configData, 1)  ' 从第2行开始（跳过表头）
        ' 检查是否执行
        If configData(i, 11) = 1 Then  ' K列：是否执行
            totalChecks = totalChecks + 1
            
            ' 获取表A和表B的工作表信息
            Dim wsA As Worksheet, wsB As Worksheet
            Dim wsARequiredFields As String, wsBRequiredFields As String
            Dim 全量比对 As Boolean
            Dim 位置A As String, 位置B As String
            
            ' 构建所需字段列表
            wsARequiredFields = 构建工作表名称(configData, i, 1, 5)  ' 表A：前5列
            wsBRequiredFields = 构建工作表名称(configData, i, 6, 10) ' 表B：6-10列
            
            全量比对 = (configData(i, 12) = 1)  ' L列：是否全量核对
            位置A = configData(i, 13)  ' M列：表A-位置
            位置B = configData(i, 14)  ' N列：表B-位置
            
            ' 查找工作表（包含所有必需字段）
            Set wsA = 查找工作表(wbA, wsARequiredFields)
            Set wsB = 查找工作表(wbB, wsBRequiredFields)
            
            If Not wsA Is Nothing And Not wsB Is Nothing Then
                ' 执行比对
                If 全量比对 Then
                    ' 全量比对
                    Call 执行全量比对(wsA, wsB, resultWs, resultRow, wsA.Name, wsB.Name)
                Else
                    ' 按位置比对
                    Call 执行位置比对(wsA, wsB, 位置A, 位置B, resultWs, resultRow, wsA.Name, wsB.Name)
                End If
                RunLog_WriteRow "2.1 表格对比", "比对", wsARequiredFields & "|" & wsBRequiredFields, "", "", "通过", "OK", ""
            Else
                ' 记录错误
                RunLog_WriteRow "2.1 表格对比", "比对", wsARequiredFields & "|" & wsBRequiredFields, "", "", "失败", "未找到匹配工作表", ""
                resultWs.Cells(resultRow, 1) = totalChecks
                resultWs.Cells(resultRow, 2) = wsARequiredFields
                resultWs.Cells(resultRow, 3) = wsBRequiredFields
                resultWs.Cells(resultRow, 4) = "工作表查找"
                resultWs.Cells(resultRow, 5) = "N/A"
                resultWs.Cells(resultRow, 6) = "N/A"
                resultWs.Cells(resultRow, 7) = "N/A"
                resultWs.Cells(resultRow, 8) = "失败"
                resultWs.Cells(resultRow, 9) = "未找到包含所有字段的工作表"
                resultWs.Rows(resultRow).Interior.Color = RGB(255, 200, 200) ' 红色背景
                resultRow = resultRow + 1
                failedChecks = failedChecks + 1
            End If
        End If
    Next i
    
    ' 关闭工作簿
    wbA.Close SaveChanges:=False
    wbB.Close SaveChanges:=False
    
    ' 计算通过率
    passedChecks = totalChecks - failedChecks
    RunLog_WriteRow "2.1 表格对比", "完成", "", "", "", "", "Done 共" & totalChecks & "项 通过" & passedChecks & " 失败" & failedChecks, CStr(Round(Timer - t0, 2))
    
    ' 添加统计信息
    ' resultWs.Cells(resultRow, 1) = "统计信息:"
    'resultWs.Cells(resultRow, 2) = "总比对项: " & totalChecks
    'resultWs.Cells(resultRow, 3) = "通过项: " & passedChecks
    'resultWs.Cells(resultRow, 4) = "失败项: " & failedChecks
    'resultWs.Cells(resultRow, 5) = "通过率: " & Format(passedChecks / totalChecks, "0.00%")
    'resultWs.Range("A" & resultRow & ":E" & resultRow).Font.Bold = True
    'resultWs.Range("A" & resultRow & ":E" & resultRow).Interior.Color = RGB(200, 230, 255)
    
    ' 调整列宽
    resultWs.Columns.AutoFit
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    
    ' 显示结果
    'MsgBox "表格比对完成！" & vbCrLf & _
    '       "总比对项: " & totalChecks & vbCrLf & _
    '       "通过项: " & passedChecks & vbCrLf & _
    '       "失败项: " & failedChecks & vbCrLf & _
    '       "通过率: " & Format(passedChecks / totalChecks, "0.00%"), vbInformation
End Sub

Function 读取配置表(ws As Worksheet) As Variant
    Dim lastRow As Long, lastCol As Long
    Dim dataRange As Range
    
    ' 获取数据范围
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    
    If lastRow < 2 Then
        读取配置表 = Array(Array(""))
        Exit Function
    End If
    
    Set dataRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    读取配置表 = dataRange.value
End Function

Function 构建工作表名称(configData As Variant, rowNum As Long, startCol As Long, endCol As Long) As String
    Dim i As Long
    Dim requiredFields As Collection
    Dim wsName As String
    
    Set requiredFields = New Collection
    
    ' 收集所有非空字段
    For i = startCol To endCol
        If Not IsEmpty(configData(rowNum, i)) And configData(rowNum, i) <> "" Then
            requiredFields.Add CStr(configData(rowNum, i))
        End If
    Next i
    
    ' 如果没有任何字段，返回空字符串
    If requiredFields.count = 0 Then
        构建工作表名称 = ""
        Exit Function
    End If
    
    ' 返回字段列表（用于查找包含所有字段的工作表）
    wsName = ""
    For i = 1 To requiredFields.count
        If wsName <> "" Then
            wsName = wsName & ","
        End If
        wsName = wsName & requiredFields.Item(i)
    Next i
    
    构建工作表名称 = wsName
End Function

Function 查找工作表(wb As Workbook, requiredFieldsStr As String) As Worksheet
    Dim ws As Worksheet
    Dim requiredFields() As String
    Dim i As Long
    Dim allFieldsPresent As Boolean
    
    ' 如果字段字符串为空，返回Nothing
    If requiredFieldsStr = "" Then
        Set 查找工作表 = Nothing
        Exit Function
    End If
    
    ' 解析所需字段
    requiredFields = Split(requiredFieldsStr, ",")
    
    ' 遍历所有工作表，查找包含所有字段的工作表
    For Each ws In wb.Worksheets
        allFieldsPresent = True
        
        ' 检查工作表名称是否包含所有必需字段
        For i = 0 To UBound(requiredFields)
            If InStr(1, ws.Name, requiredFields(i), vbTextCompare) = 0 Then
                allFieldsPresent = False
                Exit For
            End If
        Next i
        
        ' 如果找到包含所有字段的工作表，返回它
        If allFieldsPresent Then
            Set 查找工作表 = ws
            Exit Function
        End If
    Next ws
    
    ' 如果没有找到匹配的工作表，返回Nothing
    Set 查找工作表 = Nothing
End Function

Sub 执行全量比对(wsA As Worksheet, wsB As Worksheet, resultWs As Worksheet, ByRef resultRow As Long, wsAName As String, wsBName As String)
    Dim usedRangeA As Range, usedRangeB As Range
    Dim lastRowA As Long, lastColA As Long
    Dim lastRowB As Long, lastColB As Long
    Dim i As Long, j As Long
    Dim cellA As Range, cellB As Range
    Dim passedChecks As Long, failedChecks As Long
    
    ' 获取使用范围
    On Error Resume Next
    Set usedRangeA = wsA.UsedRange
    Set usedRangeB = wsB.UsedRange
    On Error GoTo 0
    
    If usedRangeA Is Nothing Or usedRangeB Is Nothing Then
        ' 记录错误
        resultWs.Cells(resultRow, 1) = resultRow - 1
        resultWs.Cells(resultRow, 2) = wsAName
        resultWs.Cells(resultRow, 3) = wsBName
        resultWs.Cells(resultRow, 4) = "全量比对"
        resultWs.Cells(resultRow, 5) = "N/A"
        resultWs.Cells(resultRow, 6) = "N/A"
        resultWs.Cells(resultRow, 7) = "N/A"
        resultWs.Cells(resultRow, 8) = "失败"
        resultWs.Cells(resultRow, 9) = "工作表为空"
        resultWs.Rows(resultRow).Interior.Color = RGB(255, 200, 200)
        resultRow = resultRow + 1
        Exit Sub
    End If
    
    lastRowA = usedRangeA.Rows.count
    lastColA = usedRangeA.Columns.count
    lastRowB = usedRangeB.Rows.count
    lastColB = usedRangeB.Columns.count
    
    ' 取较小范围
    Dim maxRow As Long, maxCol As Long
    maxRow = Application.WorksheetFunction.Min(lastRowA, lastRowB)
    maxCol = Application.WorksheetFunction.Min(lastColA, lastColB)
    
    passedChecks = 0
    failedChecks = 0
    
    ' 遍历所有单元格
    For i = 1 To maxRow
        For j = 1 To maxCol
            Set cellA = wsA.Cells(i, j)
            Set cellB = wsB.Cells(i, j)
            
            ' 比对单元格
            If 比对单元格(cellA, cellB) Then
                passedChecks = passedChecks + 1
                ' 记录成功的
                resultWs.Cells(resultRow, 1) = resultRow - 1
                resultWs.Cells(resultRow, 2) = wsAName
                resultWs.Cells(resultRow, 3) = wsBName
                resultWs.Cells(resultRow, 4) = "全量比对"
                resultWs.Cells(resultRow, 5) = cellA.Address(False, False)
                resultWs.Cells(resultRow, 6) = 获取单元格值(cellA)
                resultWs.Cells(resultRow, 7) = 获取单元格值(cellB)
                resultWs.Cells(resultRow, 8) = "成功"
                resultWs.Cells(resultRow, 9) = "单元格值相同"
                resultWs.Rows(resultRow).Interior.Color = RGB(355, 300, 300)
                resultRow = resultRow + 1
            Else
                failedChecks = failedChecks + 1
                
                ' 记录差异
                resultWs.Cells(resultRow, 1) = resultRow - 1
                resultWs.Cells(resultRow, 2) = wsAName
                resultWs.Cells(resultRow, 3) = wsBName
                resultWs.Cells(resultRow, 4) = "全量比对"
                resultWs.Cells(resultRow, 5) = cellA.Address(False, False)
                resultWs.Cells(resultRow, 6) = 获取单元格值(cellA)
                resultWs.Cells(resultRow, 7) = 获取单元格值(cellB)
                resultWs.Cells(resultRow, 8) = "失败"
                resultWs.Cells(resultRow, 9) = "单元格值不匹配"
                resultWs.Rows(resultRow).Interior.Color = RGB(255, 200, 200)
                resultRow = resultRow + 1
            End If
        Next j
    Next i
    
    ' 记录统计
    resultWs.Cells(resultRow, 1) = resultRow - 1
    resultWs.Cells(resultRow, 2) = wsAName
    resultWs.Cells(resultRow, 3) = wsBName
    resultWs.Cells(resultRow, 4) = "全量比对统计"
    resultWs.Cells(resultRow, 5) = "范围: A1:" & Chr(64 + maxCol) & maxRow
    resultWs.Cells(resultRow, 6) = "通过: " & passedChecks
    resultWs.Cells(resultRow, 7) = "失败: " & failedChecks
    resultWs.Cells(resultRow, 8) = IIf(failedChecks = 0, "通过", "失败")
    resultWs.Cells(resultRow, 9) = "总比对: " & (passedChecks + failedChecks)
    resultWs.Rows(resultRow).Interior.Color = IIf(failedChecks = 0, RGB(200, 255, 200), RGB(255, 200, 200))
    resultRow = resultRow + 1
End Sub

Sub 执行位置比对(wsA As Worksheet, wsB As Worksheet, 位置A As String, 位置B As String, resultWs As Worksheet, ByRef resultRow As Long, wsAName As String, wsBName As String)
    Dim cellA As Range, cellB As Range
    
    ' 获取指定位置的单元格
    On Error Resume Next
    Set cellA = wsA.Range(位置A)
    Set cellB = wsB.Range(位置B)
    On Error GoTo 0
    
    If cellA Is Nothing Or cellB Is Nothing Then
        ' 记录错误
        resultWs.Cells(resultRow, 1) = resultRow - 1
        resultWs.Cells(resultRow, 2) = wsAName
        resultWs.Cells(resultRow, 3) = wsBName
        resultWs.Cells(resultRow, 4) = "位置比对"
        resultWs.Cells(resultRow, 5) = 位置A & " vs " & 位置B
        resultWs.Cells(resultRow, 6) = "N/A"
        resultWs.Cells(resultRow, 7) = "N/A"
        resultWs.Cells(resultRow, 8) = "失败"
        resultWs.Cells(resultRow, 9) = "指定位置不存在"
        resultWs.Rows(resultRow).Interior.Color = RGB(255, 200, 200)
        resultRow = resultRow + 1
        Exit Sub
    End If
    
    ' 比对单元格
    If 比对单元格(cellA, cellB) Then
        resultWs.Cells(resultRow, 1) = resultRow - 1
        resultWs.Cells(resultRow, 2) = wsAName
        resultWs.Cells(resultRow, 3) = wsBName
        resultWs.Cells(resultRow, 4) = "位置比对"
        resultWs.Cells(resultRow, 5) = 位置A & " vs " & 位置B
        resultWs.Cells(resultRow, 6) = 获取单元格值(cellA)
        resultWs.Cells(resultRow, 7) = 获取单元格值(cellB)
        resultWs.Cells(resultRow, 8) = "通过"
        resultWs.Cells(resultRow, 9) = "单元格值匹配"
        resultWs.Rows(resultRow).Interior.Color = RGB(200, 255, 200)
    Else
        resultWs.Cells(resultRow, 1) = resultRow - 1
        resultWs.Cells(resultRow, 2) = wsAName
        resultWs.Cells(resultRow, 3) = wsBName
        resultWs.Cells(resultRow, 4) = "位置比对"
        resultWs.Cells(resultRow, 5) = 位置A & " vs " & 位置B
        resultWs.Cells(resultRow, 6) = 获取单元格值(cellA)
        resultWs.Cells(resultRow, 7) = 获取单元格值(cellB)
        resultWs.Cells(resultRow, 8) = "失败"
        resultWs.Cells(resultRow, 9) = "单元格值不匹配"
        resultWs.Rows(resultRow).Interior.Color = RGB(255, 200, 200)
    End If
    resultRow = resultRow + 1
End Sub

Function 比对单元格(cellA As Range, cellB As Range) As Boolean
    Dim valueA As Variant, valueB As Variant
    
    valueA = 获取单元格值(cellA)
    valueB = 获取单元格值(cellB)
    
    ' 处理空值
    If IsEmpty(valueA) And IsEmpty(valueB) Then
        比对单元格 = True
        Exit Function
    ElseIf IsEmpty(valueA) Or IsEmpty(valueB) Then
        比对单元格 = False
        Exit Function
    End If
    
    ' 处理错误值
    If IsError(valueA) Or IsError(valueB) Then
        比对单元格 = False
        Exit Function
    End If
    
    ' 数值比较
    If IsNumeric(valueA) And IsNumeric(valueB) Then
        比对单元格 = Abs(CDbl(valueA) - CDbl(valueB)) < 0.0001
    Else
        ' 文本比较（不区分大小写）
        比对单元格 = (Trim(CStr(valueA)) = Trim(CStr(valueB)))
    End If
End Function

Function 获取单元格值(cell As Range) As Variant
    On Error Resume Next
    If cell Is Nothing Then
        获取单元格值 = "N/A"
    Else
        获取单元格值 = cell.value
    End If
    On Error GoTo 0
End Function

