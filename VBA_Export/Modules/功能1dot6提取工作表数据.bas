Attribute VB_Name = "功能1dot6提取工作表数据"
Sub 批量提取工作表数据()
    Dim fd As FileDialog
    Dim configSheet As Worksheet
    Dim lastRow As Long, i As Long
    Dim outputDict As Object
    Dim processedCount As Long
    Dim fileItem As Variant
    Dim t0 As Double
    
    t0 = Timer
    RunLog_WriteRow "1.6 提取工作表", "开始", "", "", "", "", "开始", ""
    
    ' 初始化
    processedCount = 0
    
    ' 在当前工作簿中查找配置表
    On Error Resume Next
    Set configSheet = ThisWorkbook.Worksheets("工作表提取")
    On Error GoTo 0
    
    If configSheet Is Nothing Then
        RunLog_WriteRow "1.6 提取工作表", "完成", "", "", "", "失败", "找不到'工作表提取'配置表", CStr(Round(Timer - t0, 2))
        MsgBox "当前工作簿中找不到名为'工作表提取'的工作表！" & vbCrLf & _
               "请确保配置表的工作表名称正确。", vbExclamation
        Exit Sub
    End If
    
    ' 获取配置数据
    lastRow = configSheet.Cells(configSheet.Rows.count, "A").End(xlUp).row
    
    If lastRow < 2 Then
        RunLog_WriteRow "1.6 提取工作表", "完成", "", "", "", "失败", "配置表无数据", CStr(Round(Timer - t0, 2))
        MsgBox "配置表中没有数据！", vbExclamation
        Exit Sub
    End If
    
    ' 选择源文件
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "请选择要提取的Excel文件（可多选）"
    fd.Filters.Add "Excel文件", "*.xls; *.xlsx; *.xlsm"
    fd.AllowMultiSelect = True
    
    If fd.Show = -1 Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        
        ' 创建字典来管理输出文件
        Set outputDict = CreateObject("Scripting.Dictionary")
        
        ' [优化A] 预读所有配置行，避免每条配置都重复打开同一文件
        Dim iCfg As Long
        Dim cfgSkip() As Boolean
        Dim cfgFullExtract() As Boolean
        Dim cfgSheetInfo() As Variant
        Dim cfgOutputFile() As String
        Dim cfgRows() As String
        Dim cfgCols() As String
        ReDim cfgSkip(2 To lastRow)
        ReDim cfgFullExtract(2 To lastRow)
        ReDim cfgSheetInfo(2 To lastRow, 1 To 5)
        ReDim cfgOutputFile(2 To lastRow)
        ReDim cfgRows(2 To lastRow)
        ReDim cfgCols(2 To lastRow)
        For iCfg = 2 To lastRow
            With configSheet
                cfgSkip(iCfg) = (.Cells(iCfg, "F").value = "是" Or .Cells(iCfg, "F").value = "1" Or .Cells(iCfg, "F").value = True)
                If Not cfgSkip(iCfg) Then
                    cfgSheetInfo(iCfg, 1) = CStr(.Cells(iCfg, "A").value)
                    cfgSheetInfo(iCfg, 2) = CStr(.Cells(iCfg, "B").value)
                    cfgSheetInfo(iCfg, 3) = CStr(.Cells(iCfg, "C").value)
                    cfgSheetInfo(iCfg, 4) = CStr(.Cells(iCfg, "D").value)
                    cfgSheetInfo(iCfg, 5) = CStr(.Cells(iCfg, "E").value)
                    cfgFullExtract(iCfg) = (.Cells(iCfg, "G").value = "是" Or .Cells(iCfg, "G").value = "1" Or .Cells(iCfg, "G").value = True)
                    cfgRows(iCfg) = CStr(.Cells(iCfg, "H").value)
                    cfgCols(iCfg) = CStr(.Cells(iCfg, "I").value)
                    cfgOutputFile(iCfg) = CStr(.Cells(iCfg, "J").value)
                    If cfgOutputFile(iCfg) = "" Then
                        cfgOutputFile(iCfg) = "提取结果_" & Format(Now, "yyyymmdd_hhmmss")
                    End If
                End If
            End With
        Next iCfg
        
        ' 外层：每个源文件只打开一次
        Dim sourceWb As Workbook
        Dim sourceWs As Worksheet
        Dim targetWs As Worksheet
        Dim targetWb As Workbook
        Dim matchedSheets As Collection
        Dim wsItem As Variant
        Dim sheetInfoArr(1 To 5) As String
        
        For Each fileItem In fd.SelectedItems
            On Error Resume Next
            Set sourceWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True)
            If Err.Number <> 0 Then
                Debug.Print "无法打开文件: " & fileItem
                On Error GoTo 0
                GoTo NextFile
            End If
            On Error GoTo 0
            
            ' 内层：对该文件应用所有配置行
            For iCfg = 2 To lastRow
                If Not cfgSkip(iCfg) Then
                    sheetInfoArr(1) = cfgSheetInfo(iCfg, 1)
                    sheetInfoArr(2) = cfgSheetInfo(iCfg, 2)
                    sheetInfoArr(3) = cfgSheetInfo(iCfg, 3)
                    sheetInfoArr(4) = cfgSheetInfo(iCfg, 4)
                    sheetInfoArr(5) = cfgSheetInfo(iCfg, 5)
                    Set matchedSheets = 查找所有匹配工作表(sourceWb, sheetInfoArr)
                    If matchedSheets.count > 0 Then
                        For Each wsItem In matchedSheets
                            Set sourceWs = wsItem
                            If Not outputDict.Exists(cfgOutputFile(iCfg)) Then
                                Set targetWb = Workbooks.Add
                                targetWb.SaveAs sourceWb.path & "\" & cfgOutputFile(iCfg) & ".xlsx"
                                outputDict.Add cfgOutputFile(iCfg), targetWb
                            Else
                                Set targetWb = outputDict(cfgOutputFile(iCfg))
                            End If
                            On Error Resume Next
                            Set targetWs = targetWb.Worksheets.Add(After:=targetWb.Worksheets(targetWb.Worksheets.count))
                            targetWs.Name = 获取唯一工作表名称(targetWb, sourceWs.Name)
                            On Error GoTo 0
                            If cfgFullExtract(iCfg) Then
                                整表提取 sourceWs, targetWs
                            Else
                                If Len(Trim$(cfgRows(iCfg))) > 0 Or Len(Trim$(cfgCols(iCfg))) > 0 Then
                                    ExtractPartialData sourceWs, targetWs, cfgRows(iCfg), cfgCols(iCfg)
                                Else
                                    整表提取 sourceWs, targetWs
                                End If
                            End If
                            删除空白行列 targetWs
                            processedCount = processedCount + 1
                            Debug.Print "已提取: " & sourceWb.Name & " - " & sourceWs.Name & " -> " & cfgOutputFile(iCfg)
                        Next wsItem
                    End If
                End If
            Next iCfg
            
            sourceWb.Close SaveChanges:=False
NextFile:
            Set sourceWb = Nothing
        Next fileItem
        
        ' 保存所有输出文件
        Dim outputKey As Variant
        For Each outputKey In outputDict.keys
            outputDict(outputKey).Save
            RunLog_WriteRow "1.6 提取工作表", "输出文件", CStr(outputKey) & ".xlsx", "", "", "成功", "已保存", ""
            outputDict(outputKey).Close
        Next outputKey
        
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        RunLog_WriteRow "1.6 提取工作表", "完成", "", "", "", "", "Done 已提取 " & processedCount & " 个工作表", CStr(Round(Timer - t0, 2))
        ' 显示结果
        Dim msg As String
        msg = "提取完成！" & vbCrLf & vbCrLf
        
        If processedCount > 0 Then
            msg = msg & "已提取 " & processedCount & " 个工作表" & vbCrLf
            msg = msg & "输出文件: " & outputDict.count & " 个" & vbCrLf & vbCrLf
            msg = msg & "输出文件列表:"
            For Each outputKey In outputDict.keys
                msg = msg & vbCrLf & "? " & outputKey & ".xlsx"
            Next outputKey
        Else
            msg = msg & "没有找到匹配的工作表可提取。"
        End If
        
        MsgBox msg, vbInformation
        
    Else
        RunLog_WriteRow "1.6 提取工作表", "完成", "", "", "", "", "未选择源文件", CStr(Round(Timer - t0, 2))
        MsgBox "未选择源文件", vbInformation
    End If
    
    Set fd = Nothing
End Sub

' 查找所有匹配的工作表（新增函数：不只匹配第一个）
Function 查找所有匹配工作表(wb As Workbook, sheetInfo As Variant) As Collection
    Dim ws As Worksheet
    Dim targetName As String
    Dim matchedSheets As Collection
    
    Set matchedSheets = New Collection
    Dim addedNames As Object
    Set addedNames = CreateObject("Scripting.Dictionary")
    
    ' 构建目标工作表名称（用于精确匹配）
    targetName = 构建工作表名称(sheetInfo)
    
    ' 先尝试精确匹配
    On Error Resume Next
    Set ws = wb.Worksheets(targetName)
    If Not ws Is Nothing Then
        matchedSheets.Add ws
        addedNames(ws.Name) = 1
    End If
    On Error GoTo 0
    
    ' 然后进行模糊匹配
    For Each ws In wb.Worksheets
        ' 检查是否包含所有关键词，并且不是已经添加的工作表
        If 工作表包含所有关键词(ws, sheetInfo) Then
            ' 检查是否已经添加过这个工作表
            If Not addedNames.Exists(ws.Name) Then
                matchedSheets.Add ws
                addedNames(ws.Name) = 1
            End If
        End If
    Next ws
    
    Set 查找所有匹配工作表 = matchedSheets
End Function

' 构建工作表名称
Function 构建工作表名称(sheetInfo As Variant) As String
    Dim result As String
    Dim i As Integer
    
    For i = 1 To 5
        If Len(Trim(sheetInfo(i))) > 0 Then
            If Len(result) > 0 Then
                result = result & "-" & sheetInfo(i)
            Else
                result = sheetInfo(i)
            End If
        End If
    Next i
    
    构建工作表名称 = result
End Function

' 获取唯一的工作表名称（避免重复）
Function 获取唯一工作表名称(targetWb As Workbook, baseName As String) As String
    Dim finalName As String
    Dim suffix As Long
    finalName = baseName
    
    ' Excel工作表名最多31个字符
    If Len(finalName) > 31 Then
        finalName = Left(finalName, 31)
    End If
    
    ' 确保名称不包含非法字符
    finalName = Replace(finalName, ":", "_")
    finalName = Replace(finalName, "\", "_")
    finalName = Replace(finalName, "/", "_")
    finalName = Replace(finalName, "?", "_")
    finalName = Replace(finalName, "*", "_")
    finalName = Replace(finalName, "[", "_")
    finalName = Replace(finalName, "]", "_")
    
    ' 检查名称是否已存在，如果存在则添加数字后缀
    suffix = 1
    Dim originalName As String
    originalName = finalName
    
    Do While 工作表名称已存在(targetWb, finalName)
        suffix = suffix + 1
        If Len(originalName) > 28 Then
            finalName = Left(originalName, 28) & "_" & suffix
        Else
            finalName = originalName & "_" & suffix
        End If
        
        If suffix > 100 Then
            ' 避免无限循环
            finalName = originalName & "_" & Format(Now, "hhmmss")
            Exit Do
        End If
    Loop
    
    获取唯一工作表名称 = finalName
End Function

' 检查工作表名称是否已存在
Function 工作表名称已存在(wb As Workbook, sheetName As String) As Boolean
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = wb.Worksheets(sheetName)
    工作表名称已存在 = Not (ws Is Nothing)
    On Error GoTo 0
End Function

' 检查工作表是否包含所有关键词
Function 工作表包含所有关键词(ws As Worksheet, keywords As Variant) As Boolean
    Dim i As Integer
    Dim wsName As String
    
    wsName = ws.Name
    
    工作表包含所有关键词 = True
    
    For i = 1 To 5
        If Len(Trim(keywords(i))) > 0 Then
            If InStr(1, wsName, keywords(i), vbTextCompare) = 0 Then
                工作表包含所有关键词 = False
                Exit Function
            End If
        End If
    Next i
End Function

' 整表提取
Sub 整表提取(sourceWs As Worksheet, targetWs As Worksheet)
    Dim lastRow As Long, lastCol As Long
    
    ' 获取源工作表的实际数据范围
    lastRow = 获取最后有数据的行(sourceWs)
    lastCol = 获取最后有数据的列(sourceWs)
    
    If lastRow > 0 And lastCol > 0 Then
        ' 复制整个数据范围
        sourceWs.Range(sourceWs.Cells(1, 1), sourceWs.Cells(lastRow, lastCol)).Copy
        targetWs.Cells(1, 1).PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
    End If
End Sub

' 按行列提取（支持列的离散/范围语法：如 "B,BO,BQ,BP" 或 "B,BO:BP"）
Sub ExtractPartialData(sourceWs As Worksheet, targetWs As Worksheet, rowsStr As String, colsStr As String)
    Dim rowIndexes As Collection
    Dim colIndexes As Collection
    Dim sourceRow As Variant, sourceCol As Variant
    Dim targetRow As Long, targetCol As Long
    Dim lastRow As Long, lastCol As Long

    lastRow = GetLastUsedRow(sourceWs)
    lastCol = GetLastUsedCol(sourceWs)
    If lastRow <= 0 Or lastCol <= 0 Then Exit Sub

    Set rowIndexes = ParseRowIndexes(rowsStr, lastRow)
    Set colIndexes = ParseColIndexes(colsStr, lastCol)

    If rowIndexes.Count = 0 Or colIndexes.Count = 0 Then Exit Sub

    ' [优化B] 读入二维数组，整块写出，消除逐格剪贴板操作
    Dim nR As Long, nC As Long, ri As Long, ci As Long
    Dim dataArr() As Variant
    nR = rowIndexes.Count
    nC = colIndexes.Count
    If nR = 0 Or nC = 0 Then Exit Sub
    ReDim dataArr(1 To nR, 1 To nC)
    ri = 0
    For Each sourceRow In rowIndexes
        ri = ri + 1
        ci = 0
        For Each sourceCol In colIndexes
            ci = ci + 1
            dataArr(ri, ci) = sourceWs.Cells(CLng(sourceRow), CLng(sourceCol)).Value
        Next sourceCol
    Next sourceRow
    targetWs.Cells(1, 1).Resize(nR, nC).Value = dataArr
End Sub

Private Function ParseRowIndexes(rowsStr As String, maxRow As Long) As Collection
    Dim result As New Collection
    Dim norm As String
    Dim segs As Variant
    Dim seg As Variant
    Dim i As Long

    norm = NormalizeRangeSpec(rowsStr)
    If Len(norm) = 0 Then
        For i = 1 To maxRow
            result.Add i
        Next i
        Set ParseRowIndexes = result
        Exit Function
    End If

    segs = Split(norm, ",")
    For Each seg In segs
        AddIndexSegment CStr(seg), maxRow, result, True
    Next seg

    If result.Count = 0 Then
        For i = 1 To maxRow
            result.Add i
        Next i
    End If

    Set ParseRowIndexes = result
End Function

Private Function ParseColIndexes(colsStr As String, maxCol As Long) As Collection
    Dim result As New Collection
    Dim norm As String
    Dim segs As Variant
    Dim seg As Variant
    Dim i As Long

    norm = NormalizeRangeSpec(colsStr)
    If Len(norm) = 0 Then
        For i = 1 To maxCol
            result.Add i
        Next i
        Set ParseColIndexes = result
        Exit Function
    End If

    segs = Split(norm, ",")
    For Each seg In segs
        AddIndexSegment CStr(seg), maxCol, result, False
    Next seg

    If result.Count = 0 Then
        For i = 1 To maxCol
            result.Add i
        Next i
    End If

    Set ParseColIndexes = result
End Function

Private Sub AddIndexSegment(seg As String, maxValue As Long, ByRef result As Collection, isRow As Boolean)
    Dim parts As Variant
    Dim startNo As Long, endNo As Long, n As Long, tmpNo As Long

    seg = Trim$(seg)
    If Len(seg) = 0 Then Exit Sub

    If InStr(1, seg, ":", vbTextCompare) > 0 Then
        parts = Split(seg, ":")
        If UBound(parts) >= 1 Then
            startNo = TokenToIndex(CStr(parts(0)), isRow)
            endNo = TokenToIndex(CStr(parts(1)), isRow)
            If startNo > 0 And endNo > 0 Then
                If startNo > endNo Then
                    tmpNo = startNo
                    startNo = endNo
                    endNo = tmpNo
                End If
                If startNo < 1 Then startNo = 1
                If endNo > maxValue Then endNo = maxValue
                For n = startNo To endNo
                    result.Add n
                Next n
            End If
        End If
    Else
        startNo = TokenToIndex(seg, isRow)
        If startNo >= 1 And startNo <= maxValue Then
            result.Add startNo
        End If
    End If
End Sub

Private Function TokenToIndex(token As String, isRow As Boolean) As Long
    token = Trim$(token)
    If Len(token) = 0 Then
        TokenToIndex = 0
        Exit Function
    End If

    If IsNumeric(token) Then
        TokenToIndex = CLng(token)
    ElseIf isRow Then
        TokenToIndex = 0
    Else
        TokenToIndex = ColumnLetterToNumber(token)
    End If
End Function

Private Function ColumnLetterToNumber(colLetter As String) As Long
    Dim i As Long
    Dim result As Long
    Dim ch As String

    colLetter = UCase$(Trim$(colLetter))
    result = 0

    For i = 1 To Len(colLetter)
        ch = Mid$(colLetter, i, 1)
        If ch < "A" Or ch > "Z" Then
            ColumnLetterToNumber = 0
            Exit Function
        End If
        result = result * 26 + (Asc(ch) - 64)
    Next i

    ColumnLetterToNumber = result
End Function

Private Function NormalizeRangeSpec(text As String) As String
    Dim s As String
    s = CStr(text)
    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, ChrW(&HFF0C), ",")
    s = Replace(s, ChrW(&H3001), ",")
    s = Replace(s, ChrW(&HFF1B), ",")
    s = Replace(s, ";", ",")
    s = Replace(s, ChrW(&HFF1A), ":")
    NormalizeRangeSpec = s
End Function

Private Function GetLastUsedRow(ws As Worksheet) As Long
    On Error GoTo Fail
    Dim rng As Range
    Set rng = ws.Cells.Find("*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If rng Is Nothing Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = rng.Row
    End If
    Exit Function
Fail:
    GetLastUsedRow = 0
End Function

Private Function GetLastUsedCol(ws As Worksheet) As Long
    On Error GoTo Fail
    Dim rng As Range
    Set rng = ws.Cells.Find("*", LookIn:=xlValues, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If rng Is Nothing Then
        GetLastUsedCol = 0
    Else
        GetLastUsedCol = rng.Column
    End If
    Exit Function
Fail:
    GetLastUsedCol = 0
End Function

' 删除空白行列，使表格紧凑
Sub 删除空白行列(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim row As Long, col As Long
    Dim hasData As Boolean
    
    ' 获取实际数据范围
    lastRow = 获取最后有数据的行(ws)
    lastCol = 获取最后有数据的列(ws)
    
    If lastRow = 0 Or lastCol = 0 Then Exit Sub
    
    ' [优化C] 先收集空行，最后 Union 一次删除
    Dim delRows As Range
    For row = lastRow To 1 Step -1
        hasData = False
        For col = 1 To lastCol
            If Len(Trim(CStr(ws.Cells(row, col).Value))) > 0 Then
                hasData = True
                Exit For
            End If
        Next col
        If Not hasData Then
            If delRows Is Nothing Then
                Set delRows = ws.Rows(row)
            Else
                Set delRows = Union(delRows, ws.Rows(row))
            End If
        End If
    Next row
    If Not delRows Is Nothing Then delRows.Delete Shift:=xlUp
    Set delRows = Nothing
    
    ' 重新获取数据范围（因为删除了行）
    lastRow = 获取最后有数据的行(ws)
    lastCol = 获取最后有数据的列(ws)
    
    If lastRow = 0 Or lastCol = 0 Then Exit Sub
    
    ' [优化C] 先收集空列，最后 Union 一次删除
    Dim delCols As Range
    For col = lastCol To 1 Step -1
        hasData = False
        For row = 1 To lastRow
            If Len(Trim(CStr(ws.Cells(row, col).Value))) > 0 Then
                hasData = True
                Exit For
            End If
        Next row
        If Not hasData Then
            If delCols Is Nothing Then
                Set delCols = ws.Columns(col)
            Else
                Set delCols = Union(delCols, ws.Columns(col))
            End If
        End If
    Next col
    If Not delCols Is Nothing Then delCols.Delete Shift:=xlToLeft
    Set delCols = Nothing
End Sub

' 获取最后有数据的行
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

' 获取最后有数据的列
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

' 列字母转数字（如"A"转1，"AA"转27）
Function 列字母转数字(colLetter As String) As Long
    Dim i As Long
    Dim result As Long
    
    colLetter = UCase(Trim(colLetter))
    result = 0
    
    For i = 1 To Len(colLetter)
        result = result * 26 + (Asc(Mid(colLetter, i, 1)) - 64)
    Next i
    
    列字母转数字 = result
End Function

' 创建配置表示例
Sub 创建配置表示例()
    Dim ws As Worksheet
    On Error Resume Next
    ThisWorkbook.Worksheets("工作表提取").Delete
    On Error GoTo 0
    
    Set ws = ThisWorkbook.Worksheets.Add
    ws.Name = "工作表提取"
    
    ' 设置表头
    ws.Cells(1, 1) = "表A-币种"
    ws.Cells(1, 2) = "表A-地区"
    ws.Cells(1, 3) = "表A-机构"
    ws.Cells(1, 4) = "表A-类型"
    ws.Cells(1, 5) = "表A-名称"
    ws.Cells(1, 6) = "是否禁用"
    ws.Cells(1, 7) = "是否整表提取"
    ws.Cells(1, 8) = "提取行"
    ws.Cells(1, 9) = "提取列"
    ws.Cells(1, 10) = "输出文件名"
    
    ' 示例数据
    ws.Cells(2, 1) = "人民币"
    ws.Cells(2, 2) = "惠州市"
    ws.Cells(2, 3) = "全金融机构"
    ws.Cells(2, 4) = "信贷"
    ws.Cells(2, 5) = "收支表"
    ws.Cells(2, 6) = "否"
    ws.Cells(2, 7) = "是"
    ws.Cells(2, 8) = ""
    ws.Cells(2, 9) = ""
    ws.Cells(2, 10) = "惠州市信贷数据"
    
    ws.Cells(3, 1) = "外币"
    ws.Cells(3, 2) = "深圳市"
    ws.Cells(3, 3) = "商业银行"
    ws.Cells(3, 4) = "存款"
    ws.Cells(3, 5) = "明细表"
    ws.Cells(3, 6) = "否"
    ws.Cells(3, 7) = "否"
    ws.Cells(3, 8) = "2,3,4,5"
    ws.Cells(3, 9) = "A,B,C,D"
    ws.Cells(3, 10) = "深圳市存款数据"
    
    ws.Cells(4, 1) = "本外币"
    ws.Cells(4, 2) = "广州市"
    ws.Cells(4, 3) = "国有银行"
    ws.Cells(4, 4) = "贷款"
    ws.Cells(4, 5) = "统计表"
    ws.Cells(4, 6) = "是"
    ws.Cells(4, 7) = "否"
    ws.Cells(4, 8) = ""
    ws.Cells(4, 9) = ""
    ws.Cells(4, 10) = "广州市贷款数据"
    
    ' 设置格式
    With ws.Rows(1)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(200, 220, 240)
    End With
    
    ws.Columns("A:J").AutoFit
    
    MsgBox "配置表示例已创建在当前工作簿中！" & vbCrLf & _
           "工作表名称: 工作表提取" & vbCrLf & _
           "请根据实际需求修改配置数据。", vbInformation
End Sub

