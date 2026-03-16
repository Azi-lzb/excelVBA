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
        
        ' 创建字典来管理输出文件
        Set outputDict = CreateObject("Scripting.Dictionary")
        
        ' 处理每个配置行
        For i = 2 To lastRow
            Dim shouldSkip As Boolean
            Dim shouldFullExtract As Boolean
            Dim sheetInfo(1 To 5) As String
            Dim outputFileName As String
            Dim extractRows As String
            Dim extractCols As String
            
            ' 读取配置
            With configSheet
                ' 是否禁用 - F列
                shouldSkip = (.Cells(i, "F").value = "是" Or .Cells(i, "F").value = "1" Or .Cells(i, "F").value = True)
                
                ' 如果禁用则跳过此配置
                If Not shouldSkip Then
                    ' 前五列确定sheet名
                    sheetInfo(1) = CStr(.Cells(i, "A").value) ' 币种
                    sheetInfo(2) = CStr(.Cells(i, "B").value) ' 地区
                    sheetInfo(3) = CStr(.Cells(i, "C").value) ' 机构
                    sheetInfo(4) = CStr(.Cells(i, "D").value) ' 类型
                    sheetInfo(5) = CStr(.Cells(i, "E").value) ' 名称
                    
                    ' 是否整表提取 - G列
                    shouldFullExtract = (.Cells(i, "G").value = "是" Or .Cells(i, "G").value = "1" Or .Cells(i, "G").value = True)
                    
                    ' 提取行列配置
                    extractRows = CStr(.Cells(i, "H").value) ' 提取行
                    extractCols = CStr(.Cells(i, "I").value) ' 提取列
                    
                    ' 输出文件名 - J列
                    outputFileName = CStr(.Cells(i, "J").value)
                    If outputFileName = "" Then
                        outputFileName = "提取结果_" & Format(Now, "yyyymmdd_hhmmss")
                    End If
                    
                    ' 在所有源文件中查找匹配的sheet
                    For Each fileItem In fd.SelectedItems
                        Dim sourceWb As Workbook
                        Dim sourceWs As Worksheet
                        Dim targetWs As Worksheet
                        Dim targetWb As Workbook
                        
                        On Error Resume Next
                        Set sourceWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True)
                        If Err.Number <> 0 Then
                            Debug.Print "无法打开文件: " & fileItem
                            On Error GoTo 0
                        Else
                            On Error GoTo 0
                            
                            ' 查找所有匹配的工作表（修复：不只匹配第一个）
                            Dim matchedSheets As Collection
                            Set matchedSheets = 查找所有匹配工作表(sourceWb, sheetInfo)
                            
                            ' 处理每一个匹配的工作表
                            If matchedSheets.count > 0 Then
                                Dim wsItem As Variant
                                
                                For Each wsItem In matchedSheets
                                    Set sourceWs = wsItem
                                    
                                    ' 获取或创建目标工作簿
                                    If Not outputDict.Exists(outputFileName) Then
                                        Set targetWb = Workbooks.Add
                                        targetWb.SaveAs sourceWb.path & "\" & outputFileName & ".xlsx"
                                        outputDict.Add outputFileName, targetWb
                                    Else
                                        Set targetWb = outputDict(outputFileName)
                                    End If
                                    
                                    ' 在目标工作簿中创建新工作表（使用源工作表的原名）
                                    On Error Resume Next
                                    Set targetWs = targetWb.Worksheets.Add(After:=targetWb.Worksheets(targetWb.Worksheets.count))
                                    targetWs.Name = 获取唯一工作表名称(targetWb, sourceWs.Name)
                                    On Error GoTo 0
                                    
                                    ' 提取数据
                                    If shouldFullExtract Then
                                        ' 整表提取
                                        整表提取 sourceWs, targetWs
                                    Else
                                        ' 按行列提取
                                        If extractRows <> "" And extractCols <> "" Then
                                            按行列提取 sourceWs, targetWs, extractRows, extractCols
                                        Else
                                            ' 如果没有指定行列，则整表提取
                                            整表提取 sourceWs, targetWs
                                        End If
                                    End If
                                    
                                    ' 删除空白行列，使表格紧凑
                                    删除空白行列 targetWs
                                    
                                    processedCount = processedCount + 1
                                    Debug.Print "已提取: " & sourceWb.Name & " - " & sourceWs.Name & " -> " & outputFileName
                                Next wsItem
                            End If
                            
                            sourceWb.Close SaveChanges:=False
                        End If
                    Next fileItem
                End If
            End With
        Next i
        
        ' 保存所有输出文件
        Dim outputKey As Variant
        For Each outputKey In outputDict.keys
            outputDict(outputKey).Save
            RunLog_WriteRow "1.6 提取工作表", "输出文件", CStr(outputKey) & ".xlsx", "", "", "成功", "已保存", ""
            outputDict(outputKey).Close
        Next outputKey
        
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
    
    ' 构建目标工作表名称（用于精确匹配）
    targetName = 构建工作表名称(sheetInfo)
    
    ' 先尝试精确匹配
    On Error Resume Next
    Set ws = wb.Worksheets(targetName)
    If Not ws Is Nothing Then
        matchedSheets.Add ws
    End If
    On Error GoTo 0
    
    ' 然后进行模糊匹配
    For Each ws In wb.Worksheets
        ' 检查是否包含所有关键词，并且不是已经添加的工作表
        If 工作表包含所有关键词(ws, sheetInfo) Then
            ' 检查是否已经添加过这个工作表
            Dim alreadyAdded As Boolean
            alreadyAdded = False
            
            Dim existingWs As Variant
            For Each existingWs In matchedSheets
                If existingWs.Name = ws.Name Then
                    alreadyAdded = True
                    Exit For
                End If
            Next existingWs
            
            If Not alreadyAdded Then
                matchedSheets.Add ws
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
Sub 按行列提取(sourceWs As Worksheet, targetWs As Worksheet, rowsStr As String, colsStr As String)
    Dim rowArray As Variant
    Dim colArray As Variant
    Dim i As Long, j As Long
    Dim targetRow As Long, targetCol As Long
    Dim seg As String
    Dim parts As Variant
    Dim sourceRow As Long
    Dim sourceCol As Long
    Dim startCol As Long, endCol As Long, tmpCol As Long
    
    ' 解析行字符串（如"2,3,4"）
    rowArray = Split(Replace(rowsStr, " ", ""), ",")
    
    ' 解析列字符串（如"A,B,C"，也支持 B,BO,BQ,BP 或 B,BO:BP 这样的组合）
    colArray = Split(Replace(colsStr, " ", ""), ",")
    
    targetRow = 1
    For i = 0 To UBound(rowArray)
        If IsNumeric(rowArray(i)) Then
            sourceRow = CLng(rowArray(i))
            
            targetCol = 1
            For j = 0 To UBound(colArray)
                seg = Trim(CStr(colArray(j)))
                If Len(seg) > 0 Then
                    ' 支持单列：B；范围：BO:BP；以及组合：B,BO,BQ,BP / B,BO:BP
                    If InStr(1, seg, ":", vbTextCompare) > 0 Then
                        parts = Split(seg, ":")
                        If UBound(parts) >= 1 Then
                            startCol = 列字母转数字(CStr(parts(0)))
                            endCol = 列字母转数字(CStr(parts(1)))
                            If startCol > 0 And endCol > 0 Then
                                If startCol > endCol Then
                                    tmpCol = startCol
                                    startCol = endCol
                                    endCol = tmpCol
                                End If
                                For sourceCol = startCol To endCol
                                    sourceWs.Cells(sourceRow, sourceCol).Copy
                                    targetWs.Cells(targetRow, targetCol).PasteSpecial Paste:=xlPasteAll
                                    targetCol = targetCol + 1
                                Next sourceCol
                            End If
                        End If
                    Else
                        sourceCol = 列字母转数字(seg)
                        If sourceCol > 0 Then
                            sourceWs.Cells(sourceRow, sourceCol).Copy
                            targetWs.Cells(targetRow, targetCol).PasteSpecial Paste:=xlPasteAll
                            targetCol = targetCol + 1
                        End If
                    End If
                End If
            Next j
            
            targetRow = targetRow + 1
        End If
    Next i
    
    Application.CutCopyMode = False
End Sub

' 删除空白行列，使表格紧凑
Sub 删除空白行列(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long
    Dim row As Long, col As Long
    Dim hasData As Boolean
    
    ' 获取实际数据范围
    lastRow = 获取最后有数据的行(ws)
    lastCol = 获取最后有数据的列(ws)
    
    If lastRow = 0 Or lastCol = 0 Then Exit Sub
    
    ' 删除空行
    For row = lastRow To 1 Step -1
        hasData = False
        For col = 1 To lastCol
            If Len(Trim(CStr(ws.Cells(row, col).value))) > 0 Then
                hasData = True
                Exit For
            End If
        Next col
        
        If Not hasData Then
            ws.Rows(row).Delete
        End If
    Next row
    
    ' 重新获取数据范围（因为删除了行）
    lastRow = 获取最后有数据的行(ws)
    lastCol = 获取最后有数据的列(ws)
    
    If lastRow = 0 Or lastCol = 0 Then Exit Sub
    
    ' 删除空列
    For col = lastCol To 1 Step -1
        hasData = False
        For row = 1 To lastRow
            If Len(Trim(CStr(ws.Cells(row, col).value))) > 0 Then
                hasData = True
                Exit For
            End If
        Next row
        
        If Not hasData Then
            ws.Columns(col).Delete
        End If
    Next col
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

