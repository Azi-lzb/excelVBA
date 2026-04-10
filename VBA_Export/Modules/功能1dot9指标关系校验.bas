Attribute VB_Name = "功能1dot9指标关系校验"
Option Explicit

Private Const LOG_KEY As String = "指标关系校验"
Private Const RULE_SHEET_NAME As String = "指标校验规则"
Private Const RESULT_SHEET_NAME As String = "指标校验结果"

Private Enum RuleCols
    rcEnabled = 1
    rcFormCode = 2
    rcSheetKeyword = 3
    rcCodeCol = 4
    rcValueCol = 5
    rcGroup1 = 6
    rcGroup2 = 7
    rcMainCode = 8
    rcRelation = 9
    rcCompareCodes = 10
    rcTolerance = 11
    rcMessage = 12
    rcRemark = 13
    rcExample = 14
End Enum

Private Enum ResultCols
    rsExecTime = 1
    rsTargetBook = 2
    rsFormCode = 3
    rsSheetName = 4
    rsGroup1 = 5
    rsGroup2 = 6
    rsMainCode = 7
    rsRelation = 8
    rsCompareCodes = 9
    rsStatus = 10
    rsDiff = 11
    rsDetail = 12
End Enum

Public Sub 初始化指标校验配置()
    Dim wsRule As Worksheet

    Set wsRule = EnsureRuleSheet()
    InitRuleSheetHeader wsRule

    MsgBox "指标校验配置表头已更新。", vbInformation
End Sub

Public Sub 执行指标关系校验()
    Dim t0 As Double
    Dim targetPath As String
    Dim targetWb As Workbook
    Dim wsRule As Worksheet
    Dim wsResult As Worksheet
    Dim resultWb As Workbook
    Dim cacheSheet As Object
    Dim cacheMap As Object
    Dim lastRow As Long
    Dim rowNo As Long
    Dim resultRow As Long
    Dim passCount As Long
    Dim failCount As Long
    Dim skipCount As Long

    t0 = Timer
    RunLog_WriteRow LOG_KEY, "开始", "", "", "", "", "开始", ""

    Set wsRule = EnsureRuleSheet()
    InitRuleSheetHeader wsRule

    targetPath = PickTargetWorkbookPath("请选择要执行指标校验的工作簿")
    If targetPath = "" Then
        RunLog_WriteRow LOG_KEY, "结束", "", "", "", "", "已取消", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrHandler

    Set cacheSheet = CreateObject("Scripting.Dictionary")
    cacheSheet.CompareMode = vbTextCompare
    Set cacheMap = CreateObject("Scripting.Dictionary")
    cacheMap.CompareMode = vbTextCompare

    Set resultWb = CreateResultWorkbook(RESULT_SHEET_NAME, wsResult)
    InitResultHeader wsResult
    Set targetWb = Workbooks.Open(targetPath, ReadOnly:=True, UpdateLinks:=0)
    lastRow = GetLastUsedRow(wsRule)
    resultRow = 2

    For rowNo = 2 To lastRow
        If IsRuleRowEmpty(wsRule, rowNo) Then
            skipCount = skipCount + 1
        ElseIf Not IsEnabledValue(wsRule.Cells(rowNo, rcEnabled).Value) Then
            skipCount = skipCount + 1
        Else
            ProcessOneRule wsRule, rowNo, wsResult, resultRow, targetWb, cacheSheet, cacheMap, passCount, failCount, skipCount
        End If
    Next rowNo

    wsResult.Columns("A:L").AutoFit
    SafeCloseWorkbook targetWb

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    RunLog_WriteRow LOG_KEY, "完成", Dir$(targetPath), CStr(passCount), CStr(failCount), "完成", "跳过规则=" & skipCount, CStr(Round(Timer - t0, 2))
    MsgBox BuildSummaryMessage(targetPath, passCount, failCount, skipCount), IIf(failCount > 0, vbExclamation, vbInformation), "指标关系校验"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    SafeCloseWorkbook targetWb
    RunLog_WriteRow LOG_KEY, "完成", targetPath, "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "执行失败：" & Err.Number & " " & Err.Description, vbCritical, "指标关系校验"
End Sub

Private Sub ProcessOneRule(ByVal wsRule As Worksheet, ByVal rowNo As Long, ByVal wsResult As Worksheet, ByRef resultRow As Long, ByVal targetWb As Workbook, ByVal cacheSheet As Object, ByVal cacheMap As Object, ByRef passCount As Long, ByRef failCount As Long, ByRef skipCount As Long)
    Dim formCode As String
    Dim sheetKeyword As String
    Dim codeCol As Long
    Dim valueCol As Long
    Dim group1 As String
    Dim group2 As String
    Dim mainCode As String
    Dim relationType As String
    Dim compareCodes As String
    Dim tolerance As Double
    Dim failMessage As String
    Dim targetWs As Worksheet
    Dim targetSheets As Collection
    Dim valueMap As Object
    Dim diffValue As Double
    Dim detailText As String
    Dim statusText As String
    Dim ok As Boolean
    Dim item As Variant
    Dim currentValueCol As Long

    On Error GoTo RuleErr

    formCode = Trim$(CStr(wsRule.Cells(rowNo, rcFormCode).Value))
    sheetKeyword = Trim$(CStr(wsRule.Cells(rowNo, rcSheetKeyword).Value))
    codeCol = ParseColumnSpec(wsRule.Cells(rowNo, rcCodeCol).Value, 2)
    valueCol = ParseColumnSpec(wsRule.Cells(rowNo, rcValueCol).Value, 0)
    group1 = Trim$(CStr(wsRule.Cells(rowNo, rcGroup1).Value))
    group2 = Trim$(CStr(wsRule.Cells(rowNo, rcGroup2).Value))
    mainCode = NormalizeCode(CStr(wsRule.Cells(rowNo, rcMainCode).Value))
    relationType = UCase$(Trim$(CStr(wsRule.Cells(rowNo, rcRelation).Value)))
    compareCodes = Trim$(CStr(wsRule.Cells(rowNo, rcCompareCodes).Value))
    tolerance = ParseTolerance(wsRule.Cells(rowNo, rcTolerance).Value)
    failMessage = Trim$(CStr(wsRule.Cells(rowNo, rcMessage).Value))

    If formCode = "" Then Err.Raise vbObjectError + 3001, , "规则行缺少表单编码"
    If mainCode = "" Then Err.Raise vbObjectError + 3002, , "规则行缺少主指标"
    If relationType = "" Then Err.Raise vbObjectError + 3003, , "规则行缺少关系类型"
    If compareCodes = "" Then Err.Raise vbObjectError + 3004, , "规则行缺少比较指标"
    Set targetSheets = ResolveSourceSheetsCached(targetWb, formCode, sheetKeyword, codeCol, cacheSheet)
    If targetSheets Is Nothing Or targetSheets.Count = 0 Then
        skipCount = skipCount + 1
        RunLog_WriteRow LOG_KEY, "跳过规则", formCode & "|" & mainCode, "", "", "跳过", "未找到匹配工作表", ""
        Exit Sub
    End If

    For Each item In targetSheets
        Set targetWs = item
        currentValueCol = valueCol
        If currentValueCol <= 0 Then
            currentValueCol = DetectValueColumn(targetWs, codeCol)
        End If
        If currentValueCol <= 0 Then Err.Raise vbObjectError + 3005, , "规则行缺少有效的取值列，且自动探测失败"

        Set valueMap = BuildValueMapCached(targetWs, codeCol, currentValueCol, cacheMap)
        ok = EvaluateRule(valueMap, mainCode, relationType, compareCodes, tolerance, diffValue, detailText)

        If ok Then
            statusText = "通过"
            passCount = passCount + 1
        Else
            statusText = "不通过"
            If failMessage <> "" Then
                detailText = failMessage & "；" & detailText
            End If
            failCount = failCount + 1
        End If

        WriteResultRow wsResult, resultRow, targetWb.Name, formCode, targetWs.Name, group1, group2, mainCode, relationType, compareCodes, statusText, diffValue, detailText
        resultRow = resultRow + 1
    Next item
    Exit Sub

RuleErr:
    failCount = failCount + 1
    WriteResultRow wsResult, resultRow, targetWb.Name, formCode, "", group1, group2, mainCode, relationType, compareCodes, "失败", 0, "规则第 " & rowNo & " 行：" & Err.Description
    resultRow = resultRow + 1
End Sub

Private Function ResolveSourceSheetsCached(ByVal wb As Workbook, ByVal formCode As String, ByVal sheetKeyword As String, ByVal codeCol As Long, ByVal cacheSheet As Object) As Collection
    Dim cacheKey As String

    cacheKey = wb.Name & "|" & formCode & "|" & sheetKeyword & "|" & CStr(codeCol)
    If cacheSheet.Exists(cacheKey) Then
        Set ResolveSourceSheetsCached = cacheSheet(cacheKey)
        Exit Function
    End If

    Set ResolveSourceSheetsCached = ResolveSourceSheets(wb, formCode, sheetKeyword, codeCol)
    Set cacheSheet(cacheKey) = ResolveSourceSheetsCached
End Function

Private Function ResolveSourceSheets(ByVal wb As Workbook, ByVal formCode As String, ByVal sheetKeyword As String, ByVal codeCol As Long) As Collection
    Dim ws As Worksheet
    Dim result As Collection
    Dim keys As Object

    Set result = New Collection
    Set keys = CreateObject("Scripting.Dictionary")
    keys.CompareMode = vbTextCompare

    If Trim$(sheetKeyword) <> "" Then
        For Each ws In wb.Worksheets
            If ContainsText(ws.Name, sheetKeyword) Then
                AddSheetUnique result, keys, ws
            End If
        Next ws
    End If

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, formCode, vbTextCompare) = 0 Or ContainsText(ws.Name, formCode) Then
            AddSheetUnique result, keys, ws
        End If
    Next ws

    For Each ws In wb.Worksheets
        If WorksheetContainsFormCode(ws, formCode, codeCol) Then
            AddSheetUnique result, keys, ws
        End If
    Next ws

    If result.Count > 0 Then
        Set ResolveSourceSheets = result
        Exit Function
    End If

    If wb.Worksheets.Count = 1 Then
        AddSheetUnique result, keys, wb.Worksheets(1)
        Set ResolveSourceSheets = result
        Exit Function
    End If

    Set ResolveSourceSheets = result
End Function

Private Function WorksheetContainsFormCode(ByVal ws As Worksheet, ByVal formCode As String, ByVal codeCol As Long) As Boolean
    Dim lastRow As Long
    Dim rowNo As Long

    lastRow = GetLastUsedRow(ws)
    If lastRow = 0 Then Exit Function

    For rowNo = 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(rowNo, 1).Value)), formCode, vbTextCompare) = 0 Then
            WorksheetContainsFormCode = True
            Exit Function
        End If
    Next rowNo

    If codeCol > 0 Then
        For rowNo = 1 To lastRow
            If StrComp(Trim$(CStr(ws.Cells(rowNo, codeCol).Value)), formCode, vbTextCompare) = 0 Then
                WorksheetContainsFormCode = True
                Exit Function
            End If
        Next rowNo
    End If
End Function

Private Sub AddSheetUnique(ByRef sheets As Collection, ByVal keys As Object, ByVal ws As Worksheet)
    Dim oneKey As String

    oneKey = ws.Parent.Name & "|" & ws.Name
    If keys.Exists(oneKey) Then Exit Sub
    keys(oneKey) = True
    sheets.Add ws
End Sub

Private Function BuildValueMapCached(ByVal ws As Worksheet, ByVal codeCol As Long, ByVal valueCol As Long, ByVal cacheMap As Object) As Object
    Dim cacheKey As String

    cacheKey = ws.Parent.Name & "|" & ws.Name & "|" & CStr(codeCol) & "|" & CStr(valueCol)
    If cacheMap.Exists(cacheKey) Then
        Set BuildValueMapCached = cacheMap(cacheKey)
        Exit Function
    End If

    Set BuildValueMapCached = BuildIndicatorValueMap(ws, codeCol, valueCol)
    Set cacheMap(cacheKey) = BuildValueMapCached
End Function

Private Function BuildIndicatorValueMap(ByVal ws As Worksheet, ByVal codeCol As Long, ByVal valueCol As Long) As Object
    Dim dict As Object
    Dim duplicateCodes As Collection
    Dim lastRow As Long
    Dim rowNo As Long
    Dim codeText As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Set duplicateCodes = New Collection
    lastRow = GetLastUsedRow(ws)

    For rowNo = 1 To lastRow
        codeText = NormalizeCode(CStr(ws.Cells(rowNo, codeCol).Value))
        If codeText <> "" Then
            If IsHeaderCode(codeText) Then
            ElseIf dict.Exists(codeText) Then
                duplicateCodes.Add codeText
            Else
                dict(codeText) = ParseCellNumber(ws.Cells(rowNo, valueCol).Value)
            End If
        End If
    Next rowNo

    If duplicateCodes.Count > 0 Then
        Err.Raise vbObjectError + 3021, , "存在重复指标编码：" & JoinCollection(duplicateCodes, "、")
    End If

    Set BuildIndicatorValueMap = dict
End Function

Private Function EvaluateRule(ByVal valueMap As Object, ByVal mainCode As String, ByVal relationType As String, ByVal compareCodesText As String, ByVal tolerance As Double, ByRef diffValue As Double, ByRef detailText As String) As Boolean
    Dim compareCodes As Collection
    Dim compareValues As Collection
    Dim oneCode As Variant
    Dim mainValue As Double
    Dim calcValue As Double
    Dim sumValue As Double
    Dim maxDiff As Double
    Dim idx As Long
    Dim currentValue As Double

    Set compareCodes = SplitCodes(compareCodesText)
    If compareCodes.Count = 0 Then Err.Raise vbObjectError + 3031, , "比较指标为空"
    If Not valueMap.Exists(mainCode) Then Err.Raise vbObjectError + 3032, , "主指标未找到：" & mainCode

    mainValue = CDbl(valueMap(mainCode))
    Set compareValues = New Collection

    For Each oneCode In compareCodes
        If Not valueMap.Exists(CStr(oneCode)) Then
            Err.Raise vbObjectError + 3033, , "比较指标未找到：" & CStr(oneCode)
        End If
        compareValues.Add CDbl(valueMap(CStr(oneCode)))
    Next oneCode

    Select Case relationType
        Case "ALL_EQUAL"
            maxDiff = 0
            For idx = 1 To compareValues.Count
                currentValue = CDbl(compareValues(idx))
                If Abs(mainValue - currentValue) > maxDiff Then maxDiff = Abs(mainValue - currentValue)
            Next idx
            diffValue = maxDiff
            EvaluateRule = (maxDiff <= tolerance)
            detailText = "主指标=" & FormatNumberText(mainValue) & "；比较值=" & JoinNumberCollection(compareValues)
        Case "EQUAL"
            diffValue = mainValue - CDbl(compareValues(1))
            EvaluateRule = (Abs(diffValue) <= tolerance)
            detailText = "主指标=" & FormatNumberText(mainValue) & "；比较值=" & FormatNumberText(CDbl(compareValues(1)))
        Case "SUM"
            sumValue = 0
            For idx = 1 To compareValues.Count
                sumValue = sumValue + CDbl(compareValues(idx))
            Next idx
            diffValue = mainValue - sumValue
            EvaluateRule = (Abs(diffValue) <= tolerance)
            detailText = "主指标=" & FormatNumberText(mainValue) & "；比较项合计=" & FormatNumberText(sumValue)
        Case "DIFF"
            calcValue = CDbl(compareValues(1))
            If compareValues.Count >= 2 Then
                For idx = 2 To compareValues.Count
                    calcValue = calcValue - CDbl(compareValues(idx))
                Next idx
            End If
            diffValue = mainValue - calcValue
            EvaluateRule = (Abs(diffValue) <= tolerance)
            detailText = "主指标=" & FormatNumberText(mainValue) & "；计算值=" & FormatNumberText(calcValue)
        Case "GTE_SUM"
            sumValue = 0
            For idx = 1 To compareValues.Count
                sumValue = sumValue + CDbl(compareValues(idx))
            Next idx
            diffValue = mainValue - sumValue
            EvaluateRule = (diffValue >= -tolerance)
            detailText = "主指标=" & FormatNumberText(mainValue) & "；比较项合计=" & FormatNumberText(sumValue) & "；要求主指标大于等于比较项合计"
        Case Else
            Err.Raise vbObjectError + 3034, , "暂不支持的关系类型：" & relationType
    End Select

    detailText = detailText & "；差额=" & FormatNumberText(diffValue) & "；容差=" & FormatNumberText(tolerance)
End Function

Private Sub WriteResultRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal bookName As String, ByVal formCode As String, ByVal sheetName As String, ByVal group1 As String, ByVal group2 As String, ByVal mainCode As String, ByVal relationType As String, ByVal compareCodes As String, ByVal statusText As String, ByVal diffValue As Double, ByVal detailText As String)
    ws.Cells(rowNo, rsExecTime).Value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(rowNo, rsTargetBook).Value = bookName
    ws.Cells(rowNo, rsFormCode).Value = formCode
    ws.Cells(rowNo, rsSheetName).Value = sheetName
    ws.Cells(rowNo, rsGroup1).Value = group1
    ws.Cells(rowNo, rsGroup2).Value = group2
    ws.Cells(rowNo, rsMainCode).Value = mainCode
    ws.Cells(rowNo, rsRelation).Value = relationType
    ws.Cells(rowNo, rsCompareCodes).Value = compareCodes
    ws.Cells(rowNo, rsStatus).Value = statusText
    ws.Cells(rowNo, rsDiff).Value = diffValue
    ws.Cells(rowNo, rsDetail).Value = detailText
End Sub

Private Function EnsureRuleSheet() As Worksheet
    On Error Resume Next
    Set EnsureRuleSheet = ThisWorkbook.Worksheets(RULE_SHEET_NAME)
    On Error GoTo 0

    If EnsureRuleSheet Is Nothing Then
        Set EnsureRuleSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureRuleSheet.Name = RULE_SHEET_NAME
    End If
End Function

Private Sub InitRuleSheetHeader(ByVal ws As Worksheet)
    ws.Cells(1, rcEnabled).Value = "是否启用"
    ws.Cells(1, rcFormCode).Value = "表单编码"
    ws.Cells(1, rcSheetKeyword).Value = "工作表名关键字"
    ws.Cells(1, rcCodeCol).Value = "指标编码列"
    ws.Cells(1, rcValueCol).Value = "取值列"
    ws.Cells(1, rcGroup1).Value = "一级分组"
    ws.Cells(1, rcGroup2).Value = "二级分组"
    ws.Cells(1, rcMainCode).Value = "主指标"
    ws.Cells(1, rcRelation).Value = "关系类型"
    ws.Cells(1, rcCompareCodes).Value = "比较指标"
    ws.Cells(1, rcTolerance).Value = "容差"
    ws.Cells(1, rcMessage).Value = "错误提示"
    ws.Cells(1, rcRemark).Value = "备注"
    ws.Cells(1, rcExample).Value = "示例表达式"
    ws.Rows(1).Font.Bold = True

    AddOrReplaceComment ws.Cells(1, rcEnabled), "程序依赖字段。Y 表示启用，N 表示跳过。"
    AddOrReplaceComment ws.Cells(1, rcFormCode), "程序依赖字段。用于定位目标表单，例如 A3327。"
    AddOrReplaceComment ws.Cells(1, rcSheetKeyword), "程序依赖字段。可为空。填写后优先按工作表名称包含关键字定位。"
    AddOrReplaceComment ws.Cells(1, rcCodeCol), "程序依赖字段。指标编码所在列，可填 B 或 2。"
    AddOrReplaceComment ws.Cells(1, rcValueCol), "程序依赖字段。实际取值所在列，可填 C、D、AA、列号，或填 AUTO 自动探测。"
    AddOrReplaceComment ws.Cells(1, rcGroup1), "仅用于人工分类，不参与计算。"
    AddOrReplaceComment ws.Cells(1, rcGroup2), "仅用于人工分类，不参与计算。"
    AddOrReplaceComment ws.Cells(1, rcMainCode), "程序依赖字段。主指标编码，必须与源表完全一致。"
    AddOrReplaceComment ws.Cells(1, rcRelation), "程序依赖字段。当前支持 ALL_EQUAL、EQUAL、SUM、DIFF、GTE_SUM。"
    AddOrReplaceComment ws.Cells(1, rcCompareCodes), "程序依赖字段。多个指标请用英文逗号分隔。"
    AddOrReplaceComment ws.Cells(1, rcTolerance), "程序依赖字段。允许误差，通常填 0。"
    AddOrReplaceComment ws.Cells(1, rcMessage), "程序依赖字段。校验失败时显示给用户的提示。"
    AddOrReplaceComment ws.Cells(1, rcRemark), "仅用于维护说明，不参与计算。"
    AddOrReplaceComment ws.Cells(1, rcExample), "仅用于人工查看，不参与计算。"
End Sub

Private Function CreateResultWorkbook(ByVal resultSheetName As String, ByRef resultWs As Worksheet) As Workbook
    Set CreateResultWorkbook = Workbooks.Add(xlWBATWorksheet)
    Set resultWs = CreateResultWorkbook.Worksheets(1)
    On Error Resume Next
    resultWs.Name = resultSheetName
    On Error GoTo 0
End Function

Private Sub InitResultHeader(ByVal ws As Worksheet)
    ws.Cells(1, rsExecTime).Value = "执行时间"
    ws.Cells(1, rsTargetBook).Value = "目标文件"
    ws.Cells(1, rsFormCode).Value = "表单编码"
    ws.Cells(1, rsSheetName).Value = "工作表名"
    ws.Cells(1, rsGroup1).Value = "一级分组"
    ws.Cells(1, rsGroup2).Value = "二级分组"
    ws.Cells(1, rsMainCode).Value = "主指标"
    ws.Cells(1, rsRelation).Value = "关系类型"
    ws.Cells(1, rsCompareCodes).Value = "比较指标"
    ws.Cells(1, rsStatus).Value = "校验结果"
    ws.Cells(1, rsDiff).Value = "差额"
    ws.Cells(1, rsDetail).Value = "明细说明"
    ws.Rows(1).Font.Bold = True
End Sub

Private Function PickTargetWorkbookPath(ByVal dialogTitle As String) As String
    Dim fd As FileDialog

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = dialogTitle
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"
        If .Show = -1 Then
            PickTargetWorkbookPath = CStr(.SelectedItems(1))
        Else
            PickTargetWorkbookPath = ""
        End If
    End With
End Function

Private Function SplitCodes(ByVal textValue As String) As Collection
    Dim result As New Collection
    Dim arr As Variant
    Dim i As Long
    Dim oneText As String

    textValue = Replace(textValue, "，", ",")
    textValue = Replace(textValue, "；", ",")
    textValue = Replace(textValue, ";", ",")
    arr = Split(textValue, ",")

    For i = LBound(arr) To UBound(arr)
        oneText = NormalizeCode(CStr(arr(i)))
        If oneText <> "" Then result.Add oneText
    Next i

    Set SplitCodes = result
End Function

Private Function JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String
    Dim i As Long

    For i = 1 To items.Count
        If JoinCollection <> "" Then JoinCollection = JoinCollection & delimiter
        JoinCollection = JoinCollection & CStr(items(i))
    Next i
End Function

Private Function JoinNumberCollection(ByVal items As Collection) As String
    Dim i As Long

    For i = 1 To items.Count
        If JoinNumberCollection <> "" Then JoinNumberCollection = JoinNumberCollection & "、"
        JoinNumberCollection = JoinNumberCollection & FormatNumberText(CDbl(items(i)))
    Next i
End Function

Private Function NormalizeCode(ByVal rawText As String) As String
    NormalizeCode = UCase$(Trim$(CStr(rawText)))
End Function

Private Function ContainsText(ByVal sourceText As String, ByVal keyword As String) As Boolean
    ContainsText = (InStr(1, CStr(sourceText), CStr(keyword), vbTextCompare) > 0)
End Function

Private Function ParseColumnSpec(ByVal rawValue As Variant, ByVal defaultValue As Long) As Long
    Dim textValue As String

    textValue = Trim$(CStr(rawValue))
    If textValue = "" Then
        ParseColumnSpec = defaultValue
        Exit Function
    End If

    If UCase$(textValue) = "AUTO" Then
        ParseColumnSpec = 0
        Exit Function
    End If

    If IsNumeric(textValue) Then
        ParseColumnSpec = CLng(textValue)
    Else
        ParseColumnSpec = ColLetterToNumber(UCase$(textValue))
    End If
End Function

Private Function ColLetterToNumber(ByVal colText As String) As Long
    Dim i As Long

    For i = 1 To Len(colText)
        If Mid$(colText, i, 1) < "A" Or Mid$(colText, i, 1) > "Z" Then
            Err.Raise vbObjectError + 3041, , "列标识无效：" & colText
        End If
        ColLetterToNumber = ColLetterToNumber * 26 + Asc(Mid$(colText, i, 1)) - 64
    Next i
End Function

Private Function ParseTolerance(ByVal rawValue As Variant) As Double
    If IsNumeric(rawValue) Then
        ParseTolerance = CDbl(rawValue)
    ElseIf Trim$(CStr(rawValue)) = "" Then
        ParseTolerance = 0
    Else
        ParseTolerance = CDbl(Val(CStr(rawValue)))
    End If
End Function

Private Function ParseCellNumber(ByVal rawValue As Variant) As Double
    Dim textValue As String

    If IsError(rawValue) Or IsEmpty(rawValue) Then Exit Function
    If IsNumeric(rawValue) Then
        ParseCellNumber = CDbl(rawValue)
        Exit Function
    End If

    textValue = Trim$(CStr(rawValue))
    If textValue = "" Or textValue = "-" Then Exit Function
    textValue = Replace(textValue, ",", "")
    If IsNumeric(textValue) Then ParseCellNumber = CDbl(textValue)
End Function

Private Function FormatNumberText(ByVal valueNumber As Double) As String
    FormatNumberText = Format(valueNumber, "0.00############")
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim rng As Range

    On Error Resume Next
    Set rng = ws.Cells.Find("*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If rng Is Nothing Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = rng.Row
    End If
End Function

Private Function GetLastUsedCol(ByVal ws As Worksheet) As Long
    Dim rng As Range

    On Error Resume Next
    Set rng = ws.Cells.Find("*", LookIn:=xlValues, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If rng Is Nothing Then
        GetLastUsedCol = 0
    Else
        GetLastUsedCol = rng.Column
    End If
End Function

Private Function DetectValueColumn(ByVal ws As Worksheet, ByVal codeCol As Long) As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim colNo As Long
    Dim rowNo As Long
    Dim numericCount As Long
    Dim textCount As Long
    Dim codeText As String
    Dim cellText As String

    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedCol(ws)
    If lastRow = 0 Or lastCol = 0 Then Exit Function

    For colNo = codeCol + 1 To lastCol
        numericCount = 0
        textCount = 0

        For rowNo = 2 To lastRow
            codeText = NormalizeCode(CStr(ws.Cells(rowNo, codeCol).Value))
            If codeText <> "" And Not IsHeaderCode(codeText) Then
                If IsNumeric(ws.Cells(rowNo, colNo).Value) Then
                    numericCount = numericCount + 1
                Else
                    cellText = Trim$(CStr(ws.Cells(rowNo, colNo).Value))
                    If cellText <> "" And cellText <> "-" Then
                        textCount = textCount + 1
                    End If
                End If
            End If
        Next rowNo

        If numericCount > 0 And textCount = 0 Then
            DetectValueColumn = colNo
            Exit Function
        End If
    Next colNo
End Function

Private Function IsEnabledValue(ByVal rawValue As Variant) As Boolean
    Dim textValue As String

    textValue = UCase$(Trim$(CStr(rawValue)))
    IsEnabledValue = (textValue = "Y" Or textValue = "1" Or textValue = "TRUE" Or textValue = "是")
End Function

Private Function IsRuleRowEmpty(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean
    IsRuleRowEmpty = (Trim$(CStr(ws.Cells(rowNo, rcFormCode).Value)) = "" And Trim$(CStr(ws.Cells(rowNo, rcMainCode).Value)) = "" And Trim$(CStr(ws.Cells(rowNo, rcCompareCodes).Value)) = "")
End Function

Private Function IsHeaderCode(ByVal codeText As String) As Boolean
    IsHeaderCode = (codeText = "指标编码" Or codeText = "INDEXCODE")
End Function

Private Sub AddOrReplaceComment(ByVal targetCell As Range, ByVal commentText As String)
    On Error Resume Next
    If Not targetCell.Comment Is Nothing Then targetCell.Comment.Delete
    targetCell.AddComment commentText
    On Error GoTo 0
End Sub

Private Function BuildSummaryMessage(ByVal targetPath As String, ByVal passCount As Long, ByVal failCount As Long, ByVal skipCount As Long) As String
    BuildSummaryMessage = "文件：" & Dir$(targetPath)
    BuildSummaryMessage = BuildSummaryMessage & vbCrLf & "通过规则：" & passCount
    BuildSummaryMessage = BuildSummaryMessage & vbCrLf & "失败规则：" & failCount
    BuildSummaryMessage = BuildSummaryMessage & vbCrLf & "跳过规则：" & skipCount
    BuildSummaryMessage = BuildSummaryMessage & vbCrLf & vbCrLf & "结果已写入新工作簿：" & RESULT_SHEET_NAME
End Function

Private Sub SafeCloseWorkbook(ByRef wb As Workbook)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If
    On Error GoTo 0
End Sub
