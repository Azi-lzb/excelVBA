Attribute VB_Name = "功能3dot14批量转时序"
Option Explicit

Private Const LOG_KEY As String = "3.8 批量转时序数据"
Private Const RULE_SHEET_NAME As String = "时序提取规则"
Private Const RESULT_SHEET_NAME As String = "时序提取结果"

Private Enum RuleCols
    rcEnabled = 1
    rcRuleName = 2
    rcBookKeywords = 3
    rcSheetKeywords = 4
    rcRowHeaderCols = 5
    rcColHeaderRows = 6
    rcRequiredColPaths = 7
    rcRequiredRowPaths = 8
    rcStartRow = 9
    rcEndRow = 10
    rcStartCol = 11
    rcEndCol = 12
    rcSkipKeywords = 13
    rcRemark = 14
End Enum

Private Enum ResultCols
    rsExecTime = 1
    rsSourceBook = 2
    rsSourceSheet = 3
    rsRuleName = 4
    rsFileModified = 5
    rsDataDate = 6
    rsDateSource = 7
    rsRowPath = 8
    rsColPath = 9
    rsValue = 10
    rsCellAddress = 11
End Enum

Public Sub 初始化时序提取配置()
    Dim wsRule As Worksheet

    Set wsRule = EnsureRuleSheet()
    InitRuleHeader wsRule

    MsgBox "时序提取配置表头已更新。", vbInformation
End Sub

Public Sub 批量转时序数据()
    Dim t0 As Double
    Dim wsRule As Worksheet
    Dim wsResult As Worksheet
    Dim resultWb As Workbook
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim targetWb As Workbook
    Dim lastRuleRow As Long
    Dim ruleRow As Long
    Dim resultRow As Long
    Dim hitBooks As Long
    Dim hitSheets As Long
    Dim outputRows As Long
    Dim skipRules As Long
    Dim skipBooks As Long
    Dim duplicateRows As Long
    Dim fileModified As String
    Dim dataDateText As String
    Dim dateSource As String
    Dim duplicateMap As Object

    t0 = Timer
    RunLog_WriteRow LOG_KEY, "开始", "", "", "", "", "开始", ""

    Set wsRule = EnsureRuleSheet()

    InitRuleHeader wsRule

    If MsgBox("请先确认目标工作簿已完成表格校验且无错误。" & vbCrLf & "是否继续执行时序提取？", vbQuestion + vbYesNo, "批量转时序数据") <> vbYes Then
        RunLog_WriteRow LOG_KEY, "结束", "", "", "", "", "用户取消，未确认表格校验", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "请选择要转时序的工作簿，可多选。"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"
        If .Show <> -1 Then
            RunLog_WriteRow LOG_KEY, "结束", "", "", "", "", "已取消", CStr(Round(Timer - t0, 2))
            Exit Sub
        End If
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrHandler

    Set resultWb = CreateResultWorkbook(RESULT_SHEET_NAME, wsResult)
    InitResultHeader wsResult
    Set duplicateMap = CreateObject("Scripting.Dictionary")
    duplicateMap.CompareMode = vbTextCompare
    resultRow = 2
    lastRuleRow = GetLastUsedRow(wsRule)

    For Each fileItem In fd.SelectedItems
        fileModified = ""
        On Error Resume Next
        fileModified = Format(FileDateTime(CStr(fileItem)), "yyyy/mm/dd hh:nn:ss")
        On Error GoTo ErrHandler

        Set targetWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True, UpdateLinks:=0)
        If Not ResolveWorkbookDataDate(targetWb, dataDateText, dateSource) Then
            skipBooks = skipBooks + 1
            RunLog_WriteRow LOG_KEY, "跳过工作簿", targetWb.Name, "", "", "跳过", "未识别数据日期", ""
            SafeCloseWorkbook targetWb
            GoTo NextBook
        End If

        hitBooks = hitBooks + 1

        For ruleRow = 2 To lastRuleRow
            If ShouldSkipRuleRow(wsRule, ruleRow) Then
                skipRules = skipRules + 1
            Else
                ProcessOneExtractRule wsRule, ruleRow, wsResult, resultRow, targetWb, fileModified, dataDateText, dateSource, duplicateMap, hitSheets, outputRows, skipRules, duplicateRows
            End If
        Next ruleRow

        SafeCloseWorkbook targetWb
NextBook:
    Next fileItem

    wsResult.Columns("A:K").AutoFit
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    RunLog_WriteRow LOG_KEY, "结束", CStr(hitBooks), CStr(hitSheets), CStr(outputRows), "完成", "跳过规则=" & skipRules & "，跳过工作簿=" & skipBooks & "，重复记录=" & duplicateRows, CStr(Round(Timer - t0, 2))
    MsgBox BuildSummaryMessage(hitBooks, hitSheets, outputRows, skipRules, skipBooks, duplicateRows), vbInformation, "批量转时序数据"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    SafeCloseWorkbook targetWb
    RunLog_WriteRow LOG_KEY, "结束", "", "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "执行失败：" & Err.Number & " " & Err.Description, vbCritical, "批量转时序数据"
End Sub

Private Sub ProcessOneExtractRule(ByVal wsRule As Worksheet, ByVal ruleRow As Long, ByVal wsResult As Worksheet, ByRef resultRow As Long, ByVal targetWb As Workbook, ByVal fileModified As String, ByVal dataDateText As String, ByVal dateSource As String, ByVal duplicateMap As Object, ByRef hitSheets As Long, ByRef outputRows As Long, ByRef skipRules As Long, ByRef duplicateRows As Long)
    Dim ruleName As String
    Dim bookKeywords As String
    Dim sheetKeywords As String
    Dim rowHeaderCols As Collection
    Dim colHeaderRows As Collection
    Dim dataStartRow As Long
    Dim dataEndRow As Long
    Dim dataStartCol As Long
    Dim dataEndCol As Long
    Dim requiredColPaths As String
    Dim requiredRowPaths As String
    Dim skipKeywords As String
    Dim matchedSheets As Collection
    Dim item As Variant

    ruleName = Trim$(CStr(wsRule.Cells(ruleRow, rcRuleName).Value))
    bookKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcBookKeywords).Value))
    sheetKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcSheetKeywords).Value))
    Set rowHeaderCols = ParseColumnCollection(CStr(wsRule.Cells(ruleRow, rcRowHeaderCols).Value))
    Set colHeaderRows = ParseNumberCollection(CStr(wsRule.Cells(ruleRow, rcColHeaderRows).Value))
    requiredColPaths = Trim$(CStr(wsRule.Cells(ruleRow, rcRequiredColPaths).Value))
    requiredRowPaths = Trim$(CStr(wsRule.Cells(ruleRow, rcRequiredRowPaths).Value))
    dataStartRow = ParseLongValue(wsRule.Cells(ruleRow, rcStartRow).Value, 0)
    dataEndRow = ParseLongValue(wsRule.Cells(ruleRow, rcEndRow).Value, 0)
    dataStartCol = ParseColumnSpec(wsRule.Cells(ruleRow, rcStartCol).Value, 0)
    dataEndCol = ParseColumnSpec(wsRule.Cells(ruleRow, rcEndCol).Value, 0)
    skipKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcSkipKeywords).Value))

    If ruleName = "" Then
        ruleName = "规则" & CStr(ruleRow)
    End If

    If bookKeywords <> "" Then
        If Not MatchAllKeywords(targetWb.Name, bookKeywords) Then
            Exit Sub
        End If
    End If

    If colHeaderRows.Count = 0 Then
        Exit Sub
    End If

    If dataStartRow <= 0 Or dataStartCol <= 0 Then
        Exit Sub
    End If

    Set matchedSheets = MatchWorksheets(targetWb, sheetKeywords)
    If matchedSheets Is Nothing Or matchedSheets.Count = 0 Then
        skipRules = skipRules + 1
        RunLog_WriteRow LOG_KEY, "跳过规则", targetWb.Name & "|" & ruleName, "", "", "跳过", "未匹配到工作表", ""
        Exit Sub
    End If

    For Each item In matchedSheets
        ExtractSheetToTimeline item, wsResult, resultRow, targetWb.Name, fileModified, dataDateText, dateSource, ruleName, rowHeaderCols, colHeaderRows, requiredColPaths, requiredRowPaths, dataStartRow, dataEndRow, dataStartCol, dataEndCol, skipKeywords, duplicateMap, hitSheets, skipRules, outputRows, duplicateRows
    Next item
End Sub

Private Sub ExtractSheetToTimeline(ByVal ws As Worksheet, ByVal wsResult As Worksheet, ByRef resultRow As Long, ByVal sourceBookName As String, ByVal fileModified As String, ByVal dataDateText As String, ByVal dateSource As String, ByVal ruleName As String, ByVal rowHeaderCols As Collection, ByVal colHeaderRows As Collection, ByVal requiredColPaths As String, ByVal requiredRowPaths As String, ByVal dataStartRow As Long, ByVal dataEndRow As Long, ByVal dataStartCol As Long, ByVal dataEndCol As Long, ByVal skipKeywords As String, ByVal duplicateMap As Object, ByRef hitSheets As Long, ByRef skipRules As Long, ByRef outputRows As Long, ByRef duplicateRows As Long)
    Dim actualEndRow As Long
    Dim actualEndCol As Long
    Dim headerText() As String
    Dim rowPathMap As Object
    Dim colPathMap As Object
    Dim rowNo As Long
    Dim colNo As Long
    Dim rowPath As String
    Dim colPath As String
    Dim cellValue As Variant
    Dim valuesWritten As Long
    Dim oneKey As String

    actualEndRow = dataEndRow
    If actualEndRow <= 0 Then
        actualEndRow = GetLastUsedRow(ws)
    End If

    actualEndCol = dataEndCol
    If actualEndCol <= 0 Then
        actualEndCol = GetLastUsedCol(ws)
    End If

    If actualEndRow < dataStartRow Or actualEndCol < dataStartCol Then
        Exit Sub
    End If

    ReDim headerText(1 To colHeaderRows.Count, dataStartCol To actualEndCol)
    BuildHeaderCache ws, colHeaderRows, dataStartCol, actualEndCol, headerText
    Set rowPathMap = BuildUniqueRowPathMap(sourceBookName, ws.Name, ruleName, ws, rowHeaderCols, dataStartRow, actualEndRow, skipKeywords)
    Set colPathMap = BuildUniqueColPathMap(sourceBookName, ws.Name, ruleName, colHeaderRows, headerText, dataStartCol, actualEndCol)
    If Not ValidateRequiredAnchors(ws, rowHeaderCols, headerText, dataStartRow, actualEndRow, dataStartCol, actualEndCol, requiredColPaths, requiredRowPaths, skipKeywords) Then
        skipRules = skipRules + 1
        RunLog_WriteRow LOG_KEY, "跳过规则", sourceBookName & "|" & ws.Name & "|" & ruleName, "", "", "跳过", "未命中必含列头或必含行头", ""
        Exit Sub
    End If

    hitSheets = hitSheets + 1
    For rowNo = dataStartRow To actualEndRow
        rowPath = GetUniquePathFromMap(rowPathMap, rowNo)
        If rowPath <> "" Then
            If Not MatchAnyKeyword(rowPath, skipKeywords) Then
                For colNo = dataStartCol To actualEndCol
                    cellValue = ws.Cells(rowNo, colNo).Value
                    If IsOutputValue(cellValue) Then
                        colPath = GetUniquePathFromMap(colPathMap, colNo)
                        If colPath <> "" Then
                            oneKey = BuildDuplicateKey(sourceBookName, ws.Name, dataDateText, rowPath, colPath, cellValue)
                            If duplicateMap.Exists(oneKey) Then
                                duplicateRows = duplicateRows + 1
                                RunLog_WriteRow LOG_KEY, "重复记录", sourceBookName & "|" & ws.Name & "|" & ruleName, dataDateText, rowPath, "跳过", "当前=" & ws.Cells(rowNo, colNo).Address(False, False) & "|列头=" & colPath & "|首次=" & CStr(duplicateMap(oneKey)), ""
                            Else
                                duplicateMap(oneKey) = "规则=" & ruleName & "|单元格=" & ws.Cells(rowNo, colNo).Address(False, False) & "|行头=" & rowPath & "|列头=" & colPath
                                WriteTimelineRow wsResult, resultRow, sourceBookName, ws.Name, ruleName, fileModified, dataDateText, dateSource, rowPath, colPath, cellValue, ws.Cells(rowNo, colNo).Address(False, False)
                                resultRow = resultRow + 1
                                valuesWritten = valuesWritten + 1
                            End If
                        End If
                    End If
                Next colNo
            End If
        End If
    Next rowNo

    outputRows = outputRows + valuesWritten
    RunLog_WriteRow LOG_KEY, "提取工作表", sourceBookName & "|" & ws.Name & "|" & ruleName, "", CStr(valuesWritten), "完成", "OK", ""
End Sub

Private Sub BuildHeaderCache(ByVal ws As Worksheet, ByVal headerRows As Collection, ByVal startCol As Long, ByVal endCol As Long, ByRef headerText() As String)
    Dim idx As Long
    Dim rowNo As Long
    Dim colNo As Long

    For idx = 1 To headerRows.Count
        rowNo = CLng(headerRows(idx))
        For colNo = startCol To endCol
            headerText(idx, colNo) = GetMergedAwareText(ws.Cells(rowNo, colNo))
        Next colNo
    Next idx
End Sub

Private Function BuildRowPath(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal rowHeaderCols As Collection) As String
    Dim idx As Long
    Dim colNo As Long
    Dim oneText As String

    If rowHeaderCols Is Nothing Or rowHeaderCols.Count = 0 Then
        Exit Function
    End If

    For idx = 1 To rowHeaderCols.Count
        colNo = CLng(rowHeaderCols(idx))
        oneText = GetMergedAwareText(ws.Cells(rowNo, colNo))
        If oneText <> "" Then
            If BuildRowPath <> "" Then
                BuildRowPath = BuildRowPath & "_"
            End If
            BuildRowPath = BuildRowPath & oneText
        End If
    Next idx
End Function

Private Function BuildColPath(ByRef headerText() As String, ByVal colNo As Long) As String
    Dim idx As Long
    Dim oneText As String

    For idx = LBound(headerText, 1) To UBound(headerText, 1)
        oneText = Trim$(CStr(headerText(idx, colNo)))
        If oneText <> "" Then
            If BuildColPath <> "" Then
                BuildColPath = BuildColPath & "_"
            End If
            BuildColPath = BuildColPath & oneText
        End If
    Next idx
End Function

Private Function BuildUniqueRowPathMap(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal ws As Worksheet, ByVal rowHeaderCols As Collection, ByVal startRow As Long, ByVal endRow As Long, ByVal skipKeywords As String) As Object
    Set BuildUniqueRowPathMap = BuildUniquePathMapForRows(sourceBookName, sourceSheetName, ruleName, ws, rowHeaderCols, startRow, endRow, skipKeywords)
End Function

Private Function BuildUniqueColPathMap(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal colHeaderRows As Collection, ByRef headerText() As String, ByVal startCol As Long, ByVal endCol As Long) As Object
    Set BuildUniqueColPathMap = BuildUniquePathMapForCols(sourceBookName, sourceSheetName, ruleName, colHeaderRows, headerText, startCol, endCol)
End Function

Private Function BuildUniquePathMapForRows(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal ws As Worksheet, ByVal rowHeaderCols As Collection, ByVal startRow As Long, ByVal endRow As Long, ByVal skipKeywords As String) As Object
    Dim countMap As Object
    Dim seqMap As Object
    Dim result As Object
    Dim rowNo As Long
    Dim basePath As String
    Dim seq As Long
    Dim renamedPath As String

    Set countMap = CreateObject("Scripting.Dictionary")
    countMap.CompareMode = vbTextCompare
    Set seqMap = CreateObject("Scripting.Dictionary")
    seqMap.CompareMode = vbTextCompare
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    For rowNo = startRow To endRow
        basePath = BuildRowPath(ws, rowNo, rowHeaderCols)
        If basePath <> "" Then
            If Not MatchAnyKeyword(basePath, skipKeywords) Then
                If countMap.Exists(basePath) Then
                    countMap(basePath) = CLng(countMap(basePath)) + 1
                Else
                    countMap(basePath) = 1
                End If
            End If
        End If
    Next rowNo

    For rowNo = startRow To endRow
        basePath = BuildRowPath(ws, rowNo, rowHeaderCols)
        If basePath <> "" Then
            If Not MatchAnyKeyword(basePath, skipKeywords) Then
                If CLng(countMap(basePath)) > 1 Then
                    If seqMap.Exists(basePath) Then
                        seq = CLng(seqMap(basePath)) + 1
                    Else
                        seq = 1
                    End If
                    seqMap(basePath) = seq
                    renamedPath = basePath & "_" & CStr(seq)
                    result(CStr(rowNo)) = renamedPath
                    RunLog_WriteRow LOG_KEY, "行头重命名", sourceBookName & "|" & sourceSheetName & "|" & ruleName, basePath, renamedPath, "完成", "位置=R" & CStr(rowNo) & "|从=" & basePath & "|到=" & renamedPath, ""
                Else
                    result(CStr(rowNo)) = basePath
                End If
            End If
        End If
    Next rowNo

    Set BuildUniquePathMapForRows = result
End Function

Private Function BuildUniquePathMapForCols(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal colHeaderRows As Collection, ByRef headerText() As String, ByVal startCol As Long, ByVal endCol As Long) As Object
    Dim countMap As Object
    Dim seqMap As Object
    Dim result As Object
    Dim colNo As Long
    Dim basePath As String
    Dim seq As Long
    Dim renamedPath As String

    Set countMap = CreateObject("Scripting.Dictionary")
    countMap.CompareMode = vbTextCompare
    Set seqMap = CreateObject("Scripting.Dictionary")
    seqMap.CompareMode = vbTextCompare
    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare

    For colNo = startCol To endCol
        basePath = BuildColPath(headerText, colNo)
        If basePath <> "" Then
            If countMap.Exists(basePath) Then
                countMap(basePath) = CLng(countMap(basePath)) + 1
            Else
                countMap(basePath) = 1
            End If
        End If
    Next colNo

    For colNo = startCol To endCol
        basePath = BuildColPath(headerText, colNo)
        If basePath <> "" Then
            If CLng(countMap(basePath)) > 1 Then
                If seqMap.Exists(basePath) Then
                    seq = CLng(seqMap(basePath)) + 1
                Else
                    seq = 1
                End If
                seqMap(basePath) = seq
                renamedPath = basePath & "_" & CStr(seq)
                result(CStr(colNo)) = renamedPath
                RunLog_WriteRow LOG_KEY, "列头重命名", sourceBookName & "|" & sourceSheetName & "|" & ruleName, basePath, renamedPath, "完成", "位置=列" & ColumnNumberToLetter(colNo) & " 行" & BuildHeaderRowsLabel(colHeaderRows) & "|从=" & basePath & "|到=" & renamedPath, ""
            Else
                result(CStr(colNo)) = basePath
            End If
        End If
    Next colNo

    Set BuildUniquePathMapForCols = result
End Function

Private Function GetUniquePathFromMap(ByVal pathMap As Object, ByVal keyNo As Long) As String
    Dim mapKey As String

    If pathMap Is Nothing Then
        Exit Function
    End If

    mapKey = CStr(keyNo)
    If pathMap.Exists(mapKey) Then
        GetUniquePathFromMap = CStr(pathMap(mapKey))
    End If
End Function

Private Function BuildHeaderRowsLabel(ByVal colHeaderRows As Collection) As String
    Dim idx As Long

    If colHeaderRows Is Nothing Then
        Exit Function
    End If

    For idx = 1 To colHeaderRows.Count
        If BuildHeaderRowsLabel <> "" Then
            BuildHeaderRowsLabel = BuildHeaderRowsLabel & ","
        End If
        BuildHeaderRowsLabel = BuildHeaderRowsLabel & CStr(colHeaderRows(idx))
    Next idx
End Function

Private Function ColumnNumberToLetter(ByVal colNo As Long) As String
    Dim n As Long

    n = colNo
    Do While n > 0
        ColumnNumberToLetter = Chr$(((n - 1) Mod 26) + 65) & ColumnNumberToLetter
        n = (n - 1) \ 26
    Loop
End Function

Private Function GetMergedAwareText(ByVal targetCell As Range) As String
    Dim rawText As String

    On Error Resume Next
    If targetCell.MergeCells Then
        rawText = CStr(targetCell.MergeArea.Cells(1, 1).Value)
    Else
        rawText = CStr(targetCell.Value)
    End If
    On Error GoTo 0

    GetMergedAwareText = NormalizeText(rawText)
End Function

Private Sub WriteTimelineRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal fileModified As String, ByVal dataDateText As String, ByVal dateSource As String, ByVal rowPath As String, ByVal colPath As String, ByVal valueText As Variant, ByVal cellAddress As String)
    ws.Cells(rowNo, rsExecTime).Value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(rowNo, rsSourceBook).Value = sourceBookName
    ws.Cells(rowNo, rsSourceSheet).Value = sourceSheetName
    ws.Cells(rowNo, rsRuleName).Value = ruleName
    ws.Cells(rowNo, rsFileModified).Value = fileModified
    ws.Cells(rowNo, rsDataDate).Value = dataDateText
    ws.Cells(rowNo, rsDateSource).Value = dateSource
    ws.Cells(rowNo, rsRowPath).Value = rowPath
    ws.Cells(rowNo, rsColPath).Value = colPath
    ws.Cells(rowNo, rsValue).Value = valueText
    ws.Cells(rowNo, rsCellAddress).Value = cellAddress
End Sub

Private Function MatchWorksheets(ByVal wb As Workbook, ByVal sheetKeywords As String) As Collection
    Dim result As New Collection
    Dim keys As Object
    Dim ws As Worksheet

    Set keys = CreateObject("Scripting.Dictionary")
    keys.CompareMode = vbTextCompare

    For Each ws In wb.Worksheets
        If Trim$(sheetKeywords) = "" Then
            AddSheetUnique result, keys, ws
        ElseIf MatchAllKeywords(ws.Name, sheetKeywords) Then
            AddSheetUnique result, keys, ws
        End If
    Next ws

    Set MatchWorksheets = result
End Function

Private Function MatchAllKeywords(ByVal sourceText As String, ByVal keywordText As String) As Boolean
    Dim arr As Variant
    Dim i As Long
    Dim token As String

    sourceText = CStr(sourceText)
    keywordText = Replace(keywordText, "，", ";")
    keywordText = Replace(keywordText, "；", ";")
    keywordText = Replace(keywordText, ",", ";")
    arr = Split(keywordText, ";")

    MatchAllKeywords = True
    For i = LBound(arr) To UBound(arr)
        token = Trim$(CStr(arr(i)))
        If token <> "" Then
            If InStr(1, sourceText, token, vbTextCompare) = 0 Then
                MatchAllKeywords = False
                Exit Function
            End If
        End If
    Next i
End Function

Private Function MatchAnyKeyword(ByVal sourceText As String, ByVal keywordText As String) As Boolean
    Dim arr As Variant
    Dim i As Long
    Dim token As String

    keywordText = Replace(keywordText, "，", ";")
    keywordText = Replace(keywordText, "；", ";")
    keywordText = Replace(keywordText, ",", ";")
    arr = Split(keywordText, ";")

    For i = LBound(arr) To UBound(arr)
        token = Trim$(CStr(arr(i)))
        If token <> "" Then
            If InStr(1, sourceText, token, vbTextCompare) > 0 Then
                MatchAnyKeyword = True
                Exit Function
            End If
        End If
    Next i
End Function

Private Function ValidateRequiredAnchors(ByVal ws As Worksheet, ByVal rowHeaderCols As Collection, ByRef headerText() As String, ByVal dataStartRow As Long, ByVal actualEndRow As Long, ByVal dataStartCol As Long, ByVal actualEndCol As Long, ByVal requiredColPaths As String, ByVal requiredRowPaths As String, ByVal skipKeywords As String) As Boolean
    Dim requiredCols As Collection
    Dim requiredRows As Collection
    Dim idx As Long
    Dim colNo As Long
    Dim rowNo As Long
    Dim found As Boolean
    Dim candidate As String

    Set requiredCols = SplitTokens(requiredColPaths)
    For idx = 1 To requiredCols.Count
        found = False
        For colNo = dataStartCol To actualEndCol
            candidate = BuildColPath(headerText, colNo)
            If StrComp(candidate, CStr(requiredCols(idx)), vbTextCompare) = 0 Then
                found = True
                Exit For
            End If
        Next colNo
        If Not found Then
            Exit Function
        End If
    Next idx

    Set requiredRows = SplitTokens(requiredRowPaths)
    For idx = 1 To requiredRows.Count
        found = False
        For rowNo = dataStartRow To actualEndRow
            candidate = BuildRowPath(ws, rowNo, rowHeaderCols)
            If candidate <> "" Then
                If Not MatchAnyKeyword(candidate, skipKeywords) Then
                    If StrComp(candidate, CStr(requiredRows(idx)), vbTextCompare) = 0 Then
                        found = True
                        Exit For
                    End If
                End If
            End If
        Next rowNo
        If Not found Then
            Exit Function
        End If
    Next idx

    ValidateRequiredAnchors = True
End Function

Private Function SplitTokens(ByVal rawText As String) As Collection
    Dim result As New Collection
    Dim arr As Variant
    Dim token As String
    Dim i As Long

    rawText = Replace(rawText, "，", ";")
    rawText = Replace(rawText, ",", ";")
    arr = Split(rawText, ";")

    For i = LBound(arr) To UBound(arr)
        token = Trim$(CStr(arr(i)))
        If token <> "" Then
            result.Add token
        End If
    Next i

    Set SplitTokens = result
End Function

Private Function ShouldSkipRuleRow(ByVal wsRule As Worksheet, ByVal rowNo As Long) As Boolean
    If Not IsEnabledValue(wsRule.Cells(rowNo, rcEnabled).Value) Then
        ShouldSkipRuleRow = True
        Exit Function
    End If

    If Trim$(CStr(wsRule.Cells(rowNo, rcRuleName).Value)) = "" And Trim$(CStr(wsRule.Cells(rowNo, rcSheetKeywords).Value)) = "" Then
        ShouldSkipRuleRow = True
    End If
End Function

Private Function EnsureRuleSheet() As Worksheet
    On Error Resume Next
    Set EnsureRuleSheet = ThisWorkbook.Worksheets(RULE_SHEET_NAME)
    On Error GoTo 0

    If EnsureRuleSheet Is Nothing Then
        Set EnsureRuleSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureRuleSheet.Name = RULE_SHEET_NAME
    End If
End Function

Private Sub InitRuleHeader(ByVal ws As Worksheet)
    ws.Cells(1, rcEnabled).Value = "是否启用"
    ws.Cells(1, rcRuleName).Value = "规则名称"
    ws.Cells(1, rcBookKeywords).Value = "工作簿关键字"
    ws.Cells(1, rcSheetKeywords).Value = "工作表关键字"
    ws.Cells(1, rcRowHeaderCols).Value = "行头列"
    ws.Cells(1, rcColHeaderRows).Value = "列表头行"
    ws.Cells(1, rcRequiredColPaths).Value = "必含列头"
    ws.Cells(1, rcRequiredRowPaths).Value = "必含行头"
    ws.Cells(1, rcStartRow).Value = "数据起始行"
    ws.Cells(1, rcEndRow).Value = "数据结束行"
    ws.Cells(1, rcStartCol).Value = "数据起始列"
    ws.Cells(1, rcEndCol).Value = "数据结束列"
    ws.Cells(1, rcSkipKeywords).Value = "跳过关键字"
    ws.Cells(1, rcRemark).Value = "备注"
    ws.Rows(1).Font.Bold = True
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
    ws.Cells(1, rsSourceBook).Value = "源文件"
    ws.Cells(1, rsSourceSheet).Value = "工作表名"
    ws.Cells(1, rsRuleName).Value = "规则名称"
    ws.Cells(1, rsFileModified).Value = "源文件修改时间"
    ws.Cells(1, rsDataDate).Value = "数据日期"
    ws.Cells(1, rsDateSource).Value = "日期来源"
    ws.Cells(1, rsRowPath).Value = "行头路径"
    ws.Cells(1, rsColPath).Value = "列头路径"
    ws.Cells(1, rsValue).Value = "数值"
    ws.Cells(1, rsCellAddress).Value = "单元格地址"
    ws.Rows(1).Font.Bold = True
End Sub

Private Function ParseNumberCollection(ByVal rawText As String) As Collection
    Dim result As New Collection
    Dim arr As Variant
    Dim i As Long
    Dim token As String

    rawText = Replace(rawText, "，", ",")
    rawText = Replace(rawText, "；", ",")
    arr = Split(rawText, ",")

    For i = LBound(arr) To UBound(arr)
        token = Trim$(CStr(arr(i)))
        If token <> "" Then
            result.Add CLng(token)
        End If
    Next i

    Set ParseNumberCollection = result
End Function

Private Function ParseColumnCollection(ByVal rawText As String) As Collection
    Dim result As New Collection
    Dim arr As Variant
    Dim i As Long
    Dim token As String

    rawText = Replace(rawText, "，", ",")
    rawText = Replace(rawText, "；", ",")
    arr = Split(rawText, ",")

    For i = LBound(arr) To UBound(arr)
        token = Trim$(CStr(arr(i)))
        If token <> "" Then
            result.Add ParseColumnSpec(token, 0)
        End If
    Next i

    Set ParseColumnCollection = result
End Function

Private Function ParseColumnSpec(ByVal rawValue As Variant, ByVal defaultValue As Long) As Long
    Dim textValue As String

    textValue = Trim$(CStr(rawValue))
    If textValue = "" Then
        ParseColumnSpec = defaultValue
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
            Err.Raise vbObjectError + 5101, , "列表达式无效：" & colText
        End If
        ColLetterToNumber = ColLetterToNumber * 26 + Asc(Mid$(colText, i, 1)) - 64
    Next i
End Function

Private Function ParseLongValue(ByVal rawValue As Variant, ByVal defaultValue As Long) As Long
    If IsNumeric(rawValue) Then
        ParseLongValue = CLng(rawValue)
    Else
        ParseLongValue = defaultValue
    End If
End Function

Private Function IsEnabledValue(ByVal rawValue As Variant) As Boolean
    Dim textValue As String

    textValue = UCase$(Trim$(CStr(rawValue)))
    IsEnabledValue = (textValue = "Y" Or textValue = "1" Or textValue = "TRUE" Or textValue = "是")
End Function

Private Function NormalizeText(ByVal rawText As String) As String
    Dim textValue As String

    textValue = CStr(rawText)
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    textValue = Replace(textValue, ChrW(&H3000), " ")
    textValue = Replace(textValue, Chr(160), " ")
    textValue = Trim$(textValue)

    Do While InStr(textValue, "  ") > 0
        textValue = Replace(textValue, "  ", " ")
    Loop

    NormalizeText = textValue
End Function

Private Function IsOutputValue(ByVal rawValue As Variant) As Boolean
    Dim textValue As String

    If IsError(rawValue) Or IsEmpty(rawValue) Then
        Exit Function
    End If

    If IsNumeric(rawValue) Then
        IsOutputValue = True
        Exit Function
    End If

    textValue = Trim$(CStr(rawValue))
    IsOutputValue = (textValue <> "")
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

Private Sub AddSheetUnique(ByRef sheets As Collection, ByVal keys As Object, ByVal ws As Worksheet)
    Dim oneKey As String

    oneKey = ws.Parent.Name & "|" & ws.Name
    If keys.Exists(oneKey) Then
        Exit Sub
    End If

    keys(oneKey) = True
    sheets.Add ws
End Sub

Private Sub SafeCloseWorkbook(ByRef wb As Workbook)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
        Set wb = Nothing
    End If
    On Error GoTo 0
End Sub

Private Function BuildSummaryMessage(ByVal hitBooks As Long, ByVal hitSheets As Long, ByVal outputRows As Long, ByVal skipRules As Long, ByVal skipBooks As Long, ByVal duplicateRows As Long) As String
    BuildSummaryMessage = "处理文件数：" & hitBooks
    BuildSummaryMessage = BuildSummaryMessage & vbCrLf & "处理工作表数：" & hitSheets
    BuildSummaryMessage = BuildSummaryMessage & vbCrLf & "输出记录数：" & outputRows
    BuildSummaryMessage = BuildSummaryMessage & vbCrLf & "跳过规则数：" & skipRules
    BuildSummaryMessage = BuildSummaryMessage & vbCrLf & "跳过工作簿数：" & skipBooks
    BuildSummaryMessage = BuildSummaryMessage & vbCrLf & "已自动筛除重复记录数：" & duplicateRows
    BuildSummaryMessage = BuildSummaryMessage & vbCrLf & vbCrLf & "结果已写入新工作簿：" & RESULT_SHEET_NAME
End Function

Private Function BuildDuplicateKey(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal dataDateText As String, ByVal rowPath As String, ByVal colPath As String, ByVal rawValue As Variant) As String
    BuildDuplicateKey = sourceBookName & "|" & sourceSheetName & "|" & dataDateText & "|" & rowPath & "|" & colPath & "|" & NormalizeValueText(rawValue)
End Function

Private Function NormalizeValueText(ByVal rawValue As Variant) As String
    If IsError(rawValue) Or IsEmpty(rawValue) Then
        Exit Function
    End If

    If IsNumeric(rawValue) Then
        NormalizeValueText = Format(CDbl(rawValue), "0.###############")
    Else
        NormalizeValueText = Trim$(CStr(rawValue))
    End If
End Function

Private Function ResolveWorkbookDataDate(ByVal wb As Workbook, ByRef dataDateText As String, ByRef dateSource As String) As Boolean
    Dim ws As Worksheet

    If TryParseDateFromText(wb.Name, dataDateText) Then
        dateSource = "工作簿名"
        ResolveWorkbookDataDate = True
        Exit Function
    End If

    For Each ws In wb.Worksheets
        If TryParseDateFromText(ws.Name, dataDateText) Then
            dateSource = "工作表名"
            ResolveWorkbookDataDate = True
            Exit Function
        End If
    Next ws
End Function

Private Function TryParseDateFromText(ByVal sourceText As String, ByRef normalizedDate As String) As Boolean
    Dim re As Object
    Dim matches As Object
    Dim yearNo As Long
    Dim monthNo As Long
    Dim dayNo As Long

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    re.Pattern = "((19|20)\d{2})[.\-_/年]\s*(\d{1,2})[.\-_/月]\s*(\d{1,2})(?:日)?"
    If re.Test(sourceText) Then
        Set matches = re.Execute(sourceText)
        yearNo = CLng(matches(0).SubMatches(0))
        monthNo = CLng(matches(0).SubMatches(2))
        dayNo = CLng(matches(0).SubMatches(3))
        normalizedDate = BuildIsoDate(yearNo, monthNo, dayNo)
        TryParseDateFromText = (normalizedDate <> "")
        Exit Function
    End If

    re.Pattern = "((19|20)\d{2})(\d{2})(\d{2})"
    If re.Test(sourceText) Then
        Set matches = re.Execute(sourceText)
        yearNo = CLng(matches(0).SubMatches(0))
        monthNo = CLng(matches(0).SubMatches(2))
        dayNo = CLng(matches(0).SubMatches(3))
        normalizedDate = BuildIsoDate(yearNo, monthNo, dayNo)
        TryParseDateFromText = (normalizedDate <> "")
        Exit Function
    End If

    re.Pattern = "((19|20)\d{2})[.\-_/年]\s*(\d{1,2})(?:月)?"
    If re.Test(sourceText) Then
        Set matches = re.Execute(sourceText)
        yearNo = CLng(matches(0).SubMatches(0))
        monthNo = CLng(matches(0).SubMatches(2))
        dayNo = Day(DateSerial(yearNo, monthNo + 1, 0))
        normalizedDate = BuildIsoDate(yearNo, monthNo, dayNo)
        TryParseDateFromText = (normalizedDate <> "")
        Exit Function
    End If

    re.Pattern = "((19|20)\d{2})(\d{2})(?!\d)"
    If re.Test(sourceText) Then
        Set matches = re.Execute(sourceText)
        yearNo = CLng(matches(0).SubMatches(0))
        monthNo = CLng(matches(0).SubMatches(2))
        dayNo = Day(DateSerial(yearNo, monthNo + 1, 0))
        normalizedDate = BuildIsoDate(yearNo, monthNo, dayNo)
        TryParseDateFromText = (normalizedDate <> "")
    End If
End Function

Private Function BuildIsoDate(ByVal yearNo As Long, ByVal monthNo As Long, ByVal dayNo As Long) As String
    Dim oneDate As Date

    On Error Resume Next
    oneDate = DateSerial(yearNo, monthNo, dayNo)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    If Year(oneDate) <> yearNo Or Month(oneDate) <> monthNo Or Day(oneDate) <> dayNo Then
        Exit Function
    End If

    BuildIsoDate = Format(oneDate, "yyyy-mm-dd")
End Function
