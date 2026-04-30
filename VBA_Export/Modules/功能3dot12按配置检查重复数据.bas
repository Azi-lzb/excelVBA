Attribute VB_Name = "功能3dot12按配置检查重复数据"
Option Explicit

Private Const CONFIG_SHEET_NAME As String = "去重追加数据配置"
Private Const CONFIG_SHEET_NAME_LEGACY As String = "按配置查重"
Private Const PRECHECK_LOG_KEY As String = "3.11.6 按配置预校验"
Private Const HEADER_ROW As Long = 1
Private Const DATA_START_ROW As Long = 2
Private Const KEY_SEP As String = "|#|"

Public Sub InitDedupConfigSheet()
    Application.Run "功能3dot12_初始化按配置查重"
End Sub

Public Sub ExecuteDedupCheckByConfig()
    Application.Run "功能3dot12_按配置检查重复数据"
End Sub

Public Sub ExecuteDedupPrecheckByConfig()
    Application.Run "功能3dot12_按配置预校验"
End Sub

Public Sub 功能3dot12_初始化按配置查重()
    Dim ws As Worksheet

    Set ws = GetOrCreateConfigSheet()
    InitConfigHeader ws
    WriteConfigExample ws
    ws.Columns("A:H").AutoFit
    ws.Activate

    MsgBox "按配置查重配置表已初始化。", vbInformation, "按配置检查重复数据"
End Sub

Public Sub 功能3dot12_按配置检查重复数据()
    Dim wsCfg As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim enabled As Boolean
    Dim wbPathText As String
    Dim sheetName As String
    Dim dedupeColsText As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim dedupeCols As Collection
    Dim dupCount As Long
    Dim totalDup As Long
    Dim hitTask As Long
    Dim skipTask As Long
    Dim msg As String
    Dim wbCache As Object
    Dim openedByCode As Object
    Dim modified As Object
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation

    Set wsCfg = FindConfigSheet()
    If wsCfg Is Nothing Then
        功能3dot12_初始化按配置查重
        MsgBox "未找到配置表，已为你创建。请先填写后再执行。", vbExclamation, "按配置检查重复数据"
        Exit Sub
    End If

    lastRow = GetLastUsedRow(wsCfg)
    If lastRow < 2 Then
        MsgBox "配置表为空，请先填写配置。", vbExclamation, "按配置检查重复数据"
        Exit Sub
    End If

    Set wbCache = CreateObject("Scripting.Dictionary")
    Set openedByCode = CreateObject("Scripting.Dictionary")
    Set modified = CreateObject("Scripting.Dictionary")

    CaptureAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc
    BeginFastMode
    On Error GoTo FailHandler

    For r = 2 To lastRow
        enabled = IsTruthyValue(wsCfg.cells(r, 1).Value2)
        If Not enabled Then GoTo nextRow

        wbPathText = NormalizeText(wsCfg.cells(r, 2).Value2)
        sheetName = NormalizeText(wsCfg.cells(r, 3).Value2)
        dedupeColsText = NormalizeText(wsCfg.cells(r, 4).Value2)

        If Len(wbPathText) = 0 Or Len(sheetName) = 0 Then
            skipTask = skipTask + 1
            GoTo nextRow
        End If

        Set wb = AcquireWorkbookByPath(wbPathText, wbCache, openedByCode, msg)
        If wb Is Nothing Then
            skipTask = skipTask + 1
            GoTo nextRow
        End If

        Set ws = GetWorksheetByName(wb, sheetName)
        If ws Is Nothing Then
            skipTask = skipTask + 1
            GoTo nextRow
        End If

        Set dedupeCols = ParseIndexCollection(dedupeColsText)
        If dedupeCols Is Nothing Then
            Set dedupeCols = BuildAllUsedColumnCollection(ws)
        ElseIf dedupeCols.count = 0 Then
            Set dedupeCols = BuildAllUsedColumnCollection(ws)
        Else
            Set dedupeCols = FilterColumnsByWorksheet(dedupeCols, ws)
            If dedupeCols Is Nothing Then
                Set dedupeCols = BuildAllUsedColumnCollection(ws)
            ElseIf dedupeCols.count = 0 Then
                Set dedupeCols = BuildAllUsedColumnCollection(ws)
            End If
        End If
        If dedupeCols Is Nothing Then
            skipTask = skipTask + 1
            GoTo nextRow
        End If
        If dedupeCols.count = 0 Then
            skipTask = skipTask + 1
            GoTo nextRow
        End If

        dupCount = MarkDuplicateRowsByIndexes(ws, dedupeCols)
        If dupCount > 0 Then
            MarkModifiedPath modified, NormalizeText(wb.fullName)
        End If
        totalDup = totalDup + dupCount
        hitTask = hitTask + 1

nextRow:
    Next r

    SaveModifiedWorkbooks wbCache, modified, openedByCode
    CloseOpenedWorkbooks wbCache, openedByCode
    RestoreAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc

    MsgBox "按配置检查重复完成。" & vbCrLf & _
           "执行任务数：" & hitTask & vbCrLf & _
           "跳过任务数：" & skipTask & vbCrLf & _
           "重复标红行数：" & totalDup, vbInformation, "按配置检查重复数据"
    Exit Sub

FailHandler:
    SaveModifiedWorkbooks wbCache, modified, openedByCode
    CloseOpenedWorkbooks wbCache, openedByCode
    RestoreAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc
    MsgBox "执行失败：" & CStr(Err.Number) & " " & Err.Description, vbCritical, "按配置检查重复数据"
End Sub

Public Sub 功能3dot12_按配置预校验()
    Dim wsCfg As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim enabled As Boolean
    Dim srcWbPath As String
    Dim srcSheetName As String
    Dim dedupeColsText As String
    Dim tgtWbPath As String
    Dim tgtSheetName As String
    Dim execMode As String
    Dim srcWb As Workbook
    Dim tgtWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtWs As Worksheet
    Dim dedupeCols As Collection
    Dim wbCache As Object
    Dim openedByCode As Object
    Dim msg As String
    Dim okCount As Long
    Dim warnCount As Long
    Dim failCount As Long
    Dim skipCount As Long
    Dim hasTarget As Boolean
    Dim precheckMsg As String
    Dim t0 As Double

    t0 = Timer
    RunLog_WriteRow PRECHECK_LOG_KEY, "开始", "", "", "", "", "开始", ""

    Set wsCfg = FindConfigSheet()
    If wsCfg Is Nothing Then
        功能3dot12_初始化按配置查重
        RunLog_WriteRow PRECHECK_LOG_KEY, "结束", "", "", "", "跳过", "未找到配置表，已初始化", CStr(Round(Timer - t0, 2))
        MsgBox "未找到配置表，已为你创建。请先填写后再校验。", vbExclamation, "按配置预校验"
        Exit Sub
    End If

    lastRow = GetLastUsedRow(wsCfg)
    If lastRow < 2 Then
        RunLog_WriteRow PRECHECK_LOG_KEY, "结束", "", "", "", "跳过", "配置表为空", CStr(Round(Timer - t0, 2))
        MsgBox "配置表为空，请先填写配置。", vbExclamation, "按配置预校验"
        Exit Sub
    End If

    Set wbCache = CreateObject("Scripting.Dictionary")
    Set openedByCode = CreateObject("Scripting.Dictionary")

    On Error GoTo FailHandler

    For r = 2 To lastRow
        enabled = IsTruthyValue(wsCfg.Cells(r, 1).Value2)
        If Not enabled Then
            skipCount = skipCount + 1
            GoTo NextRow
        End If

        srcWbPath = NormalizeText(wsCfg.Cells(r, 2).Value2)
        srcSheetName = NormalizeText(wsCfg.Cells(r, 3).Value2)
        dedupeColsText = NormalizeText(wsCfg.Cells(r, 4).Value2)
        tgtWbPath = NormalizeText(wsCfg.Cells(r, 5).Value2)
        tgtSheetName = NormalizeText(wsCfg.Cells(r, 6).Value2)
        execMode = NormalizeText(wsCfg.Cells(r, 7).Value2)
        hasTarget = (Len(tgtWbPath) > 0 Or Len(tgtSheetName) > 0)

        If Len(srcWbPath) = 0 Or Len(srcSheetName) = 0 Then
            failCount = failCount + 1
            RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcSheetName, "", "失败", "源工作簿或源工作表为空", ""
            GoTo NextRow
        End If

        Set srcWb = AcquireWorkbookByPath(srcWbPath, wbCache, openedByCode, msg)
        If srcWb Is Nothing Then
            failCount = failCount + 1
            RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWbPath, "", "失败", "源工作簿不可打开：" & msg, ""
            GoTo NextRow
        End If

        Set srcWs = GetWorksheetByName(srcWb, srcSheetName)
        If srcWs Is Nothing Then
            failCount = failCount + 1
            RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWb.Name, "", "失败", "源工作表不存在", ""
            GoTo NextRow
        End If

        If Len(dedupeColsText) > 0 Then
            Set dedupeCols = ParseIndexCollection(dedupeColsText)
            If dedupeCols Is Nothing Then
                failCount = failCount + 1
                RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWs.Name, "", "失败", "标识列序号格式错误", ""
                GoTo NextRow
            End If
            Set dedupeCols = FilterColumnsByWorksheet(dedupeCols, srcWs)
            If dedupeCols Is Nothing Or dedupeCols.count = 0 Then
                failCount = failCount + 1
                RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWs.Name, "", "失败", "标识列越界或为空", ""
                GoTo NextRow
            End If
        End If

        If hasTarget Then
            If Len(tgtWbPath) = 0 Or Len(tgtSheetName) = 0 Then
                failCount = failCount + 1
                RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWs.Name, "", "失败", "目标工作簿/目标工作表需同时填写", ""
                GoTo NextRow
            End If

            If Not ValidateTargetPathForPrecheck(tgtWbPath, msg) Then
                failCount = failCount + 1
                RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWs.Name, "", "失败", msg, ""
                GoTo NextRow
            End If

            If Not FileExists(ResolveWorkbookPath(tgtWbPath)) Then
                warnCount = warnCount + 1
                RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWs.Name, tgtSheetName, "提示", "目标工作簿不存在，3.11.5执行时将自动创建", ""
            Else
                Set tgtWb = AcquireWorkbookByPath(tgtWbPath, wbCache, openedByCode, msg)
                If tgtWb Is Nothing Then
                    failCount = failCount + 1
                    RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWs.Name, "", "失败", "目标工作簿不可打开：" & msg, ""
                    GoTo NextRow
                End If

                Set tgtWs = GetWorksheetByName(tgtWb, tgtSheetName)
                If tgtWs Is Nothing Then
                    warnCount = warnCount + 1
                    RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWs.Name, tgtSheetName, "提示", "目标工作表不存在，执行追加时将自动新建", ""
                Else
                    If Not HeadersCompatibleForAppendCheck(srcWs, tgtWs) Then
                        warnCount = warnCount + 1
                        RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWs.Name, tgtWs.Name, "提示", "源/目标表头不一致，3.11.5会跳过", ""
                    End If
                End If
            End If
        End If

        If execMode <> "" Then
            If execMode <> "1" And execMode <> "2" And execMode <> "3" Then
                warnCount = warnCount + 1
                RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWs.Name, "", "提示", "执行模式建议为1/2/3（其他值将按默认处理）", ""
            End If
        End If

        okCount = okCount + 1
        RunLog_WriteRow PRECHECK_LOG_KEY, "校验", "第" & r & "行", srcWs.Name, IIf(hasTarget, tgtSheetName, ""), "成功", "通过", ""
NextRow:
        Set srcWs = Nothing
        Set tgtWs = Nothing
        Set srcWb = Nothing
        Set tgtWb = Nothing
        Set dedupeCols = Nothing
    Next r

    CloseOpenedWorkbooks wbCache, openedByCode
    precheckMsg = "按配置预校验完成。" & vbCrLf & _
                  "通过：" & okCount & vbCrLf & _
                  "提示：" & warnCount & vbCrLf & _
                  "失败：" & failCount & vbCrLf & _
                  "未启用跳过：" & skipCount
    RunLog_WriteRow PRECHECK_LOG_KEY, "结束", CStr(okCount), CStr(warnCount), CStr(failCount), "完成", "跳过=" & skipCount, CStr(Round(Timer - t0, 2))
    MsgBox precheckMsg, vbInformation, "按配置预校验"
    Exit Sub

FailHandler:
    CloseOpenedWorkbooks wbCache, openedByCode
    RunLog_WriteRow PRECHECK_LOG_KEY, "结束", "", "", "", "失败", CStr(Err.Number) & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "预校验失败：" & CStr(Err.Number) & " " & Err.Description, vbCritical, "按配置预校验"
End Sub

Private Function ValidateTargetPathForPrecheck(ByVal rawPath As String, ByRef messageText As String) As Boolean
    Dim resolvedPath As String
    Dim parentPath As String
    Dim p As Long

    messageText = ""
    resolvedPath = ResolveWorkbookPath(rawPath)
    If Len(resolvedPath) = 0 Then
        messageText = "目标工作簿路径为空"
        Exit Function
    End If
    If IsDirectoryPath(resolvedPath) Then
        messageText = "目标工作簿路径是文件夹"
        Exit Function
    End If
    If Not IsSupportedWorkbookFilePath(resolvedPath) Then
        messageText = "目标工作簿文件类型不支持"
        Exit Function
    End If

    p = InStrRev(resolvedPath, "\")
    If p > 1 Then
        parentPath = Left$(resolvedPath, p - 1)
        If Not IsDirectoryPath(parentPath) Then
            messageText = "目标工作簿所在目录不存在"
            Exit Function
        End If
    End If

    ValidateTargetPathForPrecheck = True
End Function

Private Function HeadersCompatibleForAppendCheck(ByVal srcWs As Worksheet, ByVal tgtWs As Worksheet) As Boolean
    Dim tgtLastRow As Long
    Dim srcLastCol As Long
    Dim tgtLastCol As Long
    Dim c As Long
    Dim srcHeader As String
    Dim tgtHeader As String

    tgtLastRow = GetLastUsedRow(tgtWs)
    If tgtLastRow < 1 Then
        HeadersCompatibleForAppendCheck = True
        Exit Function
    End If

    srcLastCol = GetLastUsedColumn(srcWs)
    tgtLastCol = GetLastUsedColumn(tgtWs)
    If srcLastCol <= 0 Then Exit Function
    If tgtLastCol < srcLastCol Then Exit Function

    For c = 1 To srcLastCol
        srcHeader = NormalizeText(srcWs.Cells(1, c).Value2)
        tgtHeader = NormalizeText(tgtWs.Cells(1, c).Value2)
        If StrComp(srcHeader, tgtHeader, vbTextCompare) <> 0 Then Exit Function
    Next c

    HeadersCompatibleForAppendCheck = True
End Function

Private Function FindConfigSheet() As Worksheet
    On Error Resume Next
    Set FindConfigSheet = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    If FindConfigSheet Is Nothing Then
        Set FindConfigSheet = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME_LEGACY)
    End If
    On Error GoTo 0
End Function

Private Function GetOrCreateConfigSheet() As Worksheet
    Dim ws As Worksheet

    Set ws = FindConfigSheet()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = CONFIG_SHEET_NAME
    ElseIf StrComp(ws.Name, CONFIG_SHEET_NAME_LEGACY, vbTextCompare) = 0 Then
        On Error Resume Next
        ws.Name = CONFIG_SHEET_NAME
        On Error GoTo 0
    End If
    Set GetOrCreateConfigSheet = ws
End Function

Private Sub InitConfigHeader(ByVal ws As Worksheet)
    ws.cells(1, 1).value = "是否启用"
    ws.cells(1, 2).value = "源数据工作簿"
    ws.cells(1, 3).value = "源数据工作表"
    ws.cells(1, 4).value = "标识列序号"
    ws.cells(1, 5).value = "目标工作簿"
    ws.cells(1, 6).value = "目标工作表"
    ws.cells(1, 7).value = "执行模式"
    ws.cells(1, 8).value = "备注"

    ws.rows(1).Font.Bold = True
    ws.rows(1).Interior.Color = RGB(221, 235, 247)

    SetHeaderComment ws.cells(1, 7), "执行模式：1=正常执行；2=仅校验不写入；3=备份后执行。"
End Sub

Private Sub WriteConfigExample(ByVal ws As Worksheet)
    If GetLastUsedRow(ws) > 1 Then Exit Sub

    ws.cells(2, 1).value = "N"
    ws.cells(2, 2).value = "C:\Users\AZI\Desktop\demo_source.xlsx"
    ws.cells(2, 3).value = "源数据"
    ws.cells(2, 4).value = "1;2;5"
    ws.cells(2, 5).value = "C:\Users\AZI\Desktop\demo_target.xlsx"
    ws.cells(2, 6).value = "汇总结果"
    ws.cells(2, 7).value = "1"
    ws.cells(2, 8).value = "示例：按第1、2、5列联合去重。留空则默认全列。"
End Sub

Private Sub SetHeaderComment(ByVal targetCell As Range, ByVal commentText As String)
    On Error Resume Next
    If targetCell.Comment Is Nothing Then
        targetCell.AddComment commentText
    Else
        targetCell.Comment.text text:=commentText
    End If
    On Error GoTo 0
End Sub

Private Function MarkDuplicateRowsByIndexes(ByVal ws As Worksheet, ByVal dedupeCols As Collection) As Long
    Dim seen As Object
    Dim firstCol As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim rowOffset As Long
    Dim rowKey As String
    Dim dataRange As Range
    Dim dataArr As Variant
    Dim relCols As Collection
    Dim isNotBlank As Boolean

    If dedupeCols Is Nothing Then Exit Function
    If dedupeCols.Count = 0 Then Exit Function

    firstCol = GetFirstUsedColumn(ws)
    lastCol = GetLastUsedColumn(ws)
    lastRow = GetLastUsedRow(ws)
    If lastCol < firstCol Or lastRow < DATA_START_ROW Then Exit Function

    Set relCols = BuildRelativeIndexCollection(dedupeCols, firstCol, lastCol)
    If relCols Is Nothing Then Exit Function
    If relCols.Count = 0 Then Exit Function

    Set dataRange = ws.Range(ws.Cells(DATA_START_ROW, firstCol), ws.Cells(lastRow, lastCol))
    If dataRange.Cells.CountLarge = 1 Then
        ReDim dataArr(1 To 1, 1 To 1)
        dataArr(1, 1) = dataRange.Value2
    Else
        dataArr = dataRange.Value2
    End If

    Set seen = CreateObject("Scripting.Dictionary")
    For rowOffset = 1 To UBound(dataArr, 1)
        rowKey = ""
        isNotBlank = BuildRowKeyOrBlankFromArray(dataArr, rowOffset, relCols, rowKey)
        If isNotBlank Then
            If seen.Exists(rowKey) Then
                MarkDuplicateRow ws, DATA_START_ROW + rowOffset - 1, firstCol, lastCol
                MarkDuplicateRowsByIndexes = MarkDuplicateRowsByIndexes + 1
            Else
                seen.Add rowKey, True
            End If
        End If
    Next rowOffset
End Function

Private Function BuildRelativeIndexCollection(ByVal absCols As Collection, ByVal firstCol As Long, ByVal lastCol As Long) As Collection
    Dim result As Collection
    Dim idx As Variant
    Dim absCol As Long
    Dim relCol As Long

    If absCols Is Nothing Then Exit Function
    Set result = New Collection
    For Each idx In absCols
        absCol = CLng(idx)
        If absCol >= firstCol And absCol <= lastCol Then
            relCol = absCol - firstCol + 1
            AddUniqueLongToCollection result, relCol
        End If
    Next idx
    If result.Count > 0 Then Set BuildRelativeIndexCollection = result
End Function

Private Function BuildRowKeyOrBlankFromArray(ByRef dataArr As Variant, ByVal rowOffset As Long, ByVal relCols As Collection, ByRef outKey As String) As Boolean
    Dim idx As Variant
    Dim txt As String

    outKey = ""
    For Each idx In relCols
        txt = NormalizeText(dataArr(rowOffset, CLng(idx)))
        outKey = outKey & KEY_SEP & txt
        If Len(txt) > 0 Then
            BuildRowKeyOrBlankFromArray = True
        End If
    Next idx
End Function

Private Function ParseIndexCollection(ByVal rawValue As Variant) As Collection
    Dim txt As String
    Dim tokens() As String
    Dim token As Variant
    Dim result As Collection
    Dim colIndex As Long

    txt = NormalizeTokenSeparators(NormalizeText(rawValue))
    If Len(txt) = 0 Then Exit Function

    tokens = Split(txt, ";")
    Set result = New Collection
    For Each token In tokens
        colIndex = ParseColumnIndex(Trim$(CStr(token)))
        If colIndex > 0 Then AddUniqueLongToCollection result, colIndex
    Next token

    If result.count > 0 Then Set ParseIndexCollection = result
End Function

Private Function FilterColumnsByWorksheet(ByVal sourceCols As Collection, ByVal ws As Worksheet) As Collection
    Dim result As Collection
    Dim lastCol As Long
    Dim idx As Variant

    If sourceCols Is Nothing Then Exit Function
    lastCol = GetLastUsedColumn(ws)
    If lastCol < 1 Then Exit Function

    Set result = New Collection
    For Each idx In sourceCols
        If CLng(idx) >= 1 And CLng(idx) <= lastCol Then
            AddUniqueLongToCollection result, CLng(idx)
        End If
    Next idx

    If result.count > 0 Then Set FilterColumnsByWorksheet = result
End Function

Private Function BuildAllUsedColumnCollection(ByVal ws As Worksheet) As Collection
    Dim result As Collection
    Dim firstCol As Long
    Dim lastCol As Long
    Dim c As Long

    firstCol = GetFirstUsedColumn(ws)
    lastCol = GetLastUsedColumn(ws)
    If lastCol < firstCol Then Exit Function

    Set result = New Collection
    For c = firstCol To lastCol
        AddUniqueLongToCollection result, c
    Next c
    If result.count > 0 Then Set BuildAllUsedColumnCollection = result
End Function

Private Function BuildRowKeyByColumns(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndexes As Collection) As String
    Dim idx As Variant
    Dim parts As String

    For Each idx In colIndexes
        parts = parts & KEY_SEP & NormalizeText(ws.cells(rowIndex, CLng(idx)).Value2)
    Next idx
    BuildRowKeyByColumns = parts
End Function

Private Function RowIsBlankByColumns(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndexes As Collection) As Boolean
    Dim idx As Variant
    For Each idx In colIndexes
        If Len(NormalizeText(ws.cells(rowIndex, CLng(idx)).Value2)) > 0 Then
            Exit Function
        End If
    Next idx
    RowIsBlankByColumns = True
End Function

Private Sub MarkDuplicateRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal firstCol As Long, ByVal lastCol As Long)
    ws.Range(ws.cells(rowIndex, firstCol), ws.cells(rowIndex, lastCol)).Interior.Color = RGB(255, 199, 206)
End Sub

Private Function AcquireWorkbookByPath(ByVal rawPath As String, ByVal wbCache As Object, ByVal openedByCode As Object, ByRef messageText As String) As Workbook
    Dim resolvedPath As String
    Dim wb As Workbook
    Dim openWb As Workbook

    messageText = ""
    resolvedPath = ResolveWorkbookPath(rawPath)
    If Len(resolvedPath) = 0 Then
        messageText = "源工作簿路径为空"
        Exit Function
    End If
    If IsDirectoryPath(resolvedPath) Then
        messageText = "源工作簿路径是文件夹"
        Exit Function
    End If
    If Not IsSupportedWorkbookFilePath(resolvedPath) Then
        messageText = "源工作簿文件类型不支持"
        Exit Function
    End If
    If Not FileExists(resolvedPath) Then
        messageText = "源工作簿不存在"
        Exit Function
    End If

    If wbCache.Exists(resolvedPath) Then
        Set AcquireWorkbookByPath = wbCache(resolvedPath)
        Exit Function
    End If

    Set openWb = FindOpenWorkbookByFullName(resolvedPath)
    If Not openWb Is Nothing Then
        wbCache.Add resolvedPath, openWb
        openedByCode.Add resolvedPath, False
        Set AcquireWorkbookByPath = openWb
        Exit Function
    End If

    On Error GoTo OpenFail
    Set wb = Workbooks.Open(resolvedPath, ReadOnly:=False, UpdateLinks:=0, AddToMru:=False)
    wbCache.Add resolvedPath, wb
    openedByCode.Add resolvedPath, True
    Set AcquireWorkbookByPath = wb
    Exit Function

OpenFail:
    messageText = CStr(Err.Number) & " " & Err.Description
End Function

Private Sub SaveModifiedWorkbooks(ByVal wbCache As Object, ByVal modified As Object, ByVal openedByCode As Object)
    Dim key As Variant
    Dim wb As Workbook
    Dim shouldSave As Boolean

    If wbCache Is Nothing Then Exit Sub
    If modified Is Nothing Then Exit Sub
    If openedByCode Is Nothing Then Exit Sub

    For Each key In modified.keys
        If wbCache.Exists(CStr(key)) Then
            Set wb = wbCache(CStr(key))
            If Not wb Is Nothing Then
                If Not wb.ReadOnly Then
                    shouldSave = False
                    If openedByCode.Exists(CStr(key)) Then
                        shouldSave = CBool(openedByCode(CStr(key)))
                    End If
                    If Not shouldSave Then GoTo NextModifiedWorkbook
                    On Error Resume Next
                    wb.Save
                    On Error GoTo 0
                End If
            End If
        End If
NextModifiedWorkbook:
    Next key
End Sub

Private Sub CloseOpenedWorkbooks(ByVal wbCache As Object, ByVal openedByCode As Object)
    Dim key As Variant
    Dim wb As Workbook

    If wbCache Is Nothing Then Exit Sub
    If openedByCode Is Nothing Then Exit Sub

    For Each key In wbCache.keys
        If openedByCode.Exists(CStr(key)) Then
            If CBool(openedByCode(CStr(key))) Then
                Set wb = wbCache(CStr(key))
                If Not wb Is Nothing Then
                    On Error Resume Next
                    wb.Close saveChanges:=False
                    On Error GoTo 0
                End If
            End If
        End If
    Next key
End Sub

Private Sub MarkModifiedPath(ByVal modified As Object, ByVal resolvedPath As String)
    If modified Is Nothing Then Exit Sub
    If Len(resolvedPath) = 0 Then Exit Sub
    If Not modified.Exists(resolvedPath) Then modified.Add resolvedPath, True
End Sub

Private Function GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function FindOpenWorkbookByFullName(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.fullName, workbookPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
End Function

Private Function ResolveWorkbookPath(ByVal workbookPath As String) As String
    Dim txt As String

    txt = NormalizeText(workbookPath)
    If Len(txt) = 0 Then Exit Function

    If Left$(txt, 2) = "\\" Or (Len(txt) >= 2 And Mid$(txt, 2, 1) = ":") Then
        ResolveWorkbookPath = txt
    Else
        ResolveWorkbookPath = ThisWorkbook.path & "\" & txt
    End If

    Do While Len(ResolveWorkbookPath) > 0 And Right$(ResolveWorkbookPath, 1) = "\"
        ResolveWorkbookPath = Left$(ResolveWorkbookPath, Len(ResolveWorkbookPath) - 1)
    Loop
End Function

Private Function IsSupportedWorkbookFilePath(ByVal filePath As String) As Boolean
    Dim ext As String
    Dim dotPos As Long

    filePath = NormalizeText(filePath)
    If Len(filePath) = 0 Then Exit Function
    dotPos = InStrRev(filePath, ".")
    If dotPos <= 0 Then Exit Function

    ext = LCase$(Mid$(filePath, dotPos + 1))
    Select Case ext
        Case "xls", "xlsx", "xlsm", "xlsb", "csv"
            IsSupportedWorkbookFilePath = True
    End Select
End Function

Private Function IsDirectoryPath(ByVal pathText As String) As Boolean
    Dim attrValue As Long

    pathText = NormalizeText(pathText)
    If Len(pathText) = 0 Then Exit Function

    On Error Resume Next
    attrValue = GetAttr(pathText)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    IsDirectoryPath = ((attrValue And vbDirectory) = vbDirectory)
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    Dim txt As String

    txt = NormalizeText(filePath)
    If Len(txt) = 0 Then Exit Function
    If IsDirectoryPath(txt) Then Exit Function

    On Error Resume Next
    FileExists = (Len(Dir(txt, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) > 0)
    On Error GoTo 0
End Function

Private Function ParseColumnIndex(ByVal textValue As String) As Long
    Dim txt As String
    Dim i As Long
    Dim ch As String
    Dim result As Long

    txt = UCase$(NormalizeText(textValue))
    If Len(txt) = 0 Then Exit Function

    If IsNumeric(txt) Then
        ParseColumnIndex = CLng(txt)
        Exit Function
    End If

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch < "A" Or ch > "Z" Then
            ParseColumnIndex = 0
            Exit Function
        End If
        result = result * 26 + (Asc(ch) - Asc("A") + 1)
    Next i
    ParseColumnIndex = result
End Function

Private Sub AddUniqueLongToCollection(ByVal items As Collection, ByVal valueToAdd As Long)
    Dim itm As Variant

    If items Is Nothing Then Exit Sub
    For Each itm In items
        If CLng(itm) = valueToAdd Then Exit Sub
    Next itm
    items.Add valueToAdd
End Sub

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set lastCell = ws.cells.Find(What:="*", After:=ws.cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = lastCell.row
    End If
End Function

Private Function GetLastUsedColumn(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set lastCell = ws.cells.Find(What:="*", After:=ws.cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        GetLastUsedColumn = 0
    Else
        GetLastUsedColumn = lastCell.Column
    End If
End Function

Private Function GetFirstUsedColumn(ByVal ws As Worksheet) As Long
    Dim usedRange As Range

    If ws Is Nothing Then Exit Function
    On Error Resume Next
    Set usedRange = ws.usedRange
    On Error GoTo 0

    If usedRange Is Nothing Then
        GetFirstUsedColumn = 1
    Else
        GetFirstUsedColumn = usedRange.Column
    End If
    If GetFirstUsedColumn < 1 Then GetFirstUsedColumn = 1
End Function

Private Function IsTruthyValue(ByVal valueIn As Variant) As Boolean
    Dim txt As String

    txt = UCase$(NormalizeText(valueIn))
    If Len(txt) = 0 Then
        IsTruthyValue = True
    ElseIf txt = "Y" Or txt = "YES" Or txt = "TRUE" Or txt = "1" Then
        IsTruthyValue = True
    ElseIf txt = "N" Or txt = "NO" Or txt = "FALSE" Or txt = "0" Then
        IsTruthyValue = False
    Else
        IsTruthyValue = True
    End If
End Function

Private Function NormalizeTokenSeparators(ByVal tokenText As String) As String
    Dim txt As String

    txt = tokenText
    txt = Replace(txt, "；", ";")
    txt = Replace(txt, "，", ";")
    txt = Replace(txt, "、", ";")
    txt = Replace(txt, ",", ";")
    txt = Replace(txt, vbTab, ";")
    txt = Replace(txt, " ", ";")
    Do While InStr(txt, ";;") > 0
        txt = Replace(txt, ";;", ";")
    Loop
    NormalizeTokenSeparators = txt
End Function

Private Function NormalizeText(ByVal rawValue As Variant) As String
    Dim txt As String

    If IsError(rawValue) Or IsEmpty(rawValue) Then Exit Function
    txt = CStr(rawValue)
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Replace(txt, ChrW(&H3000), " ")
    txt = Replace(txt, Chr(160), " ")
    txt = Trim$(txt)

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop
    NormalizeText = txt
End Function

Private Sub CaptureAppState(ByRef prevScreenUpdating As Boolean, ByRef prevDisplayAlerts As Boolean, _
                            ByRef prevEnableEvents As Boolean, ByRef prevCalc As XlCalculation)
    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation
End Sub

Private Sub BeginFastMode()
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub RestoreAppState(ByVal prevScreenUpdating As Boolean, ByVal prevDisplayAlerts As Boolean, _
                            ByVal prevEnableEvents As Boolean, ByVal prevCalc As XlCalculation)
    Application.ScreenUpdating = prevScreenUpdating
    Application.DisplayAlerts = prevDisplayAlerts
    Application.EnableEvents = prevEnableEvents
    Application.Calculation = prevCalc
End Sub
