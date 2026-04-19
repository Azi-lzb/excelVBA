Attribute VB_Name = "功能3dot15按配置去重追加到目标"
Option Explicit

Private Const CONFIG_SHEET_NAME As String = "去重追加数据配置"
Private Const CONFIG_SHEET_NAME_LEGACY As String = "按配置查重"
Private Const TASK_LOG_SHEET_NAME As String = "去重追加任务统计"
Private Const DATA_START_ROW As Long = 2
Private Const APPEND_CHUNK_ROWS As Long = 50000
Private Const KEY_SEP As String = "|#|"
Private Const EXEC_MODE_NORMAL As String = "1"
Private Const EXEC_MODE_VALIDATE_ONLY As String = "2"
Private Const EXEC_MODE_BACKUP_THEN_RUN As String = "3"

Public Sub ExecuteDedupAppendByConfig()
    Application.Run "功能3dot15_按配置去重追加到目标"
End Sub

Public Sub 功能3dot15_按配置去重追加到目标()
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
    Dim appendedRows As Long
    Dim deletedRows As Long
    Dim totalAppended As Long
    Dim totalDeleted As Long
    Dim hitTask As Long
    Dim skipTask As Long
    Dim validateOnlyTask As Long
    Dim backupRunTask As Long
    Dim headerMismatchSkip As Long
    Dim msg As String
    Dim taskStatus As String
    Dim taskDetail As String
    Dim wbCache As Object
    Dim openedByCode As Object
    Dim modified As Object
    Dim targetKeyCache As Object
    Dim taskLogs As Collection
    Dim existingKeys As Object
    Dim keyCacheId As String
    Dim appendStartRow As Long
    Dim appendEndRow As Long
    Dim dedupeStartRow As Long
    Dim dedupeEndRow As Long
    Dim tgtResolvedPath As String
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation
    Dim sourceGroups As Object
    Dim sourceOrder As Collection
    Dim sourceGroupKey As Variant
    Dim groupedTasks As Collection
    Dim taskItem As Variant
    Dim taskRowNo As Long
    Dim processedTasks As Long

    Set wsCfg = FindConfigSheet()
    If wsCfg Is Nothing Then
        MsgBox "未找到配置表【按配置查重】。请先执行 6.11 初始化按配置查重。", vbExclamation, "按配置去重追加到目标"
        Exit Sub
    End If

    lastRow = GetLastUsedRow(wsCfg)
    If lastRow < 2 Then
        MsgBox "配置表为空，请先填写配置。", vbExclamation, "按配置去重追加到目标"
        Exit Sub
    End If

    Set wbCache = CreateObject("Scripting.Dictionary")
    Set openedByCode = CreateObject("Scripting.Dictionary")
    Set modified = CreateObject("Scripting.Dictionary")
    Set targetKeyCache = CreateObject("Scripting.Dictionary")
    Set taskLogs = New Collection
    Set sourceGroups = CreateObject("Scripting.Dictionary")
    Set sourceOrder = New Collection

    CaptureAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc
    BeginFastMode
    On Error GoTo FailHandler

    For r = 2 To lastRow
        execMode = NormalizeExecMode(wsCfg.Cells(r, 7).Value2)
        enabled = IsTruthyValue(wsCfg.Cells(r, 1).Value2)
        srcWbPath = NormalizeText(wsCfg.Cells(r, 2).Value2)
        srcSheetName = NormalizeText(wsCfg.Cells(r, 3).Value2)
        dedupeColsText = NormalizeText(wsCfg.Cells(r, 4).Value2)
        tgtWbPath = NormalizeText(wsCfg.Cells(r, 5).Value2)
        tgtSheetName = NormalizeText(wsCfg.Cells(r, 6).Value2)

        If Not enabled Then
            skipTask = skipTask + 1
            taskLogs.Add Array(Format$(Now, "yyyy-mm-dd hh:nn:ss"), CStr(r), execMode, srcWbPath, srcSheetName, tgtWbPath, tgtSheetName, "跳过", "0", "0", "未启用")
            GoTo nextConfigRow
        End If

        If Len(srcWbPath) = 0 Or Len(srcSheetName) = 0 Or Len(tgtWbPath) = 0 Or Len(tgtSheetName) = 0 Then
            skipTask = skipTask + 1
            taskLogs.Add Array(Format$(Now, "yyyy-mm-dd hh:nn:ss"), CStr(r), execMode, srcWbPath, srcSheetName, tgtWbPath, tgtSheetName, "跳过", "0", "0", "关键配置为空")
            GoTo nextConfigRow
        End If

        sourceGroupKey = LCase$(ResolveWorkbookPath(srcWbPath)) & "||" & LCase$(srcSheetName)
        If Len(sourceGroupKey) = 2 Then sourceGroupKey = LCase$(srcWbPath) & "||" & LCase$(srcSheetName)
        If Not sourceGroups.Exists(sourceGroupKey) Then
            Set groupedTasks = New Collection
            sourceGroups.Add sourceGroupKey, groupedTasks
            sourceOrder.Add sourceGroupKey
        End If
        sourceGroups(sourceGroupKey).Add Array(CLng(r), srcWbPath, srcSheetName, dedupeColsText, tgtWbPath, tgtSheetName, execMode)
nextConfigRow:
    Next r

    For Each sourceGroupKey In sourceOrder
        Set groupedTasks = sourceGroups(CStr(sourceGroupKey))
        If groupedTasks Is Nothing Then GoTo nextSourceGroup
        If groupedTasks.Count = 0 Then GoTo nextSourceGroup

        taskItem = groupedTasks(1)
        srcWbPath = CStr(taskItem(1))
        srcSheetName = CStr(taskItem(2))
        Set srcWb = AcquireWorkbookByPath(srcWbPath, False, wbCache, openedByCode, msg)
        If srcWb Is Nothing Then
            For Each taskItem In groupedTasks
                taskRowNo = CLng(taskItem(0))
                dedupeColsText = CStr(taskItem(3))
                tgtWbPath = CStr(taskItem(4))
                tgtSheetName = CStr(taskItem(5))
                execMode = CStr(taskItem(6))
                skipTask = skipTask + 1
                taskLogs.Add Array(Format$(Now, "yyyy-mm-dd hh:nn:ss"), CStr(taskRowNo), execMode, srcWbPath, srcSheetName, tgtWbPath, tgtSheetName, "失败", "0", "0", "打开源工作簿失败：" & msg)
            Next taskItem
            GoTo nextSourceGroup
        End If

        Set srcWs = GetWorksheetByName(srcWb, srcSheetName)
        If srcWs Is Nothing Then
            For Each taskItem In groupedTasks
                taskRowNo = CLng(taskItem(0))
                dedupeColsText = CStr(taskItem(3))
                tgtWbPath = CStr(taskItem(4))
                tgtSheetName = CStr(taskItem(5))
                execMode = CStr(taskItem(6))
                skipTask = skipTask + 1
                taskLogs.Add Array(Format$(Now, "yyyy-mm-dd hh:nn:ss"), CStr(taskRowNo), execMode, srcWbPath, srcSheetName, tgtWbPath, tgtSheetName, "跳过", "0", "0", "源工作表不存在")
            Next taskItem
            GoTo nextSourceGroup
        End If

        For Each taskItem In groupedTasks
            processedTasks = processedTasks + 1
            taskRowNo = CLng(taskItem(0))
            srcWbPath = CStr(taskItem(1))
            srcSheetName = CStr(taskItem(2))
            dedupeColsText = CStr(taskItem(3))
            tgtWbPath = CStr(taskItem(4))
            tgtSheetName = CStr(taskItem(5))
            execMode = CStr(taskItem(6))
            appendedRows = 0
            deletedRows = 0
            taskStatus = "成功"
            taskDetail = ""

            Set tgtWb = AcquireWorkbookByPath(tgtWbPath, True, wbCache, openedByCode, msg)
            If tgtWb Is Nothing Then
                taskStatus = "失败"
                taskDetail = "打开目标工作簿失败：" & msg
                skipTask = skipTask + 1
                GoTo writeGroupedTaskLog
            End If
            tgtResolvedPath = ResolveWorkbookPath(tgtWbPath)
            If Len(tgtResolvedPath) = 0 Then tgtResolvedPath = NormalizeText(tgtWbPath)
            Set tgtWs = EnsureWorksheetExists(tgtWb, tgtSheetName)
            If tgtWs Is Nothing Then
                taskStatus = "失败"
                taskDetail = "创建/获取目标工作表失败"
                skipTask = skipTask + 1
                GoTo writeGroupedTaskLog
            End If

            Set dedupeCols = ParseIndexCollection(dedupeColsText)
            If dedupeCols Is Nothing Then
                Set dedupeCols = BuildAllUsedColumnCollection(srcWs)
            ElseIf dedupeCols.Count = 0 Then
                Set dedupeCols = BuildAllUsedColumnCollection(srcWs)
            Else
                Set dedupeCols = FilterColumnsByWorksheet(dedupeCols, srcWs)
                If dedupeCols Is Nothing Then
                    Set dedupeCols = BuildAllUsedColumnCollection(srcWs)
                ElseIf dedupeCols.Count = 0 Then
                    Set dedupeCols = BuildAllUsedColumnCollection(srcWs)
                End If
            End If

            If Not HeadersCompatibleForAppend(srcWs, tgtWs) Then
                taskStatus = "跳过"
                taskDetail = "表头不一致"
                skipTask = skipTask + 1
                headerMismatchSkip = headerMismatchSkip + 1
                GoTo writeGroupedTaskLog
            End If

            If execMode = EXEC_MODE_VALIDATE_ONLY Then
                taskStatus = "仅校验"
                taskDetail = "仅做校验，未写入"
                validateOnlyTask = validateOnlyTask + 1
                hitTask = hitTask + 1
                GoTo writeGroupedTaskLog
            End If

            If execMode = EXEC_MODE_BACKUP_THEN_RUN Then
                If Not BackupWorksheetSnapshot(tgtWs) Then
                    taskStatus = "失败"
                    taskDetail = "备份目标工作表失败"
                    skipTask = skipTask + 1
                    GoTo writeGroupedTaskLog
                End If
                backupRunTask = backupRunTask + 1
            End If

            keyCacheId = BuildTargetCacheKey(tgtWb, tgtWs, dedupeCols)
            Set existingKeys = AcquireExistingKeySetFromCache(targetKeyCache, keyCacheId, tgtWs, dedupeCols)

            appendedRows = AppendSourceToTarget(srcWs, tgtWs, appendStartRow, appendEndRow)
            If appendedRows > 0 Then MarkModifiedPath modified, tgtResolvedPath
            totalAppended = totalAppended + appendedRows

            deletedRows = 0
            If appendedRows > 0 Then
                If Not dedupeCols Is Nothing Then
                    If dedupeCols.Count > 0 Then
                        dedupeStartRow = appendStartRow
                        dedupeEndRow = appendEndRow
                        If dedupeStartRow < DATA_START_ROW Then dedupeStartRow = DATA_START_ROW
                        If dedupeEndRow >= dedupeStartRow Then
                            deletedRows = DeleteDuplicatesInAppendedRangeByIndexes(tgtWs, dedupeCols, dedupeStartRow, dedupeEndRow, existingKeys)
                        End If
                    End If
                End If
            End If
            If deletedRows > 0 Then MarkModifiedPath modified, tgtResolvedPath
            totalDeleted = totalDeleted + deletedRows
            hitTask = hitTask + 1

writeGroupedTaskLog:
            taskLogs.Add Array(Format$(Now, "yyyy-mm-dd hh:nn:ss"), CStr(taskRowNo), execMode, srcWbPath, srcSheetName, tgtWbPath, tgtSheetName, taskStatus, CStr(appendedRows), CStr(deletedRows), taskDetail)
            Set existingKeys = Nothing
            Set dedupeCols = Nothing
            Set tgtWs = Nothing
            Set tgtWb = Nothing
            If (processedTasks Mod 200) = 0 Then DoEvents
        Next taskItem

nextSourceGroup:
        Set srcWs = Nothing
        Set srcWb = Nothing
        Set groupedTasks = Nothing
    Next sourceGroupKey

    SaveModifiedWorkbooks wbCache, modified
    CloseOpenedWorkbooks wbCache, openedByCode
    RestoreAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc
    FlushTaskStatLogs taskLogs

    MsgBox "按配置去重追加完成。" & vbCrLf & _
           "执行任务数：" & hitTask & vbCrLf & _
           "仅校验任务数：" & validateOnlyTask & vbCrLf & _
           "备份后执行任务数：" & backupRunTask & vbCrLf & _
           "跳过任务数：" & skipTask & vbCrLf & _
           "表头不一致跳过：" & headerMismatchSkip & vbCrLf & _
           "追加行数：" & totalAppended & vbCrLf & _
           "目标去重删除行数：" & totalDeleted, vbInformation, "按配置去重追加到目标"
    Exit Sub

FailHandler:
    SaveModifiedWorkbooks wbCache, modified
    CloseOpenedWorkbooks wbCache, openedByCode
    RestoreAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc
    FlushTaskStatLogs taskLogs
    MsgBox "执行失败：" & CStr(Err.Number) & " " & Err.Description, vbCritical, "按配置去重追加到目标"
End Sub

Private Function FindConfigSheet() As Worksheet
    On Error Resume Next
    Set FindConfigSheet = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    If FindConfigSheet Is Nothing Then
        Set FindConfigSheet = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME_LEGACY)
    End If
    On Error GoTo 0
End Function

Private Function AppendSourceToTarget(ByVal srcWs As Worksheet, ByVal tgtWs As Worksheet, ByRef outAppendStartRow As Long, ByRef outAppendEndRow As Long) As Long
    Dim srcLastRow As Long
    Dim srcLastCol As Long
    Dim srcStartRow As Long
    Dim copyRowCount As Long
    Dim tgtLastRow As Long
    Dim tgtStartRow As Long
    Dim tgtWriteRow As Long
    Dim srcCursor As Long
    Dim chunkEnd As Long
    Dim chunkRowCount As Long
    Dim srcRange As Range
    Dim srcArr As Variant

    outAppendStartRow = 0
    outAppendEndRow = 0

    srcLastRow = GetLastUsedRow(srcWs)
    srcLastCol = GetLastUsedColumn(srcWs)
    If srcLastRow <= 0 Or srcLastCol <= 0 Then Exit Function

    tgtLastRow = GetLastUsedRow(tgtWs)
    If tgtLastRow < 1 Then
        srcStartRow = 1
        tgtStartRow = 1
    Else
        srcStartRow = 2
        If srcLastRow < srcStartRow Then Exit Function
        tgtStartRow = tgtLastRow + 1
    End If

    copyRowCount = srcLastRow - srcStartRow + 1
    If copyRowCount <= 0 Then Exit Function

    tgtWriteRow = tgtStartRow
    srcCursor = srcStartRow
    Do While srcCursor <= srcLastRow
        chunkEnd = srcCursor + APPEND_CHUNK_ROWS - 1
        If chunkEnd > srcLastRow Then chunkEnd = srcLastRow
        chunkRowCount = chunkEnd - srcCursor + 1

        Set srcRange = srcWs.Range(srcWs.cells(srcCursor, 1), srcWs.cells(chunkEnd, srcLastCol))
        If srcRange.cells.CountLarge = 1 Then
            ReDim srcArr(1 To 1, 1 To 1)
            srcArr(1, 1) = srcRange.Value2
        Else
            srcArr = srcRange.Value2
        End If

        tgtWs.Range(tgtWs.cells(tgtWriteRow, 1), tgtWs.cells(tgtWriteRow + chunkRowCount - 1, srcLastCol)).Value2 = srcArr
        Erase srcArr
        Set srcRange = Nothing
        tgtWriteRow = tgtWriteRow + chunkRowCount
        srcCursor = chunkEnd + 1
    Loop

    outAppendStartRow = tgtStartRow
    outAppendEndRow = tgtWriteRow - 1
    AppendSourceToTarget = copyRowCount
End Function

Private Function HeadersCompatibleForAppend(ByVal srcWs As Worksheet, ByVal tgtWs As Worksheet) As Boolean
    Dim tgtLastRow As Long
    Dim srcLastCol As Long
    Dim tgtLastCol As Long
    Dim c As Long
    Dim srcHeader As String
    Dim tgtHeader As String

    tgtLastRow = GetLastUsedRow(tgtWs)
    If tgtLastRow < 1 Then
        HeadersCompatibleForAppend = True
        Exit Function
    End If

    srcLastCol = GetLastUsedColumn(srcWs)
    tgtLastCol = GetLastUsedColumn(tgtWs)
    If srcLastCol <= 0 Then Exit Function
    If tgtLastCol < srcLastCol Then Exit Function

    For c = 1 To srcLastCol
        srcHeader = NormalizeText(srcWs.Cells(1, c).Value2)
        tgtHeader = NormalizeText(tgtWs.Cells(1, c).Value2)
        If StrComp(srcHeader, tgtHeader, vbTextCompare) <> 0 Then
            Exit Function
        End If
    Next c

    HeadersCompatibleForAppend = True
End Function

Private Function DeleteDuplicatesInAppendedRangeByIndexes(ByVal ws As Worksheet, ByVal dedupeCols As Collection, ByVal appendStartRow As Long, ByVal appendEndRow As Long, ByVal seen As Object) As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim rowOffset As Long
    Dim rowKey As String
    Dim rowsToDelete As Collection
    Dim dataRange As Range
    Dim dataArr As Variant
    Dim relCols As Collection
    Dim isNotBlank As Boolean

    If ws Is Nothing Then Exit Function
    If dedupeCols Is Nothing Then Exit Function
    If dedupeCols.Count = 0 Then Exit Function
    If appendStartRow < DATA_START_ROW Then Exit Function
    If appendEndRow < appendStartRow Then Exit Function

    If seen Is Nothing Then Set seen = CreateObject("Scripting.Dictionary")

    firstCol = GetFirstUsedColumn(ws)
    lastCol = GetLastUsedColumn(ws)
    If lastCol < firstCol Then Exit Function

    Set relCols = BuildRelativeIndexCollection(dedupeCols, firstCol, lastCol)
    If relCols Is Nothing Then Exit Function
    If relCols.Count = 0 Then Exit Function

    Set dataRange = ws.Range(ws.Cells(appendStartRow, firstCol), ws.Cells(appendEndRow, lastCol))
    If dataRange.Cells.CountLarge = 1 Then
        ReDim dataArr(1 To 1, 1 To 1)
        dataArr(1, 1) = dataRange.Value2
    Else
        dataArr = dataRange.Value2
    End If

    Set rowsToDelete = New Collection
    For rowOffset = 1 To UBound(dataArr, 1)
        rowKey = ""
        isNotBlank = BuildRowKeyOrBlankFromArray(dataArr, rowOffset, relCols, rowKey)
        If isNotBlank Then
            If seen.Exists(rowKey) Then
                rowsToDelete.Add (appendStartRow + rowOffset - 1)
            Else
                seen.Add rowKey, True
            End If
        End If
    Next rowOffset

    DeleteDuplicatesInAppendedRangeByIndexes = DeleteRowsByCollectionBatch(ws, rowsToDelete)
End Function

Private Function BuildExistingKeySet(ByVal ws As Worksheet, ByVal dedupeCols As Collection, ByVal lastRowToScan As Long) As Object
    Dim firstCol As Long
    Dim lastCol As Long
    Dim rowOffset As Long
    Dim rowKey As String
    Dim dataRange As Range
    Dim dataArr As Variant
    Dim relCols As Collection
    Dim seen As Object
    Dim isNotBlank As Boolean

    Set seen = CreateObject("Scripting.Dictionary")
    If ws Is Nothing Then
        Set BuildExistingKeySet = seen
        Exit Function
    End If
    If dedupeCols Is Nothing Then
        Set BuildExistingKeySet = seen
        Exit Function
    End If
    If dedupeCols.Count = 0 Or lastRowToScan < DATA_START_ROW Then
        Set BuildExistingKeySet = seen
        Exit Function
    End If

    firstCol = GetFirstUsedColumn(ws)
    lastCol = GetLastUsedColumn(ws)
    If lastCol < firstCol Then
        Set BuildExistingKeySet = seen
        Exit Function
    End If

    Set relCols = BuildRelativeIndexCollection(dedupeCols, firstCol, lastCol)
    If relCols Is Nothing Then
        Set BuildExistingKeySet = seen
        Exit Function
    End If
    If relCols.Count = 0 Then
        Set BuildExistingKeySet = seen
        Exit Function
    End If

    Set dataRange = ws.Range(ws.Cells(DATA_START_ROW, firstCol), ws.Cells(lastRowToScan, lastCol))
    If dataRange.Cells.CountLarge = 1 Then
        ReDim dataArr(1 To 1, 1 To 1)
        dataArr(1, 1) = dataRange.Value2
    Else
        dataArr = dataRange.Value2
    End If

    For rowOffset = 1 To UBound(dataArr, 1)
        rowKey = ""
        isNotBlank = BuildRowKeyOrBlankFromArray(dataArr, rowOffset, relCols, rowKey)
        If isNotBlank Then
            If Not seen.Exists(rowKey) Then seen.Add rowKey, True
        End If
    Next rowOffset

    Set BuildExistingKeySet = seen
End Function

Private Function AcquireExistingKeySetFromCache(ByVal keyCache As Object, ByVal cacheKey As String, ByVal ws As Worksheet, ByVal dedupeCols As Collection) As Object
    Dim seen As Object

    If keyCache Is Nothing Then
        Set AcquireExistingKeySetFromCache = BuildExistingKeySet(ws, dedupeCols, GetLastUsedRow(ws))
        Exit Function
    End If

    If Len(cacheKey) = 0 Then
        Set AcquireExistingKeySetFromCache = BuildExistingKeySet(ws, dedupeCols, GetLastUsedRow(ws))
        Exit Function
    End If

    If keyCache.Exists(cacheKey) Then
        Set AcquireExistingKeySetFromCache = keyCache(cacheKey)
        Exit Function
    End If

    Set seen = BuildExistingKeySet(ws, dedupeCols, GetLastUsedRow(ws))
    keyCache.Add cacheKey, seen
    Set AcquireExistingKeySetFromCache = seen
End Function

Private Function BuildTargetCacheKey(ByVal tgtWb As Workbook, ByVal tgtWs As Worksheet, ByVal dedupeCols As Collection) As String
    Dim wbKey As String
    Dim wsKey As String
    Dim colKey As String

    If tgtWb Is Nothing Then Exit Function
    If tgtWs Is Nothing Then Exit Function

    wbKey = NormalizeText(tgtWb.fullName)
    wsKey = NormalizeText(tgtWs.Name)
    colKey = BuildDedupeColsSignature(dedupeCols)
    BuildTargetCacheKey = wbKey & "||" & wsKey & "||" & colKey
End Function

Private Function BuildDedupeColsSignature(ByVal dedupeCols As Collection) As String
    Dim arr() As Long
    Dim i As Long
    Dim j As Long
    Dim t As Long
    Dim v As Variant
    Dim sig As String

    If dedupeCols Is Nothing Then Exit Function
    If dedupeCols.count = 0 Then Exit Function

    ReDim arr(1 To dedupeCols.count)
    i = 1
    For Each v In dedupeCols
        arr(i) = CLng(v)
        i = i + 1
    Next v

    For i = 1 To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                t = arr(i)
                arr(i) = arr(j)
                arr(j) = t
            End If
        Next j
    Next i

    For i = 1 To UBound(arr)
        sig = sig & ";" & CStr(arr(i))
    Next i
    BuildDedupeColsSignature = sig
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
        outKey = outKey & "|" & CStr(LenB(txt)) & ":" & txt
        If Len(txt) > 0 Then
            BuildRowKeyOrBlankFromArray = True
        End If
    Next idx
End Function

Private Function NormalizeExecMode(ByVal modeValue As Variant) As String
    Dim txt As String

    txt = NormalizeText(modeValue)
    Select Case txt
        Case EXEC_MODE_VALIDATE_ONLY, EXEC_MODE_BACKUP_THEN_RUN
            NormalizeExecMode = txt
        Case Else
            NormalizeExecMode = EXEC_MODE_NORMAL
    End Select
End Function

Private Function BackupWorksheetSnapshot(ByVal ws As Worksheet) As Boolean
    Dim wb As Workbook
    Dim baseName As String
    Dim stamp As String
    Dim candidate As String
    Dim seq As Long

    If ws Is Nothing Then Exit Function
    Set wb = ws.Parent

    stamp = Format$(Now, "yymmdd_hhnnss")
    baseName = Left$("bak_" & ws.Name & "_" & stamp, 31)
    candidate = baseName
    seq = 1
    Do While Not GetWorksheetByName(wb, candidate) Is Nothing
        candidate = Left$(baseName, 28) & "_" & CStr(seq)
        seq = seq + 1
    Loop

    On Error GoTo BackupFail
    ws.Copy After:=wb.Worksheets(wb.Worksheets.count)
    wb.Worksheets(wb.Worksheets.count).Name = candidate
    BackupWorksheetSnapshot = True
    Exit Function

BackupFail:
    BackupWorksheetSnapshot = False
End Function

Private Sub FlushTaskStatLogs(ByVal logs As Collection)
    Dim ws As Worksheet
    Dim startRow As Long
    Dim i As Long
    Dim item As Variant
    Dim arr() As Variant

    If logs Is Nothing Then Exit Sub
    If logs.count = 0 Then Exit Sub

    Set ws = EnsureTaskLogSheet()
    If ws Is Nothing Then Exit Sub

    startRow = GetLastUsedRow(ws) + 1
    If startRow < 2 Then startRow = 2

    ReDim arr(1 To logs.count, 1 To 11)
    For i = 1 To logs.count
        item = logs(i)
        arr(i, 1) = item(0)   ' 时间
        arr(i, 2) = item(1)   ' 配置行
        arr(i, 3) = item(2)   ' 执行模式
        arr(i, 4) = item(3)   ' 源工作簿
        arr(i, 5) = item(4)   ' 源工作表
        arr(i, 6) = item(5)   ' 目标工作簿
        arr(i, 7) = item(6)   ' 目标工作表
        arr(i, 8) = item(7)   ' 状态
        arr(i, 9) = item(8)   ' 追加行数
        arr(i, 10) = item(9)  ' 删除行数
        arr(i, 11) = item(10) ' 详情
    Next i

    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + logs.count - 1, 11)).Value2 = arr
End Sub

Private Function EnsureTaskLogSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(TASK_LOG_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = TASK_LOG_SHEET_NAME
    End If

    If Len(NormalizeText(ws.Cells(1, 1).Value2)) = 0 Then
        ws.Cells(1, 1).Value = "执行时间"
        ws.Cells(1, 2).Value = "配置行号"
        ws.Cells(1, 3).Value = "执行模式"
        ws.Cells(1, 4).Value = "源工作簿"
        ws.Cells(1, 5).Value = "源工作表"
        ws.Cells(1, 6).Value = "目标工作簿"
        ws.Cells(1, 7).Value = "目标工作表"
        ws.Cells(1, 8).Value = "状态"
        ws.Cells(1, 9).Value = "追加行数"
        ws.Cells(1, 10).Value = "删除行数"
        ws.Cells(1, 11).Value = "详情"
        ws.Rows(1).Font.Bold = True
        ws.Columns("A:K").AutoFit
    End If

    Set EnsureTaskLogSheet = ws
End Function

Private Function DeleteRowsByCollectionBatch(ByVal ws As Worksheet, ByVal rowsToDelete As Collection) As Long
    Const CHUNK_SIZE As Long = 500
    Dim i As Long
    Dim deleteRange As Range
    Dim batchCount As Long

    If rowsToDelete Is Nothing Then Exit Function
    If rowsToDelete.Count = 0 Then Exit Function

    For i = rowsToDelete.Count To 1 Step -1
        If deleteRange Is Nothing Then
            Set deleteRange = ws.Rows(CLng(rowsToDelete(i)))
        Else
            Set deleteRange = Union(deleteRange, ws.Rows(CLng(rowsToDelete(i))))
        End If

        batchCount = batchCount + 1
        If batchCount >= CHUNK_SIZE Then
            DeleteRowsByCollectionBatch = DeleteRowsByCollectionBatch + batchCount
            deleteRange.EntireRow.Delete
            Set deleteRange = Nothing
            batchCount = 0
        End If
    Next i

    If Not deleteRange Is Nothing Then
        DeleteRowsByCollectionBatch = DeleteRowsByCollectionBatch + batchCount
        deleteRange.EntireRow.Delete
    End If
End Function

Private Function AcquireWorkbookByPath(ByVal rawPath As String, ByVal allowCreate As Boolean, ByVal wbCache As Object, ByVal openedByCode As Object, ByRef messageText As String) As Workbook
    Dim resolvedPath As String
    Dim wb As Workbook
    Dim openWb As Workbook
    Dim parentFolder As String

    messageText = ""
    resolvedPath = ResolveWorkbookPath(rawPath)
    If Len(resolvedPath) = 0 Then
        messageText = "工作簿路径为空"
        Exit Function
    End If

    If IsDirectoryPath(resolvedPath) Then
        messageText = "工作簿路径是文件夹"
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

    If FileExists(resolvedPath) Then
        If Not IsSupportedWorkbookFilePath(resolvedPath) Then
            messageText = "工作簿文件类型不支持"
            Exit Function
        End If
        On Error GoTo OpenFail
        Set wb = Workbooks.Open(resolvedPath, ReadOnly:=False, UpdateLinks:=0, AddToMru:=False)
        wbCache.Add resolvedPath, wb
        openedByCode.Add resolvedPath, True
        Set AcquireWorkbookByPath = wb
        Exit Function
    End If

    If Not allowCreate Then
        messageText = "工作簿不存在"
        Exit Function
    End If

    parentFolder = GetParentFolderPath(resolvedPath)
    If Len(parentFolder) > 0 Then
        If Not IsDirectoryPath(parentFolder) Then
            messageText = "目标工作簿上级目录不存在"
            Exit Function
        End If
    End If

    Set wb = Workbooks.Add(xlWBATWorksheet)
    wbCache.Add resolvedPath, wb
    openedByCode.Add resolvedPath, True
    Set AcquireWorkbookByPath = wb
    Exit Function

OpenFail:
    messageText = CStr(Err.Number) & " " & Err.Description
End Function

Private Function EnsureWorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    Set ws = GetWorksheetByName(wb, sheetName)
    If Not ws Is Nothing Then
        Set EnsureWorksheetExists = ws
        Exit Function
    End If

    If wb.Worksheets.count = 1 Then
        Set ws = wb.Worksheets(1)
        If Len(NormalizeText(ws.cells(1, 1).Value2)) = 0 Then
            On Error Resume Next
            ws.Name = sheetName
            On Error GoTo 0
            Set EnsureWorksheetExists = ws
            Exit Function
        End If
    End If

    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
    On Error Resume Next
    ws.Name = sheetName
    On Error GoTo 0
    Set EnsureWorksheetExists = ws
End Function

Private Function GetWorksheetByName(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetByName = wb.Worksheets(sheetName)
    On Error GoTo 0
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

Private Function RowIsBlankByColumns(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndexes As Collection) As Boolean
    Dim idx As Variant
    For Each idx In colIndexes
        If Len(NormalizeText(ws.cells(rowIndex, CLng(idx)).Value2)) > 0 Then
            Exit Function
        End If
    Next idx
    RowIsBlankByColumns = True
End Function

Private Function BuildRowKeyByColumns(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndexes As Collection) As String
    Dim idx As Variant
    Dim parts As String

    For Each idx In colIndexes
        parts = parts & KEY_SEP & NormalizeText(ws.cells(rowIndex, CLng(idx)).Value2)
    Next idx
    BuildRowKeyByColumns = parts
End Function

Private Sub SaveModifiedWorkbooks(ByVal wbCache As Object, ByVal modified As Object)
    Dim key As Variant
    Dim wb As Workbook

    If wbCache Is Nothing Then Exit Sub
    If modified Is Nothing Then Exit Sub

    For Each key In modified.keys
        If wbCache.Exists(CStr(key)) Then
            Set wb = wbCache(CStr(key))
            If Not wb Is Nothing Then
                If Not wb.ReadOnly Then
                    On Error Resume Next
                    If FileExists(CStr(key)) Then
                        wb.Save
                    Else
                        wb.SaveAs fileName:=CStr(key), FileFormat:=GetSaveFileFormat(CStr(key))
                    End If
                    On Error GoTo 0
                End If
            End If
        End If
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

Private Function GetParentFolderPath(ByVal filePath As String) As String
    Dim p As Long
    p = InStrRev(filePath, "\")
    If p > 0 Then GetParentFolderPath = Left$(filePath, p - 1)
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

Private Function GetSaveFileFormat(ByVal workbookPath As String) As Long
    Dim ext As String
    ext = LCase$(Mid$(workbookPath, InStrRev(workbookPath, ".") + 1))
    Select Case ext
        Case "xls": GetSaveFileFormat = 56
        Case "xlsx": GetSaveFileFormat = 51
        Case "xlsm": GetSaveFileFormat = 52
        Case "xlsb": GetSaveFileFormat = 50
        Case Else: GetSaveFileFormat = 51
    End Select
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
