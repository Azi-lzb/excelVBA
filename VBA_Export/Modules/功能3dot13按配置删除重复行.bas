Attribute VB_Name = "功能3dot13按配置删除重复行"
Option Explicit

Private Const CONFIG_SHEET_NAME As String = "去重追加数据配置"
Private Const CONFIG_SHEET_NAME_LEGACY As String = "按配置查重"
Private Const DATA_START_ROW As Long = 2
Private Const KEY_SEP As String = "|#|"

Public Sub ExecuteDedupDeleteByConfig()
    Application.Run "功能3dot13_按配置删除重复行"
End Sub

Public Sub 功能3dot13_按配置删除重复行()
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
    Dim deletedCount As Long
    Dim totalDeleted As Long
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
        MsgBox "未找到配置表【按配置查重】。请先执行 6.11 初始化按配置查重。", vbExclamation, "按配置删除重复行"
        Exit Sub
    End If

    lastRow = GetLastUsedRow(wsCfg)
    If lastRow < 2 Then
        MsgBox "配置表为空，请先填写配置。", vbExclamation, "按配置删除重复行"
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

        deletedCount = DeleteDuplicateRowsByIndexes(ws, dedupeCols)
        If deletedCount > 0 Then
            MarkModifiedPath modified, NormalizeText(wb.fullName)
        End If
        totalDeleted = totalDeleted + deletedCount
        hitTask = hitTask + 1

nextRow:
    Next r

    SaveModifiedWorkbooks wbCache, modified
    CloseOpenedWorkbooks wbCache, openedByCode
    RestoreAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc

    MsgBox "按配置删除重复完成。" & vbCrLf & _
           "执行任务数：" & hitTask & vbCrLf & _
           "跳过任务数：" & skipTask & vbCrLf & _
           "删除重复行数：" & totalDeleted, vbInformation, "按配置删除重复行"
    Exit Sub

FailHandler:
    SaveModifiedWorkbooks wbCache, modified
    CloseOpenedWorkbooks wbCache, openedByCode
    RestoreAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc
    MsgBox "执行失败：" & CStr(Err.Number) & " " & Err.Description, vbCritical, "按配置删除重复行"
End Sub

Private Function DeleteDuplicateRowsByIndexes(ByVal ws As Worksheet, ByVal dedupeCols As Collection) As Long
    Dim seen As Object
    Dim firstCol As Long
    Dim lastCol As Long
    Dim lastRow As Long
    Dim rowOffset As Long
    Dim rowKey As String
    Dim rowsToDelete As Collection
    Dim dataRange As Range
    Dim dataArr As Variant
    Dim relCols As Collection
    Dim isNotBlank As Boolean

    If dedupeCols Is Nothing Then Exit Function
    If dedupeCols.Count = 0 Then Exit Function

    firstCol = GetFirstUsedColumn(ws)
    lastCol = GetLastUsedColumn(ws)
    lastRow = GetLastUsedRow(ws)
    If lastRow < DATA_START_ROW Then Exit Function
    If lastCol < firstCol Then Exit Function

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
    Set rowsToDelete = New Collection

    For rowOffset = 1 To UBound(dataArr, 1)
        rowKey = ""
        isNotBlank = BuildRowKeyOrBlankFromArray(dataArr, rowOffset, relCols, rowKey)
        If isNotBlank Then
            If seen.Exists(rowKey) Then
                rowsToDelete.Add (DATA_START_ROW + rowOffset - 1)
            Else
                seen.Add rowKey, True
            End If
        End If
    Next rowOffset

    DeleteDuplicateRowsByIndexes = DeleteRowsByCollectionBatch(ws, rowsToDelete)
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

Private Function FindConfigSheet() As Worksheet
    On Error Resume Next
    Set FindConfigSheet = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    If FindConfigSheet Is Nothing Then
        Set FindConfigSheet = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME_LEGACY)
    End If
    On Error GoTo 0
End Function

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
                    wb.Save
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
