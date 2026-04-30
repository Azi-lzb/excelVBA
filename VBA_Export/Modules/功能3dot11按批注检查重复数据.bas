Attribute VB_Name = "功能3dot11按批注检查重复数据"
Option Explicit

Private Const MARK_SHEET_TEXT As String = "源数据"
Private Const MARK_KEY_TEXT As String = "标识列"
Private Const HEADER_ROW As Long = 1
Private Const DATA_START_ROW As Long = 2
Private Const KEY_SEP As String = "|#|"

Public Sub ExecuteDedupCheckByComment()
    Application.Run "功能3dot11_按批注检查重复数据"
End Sub

Public Sub 功能3dot11_按批注检查重复数据()
    Dim pickedPath As String
    Dim resolvedPath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim openedByCode As Boolean
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation
    Dim hitSheetCount As Long
    Dim dupRowCount As Long
    Dim errNo As Long
    Dim errDesc As String

    pickedPath = PickWorkbookPath()
    If Len(pickedPath) = 0 Then
        MsgBox "已取消选择工作簿。", vbInformation, "按批注检查重复数据"
        Exit Sub
    End If

    resolvedPath = ResolveWorkbookPath(pickedPath)
    If IsDirectoryPath(resolvedPath) Then
        MsgBox "当前选择是文件夹，请选择Excel/CSV文件。", vbExclamation, "按批注检查重复数据"
        Exit Sub
    End If
    If Not IsSupportedWorkbookFilePath(resolvedPath) Then
        MsgBox "仅支持 xls/xlsx/xlsm/xlsb/csv 文件。", vbExclamation, "按批注检查重复数据"
        Exit Sub
    End If
    If Not FileExists(resolvedPath) Then
        MsgBox "文件不存在，请重新选择。", vbExclamation, "按批注检查重复数据"
        Exit Sub
    End If

    Set wb = FindOpenWorkbookByFullName(resolvedPath)
    If wb Is Nothing Then
        On Error Resume Next
        Set wb = Workbooks.Open(resolvedPath, ReadOnly:=False, UpdateLinks:=0, AddToMru:=False)
        errNo = Err.Number
        errDesc = Err.Description
        On Error GoTo 0
        If wb Is Nothing Then
            MsgBox "打开工作簿失败：" & CStr(errNo) & " " & errDesc, vbCritical, "按批注检查重复数据"
            Exit Sub
        End If
        openedByCode = True
    End If

    On Error GoTo FailHandler
    CaptureAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc
    BeginFastMode

    For Each ws In wb.Worksheets
        If CellCommentContains(ws.cells(1, 1), MARK_SHEET_TEXT) Then
            hitSheetCount = hitSheetCount + 1
            dupRowCount = dupRowCount + MarkDuplicateRowsByComment(ws)
        End If
    Next ws

    If openedByCode And Len(wb.path) > 0 And Not wb.ReadOnly Then
        On Error Resume Next
        wb.Save
        On Error GoTo 0
    End If

    RestoreAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc
    If openedByCode Then wb.Close saveChanges:=False

    MsgBox "按批注检查重复完成。" & vbCrLf & _
           "命中Sheet数：" & hitSheetCount & vbCrLf & _
           "重复标红行数：" & dupRowCount, vbInformation, "按批注检查重复数据"
    Exit Sub

FailHandler:
    errNo = Err.Number
    errDesc = Err.Description
    RestoreAppState prevScreenUpdating, prevDisplayAlerts, prevEnableEvents, prevCalc
    If openedByCode Then
        On Error Resume Next
        wb.Close saveChanges:=False
        On Error GoTo 0
    End If
    MsgBox "执行失败：" & CStr(errNo) & " " & errDesc, vbCritical, "按批注检查重复数据"
End Sub

Private Function MarkDuplicateRowsByComment(ByVal ws As Worksheet) As Long
    Dim keyCols As Collection
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

    firstCol = GetFirstUsedColumn(ws)
    lastCol = GetLastUsedColumn(ws)
    lastRow = GetLastUsedRow(ws)
    If lastCol < firstCol Or lastRow < DATA_START_ROW Then Exit Function

    Set keyCols = FindCommentMarkedColumns(ws)
    If keyCols Is Nothing Then
        Set keyCols = BuildColumnCollection(firstCol, lastCol)
    ElseIf keyCols.count = 0 Then
        Set keyCols = BuildColumnCollection(firstCol, lastCol)
    End If
    If keyCols Is Nothing Then Exit Function
    If keyCols.Count = 0 Then Exit Function

    Set relCols = BuildRelativeIndexCollection(keyCols, firstCol, lastCol)
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
                MarkDuplicateRowsByComment = MarkDuplicateRowsByComment + 1
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

Private Function FindCommentMarkedColumns(ByVal ws As Worksheet) As Collection
    Dim result As Collection
    Dim lastCol As Long
    Dim c As Long

    Set result = New Collection
    lastCol = GetLastUsedColumn(ws)
    If lastCol < 1 Then Exit Function

    For c = 1 To lastCol
        If CellCommentContains(ws.cells(HEADER_ROW, c), MARK_KEY_TEXT) Then
            AddUniqueLongToCollection result, c
        End If
    Next c

    If result.count > 0 Then Set FindCommentMarkedColumns = result
End Function

Private Function BuildColumnCollection(ByVal firstCol As Long, ByVal lastCol As Long) As Collection
    Dim result As Collection
    Dim c As Long

    If lastCol < firstCol Then Exit Function
    Set result = New Collection
    For c = firstCol To lastCol
        AddUniqueLongToCollection result, c
    Next c
    If result.count > 0 Then Set BuildColumnCollection = result
End Function

Private Function BuildRowKeyByColumns(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal colIndexes As Collection) As String
    Dim idx As Variant
    Dim parts As String

    For Each idx In colIndexes
        parts = parts & KEY_SEP & GetCellKeyValue(ws.cells(rowIndex, CLng(idx)))
    Next idx
    BuildRowKeyByColumns = parts
End Function

Private Function GetCellKeyValue(ByVal c As Range) As String
    On Error Resume Next
    GetCellKeyValue = NormalizeText(c.Value2)
    On Error GoTo 0
End Function

Private Sub MarkDuplicateRow(ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal firstCol As Long, ByVal lastCol As Long)
    If lastCol < firstCol Then Exit Sub
    ws.Range(ws.cells(rowIndex, firstCol), ws.cells(rowIndex, lastCol)).Interior.Color = RGB(255, 199, 206)
End Sub

Private Sub AddUniqueLongToCollection(ByVal items As Collection, ByVal valueToAdd As Long)
    Dim itm As Variant
    For Each itm In items
        If CLng(itm) = valueToAdd Then Exit Sub
    Next itm
    items.Add valueToAdd
End Sub

Private Function CellCommentContains(ByVal c As Range, ByVal keywordText As String) As Boolean
    Dim commentText As String

    On Error Resume Next
    If c Is Nothing Then Exit Function
    If c.Comment Is Nothing Then Exit Function
    commentText = NormalizeText(c.Comment.text)
    On Error GoTo 0

    CellCommentContains = (InStr(1, commentText, NormalizeText(keywordText), vbTextCompare) > 0)
End Function

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

Private Function FindOpenWorkbookByFullName(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.fullName, workbookPath, vbTextCompare) = 0 Then
            Set FindOpenWorkbookByFullName = wb
            Exit Function
        End If
    Next wb
End Function

Private Function PickWorkbookPath() As String
    Dim fd As FileDialog

    On Error Resume Next
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    On Error GoTo 0
    If fd Is Nothing Then Exit Function

    With fd
        .Title = "请选择要检查重复数据的工作簿"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel/CSV 文件", "*.xls;*.xlsx;*.xlsm;*.xlsb;*.csv"
        .Filters.Add "所有文件", "*.*"
        If .Show <> -1 Then Exit Function
        PickWorkbookPath = CStr(.SelectedItems(1))
    End With
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
