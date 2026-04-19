Attribute VB_Name = "功能3dot15批量输出打印版"
Option Explicit

Private Const KEEP_LOG_KEY As String = "3.10.1 批量输出打印版（保留原格式）"
Private Const LIGHT_LOG_KEY As String = "3.10.2 批量输出打印版（轻度重排）"
Private Const FAST_LOG_KEY As String = "3.10.3 批量输出打印版（快速复制）"
Private Const CFG_KEEP_LOG_KEY As String = "3.10.4 按配置打印（保留原格式）"
Private Const CFG_LIGHT_LOG_KEY As String = "3.10.5 按配置打印（轻度重排）"
Private Const CFG_FAST_LOG_KEY As String = "3.10.6 按配置打印（快速复制）"
Private Const CFG_CHECK_LOG_KEY As String = "3.10.7 按配置打印预校验"
Private Const PRINT_CONFIG_SHEET As String = "打印配置"
Private Const SHEET_MARK_PRINT As String = "打印"
Private Const SHEET_MARK_OUTPUT As String = "输出"
Private Const REGION_KEY As String = "打印区域"
Private Const KEEP_OUTPUT_SUFFIX As String = "_打印版_保留源格式.xlsx"
Private Const LIGHT_OUTPUT_SUFFIX As String = "_打印版_轻度重排.xlsx"
Private Const FAST_OUTPUT_SUFFIX As String = "_打印版_快速.xlsx"
Private Const LARGE_RANGE_CELL_THRESHOLD As Long = 5000
Private Const LARGE_RANGE_COLUMN_THRESHOLD As Long = 25
Private Const LARGE_RANGE_ROW_THRESHOLD As Long = 200
Private Const DEFAULT_FIT_TO_PAGES_WIDE As Long = 1
Private Const DEFAULT_FIT_TO_PAGES_TALL As Long = 1
Private Const WARN_ROW_THRESHOLD As Long = 120
Private Const WARN_COL_THRESHOLD As Long = 12
Private Const WARN_CELL_THRESHOLD As Long = 6000
Private Const PRINT_MODE_KEEP As Long = 1
Private Const PRINT_MODE_LIGHT As Long = 2
Private Const PRINT_MODE_FAST As Long = 3

Public Sub 批量输出打印版_保留原格式()
    ExecuteBatchPrintExport False
End Sub

Public Sub 批量输出打印版_轻度重排()
    ExecuteBatchPrintExport True
End Sub

Public Sub 批量输出打印版_快速复制()
    ExecuteFastPrintExport
End Sub

Public Sub 按配置打印_保留原格式()
    ExecuteConfigPrintExport PRINT_MODE_KEEP
End Sub

Public Sub 按配置打印_轻度重排()
    ExecuteConfigPrintExport PRINT_MODE_LIGHT
End Sub

Public Sub 按配置打印_快速复制()
    ExecuteConfigPrintExport PRINT_MODE_FAST
End Sub
Public Sub PrintConfig_RunAllModes()
    Dim t0 As Double
    Dim errNo As Long
    Dim errDesc As String

    t0 = Timer
    On Error GoTo FailHandler
    ExecuteConfigPrintExport 0, True, True
    RunLog_WriteRow "3.10.7 按配置打印（执行全部模式）", "完成", "", "", "", "成功", "已按 1/2/3 模式执行完成", CStr(Round(Timer - t0, 2))
    MsgBox "打印配置全模式执行完成。", vbInformation, "按配置打印"
    Exit Sub

FailHandler:
    errNo = Err.Number
    errDesc = Err.Description
    RunLog_WriteRow "3.10.7 按配置打印（执行全部模式）", "结果", "", "", "", "失败", CStr(errNo) & " " & errDesc, CStr(Round(Timer - t0, 2))
    MsgBox "执行全模式失败：" & CStr(errNo) & " " & errDesc, vbCritical, "按配置打印"
End Sub


Public Sub 按配置打印_预校验()
    ExecuteConfigPrintPrecheck
End Sub

Public Sub 初始化打印配置()
    Dim ws As Worksheet

    Set ws = EnsurePrintConfigSheet()
    InitPrintConfigHeader ws
    WritePrintConfigExample ws
    ws.Columns("A:J").AutoFit
    ws.Activate

    MsgBox "打印配置已初始化。", vbInformation, "按配置打印"
End Sub

Private Sub ExecuteBatchPrintExport(ByVal useLightRelayout As Boolean)
    Dim t0 As Double
    Dim logKey As String
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim hitBooks As Long
    Dim hitSheets As Long
    Dim hitRegions As Long
    Dim outputBooks As Long
    Dim skipSheets As Long
    Dim failBooks As Long
    Dim regionHitSheets As Long
    Dim usedRangeSheets As Long
    Dim pageWarnCount As Long
    Dim customPagingSheets As Long
    Dim splitAdviceCount As Long
    Dim errNo As Long
    Dim errDesc As String

    t0 = Timer
    logKey = GetPrintLogKey(useLightRelayout)
    RunLog_WriteRow logKey, "开始", "", "", "", "", "开始", ""

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "请选择要输出打印版的源文件，可多选"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"
        If .Show <> -1 Then
            RunLog_WriteRow logKey, "结束", "", "", "", "", "用户取消", CStr(Round(Timer - t0, 2))
            Exit Sub
        End If
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    SetAskToUpdateLinksSafe False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    For Each fileItem In fd.SelectedItems
        ProcessOneSourceWorkbook CStr(fileItem), useLightRelayout, logKey, hitBooks, hitSheets, hitRegions, outputBooks, skipSheets, failBooks, regionHitSheets, usedRangeSheets, pageWarnCount, customPagingSheets, splitAdviceCount
    Next fileItem

    Application.Calculation = xlCalculationAutomatic
    SetAskToUpdateLinksSafe True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    RunLog_WriteRow logKey, "结束", CStr(hitBooks), CStr(hitSheets), CStr(hitRegions), "完成", "输出工作簿=" & outputBooks & "，打印区域=" & regionHitSheets & "，UsedRange回退=" & usedRangeSheets & "，分页预警=" & pageWarnCount & "，自定义分页=" & customPagingSheets & "，建议拆页=" & splitAdviceCount & "，跳过工作表=" & skipSheets & "，失败工作簿=" & failBooks, CStr(Round(Timer - t0, 2))
    MsgBox BuildPrintSummaryMessage(useLightRelayout, hitBooks, hitSheets, hitRegions, outputBooks, skipSheets, failBooks, regionHitSheets, usedRangeSheets, pageWarnCount, customPagingSheets, splitAdviceCount), vbInformation, "批量输出打印版"
    Exit Sub

ErrHandler:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = xlCalculationAutomatic
    SetAskToUpdateLinksSafe True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    RunLog_WriteRow logKey, "结束", "", "", "", "失败", CStr(errNo) & " " & errDesc, CStr(Round(Timer - t0, 2))
    MsgBox "执行失败：" & CStr(errNo) & " " & errDesc, vbCritical, "批量输出打印版"
End Sub

Private Sub ExecuteFastPrintExport()
    Dim t0 As Double
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim hitBooks As Long
    Dim hitSheets As Long
    Dim outputBooks As Long
    Dim skipSheets As Long
    Dim failBooks As Long
    Dim usedRangeSheets As Long
    Dim pageWarnCount As Long
    Dim customPagingSheets As Long
    Dim splitAdviceCount As Long
    Dim errNo As Long
    Dim errDesc As String

    t0 = Timer
    RunLog_WriteRow FAST_LOG_KEY, "开始", "", "", "", "", "开始", ""

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "请选择要快速输出打印版的源文件，可多选"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"
        If .Show <> -1 Then
            RunLog_WriteRow FAST_LOG_KEY, "结束", "", "", "", "", "用户取消", CStr(Round(Timer - t0, 2))
            Exit Sub
        End If
    End With

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    SetAskToUpdateLinksSafe False
    Application.Calculation = xlCalculationManual
    On Error GoTo ErrHandler

    For Each fileItem In fd.SelectedItems
        ProcessOneSourceWorkbookFast CStr(fileItem), hitBooks, hitSheets, outputBooks, skipSheets, failBooks, usedRangeSheets, pageWarnCount, customPagingSheets, splitAdviceCount
    Next fileItem

    Application.Calculation = xlCalculationAutomatic
    SetAskToUpdateLinksSafe True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    RunLog_WriteRow FAST_LOG_KEY, "结束", CStr(hitBooks), CStr(hitSheets), "", "完成", "输出工作簿=" & outputBooks & "，UsedRange回退=" & usedRangeSheets & "，分页预警=" & pageWarnCount & "，自定义分页=" & customPagingSheets & "，建议拆页=" & splitAdviceCount & "，跳过工作表=" & skipSheets & "，失败工作簿=" & failBooks, CStr(Round(Timer - t0, 2))
    MsgBox BuildFastPrintSummaryMessage(hitBooks, hitSheets, outputBooks, skipSheets, failBooks, usedRangeSheets, pageWarnCount, customPagingSheets, splitAdviceCount), vbInformation, "批量输出打印版（快速复制）"
    Exit Sub

ErrHandler:
    errNo = Err.Number
    errDesc = Err.Description
    Application.Calculation = xlCalculationAutomatic
    SetAskToUpdateLinksSafe True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    RunLog_WriteRow FAST_LOG_KEY, "结束", "", "", "", "失败", CStr(errNo) & " " & errDesc, CStr(Round(Timer - t0, 2))
    MsgBox "执行失败：" & CStr(errNo) & " " & errDesc, vbCritical, "批量输出打印版（快速复制）"
End Sub

Private Sub ProcessOneSourceWorkbook(ByVal sourcePath As String, ByVal useLightRelayout As Boolean, ByVal logKey As String, ByRef hitBooks As Long, ByRef hitSheets As Long, ByRef hitRegions As Long, ByRef outputBooks As Long, ByRef skipSheets As Long, ByRef failBooks As Long, ByRef regionHitSheets As Long, ByRef usedRangeSheets As Long, ByRef pageWarnCount As Long, ByRef customPagingSheets As Long, ByRef splitAdviceCount As Long)
    Dim srcWb As Workbook
    Dim outWb As Workbook
    Dim srcWs As Worksheet
    Dim outWs As Worksheet
    Dim regions As Collection
    Dim oneRegion As Variant
    Dim parseMessage As String
    Dim hasRegionMarker As Boolean
    Dim outSheetCount As Long
    Dim oneSheetHit As Boolean
    Dim oneBookHit As Boolean
    Dim regionIndex As Long
    Dim outputPath As String
    Dim orientationText As String
    Dim fitToPagesWide As Long
    Dim fitToPagesTall As Long
    Dim wideSpecified As Boolean
    Dim tallSpecified As Boolean
    Dim hasCustomPaging As Boolean
    Dim pagingMessage As String
    Dim pageWarnMessage As String
    Dim splitAdviceMessage As String
    Dim scanHitSheets As Long
    Dim scanRegionSheets As Long
    Dim scanUsedRangeSheets As Long
    Dim scanCustomPagingSheets As Long
    Dim scanSkipSheets As Long
    Dim scanSummary As String
    Dim errNo As Long
    Dim errDesc As String

    On Error GoTo ErrHandler

    If Not TryOpenWorkbookCompatiblePrint(sourcePath, True, srcWb, errNo, errDesc) Then
        Err.Raise errNo, , errDesc
    End If
    hitBooks = hitBooks + 1
    ScanWorkbookPrintTargets srcWb, False, scanHitSheets, scanRegionSheets, scanUsedRangeSheets, scanCustomPagingSheets, scanSkipSheets, scanSummary
    RunLog_WriteRow logKey, "预扫描摘要", srcWb.Name, CStr(scanHitSheets), CStr(scanRegionSheets), "完成", scanSummary, ""

    For Each srcWs In srcWb.Worksheets
        hasRegionMarker = False
        parseMessage = ""
        Set regions = Nothing

        If Not ShouldPrintSheet(srcWs) Then
            skipSheets = skipSheets + 1
            RunLog_WriteRow logKey, "跳过工作表", srcWb.Name & " | " & srcWs.Name, "", "", "跳过", "A1批注未包含“打印”或“输出”", ""
            GoTo NextSheet
        End If

        ParseSheetPagingOptions srcWs, fitToPagesWide, fitToPagesTall, hasCustomPaging, pagingMessage
        If hasCustomPaging Then
            customPagingSheets = customPagingSheets + 1
            RunLog_WriteRow logKey, "检测到A1分页参数", srcWb.Name & " | " & srcWs.Name, CStr(fitToPagesWide), CStr(fitToPagesTall), "成功", "", ""
        End If
        If pagingMessage <> "" Then
            RunLog_WriteRow logKey, "分页参数忽略", srcWb.Name & " | " & srcWs.Name, "", "", "跳过", pagingMessage, ""
        End If

        Set regions = ExtractPrintRegions(srcWs, hasRegionMarker, parseMessage)
        If regions Is Nothing Then
            If hasRegionMarker Then
                skipSheets = skipSheets + 1
                RunLog_WriteRow logKey, "跳过工作表", srcWb.Name & " | " & srcWs.Name, "", "", "跳过", "打印区域解析失败：" & parseMessage, ""
                GoTo NextSheet
            Else
                Set regions = BuildUsedRangeRegions(srcWs)
                If regions Is Nothing Then
                    skipSheets = skipSheets + 1
                    RunLog_WriteRow logKey, "跳过工作表", srcWb.Name & " | " & srcWs.Name, "", "", "跳过", "未找到打印区域批注且UsedRange为空", ""
                    GoTo NextSheet
                End If
                usedRangeSheets = usedRangeSheets + 1
                RunLog_WriteRow logKey, "UsedRange回退", srcWb.Name & " | " & srcWs.Name, "", "", "成功", "未设置打印区域，按UsedRange输出", ""
            End If
        Else
            regionHitSheets = regionHitSheets + 1
            RunLog_WriteRow logKey, "打印区域命中", srcWb.Name & " | " & srcWs.Name, CStr(regions.Count), "", "成功", "", ""
        End If

        oneSheetHit = False
        regionIndex = 0
        For Each oneRegion In regions
            regionIndex = regionIndex + 1
            If outWb Is Nothing Then
                Set outWb = Workbooks.Add(xlWBATWorksheet)
                outSheetCount = 0
            End If

            outSheetCount = outSheetCount + 1
            If outSheetCount = 1 Then
                Set outWs = outWb.Worksheets(1)
            Else
                Set outWs = outWb.Worksheets.Add(After:=outWb.Worksheets(outWb.Worksheets.Count))
            End If

            ExportRegionToPrintSheet srcWs, oneRegion, outWs, regionIndex, regions.Count, useLightRelayout, fitToPagesWide, fitToPagesTall, orientationText, pageWarnMessage, splitAdviceMessage
            oneSheetHit = True
            oneBookHit = True
            hitRegions = hitRegions + 1
            RunLog_WriteRow logKey, "输出页", srcWb.Name & " | " & srcWs.Name, CStr(regionIndex), orientationText, "成功", outWs.Name, ""
            If pageWarnMessage <> "" Then
                pageWarnCount = pageWarnCount + 1
                RunLog_WriteRow logKey, "页面过小预警", srcWb.Name & " | " & srcWs.Name, outWs.Name, orientationText, "提示", pageWarnMessage, ""
            End If
            If splitAdviceMessage <> "" Then
                splitAdviceCount = splitAdviceCount + 1
                RunLog_WriteRow logKey, "建议拆页", srcWb.Name & " | " & srcWs.Name, outWs.Name, CStr(fitToPagesTall), "提示", splitAdviceMessage, ""
            End If
        Next oneRegion

        If oneSheetHit Then
            hitSheets = hitSheets + 1
        End If
NextSheet:
    Next srcWs

    If oneBookHit Then
        ApplySequentialSheetFooters outWb, logKey, srcWb.Name
        outputPath = BuildAvailableOutputPath(sourcePath, useLightRelayout)
        outWb.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
        outputBooks = outputBooks + 1
        RunLog_WriteRow logKey, "保存结果", srcWb.Name, CStr(scanHitSheets), outputPath, "成功", scanSummary, ""
    Else
        RunLog_WriteRow logKey, "结束源文件", srcWb.Name, "", "", "跳过", "未找到可输出的打印页", ""
    End If

SafeExit:
    SafeCloseWorkbook outWb, False
    SafeCloseWorkbook srcWb, False
    Exit Sub

ErrHandler:
    errNo = Err.Number
    errDesc = Err.Description
    failBooks = failBooks + 1
    RunLog_WriteRow logKey, "处理源文件", sourcePath, "", "", "失败", CStr(errNo) & " " & errDesc, ""
    Resume SafeExit
End Sub

Private Sub ProcessOneSourceWorkbookFast(ByVal sourcePath As String, ByRef hitBooks As Long, ByRef hitSheets As Long, ByRef outputBooks As Long, ByRef skipSheets As Long, ByRef failBooks As Long, ByRef usedRangeSheets As Long, ByRef pageWarnCount As Long, ByRef customPagingSheets As Long, ByRef splitAdviceCount As Long)
    Dim srcWb As Workbook
    Dim outWb As Workbook
    Dim srcWs As Worksheet
    Dim outWs As Worksheet
    Dim outSheetCount As Long
    Dim oneBookHit As Boolean
    Dim outputPath As String
    Dim orientationText As String
    Dim fitToPagesWide As Long
    Dim fitToPagesTall As Long
    Dim hasCustomPaging As Boolean
    Dim pagingMessage As String
    Dim pageWarnMessage As String
    Dim splitAdviceMessage As String
    Dim scanHitSheets As Long
    Dim scanRegionSheets As Long
    Dim scanUsedRangeSheets As Long
    Dim scanCustomPagingSheets As Long
    Dim scanSkipSheets As Long
    Dim scanSummary As String
    Dim errNo As Long
    Dim errDesc As String

    On Error GoTo ErrHandler

    If Not TryOpenWorkbookCompatiblePrint(sourcePath, True, srcWb, errNo, errDesc) Then
        Err.Raise errNo, , errDesc
    End If
    hitBooks = hitBooks + 1
    ScanWorkbookPrintTargets srcWb, True, scanHitSheets, scanRegionSheets, scanUsedRangeSheets, scanCustomPagingSheets, scanSkipSheets, scanSummary
    RunLog_WriteRow FAST_LOG_KEY, "预扫描摘要", srcWb.Name, CStr(scanHitSheets), CStr(scanUsedRangeSheets), "完成", scanSummary, ""

    For Each srcWs In srcWb.Worksheets
        If Not ShouldPrintSheet(srcWs) Then
            skipSheets = skipSheets + 1
            RunLog_WriteRow FAST_LOG_KEY, "跳过工作表", srcWb.Name & " | " & srcWs.Name, "", "", "跳过", "A1批注未包含“打印”或“输出”", ""
            GoTo NextSheet
        End If

        ParseSheetPagingOptions srcWs, fitToPagesWide, fitToPagesTall, hasCustomPaging, pagingMessage
        If hasCustomPaging Then
            customPagingSheets = customPagingSheets + 1
            RunLog_WriteRow FAST_LOG_KEY, "检测到A1分页参数", srcWb.Name & " | " & srcWs.Name, CStr(fitToPagesWide), CStr(fitToPagesTall), "成功", "", ""
        End If
        If pagingMessage <> "" Then
            RunLog_WriteRow FAST_LOG_KEY, "分页参数忽略", srcWb.Name & " | " & srcWs.Name, "", "", "跳过", pagingMessage, ""
        End If

        If GetLastUsedRow(srcWs) = 0 Or GetLastUsedCol(srcWs) = 0 Then
            skipSheets = skipSheets + 1
            RunLog_WriteRow FAST_LOG_KEY, "跳过工作表", srcWb.Name & " | " & srcWs.Name, "", "", "跳过", "UsedRange为空", ""
            GoTo NextSheet
        End If

        usedRangeSheets = usedRangeSheets + 1
        RunLog_WriteRow FAST_LOG_KEY, "UsedRange回退", srcWb.Name & " | " & srcWs.Name, "", "", "成功", "快速复制按UsedRange输出", ""

        If outWb Is Nothing Then
            Set outWb = Workbooks.Add(xlWBATWorksheet)
            outSheetCount = 0
        End If

        outSheetCount = outSheetCount + 1
        If outSheetCount = 1 Then
            Set outWs = outWb.Worksheets(1)
        Else
            Set outWs = outWb.Worksheets.Add(After:=outWb.Worksheets(outWb.Worksheets.Count))
        End If

        ExportWholeSheetFast srcWs, outWs, fitToPagesWide, fitToPagesTall, orientationText, pageWarnMessage, splitAdviceMessage
        oneBookHit = True
        hitSheets = hitSheets + 1
        RunLog_WriteRow FAST_LOG_KEY, "输出页", srcWb.Name & " | " & srcWs.Name, "", orientationText, "成功", outWs.Name, ""
        If pageWarnMessage <> "" Then
            pageWarnCount = pageWarnCount + 1
            RunLog_WriteRow FAST_LOG_KEY, "页面过小预警", srcWb.Name & " | " & srcWs.Name, outWs.Name, orientationText, "提示", pageWarnMessage, ""
        End If
        If splitAdviceMessage <> "" Then
            splitAdviceCount = splitAdviceCount + 1
            RunLog_WriteRow FAST_LOG_KEY, "建议拆页", srcWb.Name & " | " & srcWs.Name, outWs.Name, CStr(fitToPagesTall), "提示", splitAdviceMessage, ""
        End If
NextSheet:
    Next srcWs

    If oneBookHit Then
        ApplySequentialSheetFooters outWb, FAST_LOG_KEY, srcWb.Name
        outputPath = BuildAvailableOutputPathWithSuffix(sourcePath, FAST_OUTPUT_SUFFIX)
        outWb.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
        outputBooks = outputBooks + 1
        RunLog_WriteRow FAST_LOG_KEY, "保存结果", srcWb.Name, CStr(scanHitSheets), outputPath, "成功", scanSummary, ""
    Else
        RunLog_WriteRow FAST_LOG_KEY, "结束源文件", srcWb.Name, "", "", "跳过", "未找到可输出的工作表", ""
    End If

SafeExit:
    SafeCloseWorkbook outWb, False
    SafeCloseWorkbook srcWb, False
    Exit Sub

ErrHandler:
    errNo = Err.Number
    errDesc = Err.Description
    failBooks = failBooks + 1
    RunLog_WriteRow FAST_LOG_KEY, "处理源文件", sourcePath, "", "", "失败", CStr(errNo) & " " & errDesc, ""
    Resume SafeExit
End Sub

Private Function ShouldPrintSheet(ByVal ws As Worksheet) As Boolean
    Dim markerText As String

    markerText = GetCellCommentText(ws.Range("A1"))
    If markerText = "" Then Exit Function

    ShouldPrintSheet = (InStr(1, markerText, SHEET_MARK_PRINT, vbTextCompare) > 0 Or InStr(1, markerText, SHEET_MARK_OUTPUT, vbTextCompare) > 0)
End Function

Private Function ExtractPrintRegions(ByVal ws As Worksheet, ByRef hasRegionMarker As Boolean, ByRef parseMessage As String) As Collection
    Dim startMap As Object
    Dim endMap As Object
    Dim cm As Comment
    Dim commentText As String
    Dim regionNo As Long
    Dim isEndMarker As Boolean
    Dim oneKey As Variant
    Dim startPos As Variant
    Dim endPos As Variant
    Dim sr As Long
    Dim er As Long
    Dim sc As Long
    Dim ec As Long
    Dim result As Collection

    Set startMap = CreateObject("Scripting.Dictionary")
    Set endMap = CreateObject("Scripting.Dictionary")
    startMap.CompareMode = vbTextCompare
    endMap.CompareMode = vbTextCompare

    For Each cm In ws.Comments
        commentText = GetCellCommentText(cm.Parent)
        If ParsePrintRegionMarker(commentText, regionNo, isEndMarker) Then
            hasRegionMarker = True
            If isEndMarker Then
                endMap(CStr(regionNo)) = Array(cm.Parent.Row, cm.Parent.Column)
            Else
                startMap(CStr(regionNo)) = Array(cm.Parent.Row, cm.Parent.Column)
            End If
        End If
    Next cm

    If Not hasRegionMarker Then Exit Function

    Set result = New Collection
    For Each oneKey In startMap.Keys
        If Not endMap.Exists(CStr(oneKey)) Then
            AppendParseMessage parseMessage, "缺少结束标记#" & CStr(oneKey)
        Else
            startPos = startMap(CStr(oneKey))
            endPos = endMap(CStr(oneKey))
            sr = CLng(startPos(0))
            er = CLng(endPos(0))
            sc = CLng(startPos(1))
            ec = CLng(endPos(1))
            NormalizeRegionBounds sr, er, sc, ec
            result.Add Array(sr, er, sc, ec, CLng(oneKey))
        End If
    Next oneKey

    For Each oneKey In endMap.Keys
        If Not startMap.Exists(CStr(oneKey)) Then
            AppendParseMessage parseMessage, "缺少开始标记" & CStr(oneKey)
        End If
    Next oneKey

    If result.Count = 0 Then Exit Function
    If parseMessage <> "" Then
        Set ExtractPrintRegions = Nothing
    Else
        Set ExtractPrintRegions = result
    End If
End Function

Private Function ParsePrintRegionMarker(ByVal commentText As String, ByRef regionNo As Long, ByRef isEndMarker As Boolean) As Boolean
    Dim hitPos As Long
    Dim suffixText As String
    Dim digitText As String
    Dim idx As Long
    Dim oneChar As String

    hitPos = InStr(1, commentText, REGION_KEY, vbTextCompare)
    If hitPos <= 0 Then Exit Function

    suffixText = Mid$(commentText, hitPos + Len(REGION_KEY))
    If Left$(suffixText, 1) = "#" Then
        isEndMarker = True
        suffixText = Mid$(suffixText, 2)
    Else
        isEndMarker = False
    End If

    For idx = 1 To Len(suffixText)
        oneChar = Mid$(suffixText, idx, 1)
        If oneChar Like "[0-9]" Then
            digitText = digitText & oneChar
        Else
            Exit For
        End If
    Next idx

    If digitText = "" Then Exit Function

    regionNo = CLng(digitText)
    ParsePrintRegionMarker = True
End Function

Private Function BuildUsedRangeRegions(ByVal ws As Worksheet) As Collection
    Dim firstRow As Long
    Dim lastRow As Long
    Dim firstCol As Long
    Dim lastCol As Long
    Dim topFirstCol As Long
    Dim topLastCol As Long
    Dim result As Collection

    firstRow = GetFirstUsedRow(ws)
    lastRow = GetLastUsedRow(ws)
    firstCol = GetFirstUsedCol(ws)
    lastCol = GetLastUsedCol(ws)

    If firstRow <= 0 Or lastRow <= 0 Or firstCol <= 0 Or lastCol <= 0 Then Exit Function
    If lastRow < firstRow Or lastCol < firstCol Then Exit Function

    If firstRow > 1 Then
        If TryGetEffectiveBoundsForRow(ws, 1, topFirstCol, topLastCol) Then
            firstRow = 1
            If topFirstCol > 0 And topFirstCol < firstCol Then firstCol = topFirstCol
            If topLastCol > lastCol Then lastCol = topLastCol
        End If
    End If

    Set result = New Collection
    result.Add Array(firstRow, lastRow, firstCol, lastCol, 1)
    Set BuildUsedRangeRegions = result
End Function

Private Function TryGetEffectiveBoundsForRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByRef firstCol As Long, ByRef lastCol As Long) As Boolean
    Dim rowRange As Range
    Dim firstCell As Range
    Dim lastCell As Range

    On Error GoTo SafeExit
    Set rowRange = ws.Rows(rowNo)
    Set firstCell = rowRange.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    If firstCell Is Nothing Then
        Set firstCell = rowRange.Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)
    End If
    If firstCell Is Nothing Then GoTo SafeExit

    Set lastCell = rowRange.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    If lastCell Is Nothing Then
        Set lastCell = rowRange.Find(What:="*", LookIn:=xlValues, LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False)
    End If
    If lastCell Is Nothing Then GoTo SafeExit

    firstCol = firstCell.Column
    lastCol = lastCell.Column
    TryGetEffectiveBoundsForRow = True
    Exit Function

SafeExit:
    TryGetEffectiveBoundsForRow = False
End Function

Private Sub ExportRegionToPrintSheet(ByVal srcWs As Worksheet, ByVal regionInfo As Variant, ByVal outWs As Worksheet, ByVal regionIndex As Long, ByVal regionCount As Long, ByVal useLightRelayout As Boolean, ByVal fitToPagesWide As Long, ByVal fitToPagesTall As Long, ByRef orientationText As String, ByRef pageWarnMessage As String, ByRef splitAdviceMessage As String)
    Dim srcRange As Range
    Dim targetRange As Range
    Dim rowCount As Long
    Dim colCount As Long

    Set srcRange = srcWs.Range(srcWs.Cells(CLng(regionInfo(0)), CLng(regionInfo(2))), srcWs.Cells(CLng(regionInfo(1)), CLng(regionInfo(3))))
    rowCount = srcRange.Rows.Count
    colCount = srcRange.Columns.Count

    outWs.Cells.Clear
    srcRange.Copy
    outWs.Range("A1").PasteSpecial xlPasteAll
    outWs.Range("A1").PasteSpecial xlPasteColumnWidths
    Application.CutCopyMode = False

    Set targetRange = outWs.Range("A1").Resize(rowCount, colCount)
    CopyRowHeights srcWs, outWs, CLng(regionInfo(0)), rowCount
    CopyColumnWidths srcWs, outWs, CLng(regionInfo(2)), colCount
    EnsureFirstRowVisibleForPrint outWs, targetRange
    srcRange.Copy
    targetRange.PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    SyncRangeColorAndFontRGB srcRange, targetRange

    outWs.DisplayPageBreaks = False

    If useLightRelayout And Not IsLargeRange(targetRange) Then
        ApplyLightRelayout outWs, targetRange
    End If

    ApplyPageSetup outWs, targetRange, useLightRelayout, fitToPagesWide, fitToPagesTall, orientationText, pageWarnMessage, splitAdviceMessage
    outWs.Name = BuildOutputSheetName(outWs.Parent, srcWs.Name, regionIndex, regionCount)
End Sub

Private Sub ExportWholeSheetFast(ByVal srcWs As Worksheet, ByVal outWs As Worksheet, ByVal fitToPagesWide As Long, ByVal fitToPagesTall As Long, ByRef orientationText As String, ByRef pageWarnMessage As String, ByRef splitAdviceMessage As String)
    Dim srcRange As Range
    Dim targetRange As Range
    Dim rowCount As Long
    Dim colCount As Long

    Set srcRange = srcWs.UsedRange
    rowCount = srcRange.Rows.Count
    colCount = srcRange.Columns.Count

    outWs.Cells.Clear
    srcRange.Copy
    outWs.Range("A1").PasteSpecial xlPasteFormats
    outWs.Range("A1").PasteSpecial xlPasteColumnWidths
    Application.CutCopyMode = False

    Set targetRange = outWs.Range("A1").Resize(rowCount, colCount)
    CopyRowHeights srcWs, outWs, srcRange.Row, rowCount
    CopyColumnWidths srcWs, outWs, srcRange.Column, colCount
    EnsureFirstRowVisibleForPrint outWs, targetRange
    srcRange.Copy
    targetRange.PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False
    SyncRangeColorAndFontRGB srcRange, targetRange
    outWs.DisplayPageBreaks = False

    outWs.Name = BuildOutputSheetName(outWs.Parent, srcWs.Name, 1, 1)
    On Error Resume Next
    outWs.Range("A1").ClearComments
    On Error GoTo 0

    ApplyPageSetup outWs, targetRange, False, fitToPagesWide, fitToPagesTall, orientationText, pageWarnMessage, splitAdviceMessage
End Sub

Private Sub ApplyPageSetup(ByVal ws As Worksheet, ByVal targetRange As Range, ByVal useLightRelayout As Boolean, ByVal fitToPagesWide As Long, ByVal fitToPagesTall As Long, ByRef orientationText As String, ByRef pageWarnMessage As String, ByRef splitAdviceMessage As String)
    Dim useLandscape As Boolean
    Dim supportsPrintCommunication As Boolean

    useLandscape = ShouldUseLandscape(targetRange)
    orientationText = IIf(useLandscape, "横向", "纵向")

    If useLightRelayout Then
        OptimizeRangeForSinglePage ws, targetRange, True, useLandscape
    End If

    supportsPrintCommunication = CanUsePrintCommunication()
    On Error Resume Next
    If supportsPrintCommunication Then
        Application.PrintCommunication = False
    End If
    On Error GoTo 0

    With ws.PageSetup
        .PrintArea = targetRange.Address
        .PaperSize = xlPaperA4
        .Orientation = IIf(useLandscape, xlLandscape, xlPortrait)
        .Zoom = False
        .FitToPagesWide = SafeFitPageValue(fitToPagesWide)
        .FitToPagesTall = SafeFitPageValue(fitToPagesTall)
        .CenterHorizontally = True
        .CenterVertically = False
        .LeftMargin = Application.CentimetersToPoints(0.8)
        .RightMargin = Application.CentimetersToPoints(0.8)
        .TopMargin = Application.CentimetersToPoints(1.2)
        .BottomMargin = Application.CentimetersToPoints(1.8)
        .HeaderMargin = Application.CentimetersToPoints(0.5)
        .FooterMargin = Application.CentimetersToPoints(0.8)
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = "第 &P 页 / 共 &N 页"
        .RightFooter = ""
    End With

    On Error Resume Next
    If supportsPrintCommunication Then
        Application.PrintCommunication = True
    End If
    On Error GoTo 0

    pageWarnMessage = BuildPageWarnMessage(targetRange, useLandscape, fitToPagesWide, fitToPagesTall, useLightRelayout)
    splitAdviceMessage = BuildSplitAdviceMessage(targetRange, fitToPagesTall)
End Sub

Private Function ShouldUseLandscape(ByVal targetRange As Range) As Boolean
    If targetRange.Columns.Count >= 8 Then
        ShouldUseLandscape = True
        Exit Function
    End If

    ShouldUseLandscape = (targetRange.Width > targetRange.Height)
End Function

Private Sub ApplyLightRelayout(ByVal ws As Worksheet, ByVal targetRange As Range)
    Dim headerRows As Long
    Dim firstDataRow As Long
    Dim col As Range

    targetRange.WrapText = True
    targetRange.VerticalAlignment = xlCenter
    targetRange.Font.Name = "宋体"
    targetRange.Font.Size = 10

    headerRows = ResolveHeaderRows(targetRange)
    If headerRows < 1 Then headerRows = 1

    With ws.Range(ws.Cells(1, 1), ws.Cells(headerRows, targetRange.Columns.Count))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ApplyHeaderShadeOnlyForNoFill ws.Range(ws.Cells(1, 1), ws.Cells(headerRows, targetRange.Columns.Count)), RGB(242, 242, 242)

    ws.Rows(1).Font.Size = 14
    If headerRows >= 2 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(headerRows, targetRange.Columns.Count)).Font.Size = 11
    End If

    If Not IsLargeRange(targetRange) Then
        For Each col In targetRange.Columns
            col.EntireColumn.AutoFit
            If col.ColumnWidth > 20 Then
                col.ColumnWidth = 20
            End If
        Next col

        targetRange.Rows.AutoFit
    End If

    firstDataRow = headerRows + 1
    If firstDataRow <= targetRange.Rows.Count Then
        ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(targetRange.Rows.Count, targetRange.Columns.Count)).Font.Size = 10
    End If
End Sub

Private Sub OptimizeRangeForSinglePage(ByVal ws As Worksheet, ByVal targetRange As Range, ByVal useLightRelayout As Boolean, ByVal useLandscape As Boolean)
    Dim col As Range
    Dim baseFontSize As Double
    Dim maxColumnWidth As Double
    Dim needsAggressiveCompression As Boolean

    needsAggressiveCompression = IsLargeRange(targetRange)
    targetRange.WrapText = True

    If useLightRelayout Then
        baseFontSize = 10
    Else
        baseFontSize = 9.5
    End If

    If useLandscape Then
        maxColumnWidth = 18
    Else
        maxColumnWidth = 12
    End If

    If needsAggressiveCompression Then
        If useLandscape Then
            maxColumnWidth = 16
        Else
            maxColumnWidth = 10
        End If

        If baseFontSize > 9 Then
            baseFontSize = 9
        End If
    End If

    targetRange.Font.Size = baseFontSize

    If needsAggressiveCompression Then
        CompressColumnsByWidth targetRange, maxColumnWidth
    Else
        targetRange.Rows.AutoFit

        For Each col In targetRange.Columns
            col.EntireColumn.AutoFit
            If col.ColumnWidth > maxColumnWidth Then
                col.ColumnWidth = maxColumnWidth
            End If
        Next col

        targetRange.Rows.AutoFit
    End If

    If Not useLandscape Then
        CompressTallRows targetRange
    End If
End Sub

Private Sub CompressColumnsByWidth(ByVal targetRange As Range, ByVal maxColumnWidth As Double)
    Dim col As Range

    For Each col In targetRange.Columns
        If col.ColumnWidth > maxColumnWidth Then
            col.ColumnWidth = maxColumnWidth
        End If
    Next col
End Sub

Private Sub CompressTallRows(ByVal targetRange As Range)
    Dim oneRow As Range

    For Each oneRow In targetRange.Rows
        If oneRow.RowHeight > 30 Then
            oneRow.RowHeight = 30
        End If
    Next oneRow
End Sub

Private Function ResolveHeaderRows(ByVal targetRange As Range) As Long
    Dim maxRows As Long
    Dim rowNo As Long
    Dim numericCount As Long
    Dim totalCount As Long
    Dim oneCell As Range
    Dim textValue As String

    maxRows = targetRange.Rows.Count
    If maxRows > 3 Then maxRows = 3

    For rowNo = 1 To maxRows
        numericCount = 0
        totalCount = 0
        For Each oneCell In targetRange.Rows(rowNo).Cells
            textValue = Trim$(CStr(oneCell.Value))
            If textValue <> "" Then
                totalCount = totalCount + 1
                If IsNumeric(oneCell.Value) Then
                    numericCount = numericCount + 1
                End If
            End If
        Next oneCell

        If totalCount = 0 Then Exit For
        If numericCount * 2 > totalCount Then Exit For
        ResolveHeaderRows = rowNo
    Next rowNo
End Function

Private Sub CopyRowHeights(ByVal srcWs As Worksheet, ByVal outWs As Worksheet, ByVal startRow As Long, ByVal rowCount As Long)
    Dim idx As Long

    For idx = 1 To rowCount
        outWs.Rows(idx).RowHeight = srcWs.Rows(startRow + idx - 1).RowHeight
    Next idx
End Sub

Private Sub CopyColumnWidths(ByVal srcWs As Worksheet, ByVal outWs As Worksheet, ByVal startCol As Long, ByVal colCount As Long)
    Dim idx As Long

    For idx = 1 To colCount
        outWs.Columns(idx).ColumnWidth = srcWs.Columns(startCol + idx - 1).ColumnWidth
    Next idx
End Sub

Private Sub EnsureFirstRowVisibleForPrint(ByVal ws As Worksheet, ByVal targetRange As Range)
    Dim firstRowNo As Long
    Dim stdHeight As Double

    If ws Is Nothing Then Exit Sub
    If targetRange Is Nothing Then Exit Sub
    If targetRange.Rows.Count <= 0 Then Exit Sub

    firstRowNo = targetRange.Row
    On Error Resume Next
    ws.Rows(firstRowNo).Hidden = False
    If ws.Rows(firstRowNo).RowHeight < 2 Then
        stdHeight = ws.StandardHeight
        If stdHeight < 2 Then stdHeight = 15
        ws.Rows(firstRowNo).RowHeight = stdHeight
    End If
    On Error GoTo 0
End Sub

Private Sub SyncRangeColorAndFontRGB(ByVal srcRange As Range, ByVal dstRange As Range)
    Dim r As Long
    Dim c As Long
    Dim rowCount As Long
    Dim colCount As Long
    Dim srcCell As Range
    Dim dstCell As Range

    If srcRange Is Nothing Then Exit Sub
    If dstRange Is Nothing Then Exit Sub

    rowCount = srcRange.Rows.Count
    colCount = srcRange.Columns.Count
    If rowCount <> dstRange.Rows.Count Or colCount <> dstRange.Columns.Count Then Exit Sub

    On Error Resume Next
    For r = 1 To rowCount
        For c = 1 To colCount
            Set srcCell = srcRange.Cells(r, c)
            Set dstCell = dstRange.Cells(r, c)
            dstCell.Interior.Pattern = srcCell.Interior.Pattern
            If srcCell.Interior.Pattern <> xlPatternNone Then
                dstCell.Interior.Color = srcCell.Interior.Color
            End If
            dstCell.Font.Color = srcCell.Font.Color
        Next c
    Next r
    On Error GoTo 0
End Sub

Private Function BuildOutputSheetName(ByVal wb As Workbook, ByVal sourceSheetName As String, ByVal regionIndex As Long, ByVal regionCount As Long) As String
    Dim baseName As String
    Dim candidate As String
    Dim suffixNo As Long

    baseName = SanitizeSheetName(sourceSheetName)
    If regionCount > 1 Then
        baseName = SanitizeSheetName(baseName & "_" & CStr(regionIndex))
    End If
    If baseName = "" Then baseName = "打印页"

    candidate = Left$(baseName, 31)
    suffixNo = 1
    Do While SheetNameExists(wb, candidate)
        suffixNo = suffixNo + 1
        candidate = Left$(baseName, 31 - Len(CStr(suffixNo)) - 1) & "_" & CStr(suffixNo)
    Loop

    BuildOutputSheetName = candidate
End Function

Private Function SanitizeSheetName(ByVal sheetName As String) As String
    Dim oneChar As Variant

    SanitizeSheetName = Trim$(sheetName)
    For Each oneChar In Array(":", "\", "/", "?", "*", "[", "]")
        SanitizeSheetName = Replace$(SanitizeSheetName, CStr(oneChar), "_")
    Next oneChar
End Function

Private Function SheetNameExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet

    For Each ws In wb.Worksheets
        If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
            SheetNameExists = True
            Exit Function
        End If
    Next ws
End Function

Private Function BuildAvailableOutputPath(ByVal sourcePath As String, ByVal useLightRelayout As Boolean) As String
    If useLightRelayout Then
        BuildAvailableOutputPath = BuildAvailableOutputPathWithSuffix(sourcePath, LIGHT_OUTPUT_SUFFIX)
    Else
        BuildAvailableOutputPath = BuildAvailableOutputPathWithSuffix(sourcePath, KEEP_OUTPUT_SUFFIX)
    End If
End Function

Private Function BuildAvailableOutputPathWithSuffix(ByVal sourcePath As String, ByVal outputSuffix As String) As String
    Dim folderPath As String
    Dim fileName As String
    Dim dotPos As Long
    Dim baseName As String
    Dim candidate As String
    Dim suffixNo As Long

    folderPath = Left$(sourcePath, InStrRev(sourcePath, "\"))
    fileName = Mid$(sourcePath, InStrRev(sourcePath, "\") + 1)
    dotPos = InStrRev(fileName, ".")
    If dotPos > 0 Then
        baseName = Left$(fileName, dotPos - 1)
    Else
        baseName = fileName
    End If

    candidate = folderPath & baseName & outputSuffix
    suffixNo = 1
    Do While Len(Dir$(candidate)) > 0
        suffixNo = suffixNo + 1
        candidate = folderPath & baseName & Left$(outputSuffix, Len(outputSuffix) - 5) & "(" & CStr(suffixNo) & ").xlsx"
    Loop

    BuildAvailableOutputPathWithSuffix = candidate
End Function

Private Function BuildPrintSummaryMessage(ByVal useLightRelayout As Boolean, ByVal hitBooks As Long, ByVal hitSheets As Long, ByVal hitRegions As Long, ByVal outputBooks As Long, ByVal skipSheets As Long, ByVal failBooks As Long, ByVal regionHitSheets As Long, ByVal usedRangeSheets As Long, ByVal pageWarnCount As Long, ByVal customPagingSheets As Long, ByVal splitAdviceCount As Long) As String
    BuildPrintSummaryMessage = IIf(useLightRelayout, "轻度重排打印版已完成。", "保留原格式打印版已完成。")
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "扫描源文件数：" & hitBooks
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "命中工作表数：" & hitSheets
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "命中打印区域数：" & hitRegions
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "输出工作簿数：" & outputBooks
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "使用打印区域的工作表：" & regionHitSheets
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "UsedRange回退工作表：" & usedRangeSheets
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "自定义分页工作表：" & customPagingSheets
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "页面过小预警页数：" & pageWarnCount
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "建议拆页页数：" & splitAdviceCount
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "跳过工作表数：" & skipSheets
    BuildPrintSummaryMessage = BuildPrintSummaryMessage & vbCrLf & "失败工作簿数：" & failBooks
End Function

Private Function BuildFastPrintSummaryMessage(ByVal hitBooks As Long, ByVal hitSheets As Long, ByVal outputBooks As Long, ByVal skipSheets As Long, ByVal failBooks As Long, ByVal usedRangeSheets As Long, ByVal pageWarnCount As Long, ByVal customPagingSheets As Long, ByVal splitAdviceCount As Long) As String
    BuildFastPrintSummaryMessage = "快速复制打印版已完成。"
    BuildFastPrintSummaryMessage = BuildFastPrintSummaryMessage & vbCrLf & "扫描源文件数：" & hitBooks
    BuildFastPrintSummaryMessage = BuildFastPrintSummaryMessage & vbCrLf & "命中工作表数：" & hitSheets
    BuildFastPrintSummaryMessage = BuildFastPrintSummaryMessage & vbCrLf & "输出工作簿数：" & outputBooks
    BuildFastPrintSummaryMessage = BuildFastPrintSummaryMessage & vbCrLf & "UsedRange回退工作表：" & usedRangeSheets
    BuildFastPrintSummaryMessage = BuildFastPrintSummaryMessage & vbCrLf & "自定义分页工作表：" & customPagingSheets
    BuildFastPrintSummaryMessage = BuildFastPrintSummaryMessage & vbCrLf & "页面过小预警页数：" & pageWarnCount
    BuildFastPrintSummaryMessage = BuildFastPrintSummaryMessage & vbCrLf & "建议拆页页数：" & splitAdviceCount
    BuildFastPrintSummaryMessage = BuildFastPrintSummaryMessage & vbCrLf & "跳过工作表数：" & skipSheets
    BuildFastPrintSummaryMessage = BuildFastPrintSummaryMessage & vbCrLf & "失败工作簿数：" & failBooks
End Function

Private Sub ApplySequentialSheetFooters(ByVal wb As Workbook, ByVal logKey As String, ByVal sourceBookName As String)
    Dim ws As Worksheet
    Dim totalPages As Long
    Dim startPageNo As Long
    Dim oneSheetPages As Long
    Dim supportsPrintCommunication As Boolean
    Dim errNo As Long
    Dim errDesc As String

    If wb Is Nothing Then Exit Sub
    If wb.Worksheets.Count <= 0 Then Exit Sub

    On Error GoTo FallbackMode

    totalPages = 0
    For Each ws In wb.Worksheets
        oneSheetPages = GetEstimatedSheetPageCount(ws)
        totalPages = totalPages + oneSheetPages
    Next ws
    If totalPages <= 0 Then totalPages = wb.Worksheets.Count

    supportsPrintCommunication = CanUsePrintCommunication()
    On Error Resume Next
    If supportsPrintCommunication Then
        Application.PrintCommunication = False
    End If

    startPageNo = 1
    For Each ws In wb.Worksheets
        oneSheetPages = GetEstimatedSheetPageCount(ws)
        If oneSheetPages <= 0 Then oneSheetPages = 1
        With ws.PageSetup
            .CenterFooter = "第 &P 页 / 共 " & CStr(totalPages) & " 页"
            .FirstPageNumber = startPageNo
        End With
        startPageNo = startPageNo + oneSheetPages
    Next ws

    If supportsPrintCommunication Then
        Application.PrintCommunication = True
    End If
    On Error GoTo 0
    RunLog_WriteRow logKey, "页码累计完成", sourceBookName, CStr(totalPages), "", "成功", "已按实际页数连续编号", ""
    Exit Sub

FallbackMode:
    errNo = Err.Number
    errDesc = Err.Description
    Err.Clear
    On Error Resume Next
    If supportsPrintCommunication Then
        Application.PrintCommunication = False
    End If
    For Each ws In wb.Worksheets
        With ws.PageSetup
            .CenterFooter = "第 &P 页 / 共 &N 页"
            .FirstPageNumber = xlAutomatic
        End With
    Next ws
    If supportsPrintCommunication Then
        Application.PrintCommunication = True
    End If
    On Error GoTo 0
    RunLog_WriteRow logKey, "页码累计降级", sourceBookName, "", "", "提示", CStr(errNo) & " " & errDesc, ""
End Sub

Private Sub ScanWorkbookPrintTargets(ByVal srcWb As Workbook, ByVal fastMode As Boolean, ByRef hitSheets As Long, ByRef regionSheets As Long, ByRef usedRangeSheets As Long, ByRef customPagingSheets As Long, ByRef skipSheets As Long, ByRef summaryText As String)
    Dim ws As Worksheet
    Dim regions As Collection
    Dim hasRegionMarker As Boolean
    Dim parseMessage As String
    Dim fitToPagesWide As Long
    Dim fitToPagesTall As Long
    Dim hasCustomPaging As Boolean
    Dim pagingMessage As String
    Dim regionList As String
    Dim usedRangeList As String
    Dim skipList As String
    Dim errNo As Long
    Dim errDesc As String

    On Error GoTo SafeExit

    For Each ws In srcWb.Worksheets
        If Not ShouldPrintSheet(ws) Then
            skipSheets = skipSheets + 1
            AppendNameList skipList, ws.Name
            GoTo NextSheet
        End If

        ParseSheetPagingOptions ws, fitToPagesWide, fitToPagesTall, hasCustomPaging, pagingMessage
        If hasCustomPaging Then
            customPagingSheets = customPagingSheets + 1
        End If

        If fastMode Then
            If GetLastUsedRow(ws) = 0 Or GetLastUsedCol(ws) = 0 Then
                skipSheets = skipSheets + 1
                AppendNameList skipList, ws.Name
            Else
                hitSheets = hitSheets + 1
                usedRangeSheets = usedRangeSheets + 1
                AppendNameList usedRangeList, ws.Name
            End If
            GoTo NextSheet
        End If

        hasRegionMarker = False
        parseMessage = ""
        Set regions = ExtractPrintRegions(ws, hasRegionMarker, parseMessage)
        If regions Is Nothing Then
            If hasRegionMarker Then
                skipSheets = skipSheets + 1
                AppendNameList skipList, ws.Name
            Else
                Set regions = BuildUsedRangeRegions(ws)
                If regions Is Nothing Then
                    skipSheets = skipSheets + 1
                    AppendNameList skipList, ws.Name
                Else
                    hitSheets = hitSheets + 1
                    usedRangeSheets = usedRangeSheets + 1
                    AppendNameList usedRangeList, ws.Name
                End If
            End If
        Else
            hitSheets = hitSheets + 1
            regionSheets = regionSheets + 1
            AppendNameList regionList, ws.Name
        End If
NextSheet:
    Next ws

    summaryText = "命中=" & hitSheets & "；打印区域=" & regionSheets & "（" & regionList & "）"
    summaryText = summaryText & "；UsedRange回退=" & usedRangeSheets & "（" & usedRangeList & "）"
    summaryText = summaryText & "；自定义分页=" & customPagingSheets
    summaryText = summaryText & "；跳过=" & skipSheets & "（" & skipList & "）"
    Exit Sub

SafeExit:
    errNo = Err.Number
    errDesc = Err.Description
    If summaryText = "" Then
        summaryText = "预扫描降级：命中=" & hitSheets & "；打印区域=" & regionSheets & "；UsedRange回退=" & usedRangeSheets & "；自定义分页=" & customPagingSheets & "；跳过=" & skipSheets
    End If
    If errNo <> 0 Then
        summaryText = summaryText & "；预扫描异常=" & CStr(errNo) & " " & errDesc
    End If
    Err.Clear
End Sub

Private Sub ParseSheetPagingOptions(ByVal ws As Worksheet, ByRef fitToPagesWide As Long, ByRef fitToPagesTall As Long, ByRef hasCustomPaging As Boolean, ByRef pagingMessage As String)
    Dim commentText As String
    Dim valueText As String

    fitToPagesWide = DEFAULT_FIT_TO_PAGES_WIDE
    fitToPagesTall = DEFAULT_FIT_TO_PAGES_TALL
    hasCustomPaging = False
    pagingMessage = ""

    On Error GoTo SafeExit

    commentText = GetCellCommentText(ws.Range("A1"))
    If commentText = "" Then Exit Sub

    If TryParsePositiveAssignment(commentText, "FitToPagesWide", valueText) Then
        If IsNumeric(valueText) And CLng(valueText) > 0 Then
            fitToPagesWide = CLng(valueText)
            hasCustomPaging = True
        Else
            AppendParseMessage pagingMessage, "FitToPagesWide 非正整数"
        End If
    End If

    If TryParsePositiveAssignment(commentText, "FitToPagesTall", valueText) Then
        If IsNumeric(valueText) And CLng(valueText) > 0 Then
            fitToPagesTall = CLng(valueText)
            hasCustomPaging = True
        Else
            AppendParseMessage pagingMessage, "FitToPagesTall 非正整数"
        End If
    End If
    Exit Sub

SafeExit:
    fitToPagesWide = DEFAULT_FIT_TO_PAGES_WIDE
    fitToPagesTall = DEFAULT_FIT_TO_PAGES_TALL
    hasCustomPaging = False
    AppendParseMessage pagingMessage, "分页参数解析失败，已回退默认 1x1"
    Err.Clear
End Sub

Private Function TryParsePositiveAssignment(ByVal text As String, ByVal keyName As String, ByRef valueText As String) As Boolean
    Dim lines As Variant
    Dim oneLine As Variant
    Dim normalizedLine As String
    Dim hitPos As Long

    On Error GoTo SafeExit

    lines = Split(Replace$(text, vbCr, vbLf), vbLf)
    For Each oneLine In lines
        normalizedLine = Trim$(CStr(oneLine))
        If normalizedLine <> "" Then
            hitPos = InStr(1, LCase$(normalizedLine), LCase$(keyName) & "=", vbTextCompare)
            If hitPos = 1 Then
                valueText = Trim$(Mid$(normalizedLine, Len(keyName) + 2))
                TryParsePositiveAssignment = True
                Exit Function
            End If

            hitPos = InStr(1, LCase$(normalizedLine), LCase$(keyName) & " =", vbTextCompare)
            If hitPos = 1 Then
                valueText = Trim$(Mid$(normalizedLine, InStr(1, normalizedLine, "=", vbTextCompare) + 1))
                TryParsePositiveAssignment = True
                Exit Function
            End If
        End If
    Next oneLine
    Exit Function

SafeExit:
    valueText = ""
    TryParsePositiveAssignment = False
    Err.Clear
End Function

Private Function SafeFitPageValue(ByVal pageValue As Long) As Long
    If pageValue <= 0 Then
        SafeFitPageValue = 1
    Else
        SafeFitPageValue = pageValue
    End If
End Function

Private Function BuildPageWarnMessage(ByVal targetRange As Range, ByVal useLandscape As Boolean, ByVal fitToPagesWide As Long, ByVal fitToPagesTall As Long, ByVal useLightRelayout As Boolean) As String
    Dim cellCount As Double

    cellCount = CDbl(targetRange.Rows.Count) * CDbl(targetRange.Columns.Count)

    If fitToPagesWide = 1 And fitToPagesTall = 1 Then
        If targetRange.Rows.Count >= WARN_ROW_THRESHOLD Then
            BuildPageWarnMessage = "行数较多，单页输出可能导致字体过小"
            Exit Function
        End If
        If Not useLandscape And targetRange.Columns.Count >= WARN_COL_THRESHOLD Then
            BuildPageWarnMessage = "纵向列数较多，单页输出可能导致可读性下降"
            Exit Function
        End If
        If cellCount >= WARN_CELL_THRESHOLD Then
            BuildPageWarnMessage = "区域较大，单页输出可能导致页面过小"
            Exit Function
        End If
    End If

    If Not useLightRelayout And fitToPagesTall = 1 And targetRange.Rows.Count >= (WARN_ROW_THRESHOLD + 40) Then
        BuildPageWarnMessage = "保留源格式模式下行数较多，建议检查打印预览"
    End If
End Function

Private Function BuildSplitAdviceMessage(ByVal targetRange As Range, ByVal fitToPagesTall As Long) As String
    If IsLargeRange(targetRange) And fitToPagesTall = DEFAULT_FIT_TO_PAGES_TALL Then
        BuildSplitAdviceMessage = "超大表仍按单页高输出，建议在A1批注中设置 FitToPagesTall = 2 或更大"
    End If
End Function

Private Function GetEstimatedSheetPageCount(ByVal ws As Worksheet) As Long
    Dim hCount As Long
    Dim vCount As Long

    On Error Resume Next
    ws.DisplayPageBreaks = True
    hCount = ws.HPageBreaks.Count
    vCount = ws.VPageBreaks.Count
    ws.DisplayPageBreaks = False
    On Error GoTo 0

    GetEstimatedSheetPageCount = (hCount + 1) * (vCount + 1)
    If GetEstimatedSheetPageCount <= 0 Then
        GetEstimatedSheetPageCount = 1
    End If
    If Err.Number <> 0 Then
        Err.Clear
        GetEstimatedSheetPageCount = 1
    End If
End Function

Private Sub AppendNameList(ByRef listText As String, ByVal itemName As String)
    If Trim$(itemName) = "" Then Exit Sub
    If InStr(1, "；" & listText & "；", "；" & itemName & "；", vbTextCompare) > 0 Then Exit Sub
    If listText <> "" Then
        listText = listText & "；"
    End If
    listText = listText & itemName
End Sub

Private Sub NormalizeRegionBounds(ByRef sr As Long, ByRef er As Long, ByRef sc As Long, ByRef ec As Long)
    Dim tempValue As Long

    If sr > er Then
        tempValue = sr
        sr = er
        er = tempValue
    End If

    If sc > ec Then
        tempValue = sc
        sc = ec
        ec = tempValue
    End If
End Sub

Private Sub AppendParseMessage(ByRef parseMessage As String, ByVal oneMessage As String)
    If parseMessage <> "" Then
        parseMessage = parseMessage & "；"
    End If
    parseMessage = parseMessage & oneMessage
End Sub

Private Sub ExecuteConfigPrintExport(ByVal runMode As Long, Optional ByVal suppressSummary As Boolean = False, Optional ByVal includeAllModes As Boolean = False)
    Dim t0 As Double
    Dim logKey As String
    Dim wsCfg As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim enabled As Boolean
    Dim modeValue As Long
    Dim srcPath As String
    Dim srcSheetName As String
    Dim regionText As String
    Dim tgtPath As String
    Dim tgtSheetName As String
    Dim srcWb As Workbook
    Dim tgtWb As Workbook
    Dim srcWs As Worksheet
    Dim tgtWs As Worksheet
    Dim regions As Collection
    Dim regionInfo As Variant
    Dim parseMessage As String
    Dim wbCache As Object
    Dim openedByCode As Object
    Dim modified As Object
    Dim targetCleared As Object
    Dim targetNextRow As Object
    Dim targetFitWide As Object
    Dim targetFitTall As Object
    Dim targetWsMap As Object
    Dim targetOwnerMap As Object
    Dim targetKey As String
    Dim sourceKey As String
    Dim ownerSourceKey As String
    Dim nextRow As Long
    Dim oneAddedRows As Long
    Dim taskAddedRows As Long
    Dim fitToPagesWide As Long
    Dim fitToPagesTall As Long
    Dim wideSpecified As Boolean
    Dim tallSpecified As Boolean
    Dim hasCustomPaging As Boolean
    Dim pagingMessage As String
    Dim oneRegionIdx As Long
    Dim hitTasks As Long
    Dim skipTasks As Long
    Dim failTasks As Long
    Dim writtenRows As Long
    Dim writtenSheets As Long
    Dim pageWarnCount As Long
    Dim splitAdviceCount As Long
    Dim orientationText As String
    Dim pageWarnMessage As String
    Dim splitAdviceMessage As String
    Dim prevScreenUpdating As Boolean
    Dim prevDisplayAlerts As Boolean
    Dim prevEnableEvents As Boolean
    Dim prevCalc As XlCalculation
    Dim errNo As Long
    Dim errDesc As String

    t0 = Timer
    logKey = GetConfigPrintLogKey(runMode)
    RunLog_WriteRow logKey, "开始", "", "", "", "", "开始", ""

    Set wsCfg = FindPrintConfigSheet()
    If wsCfg Is Nothing Then
        初始化打印配置
        RunLog_WriteRow logKey, "结束", "", "", "", "跳过", "未找到打印配置，已自动初始化", CStr(Round(Timer - t0, 2))
        MsgBox "未找到打印配置，已自动初始化，请先填写后再执行。", vbExclamation, "按配置打印"
        Exit Sub
    End If

    lastRow = GetConfigLastRow(wsCfg)
    If lastRow < 2 Then
        RunLog_WriteRow logKey, "结束", "", "", "", "跳过", "打印配置为空", CStr(Round(Timer - t0, 2))
        MsgBox "打印配置为空，请先填写配置。", vbExclamation, "按配置打印"
        Exit Sub
    End If

    Set wbCache = CreateObject("Scripting.Dictionary")
    Set openedByCode = CreateObject("Scripting.Dictionary")
    Set modified = CreateObject("Scripting.Dictionary")
    Set targetCleared = CreateObject("Scripting.Dictionary")
    Set targetNextRow = CreateObject("Scripting.Dictionary")
    Set targetFitWide = CreateObject("Scripting.Dictionary")
    Set targetFitTall = CreateObject("Scripting.Dictionary")
    Set targetWsMap = CreateObject("Scripting.Dictionary")
    Set targetOwnerMap = CreateObject("Scripting.Dictionary")

    prevScreenUpdating = Application.ScreenUpdating
    prevDisplayAlerts = Application.DisplayAlerts
    prevEnableEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    SetAskToUpdateLinksSafe False
    Application.Calculation = xlCalculationManual
    On Error GoTo FailHandler

    For r = 2 To lastRow
        enabled = IsTruthyPrintConfig(wsCfg.Cells(r, 1).Value2)
        If Not enabled Then GoTo NextRow

        modeValue = ParsePrintModeValue(wsCfg.Cells(r, 2).Value2)
        If includeAllModes Then
            If modeValue < PRINT_MODE_KEEP Or modeValue > PRINT_MODE_FAST Then GoTo NextRow
        Else
            If modeValue <> runMode Then GoTo NextRow
        End If

        srcPath = NormalizePathText(CStr(wsCfg.Cells(r, 3).Value2))
        srcSheetName = Trim$(CStr(wsCfg.Cells(r, 4).Value2))
        regionText = Trim$(CStr(wsCfg.Cells(r, 5).Value2))
        tgtPath = NormalizePathText(CStr(wsCfg.Cells(r, 6).Value2))
        tgtSheetName = Trim$(CStr(wsCfg.Cells(r, 7).Value2))
        taskAddedRows = 0

        If srcPath = "" Or srcSheetName = "" Then
            skipTasks = skipTasks + 1
            RunLog_WriteRow logKey, "按工作表汇总", "第" & r & "行", srcSheetName, "", "跳过", "源工作簿或源工作表为空", ""
            GoTo NextRow
        End If

        If tgtPath = "" Or tgtSheetName = "" Then
            skipTasks = skipTasks + 1
            RunLog_WriteRow logKey, "按工作表汇总", "第" & r & "行", srcSheetName, "", "跳过", "目标工作簿或目标工作表为空", ""
            GoTo NextRow
        End If

        targetKey = BuildTargetConfigKeyByText(tgtPath, tgtSheetName)
        sourceKey = BuildSourceConfigKeyByText(srcPath, srcSheetName)
        If targetOwnerMap.Exists(targetKey) Then
            ownerSourceKey = CStr(targetOwnerMap.Item(targetKey))
            If StrComp(ownerSourceKey, sourceKey, vbTextCompare) <> 0 Then
                skipTasks = skipTasks + 1
                RunLog_WriteRow logKey, "按工作表汇总", "第" & r & "行", srcSheetName, tgtSheetName, "跳过", "多对一禁止：同一目标工作表已绑定其他源工作表", ""
                GoTo NextRow
            End If
        Else
            PutDictValuePrint targetOwnerMap, targetKey, sourceKey
        End If

        Set srcWb = AcquireWorkbookByPathPrint(srcPath, False, True, wbCache, openedByCode, parseMessage)
        If srcWb Is Nothing Then
            failTasks = failTasks + 1
            RunLog_WriteRow logKey, "按工作表汇总", srcPath, srcSheetName, "", "失败", "打开源工作簿失败：" & parseMessage, ""
            GoTo NextRow
        End If

        Set srcWs = GetWorksheetByExactNamePrint(srcWb, srcSheetName)
        If srcWs Is Nothing Then
            skipTasks = skipTasks + 1
            RunLog_WriteRow logKey, "按工作表汇总", srcWb.Name, srcSheetName, "", "跳过", "源工作表不存在（精确匹配）", ""
            GoTo NextRow
        End If

        Set tgtWb = AcquireWorkbookByPathPrint(tgtPath, True, False, wbCache, openedByCode, parseMessage)
        If tgtWb Is Nothing Then
            failTasks = failTasks + 1
            RunLog_WriteRow logKey, "按工作表汇总", srcWb.Name, srcSheetName, "", "失败", "打开目标工作簿失败：" & parseMessage, ""
            GoTo NextRow
        End If

        Set tgtWs = EnsureWorksheetExistsPrint(tgtWb, tgtSheetName)
        If tgtWs Is Nothing Then
            failTasks = failTasks + 1
            RunLog_WriteRow logKey, "按工作表汇总", srcWb.Name, srcSheetName, "", "失败", "创建/定位目标工作表失败", ""
            GoTo NextRow
        End If

        targetKey = BuildTargetSheetKeyPrint(tgtWb, tgtWs)
        If Not targetCleared.Exists(targetKey) Then
            tgtWs.Cells.Clear
            targetCleared.Add targetKey, True
            PutDictValuePrint targetNextRow, targetKey, 1
            PutDictObjectPrint targetWsMap, targetKey, tgtWs
            MarkModifiedPathPrint modified, tgtPath
        End If

        ParseConfigPagingOptions wsCfg, r, fitToPagesWide, fitToPagesTall, wideSpecified, tallSpecified, pagingMessage
        If Not targetFitWide.Exists(targetKey) Then PutDictValuePrint targetFitWide, targetKey, DEFAULT_FIT_TO_PAGES_WIDE
        If Not targetFitTall.Exists(targetKey) Then PutDictValuePrint targetFitTall, targetKey, DEFAULT_FIT_TO_PAGES_TALL
        If wideSpecified Then PutDictValuePrint targetFitWide, targetKey, fitToPagesWide
        If tallSpecified Then PutDictValuePrint targetFitTall, targetKey, fitToPagesTall
        If pagingMessage <> "" Then
            RunLog_WriteRow logKey, "按工作表汇总", srcWb.Name, srcWs.Name, tgtWs.Name, "提示", "分页参数忽略：" & pagingMessage, ""
        End If

        If regionText = "" Then
            Set regions = BuildUsedRangeRegions(srcWs)
            If regions Is Nothing Then
                skipTasks = skipTasks + 1
                RunLog_WriteRow logKey, "按工作表汇总", srcWb.Name, srcWs.Name, tgtWs.Name, "跳过", "打印区域为空且UsedRange为空", ""
                GoTo NextRow
            End If
        Else
            parseMessage = ""
            Set regions = ParseConfigPrintRegions(srcWs, regionText, parseMessage)
            If regions Is Nothing Then
                skipTasks = skipTasks + 1
                RunLog_WriteRow logKey, "按工作表汇总", srcWb.Name, srcWs.Name, tgtWs.Name, "跳过", "打印区域解析失败：" & parseMessage, ""
                GoTo NextRow
            End If
        End If

        nextRow = CLng(targetNextRow.Item(targetKey))
        oneRegionIdx = 0
        For Each regionInfo In regions
            oneRegionIdx = oneRegionIdx + 1
            oneAddedRows = AppendConfigRegionToTarget(srcWs, regionInfo, tgtWs, nextRow, runMode = PRINT_MODE_FAST, runMode = PRINT_MODE_LIGHT)
            If oneAddedRows > 0 Then
                taskAddedRows = taskAddedRows + oneAddedRows
                nextRow = nextRow + oneAddedRows
                If oneRegionIdx < regions.Count Then
                    nextRow = nextRow + 1
                End If
            End If
        Next regionInfo

        PutDictValuePrint targetNextRow, targetKey, nextRow
        If taskAddedRows > 0 Then
            hitTasks = hitTasks + 1
            writtenRows = writtenRows + taskAddedRows
            MarkModifiedPathPrint modified, tgtPath
            RunLog_WriteRow logKey, "按工作表汇总", srcWb.Name, srcWs.Name, tgtWs.Name, "成功", "写入行数=" & taskAddedRows & " 区域数=" & regions.Count, ""
        Else
            skipTasks = skipTasks + 1
            RunLog_WriteRow logKey, "按工作表汇总", srcWb.Name, srcWs.Name, tgtWs.Name, "跳过", "区域解析成功但无有效输出", ""
        End If

NextRow:
    Next r

    ApplyConfigPrintPageSetupByTarget targetWsMap, targetFitWide, targetFitTall, runMode = PRINT_MODE_LIGHT, logKey, writtenSheets, pageWarnCount, splitAdviceCount
    ApplySequentialFootersForModifiedBooks wbCache, modified, logKey
    SaveModifiedWorkbooksPrint wbCache, modified
    CloseOpenedWorkbooksPrint wbCache, openedByCode

    Application.Calculation = prevCalc
    SetAskToUpdateLinksSafe True
    Application.EnableEvents = prevEnableEvents
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating

    RunLog_WriteRow logKey, "结束", CStr(hitTasks), CStr(writtenSheets), CStr(writtenRows), "完成", "跳过任务=" & skipTasks & "，失败任务=" & failTasks & "，分页预警=" & pageWarnCount & "，建议拆页=" & splitAdviceCount, CStr(Round(Timer - t0, 2))
    If Not suppressSummary Then
        MsgBox BuildConfigPrintSummary(runMode, hitTasks, writtenSheets, writtenRows, skipTasks, failTasks, pageWarnCount, splitAdviceCount), vbInformation, "按配置打印"
    End If
    Exit Sub

FailHandler:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    SaveModifiedWorkbooksPrint wbCache, modified
    CloseOpenedWorkbooksPrint wbCache, openedByCode
    Application.Calculation = prevCalc
    SetAskToUpdateLinksSafe True
    Application.EnableEvents = prevEnableEvents
    Application.DisplayAlerts = prevDisplayAlerts
    Application.ScreenUpdating = prevScreenUpdating
    On Error GoTo 0
    RunLog_WriteRow logKey, "结束", "", "", "", "失败", CStr(errNo) & " " & errDesc, CStr(Round(Timer - t0, 2))
    If Not suppressSummary Then
        MsgBox "执行失败：" & CStr(errNo) & " " & errDesc, vbCritical, "按配置打印"
    End If
End Sub

Private Sub ExecuteConfigPrintPrecheck()
    Dim t0 As Double
    Dim wsCfg As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim enabled As Boolean
    Dim modeValue As Long
    Dim srcPath As String
    Dim srcSheetName As String
    Dim regionText As String
    Dim tgtPath As String
    Dim tgtSheetName As String
    Dim fitToPagesWide As Long
    Dim fitToPagesTall As Long
    Dim wideSpecified As Boolean
    Dim tallSpecified As Boolean
    Dim pagingMessage As String
    Dim wbCache As Object
    Dim openedByCode As Object
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim parseMessage As String
    Dim targetOwnerMap As Object
    Dim targetKey As String
    Dim sourceKey As String
    Dim okCount As Long
    Dim warnCount As Long
    Dim failCount As Long
    Dim skipCount As Long
    Dim msg As String
    Dim errNo As Long
    Dim errDesc As String

    t0 = Timer
    RunLog_WriteRow CFG_CHECK_LOG_KEY, "开始", "", "", "", "", "开始", ""

    Set wsCfg = FindPrintConfigSheet()
    If wsCfg Is Nothing Then
        RunLog_WriteRow CFG_CHECK_LOG_KEY, "结束", "", "", "", "跳过", "未找到打印配置", CStr(Round(Timer - t0, 2))
        MsgBox "未找到打印配置，请先初始化并填写。", vbExclamation, "按配置打印预校验"
        Exit Sub
    End If

    lastRow = GetConfigLastRow(wsCfg)
    If lastRow < 2 Then
        RunLog_WriteRow CFG_CHECK_LOG_KEY, "结束", "", "", "", "跳过", "打印配置为空", CStr(Round(Timer - t0, 2))
        MsgBox "打印配置为空。", vbExclamation, "按配置打印预校验"
        Exit Sub
    End If

    Set wbCache = CreateObject("Scripting.Dictionary")
    Set openedByCode = CreateObject("Scripting.Dictionary")
    Set targetOwnerMap = CreateObject("Scripting.Dictionary")
    On Error GoTo FailHandler

    For r = 2 To lastRow
        enabled = IsTruthyPrintConfig(wsCfg.Cells(r, 1).Value2)
        If Not enabled Then
            skipCount = skipCount + 1
            GoTo NextRow
        End If

        modeValue = ParsePrintModeValue(wsCfg.Cells(r, 2).Value2)
        If modeValue < PRINT_MODE_KEEP Or modeValue > PRINT_MODE_FAST Then
            failCount = failCount + 1
            RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", "", "", "失败", "打印模式仅支持1/2/3", ""
            GoTo NextRow
        End If

        srcPath = NormalizePathText(CStr(wsCfg.Cells(r, 3).Value2))
        srcSheetName = Trim$(CStr(wsCfg.Cells(r, 4).Value2))
        regionText = Trim$(CStr(wsCfg.Cells(r, 5).Value2))
        tgtPath = NormalizePathText(CStr(wsCfg.Cells(r, 6).Value2))
        tgtSheetName = Trim$(CStr(wsCfg.Cells(r, 7).Value2))

        If srcPath = "" Or srcSheetName = "" Or tgtPath = "" Or tgtSheetName = "" Then
            failCount = failCount + 1
            RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", srcSheetName, tgtSheetName, "失败", "源/目标关键字段为空", ""
            GoTo NextRow
        End If

        targetKey = BuildTargetConfigKeyByText(tgtPath, tgtSheetName)
        sourceKey = BuildSourceConfigKeyByText(srcPath, srcSheetName)
        If targetOwnerMap.Exists(targetKey) Then
            If StrComp(CStr(targetOwnerMap.Item(targetKey)), sourceKey, vbTextCompare) <> 0 Then
                failCount = failCount + 1
                RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", srcSheetName, tgtSheetName, "失败", "多对一禁止：同一目标工作表绑定了多个源工作表", ""
                GoTo NextRow
            End If
        Else
            PutDictValuePrint targetOwnerMap, targetKey, sourceKey
        End If

        ParseConfigPagingOptions wsCfg, r, fitToPagesWide, fitToPagesTall, wideSpecified, tallSpecified, pagingMessage
        If pagingMessage <> "" Then
            warnCount = warnCount + 1
            RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", srcSheetName, tgtSheetName, "提示", pagingMessage, ""
        End If

        Set srcWb = AcquireWorkbookByPathPrint(srcPath, False, True, wbCache, openedByCode, parseMessage)
        If srcWb Is Nothing Then
            failCount = failCount + 1
            RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", srcSheetName, "", "失败", "源工作簿不可打开：" & parseMessage, ""
            GoTo NextRow
        End If

        Set srcWs = GetWorksheetByExactNamePrint(srcWb, srcSheetName)
        If srcWs Is Nothing Then
            failCount = failCount + 1
            RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", srcSheetName, "", "失败", "源工作表不存在（精确匹配）", ""
            GoTo NextRow
        End If

        If regionText <> "" Then
            parseMessage = ""
            If ParseConfigPrintRegions(srcWs, regionText, parseMessage) Is Nothing Then
                failCount = failCount + 1
                RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", srcSheetName, "", "失败", "打印区域解析失败：" & parseMessage, ""
                GoTo NextRow
            End If
        End If

        If FileExistsPrint(tgtPath) Then
            If Not IsSupportedWorkbookFilePathPrint(tgtPath) Then
                failCount = failCount + 1
                RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", "", tgtSheetName, "失败", "目标工作簿扩展名不支持", ""
                GoTo NextRow
            End If
        Else
            If Not IsCreatableWorkbookPathPrint(tgtPath) Then
                failCount = failCount + 1
                RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", "", tgtSheetName, "失败", "目标工作簿路径不可创建", ""
                GoTo NextRow
            End If
            warnCount = warnCount + 1
            RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", "", tgtSheetName, "提示", "目标工作簿不存在，执行时会自动创建", ""
        End If

        okCount = okCount + 1
        RunLog_WriteRow CFG_CHECK_LOG_KEY, "校验", "第" & r & "行", srcSheetName, tgtSheetName, "成功", "通过", ""
NextRow:
    Next r

    CloseOpenedWorkbooksPrint wbCache, openedByCode
    msg = "按配置打印预校验完成。" & vbCrLf & "通过：" & okCount & vbCrLf & "提示：" & warnCount & vbCrLf & "失败：" & failCount & vbCrLf & "未启用跳过：" & skipCount
    RunLog_WriteRow CFG_CHECK_LOG_KEY, "结束", CStr(okCount), CStr(warnCount), CStr(failCount), "完成", "跳过=" & skipCount, CStr(Round(Timer - t0, 2))
    MsgBox msg, vbInformation, "按配置打印预校验"
    Exit Sub

FailHandler:
    errNo = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    CloseOpenedWorkbooksPrint wbCache, openedByCode
    On Error GoTo 0
    RunLog_WriteRow CFG_CHECK_LOG_KEY, "结束", "", "", "", "失败", CStr(errNo) & " " & errDesc, CStr(Round(Timer - t0, 2))
    MsgBox "预校验失败：" & CStr(errNo) & " " & errDesc, vbCritical, "按配置打印预校验"
End Sub

Private Function AppendConfigRegionToTarget(ByVal srcWs As Worksheet, ByVal regionInfo As Variant, ByVal tgtWs As Worksheet, ByVal startRow As Long, ByVal useFastMode As Boolean, ByVal useLightRelayout As Boolean) As Long
    Dim srcRange As Range
    Dim destTopLeft As Range
    Dim targetRange As Range
    Dim rowCount As Long
    Dim colCount As Long

    Set srcRange = srcWs.Range(srcWs.Cells(CLng(regionInfo(0)), CLng(regionInfo(2))), srcWs.Cells(CLng(regionInfo(1)), CLng(regionInfo(3))))
    rowCount = srcRange.Rows.Count
    colCount = srcRange.Columns.Count
    If rowCount <= 0 Or colCount <= 0 Then Exit Function

    Set destTopLeft = tgtWs.Cells(startRow, 1)
    Set targetRange = destTopLeft.Resize(rowCount, colCount)

    If useFastMode Then
        srcRange.Copy
        destTopLeft.PasteSpecial xlPasteFormats
        destTopLeft.PasteSpecial xlPasteColumnWidths
        Application.CutCopyMode = False

        CopyRowHeightsToStartRow srcWs, tgtWs, CLng(regionInfo(0)), rowCount, startRow
        CopyColumnWidths srcWs, tgtWs, CLng(regionInfo(2)), colCount
        srcRange.Copy
        targetRange.PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        SyncRangeColorAndFontRGB srcRange, targetRange
    Else
        srcRange.Copy
        destTopLeft.PasteSpecial xlPasteAll
        destTopLeft.PasteSpecial xlPasteColumnWidths
        Application.CutCopyMode = False

        CopyRowHeightsToStartRow srcWs, tgtWs, CLng(regionInfo(0)), rowCount, startRow
        CopyColumnWidths srcWs, tgtWs, CLng(regionInfo(2)), colCount
        srcRange.Copy
        targetRange.PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
        SyncRangeColorAndFontRGB srcRange, targetRange

        If useLightRelayout And Not IsLargeRange(targetRange) Then
            ApplyLightRelayoutToOffset tgtWs, targetRange, startRow
        End If
    End If

    AppendConfigRegionToTarget = rowCount
End Function

Private Sub ApplyConfigPrintPageSetupByTarget(ByVal targetWsMap As Object, ByVal targetFitWide As Object, ByVal targetFitTall As Object, ByVal useLightRelayout As Boolean, ByVal logKey As String, ByRef writtenSheets As Long, ByRef pageWarnCount As Long, ByRef splitAdviceCount As Long)
    Dim oneKey As Variant
    Dim ws As Worksheet
    Dim fr As Long
    Dim lr As Long
    Dim fc As Long
    Dim lc As Long
    Dim targetRange As Range
    Dim fitToPagesWide As Long
    Dim fitToPagesTall As Long
    Dim orientationText As String
    Dim pageWarnMessage As String
    Dim splitAdviceMessage As String

    If targetWsMap Is Nothing Then Exit Sub

    For Each oneKey In targetWsMap.Keys
        Set ws = targetWsMap.Item(CStr(oneKey))
        If ws Is Nothing Then GoTo NextSheet

        fr = GetFirstUsedRow(ws)
        lr = GetLastUsedRow(ws)
        fc = GetFirstUsedCol(ws)
        lc = GetLastUsedCol(ws)
        If fr <= 0 Or lr <= 0 Or fc <= 0 Or lc <= 0 Then GoTo NextSheet

        Set targetRange = ws.Range(ws.Cells(fr, fc), ws.Cells(lr, lc))
        fitToPagesWide = DEFAULT_FIT_TO_PAGES_WIDE
        fitToPagesTall = DEFAULT_FIT_TO_PAGES_TALL
        If targetFitWide.Exists(CStr(oneKey)) Then fitToPagesWide = CLng(targetFitWide.Item(CStr(oneKey)))
        If targetFitTall.Exists(CStr(oneKey)) Then fitToPagesTall = CLng(targetFitTall.Item(CStr(oneKey)))

        ApplyPageSetup ws, targetRange, useLightRelayout, fitToPagesWide, fitToPagesTall, orientationText, pageWarnMessage, splitAdviceMessage
        writtenSheets = writtenSheets + 1
        RunLog_WriteRow logKey, "按工作表汇总", ws.Parent.Name, ws.Name, orientationText, "成功", "页面设置完成", ""
        If pageWarnMessage <> "" Then
            pageWarnCount = pageWarnCount + 1
            RunLog_WriteRow logKey, "按工作表汇总", ws.Parent.Name, ws.Name, orientationText, "提示", pageWarnMessage, ""
        End If
        If splitAdviceMessage <> "" Then
            splitAdviceCount = splitAdviceCount + 1
            RunLog_WriteRow logKey, "按工作表汇总", ws.Parent.Name, ws.Name, "", "提示", splitAdviceMessage, ""
        End If
NextSheet:
    Next oneKey
End Sub

Private Function ParseConfigPrintRegions(ByVal ws As Worksheet, ByVal rawRegionText As String, ByRef parseMessage As String) As Collection
    Dim normalized As String
    Dim tokens() As String
    Dim i As Long
    Dim oneToken As String
    Dim oneRange As Range
    Dim result As Collection

    normalized = NormalizeRegionText(rawRegionText)
    If normalized = "" Then Exit Function

    tokens = Split(normalized, ";")
    Set result = New Collection
    For i = LBound(tokens) To UBound(tokens)
        oneToken = Trim$(tokens(i))
        If oneToken <> "" Then
            Set oneRange = Nothing
            On Error Resume Next
            Set oneRange = ws.Range(oneToken)
            On Error GoTo 0
            If oneRange Is Nothing Then
                AppendParseMessage parseMessage, "非法区域：" & oneToken
            Else
                result.Add Array(oneRange.Row, oneRange.Row + oneRange.Rows.Count - 1, oneRange.Column, oneRange.Column + oneRange.Columns.Count - 1, i + 1)
            End If
        End If
    Next i

    If parseMessage <> "" Then Exit Function
    If result.Count > 0 Then
        Set ParseConfigPrintRegions = result
    End If
End Function

Private Function NormalizeRegionText(ByVal textValue As String) As String
    Dim txt As String

    txt = Trim$(textValue)
    txt = Replace(txt, "；", ";")
    txt = Replace(txt, "，", ";")
    txt = Replace(txt, vbCr, ";")
    txt = Replace(txt, vbLf, ";")
    Do While InStr(txt, ";;") > 0
        txt = Replace(txt, ";;", ";")
    Loop
    If Left$(txt, 1) = ";" Then txt = Mid$(txt, 2)
    If Right$(txt, 1) = ";" Then txt = Left$(txt, Len(txt) - 1)
    NormalizeRegionText = txt
End Function

Private Function GetConfigPrintLogKey(ByVal runMode As Long) As String
    Select Case runMode
        Case PRINT_MODE_KEEP
            GetConfigPrintLogKey = CFG_KEEP_LOG_KEY
        Case PRINT_MODE_LIGHT
            GetConfigPrintLogKey = CFG_LIGHT_LOG_KEY
        Case 0
            GetConfigPrintLogKey = "3.10.7 按配置打印（执行全模式）"
        Case Else
            GetConfigPrintLogKey = CFG_FAST_LOG_KEY
    End Select
End Function

Private Function ParsePrintModeValue(ByVal rawValue As Variant) As Long
    Dim txt As String

    txt = Trim$(CStr(rawValue))
    If txt = "" Then
        ParsePrintModeValue = 0
    ElseIf IsNumeric(txt) Then
        ParsePrintModeValue = CLng(txt)
    End If
End Function

Private Sub ParseConfigPagingOptions(ByVal wsCfg As Worksheet, ByVal rowIndex As Long, ByRef fitToPagesWide As Long, ByRef fitToPagesTall As Long, ByRef wideSpecified As Boolean, ByRef tallSpecified As Boolean, ByRef pagingMessage As String)
    Dim txtWide As String
    Dim txtTall As String

    fitToPagesWide = DEFAULT_FIT_TO_PAGES_WIDE
    fitToPagesTall = DEFAULT_FIT_TO_PAGES_TALL
    wideSpecified = False
    tallSpecified = False
    pagingMessage = ""

    txtWide = Trim$(CStr(wsCfg.Cells(rowIndex, 8).Value2))
    txtTall = Trim$(CStr(wsCfg.Cells(rowIndex, 9).Value2))

    If txtWide <> "" Then
        If IsNumeric(txtWide) And CLng(txtWide) > 0 Then
            fitToPagesWide = CLng(txtWide)
            wideSpecified = True
        Else
            AppendParseMessage pagingMessage, "FitToPagesWide 非正整数"
        End If
    End If

    If txtTall <> "" Then
        If IsNumeric(txtTall) And CLng(txtTall) > 0 Then
            fitToPagesTall = CLng(txtTall)
            tallSpecified = True
        Else
            AppendParseMessage pagingMessage, "FitToPagesTall 非正整数"
        End If
    End If
End Sub

Private Function IsTruthyPrintConfig(ByVal rawValue As Variant) As Boolean
    Dim txt As String

    txt = UCase$(Trim$(CStr(rawValue)))
    If txt = "" Then
        IsTruthyPrintConfig = True
    ElseIf txt = "Y" Or txt = "YES" Or txt = "TRUE" Or txt = "1" Then
        IsTruthyPrintConfig = True
    Else
        IsTruthyPrintConfig = False
    End If
End Function

Private Function FindPrintConfigSheet() As Worksheet
    On Error Resume Next
    Set FindPrintConfigSheet = ThisWorkbook.Worksheets(PRINT_CONFIG_SHEET)
    On Error GoTo 0
End Function

Private Function EnsurePrintConfigSheet() As Worksheet
    Dim ws As Worksheet

    Set ws = FindPrintConfigSheet()
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = PRINT_CONFIG_SHEET
    End If
    Set EnsurePrintConfigSheet = ws
End Function

Private Sub InitPrintConfigHeader(ByVal ws As Worksheet)
    ClearPrintConfigComments ws

    ws.Cells(1, 1).Value = "是否启用"
    ws.Cells(1, 2).Value = "打印模式"
    ws.Cells(1, 3).Value = "源工作簿"
    ws.Cells(1, 4).Value = "源工作表"
    ws.Cells(1, 5).Value = "源工作表打印区域"
    ws.Cells(1, 6).Value = "目标工作簿"
    ws.Cells(1, 7).Value = "目标工作表"
    ws.Cells(1, 8).Value = "FitToPagesWide"
    ws.Cells(1, 9).Value = "FitToPagesTall"
    ws.Cells(1, 10).Value = "备注"

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(221, 235, 247)

    SetPrintConfigHeaderComment ws.Cells(1, 1), "Y/N，Y 表示启用。"
    SetPrintConfigHeaderComment ws.Cells(1, 2), "打印模式：1=保留原格式(3.10.4)；2=轻度重排(3.10.5)；3=快速复制(3.10.6)。"
    SetPrintConfigHeaderComment ws.Cells(1, 3), "源工作簿路径，支持绝对/相对路径。"
    SetPrintConfigHeaderComment ws.Cells(1, 4), "源工作表名，精确匹配。"
    SetPrintConfigHeaderComment ws.Cells(1, 5), "A1范围，支持多段：A1:H30;A35:H70。留空则回退UsedRange。"
    SetPrintConfigHeaderComment ws.Cells(1, 6), "目标工作簿路径。"
    SetPrintConfigHeaderComment ws.Cells(1, 7), "目标工作表名。同一轮首次命中会清空后追加。"
    SetPrintConfigHeaderComment ws.Cells(1, 8), "分页宽度页数，正整数。留空默认1。"
    SetPrintConfigHeaderComment ws.Cells(1, 9), "分页高度页数，正整数。留空默认1。"
    SetPrintConfigHeaderComment ws.Cells(1, 10), "备注。"
End Sub

Private Sub ClearPrintConfigComments(ByVal ws As Worksheet)
    Dim c As Long

    On Error Resume Next
    For c = 1 To 10
        If Not ws.Cells(1, c).Comment Is Nothing Then ws.Cells(1, c).Comment.Delete
        If Not ws.Cells(2, c).Comment Is Nothing Then ws.Cells(2, c).Comment.Delete
    Next c
    On Error GoTo 0
End Sub

Private Sub SetPrintConfigHeaderComment(ByVal targetCell As Range, ByVal commentText As String)
    On Error Resume Next
    If targetCell.Comment Is Nothing Then
        targetCell.AddComment commentText
    Else
        targetCell.Comment.Text text:=commentText
    End If
    On Error GoTo 0
End Sub

Private Sub WritePrintConfigExample(ByVal ws As Worksheet)
    If GetConfigLastRow(ws) > 1 Then Exit Sub

    ws.Cells(2, 1).Value = "N"
    ws.Cells(2, 2).Value = "1"
    ws.Cells(2, 3).Value = "C:\Users\AZI\Desktop\source_demo.xlsx"
    ws.Cells(2, 4).Value = "打印页"
    ws.Cells(2, 5).Value = "A1:H60;A65:H120"
    ws.Cells(2, 6).Value = "C:\Users\AZI\Desktop\target_demo.xlsx"
    ws.Cells(2, 7).Value = "打印汇总"
    ws.Cells(2, 8).Value = "1"
    ws.Cells(2, 9).Value = "2"
    ws.Cells(2, 10).Value = "示例：模式1按配置打印到目标表，分页1x2。"
End Sub

Private Function GetConfigLastRow(ByVal ws As Worksheet) As Long
    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0
    If lastCell Is Nothing Then
        GetConfigLastRow = 1
    Else
        GetConfigLastRow = lastCell.Row
    End If
End Function

Private Function NormalizePathText(ByVal rawText As String) As String
    Dim txt As String

    txt = Trim$(rawText)
    If txt = "" Then Exit Function
    If Left$(txt, 2) = "\\" Or (Len(txt) >= 2 And Mid$(txt, 2, 1) = ":") Then
        NormalizePathText = txt
    Else
        NormalizePathText = ThisWorkbook.Path & "\" & txt
    End If
    Do While Len(NormalizePathText) > 0 And Right$(NormalizePathText, 1) = "\"
        NormalizePathText = Left$(NormalizePathText, Len(NormalizePathText) - 1)
    Loop
End Function

Private Function AcquireWorkbookByPathPrint(ByVal rawPath As String, ByVal allowCreate As Boolean, ByVal readOnlyOpen As Boolean, ByVal wbCache As Object, ByVal openedByCode As Object, ByRef messageText As String) As Workbook
    Dim resolvedPath As String
    Dim wb As Workbook
    Dim parentFolder As String
    Dim fullPathKey As String
    Dim openErrNo As Long
    Dim openErrDesc As String

    messageText = ""
    resolvedPath = NormalizePathText(rawPath)
    If resolvedPath = "" Then
        messageText = "路径为空"
        Exit Function
    End If

    fullPathKey = LCase$(resolvedPath)
    If wbCache.Exists(fullPathKey) Then
        Set AcquireWorkbookByPathPrint = wbCache.Item(fullPathKey)
        Exit Function
    End If

    If FileExistsPrint(resolvedPath) Then
        If Not IsSupportedWorkbookFilePathPrint(resolvedPath) Then
            messageText = "文件类型不支持"
            Exit Function
        End If

        If Not TryOpenWorkbookCompatiblePrint(resolvedPath, readOnlyOpen, wb, openErrNo, openErrDesc) Then
            Set wb = Nothing
        End If

        If wb Is Nothing Then
            If openErrNo = 0 Then openErrNo = 1004
            If openErrDesc = "" Then openErrDesc = "打开工作簿失败"
            messageText = CStr(openErrNo) & " " & openErrDesc
            Exit Function
        End If

        PutDictObjectPrint wbCache, fullPathKey, wb
        PutDictValuePrint openedByCode, fullPathKey, True
        Set AcquireWorkbookByPathPrint = wb
        Exit Function
    End If

    If Not allowCreate Then
        messageText = "文件不存在"
        Exit Function
    End If

    If Not IsCreatableWorkbookPathPrint(resolvedPath) Then
        messageText = "目标文件扩展名不支持创建"
        Exit Function
    End If

    parentFolder = GetParentFolderPathPrint(resolvedPath)
    If parentFolder <> "" Then
        If Not IsDirectoryPathPrint(parentFolder) Then
            messageText = "目标目录不存在"
            Exit Function
        End If
    End If

    Set wb = Workbooks.Add(xlWBATWorksheet)
    PutDictObjectPrint wbCache, fullPathKey, wb
    PutDictValuePrint openedByCode, fullPathKey, True
    Set AcquireWorkbookByPathPrint = wb
End Function

Private Sub PutDictObjectPrint(ByVal dictObj As Object, ByVal oneKey As String, ByVal oneObj As Object)
    If dictObj Is Nothing Then Exit Sub
    If dictObj.Exists(oneKey) Then dictObj.Remove oneKey
    dictObj.Add oneKey, oneObj
End Sub

Private Sub PutDictValuePrint(ByVal dictObj As Object, ByVal oneKey As String, ByVal oneValue As Variant)
    If dictObj Is Nothing Then Exit Sub
    If dictObj.Exists(oneKey) Then dictObj.Remove oneKey
    dictObj.Add oneKey, oneValue
End Sub

Private Function TryOpenWorkbookCompatiblePrint(ByVal workbookPath As String, ByVal readOnlyOpen As Boolean, ByRef outWb As Workbook, ByRef outErrNo As Long, ByRef outErrDesc As String) As Boolean
    Dim objWb As Object

    Set outWb = Nothing
    outErrNo = 0
    outErrDesc = ""

    On Error Resume Next
    Set outWb = Workbooks.Open(workbookPath, ReadOnly:=readOnlyOpen, UpdateLinks:=0)
    outErrNo = Err.Number
    outErrDesc = "[open-1] " & Err.Description
    If Not outWb Is Nothing Then
        On Error GoTo 0
        TryOpenWorkbookCompatiblePrint = True
        Exit Function
    End If

    Err.Clear
    Set outWb = Application.Workbooks.Open(workbookPath)
    outErrNo = Err.Number
    outErrDesc = "[open-2] " & Err.Description
    If Not outWb Is Nothing Then
        On Error GoTo 0
        TryOpenWorkbookCompatiblePrint = True
        Exit Function
    End If

    Err.Clear
    Set objWb = GetObject(workbookPath)
    outErrNo = Err.Number
    outErrDesc = "[open-3] " & Err.Description
    If Not objWb Is Nothing Then
        Set outWb = objWb
        On Error GoTo 0
        TryOpenWorkbookCompatiblePrint = True
        Exit Function
    End If

    Err.Clear
    Set outWb = Workbooks.Open(Filename:=workbookPath, ReadOnly:=readOnlyOpen, UpdateLinks:=0)
    outErrNo = Err.Number
    outErrDesc = "[open-4] " & Err.Description
    If Not outWb Is Nothing Then
        On Error GoTo 0
        TryOpenWorkbookCompatiblePrint = True
        Exit Function
    End If

    On Error GoTo 0
    If outErrNo = 0 Then outErrNo = 1004
    If outErrDesc = "" Then outErrDesc = "打开工作簿失败"
End Function

Private Function FindOpenWorkbookByFullNamePrint(ByVal workbookPath As String) As Workbook
    Dim wb As Workbook
    Dim wbFullName As String

    For Each wb In Application.Workbooks
        wbFullName = GetWorkbookFullNameSafePrint(wb)
        If wbFullName <> "" Then
            If StrComp(wbFullName, workbookPath, vbTextCompare) = 0 Then
                Set FindOpenWorkbookByFullNamePrint = wb
                Exit Function
            End If
        End If
    Next wb
End Function

Private Function GetWorkbookFullNameSafePrint(ByVal wb As Workbook) As String
    On Error Resume Next
    GetWorkbookFullNameSafePrint = CStr(wb.FullName)
    If Len(GetWorkbookFullNameSafePrint) = 0 Then
        If Len(CStr(wb.Path)) > 0 Then
            GetWorkbookFullNameSafePrint = CStr(wb.Path) & "\" & CStr(wb.Name)
        Else
            GetWorkbookFullNameSafePrint = CStr(wb.Name)
        End If
    End If
    If Err.Number <> 0 Then
        Err.Clear
        GetWorkbookFullNameSafePrint = CStr(wb.Name)
    End If
    On Error GoTo 0
End Function

Private Function BuildTargetSheetKeyPrint(ByVal wb As Workbook, ByVal ws As Worksheet) As String
    BuildTargetSheetKeyPrint = LCase$(NormalizePathText(GetWorkbookFullNameSafePrint(wb))) & "|" & LCase$(ws.Name)
End Function

Private Sub MarkModifiedPathPrint(ByVal modified As Object, ByVal pathText As String)
    Dim normalizedPath As String

    If modified Is Nothing Then Exit Sub
    normalizedPath = LCase$(NormalizePathText(pathText))
    If normalizedPath = "" Then Exit Sub
    If Not modified.Exists(normalizedPath) Then modified.Add normalizedPath, True
End Sub

Private Sub SaveModifiedWorkbooksPrint(ByVal wbCache As Object, ByVal modified As Object)
    Dim oneKey As Variant
    Dim wb As Workbook

    If wbCache Is Nothing Then Exit Sub
    If modified Is Nothing Then Exit Sub

    For Each oneKey In modified.Keys
        If wbCache.Exists(CStr(oneKey)) Then
            Set wb = wbCache.Item(CStr(oneKey))
            If Not wb Is Nothing Then
                If Not wb.ReadOnly Then
                    If FileExistsPrint(CStr(oneKey)) Then
                        On Error Resume Next
                        wb.Save
                        On Error GoTo 0
                    Else
                        On Error Resume Next
                        wb.SaveAs fileName:=CStr(oneKey), FileFormat:=GetSaveFileFormatPrint(CStr(oneKey))
                        On Error GoTo 0
                    End If
                End If
            End If
        End If
    Next oneKey
End Sub

Private Function GetWorksheetByExactNamePrint(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheetByExactNamePrint = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function EnsureWorksheetExistsPrint(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet

    Set ws = GetWorksheetByExactNamePrint(wb, sheetName)
    If ws Is Nothing Then
        If wb.Worksheets.Count = 1 Then
            Set ws = wb.Worksheets(1)
            If IsSheetEffectivelyEmptyPrint(ws) Then
                On Error Resume Next
                ws.Name = sheetName
                On Error GoTo 0
                If StrComp(ws.Name, sheetName, vbTextCompare) = 0 Then
                    Set EnsureWorksheetExistsPrint = ws
                    Exit Function
                End If
            End If
        End If
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        On Error Resume Next
        ws.Name = sheetName
        On Error GoTo 0
    End If
    Set EnsureWorksheetExistsPrint = ws
End Function

Private Function IsSheetEffectivelyEmptyPrint(ByVal ws As Worksheet) As Boolean
    Dim usedRng As Range
    Dim cellCount As Double

    On Error Resume Next
    Set usedRng = ws.UsedRange
    On Error GoTo 0
    If usedRng Is Nothing Then
        IsSheetEffectivelyEmptyPrint = True
        Exit Function
    End If
    cellCount = usedRng.Cells.CountLarge
    If cellCount = 1 Then
        IsSheetEffectivelyEmptyPrint = (Trim$(CStr(usedRng.Cells(1, 1).Value2)) = "")
    Else
        On Error Resume Next
        IsSheetEffectivelyEmptyPrint = (Application.WorksheetFunction.CountA(usedRng) = 0)
        On Error GoTo 0
    End If
End Function

Private Function BuildTargetConfigKeyByText(ByVal targetPath As String, ByVal targetSheetName As String) As String
    BuildTargetConfigKeyByText = LCase$(NormalizePathText(targetPath)) & "|" & LCase$(Trim$(targetSheetName))
End Function

Private Function BuildSourceConfigKeyByText(ByVal sourcePath As String, ByVal sourceSheetName As String) As String
    BuildSourceConfigKeyByText = LCase$(NormalizePathText(sourcePath)) & "|" & LCase$(Trim$(sourceSheetName))
End Function

Private Sub CloseOpenedWorkbooksPrint(ByVal wbCache As Object, ByVal openedByCode As Object)
    Dim oneKey As Variant
    Dim wb As Workbook

    If wbCache Is Nothing Then Exit Sub
    If openedByCode Is Nothing Then Exit Sub

    For Each oneKey In wbCache.Keys
        If openedByCode.Exists(CStr(oneKey)) Then
            If CBool(openedByCode.Item(CStr(oneKey))) Then
                Set wb = wbCache.Item(CStr(oneKey))
                If Not wb Is Nothing Then
                    On Error Resume Next
                    wb.Close SaveChanges:=False
                    On Error GoTo 0
                End If
            End If
        End If
    Next oneKey
End Sub

Private Function IsSupportedWorkbookFilePathPrint(ByVal filePath As String) As Boolean
    Dim ext As String
    Dim dotPos As Long

    dotPos = InStrRev(filePath, ".")
    If dotPos <= 0 Then Exit Function
    ext = LCase$(Mid$(filePath, dotPos + 1))
    Select Case ext
        Case "xls", "xlsx", "xlsm", "xlsb", "csv"
            IsSupportedWorkbookFilePathPrint = True
    End Select
End Function

Private Function IsCreatableWorkbookPathPrint(ByVal filePath As String) As Boolean
    Dim ext As String
    Dim dotPos As Long

    dotPos = InStrRev(filePath, ".")
    If dotPos <= 0 Then
        IsCreatableWorkbookPathPrint = True
        Exit Function
    End If
    ext = LCase$(Mid$(filePath, dotPos + 1))
    Select Case ext
        Case "xls", "xlsx", "xlsm", "xlsb"
            IsCreatableWorkbookPathPrint = True
    End Select
End Function

Private Function IsDirectoryPathPrint(ByVal pathText As String) As Boolean
    Dim attrValue As Long

    If Trim$(pathText) = "" Then Exit Function
    On Error Resume Next
    attrValue = GetAttr(pathText)
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    IsDirectoryPathPrint = ((attrValue And vbDirectory) = vbDirectory)
End Function

Private Function FileExistsPrint(ByVal filePath As String) As Boolean
    If Trim$(filePath) = "" Then Exit Function
    If IsDirectoryPathPrint(filePath) Then Exit Function
    On Error Resume Next
    FileExistsPrint = (Len(Dir(filePath, vbNormal Or vbReadOnly Or vbHidden Or vbSystem)) > 0)
    On Error GoTo 0
End Function

Private Function GetParentFolderPathPrint(ByVal filePath As String) As String
    Dim p As Long

    p = InStrRev(filePath, "\")
    If p > 0 Then GetParentFolderPathPrint = Left$(filePath, p - 1)
End Function

Private Function GetSaveFileFormatPrint(ByVal workbookPath As String) As Long
    Dim ext As String

    ext = LCase$(Mid$(workbookPath, InStrRev(workbookPath, ".") + 1))
    Select Case ext
        Case "xls"
            GetSaveFileFormatPrint = 56
        Case "xlsx"
            GetSaveFileFormatPrint = 51
        Case "xlsm"
            GetSaveFileFormatPrint = 52
        Case "xlsb"
            GetSaveFileFormatPrint = 50
        Case Else
            GetSaveFileFormatPrint = 51
    End Select
End Function

Private Sub CopyRowHeightsToStartRow(ByVal srcWs As Worksheet, ByVal outWs As Worksheet, ByVal srcStartRow As Long, ByVal rowCount As Long, ByVal destStartRow As Long)
    Dim idx As Long

    For idx = 1 To rowCount
        outWs.Rows(destStartRow + idx - 1).RowHeight = srcWs.Rows(srcStartRow + idx - 1).RowHeight
    Next idx
End Sub

Private Sub ApplySequentialFootersForModifiedBooks(ByVal wbCache As Object, ByVal modified As Object, ByVal logKey As String)
    Dim oneKey As Variant
    Dim wb As Workbook

    If wbCache Is Nothing Then Exit Sub
    If modified Is Nothing Then Exit Sub

    For Each oneKey In modified.Keys
        If wbCache.Exists(CStr(oneKey)) Then
            Set wb = wbCache.Item(CStr(oneKey))
            If Not wb Is Nothing Then
                ApplySequentialSheetFooters wb, logKey, wb.Name
            End If
        End If
    Next oneKey
End Sub

Private Sub ApplyLightRelayoutToOffset(ByVal ws As Worksheet, ByVal targetRange As Range, ByVal startRow As Long)
    Dim headerRows As Long
    Dim firstDataRow As Long
    Dim col As Range
    Dim endRow As Long
    Dim endCol As Long

    endRow = startRow + targetRange.Rows.Count - 1
    endCol = targetRange.Columns.Count

    targetRange.WrapText = True
    targetRange.VerticalAlignment = xlCenter
    targetRange.Font.Name = "宋体"
    targetRange.Font.Size = 10

    headerRows = ResolveHeaderRows(targetRange)
    If headerRows < 1 Then headerRows = 1

    With ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + headerRows - 1, endCol))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    ApplyHeaderShadeOnlyForNoFill ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + headerRows - 1, endCol)), RGB(242, 242, 242)

    ws.Rows(startRow).Font.Size = 14
    If headerRows >= 2 Then
        ws.Range(ws.Cells(startRow + 1, 1), ws.Cells(startRow + headerRows - 1, endCol)).Font.Size = 11
    End If

    If Not IsLargeRange(targetRange) Then
        For Each col In ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, endCol)).Columns
            col.EntireColumn.AutoFit
            If col.ColumnWidth > 20 Then
                col.ColumnWidth = 20
            End If
        Next col
        ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, endCol)).Rows.AutoFit
    End If

    firstDataRow = startRow + headerRows
    If firstDataRow <= endRow Then
        ws.Range(ws.Cells(firstDataRow, 1), ws.Cells(endRow, endCol)).Font.Size = 10
    End If
End Sub

Private Sub ApplyHeaderShadeOnlyForNoFill(ByVal headerRange As Range, ByVal shadeColor As Long)
    Dim cell As Range

    If headerRange Is Nothing Then Exit Sub

    On Error Resume Next
    For Each cell In headerRange.Cells
        If cell.Interior.Pattern = xlPatternNone Or cell.Interior.ColorIndex = xlColorIndexNone Then
            cell.Interior.Pattern = xlPatternSolid
            cell.Interior.Color = shadeColor
        End If
    Next cell
    On Error GoTo 0
End Sub

Private Function BuildConfigPrintSummary(ByVal runMode As Long, ByVal hitTasks As Long, ByVal writtenSheets As Long, ByVal writtenRows As Long, ByVal skipTasks As Long, ByVal failTasks As Long, ByVal pageWarnCount As Long, ByVal splitAdviceCount As Long) As String
    If runMode = 0 Then
        BuildConfigPrintSummary = "按配置打印完成（全模式）。"
    Else
        BuildConfigPrintSummary = "按配置打印完成（模式" & runMode & "）。"
    End If
    BuildConfigPrintSummary = BuildConfigPrintSummary & vbCrLf & "成功任务数：" & hitTasks
    BuildConfigPrintSummary = BuildConfigPrintSummary & vbCrLf & "输出工作表数：" & writtenSheets
    BuildConfigPrintSummary = BuildConfigPrintSummary & vbCrLf & "写入总行数：" & writtenRows
    BuildConfigPrintSummary = BuildConfigPrintSummary & vbCrLf & "跳过任务数：" & skipTasks
    BuildConfigPrintSummary = BuildConfigPrintSummary & vbCrLf & "失败任务数：" & failTasks
    BuildConfigPrintSummary = BuildConfigPrintSummary & vbCrLf & "分页预警：" & pageWarnCount
    BuildConfigPrintSummary = BuildConfigPrintSummary & vbCrLf & "建议拆页：" & splitAdviceCount
End Function

Private Function GetCellCommentText(ByVal cell As Range) As String
    On Error Resume Next
    If cell Is Nothing Then
        Exit Function
    End If

    If Not cell.Comment Is Nothing Then
        GetCellCommentText = cell.Comment.Text
        If GetCellCommentText = "" Then
            GetCellCommentText = cell.Comment.Comment.Text
        End If
    End If
    On Error GoTo 0
End Function

Private Function GetPrintLogKey(ByVal useLightRelayout As Boolean) As String
    If useLightRelayout Then
        GetPrintLogKey = LIGHT_LOG_KEY
    Else
        GetPrintLogKey = KEEP_LOG_KEY
    End If
End Function

Private Function IsLargeRange(ByVal targetRange As Range) As Boolean
    Dim cellCount As Double

    cellCount = CDbl(targetRange.Rows.Count) * CDbl(targetRange.Columns.Count)
    IsLargeRange = (cellCount >= LARGE_RANGE_CELL_THRESHOLD Or _
                    targetRange.Columns.Count >= LARGE_RANGE_COLUMN_THRESHOLD Or _
                    targetRange.Rows.Count >= LARGE_RANGE_ROW_THRESHOLD)
End Function

Private Function CanUsePrintCommunication() As Boolean
    On Error Resume Next
    Application.PrintCommunication = Application.PrintCommunication
    CanUsePrintCommunication = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Function GetFirstUsedRow(ByVal ws As Worksheet) As Long
    Dim rng As Range

    On Error Resume Next
    Set rng = ws.Cells.Find("*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    On Error GoTo 0

    If rng Is Nothing Then
        GetFirstUsedRow = 0
    Else
        GetFirstUsedRow = rng.Row
    End If
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

Private Function GetFirstUsedCol(ByVal ws As Worksheet) As Long
    Dim rng As Range

    On Error Resume Next
    Set rng = ws.Cells.Find("*", LookIn:=xlValues, SearchOrder:=xlByColumns, SearchDirection:=xlNext)
    On Error GoTo 0

    If rng Is Nothing Then
        GetFirstUsedCol = 0
    Else
        GetFirstUsedCol = rng.Column
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

Private Sub SafeCloseWorkbook(ByRef wb As Workbook, ByVal saveChanges As Boolean)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=saveChanges
        Set wb = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub SetAskToUpdateLinksSafe(ByVal enabled As Boolean)
    On Error Resume Next
    Application.AskToUpdateLinks = enabled
    On Error GoTo 0
End Sub
