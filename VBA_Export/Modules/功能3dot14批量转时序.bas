Attribute VB_Name = "功能3dot14批量转时序"







Option Explicit















Private Const LOG_KEY As String = "3.8 批量转时序数据"







Private Const RULE_SHEET_NAME As String = "时序提取规则"







Private Const RESULT_SHEET_NAME As String = "时序提取结果"







Private Const MAP_SHEET_NAME As String = "路径标准化映射"







Private Const COMPARE_LOG_KEY As String = "3.8.2 表头路径比对（按位置）"







Private Const PATH_COMPARE_LOG_KEY As String = "3.8.3 表头路径比对（按路径）"







Private Const HYBRID_COMPARE_LOG_KEY As String = "3.8.4 表头路径比对（位置后路径）"







Private Const FAST_TIMELINE_LOG_KEY As String = "3.8.5 提取时序（快速）"







Private Const FAST_TIMELINE_SLIM_LOG_KEY As String = "3.8.6 提取时序（快速精简）"







Private Const WIDE_SUMMARY_LOG_KEY As String = "3.8.7 规则汇总（宽表）"







Private Const COMPARE_RESULT_SHEET_NAME As String = "表头比对结果"







Private Const WIDE_RESULT_SHEET_NAME As String = "规则汇总结果"







Private Const TIMELINE_FAST_FLUSH_SIZE As Long = 1000



Private Const PRECHECK_LOG_KEY As String = "3.9.8 配置预校验"

Private Const WIDE_DYNAMIC_COL_WIDTH As Double = 14#









Private gTimelineFastMode As Boolean







Private gTimelineFastSlimMode As Boolean







Private gTimelineFastBuffer As Collection







Private gTimelineFastSlimBuffer As Collection







Private gTimelineFastNextRow As Long



Private gRuleValidationReady As Boolean



Private gRuleInvalidRows As Object



Private gRuleWarnRows As Object



Private gValidationRuleTotal As Long



Private gValidationRuleEnabled As Long



Private gValidationRuleValid As Long



Private gValidationRuleInvalid As Long



Private gValidationRuleWarn As Long



Private gValidationMapInvalid As Long



Private gValidationMapWarn As Long



Private gPrecheckRunning As Boolean

Private gWideVerboseLogEnabled As Boolean









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















Private Enum MapCols







    mcEnabled = 1







    mcMapName = 2







    mcRuleName = 3







    mcBookKeywords = 4







    mcSheetKeywords = 5







    mcTargetType = 6







    mcMatchMode = 7







    mcOriginalPath = 8







    mcStandardPath = 9







    mcRemark = 10







End Enum















Private Enum CompareResultCols







    crcEnabled = 1







    crcMapName = 2







    crcRuleName = 3







    crcBookKeywords = 4







    crcSheetKeywords = 5







    crcTargetType = 6







    crcMatchMode = 7







    crcOriginalPath = 8







    crcStandardPath = 9







    crcRemark = 10







    crcSourceBook = 11







    crcSourceSheet = 12







    crcTemplatePosition = 13







    crcSourcePosition = 14







    crcCompareResult = 15







End Enum















Private Enum WideResultFixedCols







    wrSourceBook = 1







    wrSourceSheet = 2







    wrDataDate = 3







    wrRowPath = 4







End Enum















Public Sub 初始化时序提取配置()







    Dim wsRule As Worksheet















    Set wsRule = EnsureRuleSheet()







    InitRuleHeader wsRule















    MsgBox "时序提取配置表头已更新。", vbInformation







End Sub















Public Sub 初始化路径标准化映射()







    Dim wsMap As Worksheet















    Set wsMap = EnsurePathMapSheet()







    InitPathMapHeader wsMap















    MsgBox "路径标准化映射表头已更新。", vbInformation







End Sub















Public Sub Execute3dot9Precheck()



    Dim summaryText As String







    Run3dot9Precheck True, summaryText



    MsgBox summaryText, vbInformation, "3.9.8 配置预校验"



End Sub







Private Sub Run3dot9Precheck(ByVal writeLog As Boolean, ByRef summaryText As String)



    Dim wsRule As Worksheet



    Dim wsMap As Worksheet



    Dim lastRuleRow As Long



    Dim lastMapRow As Long



    Dim rowNo As Long



    Dim errMsg As String







    On Error GoTo CleanFail







    If gPrecheckRunning Then



        summaryText = Build3dot9PrecheckSummaryText()



        Exit Sub



    End If



    gPrecheckRunning = True







    Set gRuleInvalidRows = CreateObject("Scripting.Dictionary")



    gRuleInvalidRows.CompareMode = vbTextCompare



    Set gRuleWarnRows = CreateObject("Scripting.Dictionary")



    gRuleWarnRows.CompareMode = vbTextCompare







    gValidationRuleTotal = 0



    gValidationRuleEnabled = 0



    gValidationRuleValid = 0



    gValidationRuleInvalid = 0



    gValidationRuleWarn = 0



    gValidationMapInvalid = 0



    gValidationMapWarn = 0







    Set wsRule = EnsureRuleSheet()



    InitRuleHeader wsRule



    lastRuleRow = GetLastUsedRow(wsRule)







    For rowNo = 2 To lastRuleRow



        If IsRuleRowMeaningful(wsRule, rowNo) Then



            gValidationRuleTotal = gValidationRuleTotal + 1



            If IsEnabledValue(wsRule.Cells(rowNo, rcEnabled).Value) Then



                gValidationRuleEnabled = gValidationRuleEnabled + 1



                If ValidateOneRuleRow(wsRule, rowNo, errMsg) Then



                    gValidationRuleValid = gValidationRuleValid + 1



                Else



                    gValidationRuleInvalid = gValidationRuleInvalid + 1



                    gRuleInvalidRows(CStr(rowNo)) = errMsg



                    If writeLog Then



                        RunLog_WriteRow PRECHECK_LOG_KEY, "规则错误", CStr(rowNo), "", "", "失败", errMsg, ""



                    End If



                End If



            End If



        End If



    Next rowNo







    Set wsMap = EnsurePathMapSheet()



    InitPathMapHeader wsMap



    lastMapRow = GetLastUsedRow(wsMap)







    For rowNo = 2 To lastMapRow



        If IsMapRowMeaningful(wsMap, rowNo) Then



            ValidateOneMapRow wsMap, rowNo, writeLog



        End If



    Next rowNo







    gRuleValidationReady = True



    summaryText = Build3dot9PrecheckSummaryText()



    If writeLog Then



        RunLog_WriteRow PRECHECK_LOG_KEY, "结果", "", "", "", "完成", summaryText, ""



    End If



    gPrecheckRunning = False



    Exit Sub



CleanFail:



    gPrecheckRunning = False



    summaryText = "预校验失败：" & Err.Number & " " & Err.Description



    If writeLog Then



        RunLog_WriteRow PRECHECK_LOG_KEY, "结果", "", "", "", "失败", summaryText, ""



    End If



End Sub







Private Sub Refresh3dot9ValidationState()



    Dim summaryText As String







    Run3dot9Precheck False, summaryText



End Sub







Private Function Build3dot9PrecheckSummaryText() As String



    Build3dot9PrecheckSummaryText = "规则总数=" & gValidationRuleTotal



    Build3dot9PrecheckSummaryText = Build3dot9PrecheckSummaryText & vbCrLf & "启用规则=" & gValidationRuleEnabled



    Build3dot9PrecheckSummaryText = Build3dot9PrecheckSummaryText & vbCrLf & "有效规则=" & gValidationRuleValid



    Build3dot9PrecheckSummaryText = Build3dot9PrecheckSummaryText & vbCrLf & "规则错误=" & gValidationRuleInvalid



    Build3dot9PrecheckSummaryText = Build3dot9PrecheckSummaryText & vbCrLf & "规则警告=" & gValidationRuleWarn



    Build3dot9PrecheckSummaryText = Build3dot9PrecheckSummaryText & vbCrLf & "映射错误=" & gValidationMapInvalid



    Build3dot9PrecheckSummaryText = Build3dot9PrecheckSummaryText & vbCrLf & "映射警告=" & gValidationMapWarn



End Function







Private Function IsRuleRowMeaningful(ByVal wsRule As Worksheet, ByVal rowNo As Long) As Boolean



    Dim colNo As Long







    For colNo = rcEnabled To rcRemark



        If Trim$(CStr(wsRule.Cells(rowNo, colNo).Value)) <> "" Then



            IsRuleRowMeaningful = True



            Exit Function



        End If



    Next colNo



End Function







Private Function ValidateOneRuleRow(ByVal wsRule As Worksheet, ByVal rowNo As Long, ByRef errMsg As String) As Boolean



    Dim colHeaderRows As Collection



    Dim rowHeaderCols As Collection



    Dim invalidCount As Long



    Dim dataStartRow As Long



    Dim dataEndRow As Long



    Dim dataStartCol As Long



    Dim dataEndCol As Long



    Dim rawStartRow As String



    Dim rawEndRow As String



    Dim rawStartCol As String



    Dim rawEndCol As String







    errMsg = ""



    rawStartRow = Trim$(CStr(wsRule.Cells(rowNo, rcStartRow).Value))



    rawEndRow = Trim$(CStr(wsRule.Cells(rowNo, rcEndRow).Value))



    rawStartCol = Trim$(CStr(wsRule.Cells(rowNo, rcStartCol).Value))



    rawEndCol = Trim$(CStr(wsRule.Cells(rowNo, rcEndCol).Value))







    invalidCount = 0



    Set colHeaderRows = ParseNumberCollection(CStr(wsRule.Cells(rowNo, rcColHeaderRows).Value), invalidCount)



    If colHeaderRows.Count = 0 Then



        errMsg = AppendErrText(errMsg, "列表头行为空或格式非法")



    End If



    If invalidCount > 0 Then



        errMsg = AppendErrText(errMsg, "列表头行存在非法项")



    End If







    invalidCount = 0



    Set rowHeaderCols = ParseColumnCollection(CStr(wsRule.Cells(rowNo, rcRowHeaderCols).Value), invalidCount)



    If rowHeaderCols.Count = 0 Then



        errMsg = AppendErrText(errMsg, "行头列为空或格式非法")



    End If



    If invalidCount > 0 Then



        errMsg = AppendErrText(errMsg, "行头列存在非法项")



    End If







    If Not TryParseLongStrict(rawStartRow, dataStartRow) Or dataStartRow <= 0 Then



        errMsg = AppendErrText(errMsg, "数据起始行非法")



    End If



    If rawEndRow <> "" Then



        If Not TryParseLongStrict(rawEndRow, dataEndRow) Then



            errMsg = AppendErrText(errMsg, "数据结束行非法")



        End If



    Else



        dataEndRow = 0



    End If







    If Not TryParseColumnStrict(rawStartCol, dataStartCol) Or dataStartCol <= 0 Then



        errMsg = AppendErrText(errMsg, "数据起始列非法")



    End If



    If rawEndCol <> "" Then



        If Not TryParseColumnStrict(rawEndCol, dataEndCol) Then



            errMsg = AppendErrText(errMsg, "数据结束列非法")



        End If



    Else



        dataEndCol = 0



    End If







    If dataEndRow > 0 And dataStartRow > 0 Then



        If dataEndRow < dataStartRow Then



            errMsg = AppendErrText(errMsg, "结束行小于起始行")



        End If



    End If







    If dataEndCol > 0 And dataStartCol > 0 Then



        If dataEndCol < dataStartCol Then



            errMsg = AppendErrText(errMsg, "结束列小于起始列")



        End If



    End If







    ValidateOneRuleRow = (errMsg = "")



End Function







Private Function AppendErrText(ByVal baseText As String, ByVal appendText As String) As String



    If baseText = "" Then



        AppendErrText = appendText



    Else



        AppendErrText = baseText & "；" & appendText



    End If



End Function







Private Function IsMapRowMeaningful(ByVal wsMap As Worksheet, ByVal rowNo As Long) As Boolean



    Dim colNo As Long







    For colNo = mcEnabled To mcRemark



        If Trim$(CStr(wsMap.Cells(rowNo, colNo).Value)) <> "" Then



            IsMapRowMeaningful = True



            Exit Function



        End If



    Next colNo



End Function







Private Sub ValidateOneMapRow(ByVal wsMap As Worksheet, ByVal rowNo As Long, ByVal writeLog As Boolean)



    Dim targetType As String



    Dim matchMode As String



    Dim originalPath As String



    Dim standardPath As String



    Dim errMsg As String







    If Not IsEnabledValue(wsMap.Cells(rowNo, mcEnabled).Value) Then Exit Sub







    targetType = LCase$(Trim$(CStr(wsMap.Cells(rowNo, mcTargetType).Value)))



    matchMode = LCase$(Trim$(CStr(wsMap.Cells(rowNo, mcMatchMode).Value)))



    originalPath = Trim$(CStr(wsMap.Cells(rowNo, mcOriginalPath).Value))



    standardPath = Trim$(CStr(wsMap.Cells(rowNo, mcStandardPath).Value))



    errMsg = ""







    If originalPath = "" Then



        errMsg = AppendErrText(errMsg, "原始路径为空")



    End If



    If standardPath = "" Then



        errMsg = AppendErrText(errMsg, "标准路径为空")



    End If



    If targetType = "" Then



        errMsg = AppendErrText(errMsg, "目标类型为空")



    End If



    If matchMode = "" Then



        errMsg = AppendErrText(errMsg, "匹配方式为空")



    End If







    If targetType <> "" Then



        If InStr(targetType, "行头") = 0 And InStr(targetType, "列头") = 0 And targetType <> "row" And targetType <> "column" Then



            gValidationMapWarn = gValidationMapWarn + 1



            If writeLog Then



                RunLog_WriteRow PRECHECK_LOG_KEY, "映射警告", CStr(rowNo), originalPath, standardPath, "警告", "目标类型非标准值：" & targetType, ""



            End If



        End If



    End If







    If matchMode <> "" Then



        If InStr(matchMode, "精确") = 0 And InStr(matchMode, "包含") = 0 And matchMode <> "exact" And matchMode <> "contains" Then



            gValidationMapWarn = gValidationMapWarn + 1



            If writeLog Then



                RunLog_WriteRow PRECHECK_LOG_KEY, "映射警告", CStr(rowNo), originalPath, standardPath, "警告", "匹配方式非标准值：" & matchMode, ""



            End If



        End If



    End If







    If errMsg <> "" Then



        gValidationMapInvalid = gValidationMapInvalid + 1



        If writeLog Then



            RunLog_WriteRow PRECHECK_LOG_KEY, "映射错误", CStr(rowNo), originalPath, standardPath, "失败", errMsg, ""



        End If



    End If



End Sub







Public Sub 执行表头位置比对()



    执行表头路径比对







End Sub















Public Sub 执行表头路径比对()







    Dim t0 As Double







    Dim wsRule As Worksheet







    Dim compareWb As Workbook







    Dim compareWs As Worksheet







    Dim templatePath As String







    Dim templateWb As Workbook







    Dim fd As FileDialog







    Dim fileItem As Variant







    Dim sourceWb As Workbook







    Dim lastRuleRow As Long







    Dim ruleRow As Long







    Dim resultRow As Long







    Dim hitBooks As Long







    Dim hitSheets As Long







    Dim diffRows As Long







    Dim sameRows As Long







    Dim skipRules As Long







    Dim skipBooks As Long







    Dim errNo As Long







    Dim errDesc As String















    t0 = Timer







    RunLog_WriteRow COMPARE_LOG_KEY, "开始", "", "", "", "", "开始", ""















    Set wsRule = EnsureRuleSheet()







    InitRuleHeader wsRule


















    templatePath = PickOneWorkbookPath("请选择作为模板的工作簿")







    If templatePath = "" Then







        RunLog_WriteRow COMPARE_LOG_KEY, "结束", "", "", "", "", "用户取消模板选择", CStr(Round(Timer - t0, 2))







        Exit Sub







    End If















    Set fd = Application.FileDialog(msoFileDialogFilePicker)







    With fd







        .Title = "请选择要进行表头比对的源工作簿，可多选"







        .AllowMultiSelect = True







        .Filters.Clear







        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"







        If .Show <> -1 Then







            RunLog_WriteRow COMPARE_LOG_KEY, "结束", "", "", "", "", "用户取消源文件选择", CStr(Round(Timer - t0, 2))







            Exit Sub







        End If







    End With















    Application.ScreenUpdating = False







    Application.DisplayAlerts = False







    On Error GoTo CompareErrHandler















    Set templateWb = Workbooks.Open(templatePath, ReadOnly:=True, UpdateLinks:=0)







    Set compareWb = CreateResultWorkbook(COMPARE_RESULT_SHEET_NAME, compareWs)







    InitCompareResultHeader compareWs







    resultRow = 2







    lastRuleRow = GetLastUsedRow(wsRule)















    For Each fileItem In fd.SelectedItems







        Set sourceWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True, UpdateLinks:=0)







        hitBooks = hitBooks + 1















        For ruleRow = 2 To lastRuleRow







            If ShouldSkipRuleRow(wsRule, ruleRow) Then







                skipRules = skipRules + 1







            Else







                CompareOneRule wsRule, ruleRow, templateWb, sourceWb, compareWs, resultRow, hitSheets, diffRows, sameRows, skipRules







            End If







        Next ruleRow















        SafeCloseWorkbook sourceWb







    Next fileItem















    On Error Resume Next







    compareWs.Columns("A:O").AutoFit







    Application.DisplayAlerts = True







    Application.ScreenUpdating = True







    SafeCloseWorkbook templateWb







    RunLog_WriteRow COMPARE_LOG_KEY, "结束", CStr(hitBooks), CStr(hitSheets), CStr(diffRows), "完成", "一致=" & sameRows & "，跳过规则=" & skipRules & "，跳过工作簿=" & skipBooks, CStr(Round(Timer - t0, 2))







    On Error GoTo 0







    MsgBox "表头路径比对完成。" & vbCrLf & "比对工作簿数：" & hitBooks & vbCrLf & "命中工作表数：" & hitSheets & vbCrLf & "一致条数：" & sameRows & vbCrLf & "差异条数：" & diffRows, vbInformation, "表头路径比对"







    Exit Sub















CompareErrHandler:







    errNo = Err.Number







    errDesc = Err.Description







    Application.DisplayAlerts = True







    Application.ScreenUpdating = True







    SafeCloseWorkbook sourceWb







    SafeCloseWorkbook templateWb







    RunLog_WriteRow COMPARE_LOG_KEY, "结束", "", "", "", "失败", CStr(errNo) & " " & errDesc, CStr(Round(Timer - t0, 2))







    MsgBox "执行失败：" & CStr(errNo) & " " & errDesc, vbCritical, "表头路径比对"







End Sub















Public Sub 执行表头路径集合比对()







    Dim t0 As Double







    Dim wsRule As Worksheet







    Dim compareWb As Workbook







    Dim compareWs As Worksheet







    Dim templatePath As String







    Dim templateWb As Workbook







    Dim fd As FileDialog







    Dim fileItem As Variant







    Dim sourceWb As Workbook







    Dim lastRuleRow As Long







    Dim ruleRow As Long







    Dim resultRow As Long







    Dim hitBooks As Long







    Dim hitSheets As Long







    Dim diffRows As Long







    Dim sameRows As Long







    Dim skipRules As Long







    Dim skipBooks As Long







    Dim errNo As Long







    Dim errDesc As String















    t0 = Timer







    RunLog_WriteRow PATH_COMPARE_LOG_KEY, "开始", "", "", "", "", "开始", ""















    Set wsRule = EnsureRuleSheet()







    InitRuleHeader wsRule


















    templatePath = PickOneWorkbookPath("请选择作为模板的工作簿")







    If templatePath = "" Then







        RunLog_WriteRow PATH_COMPARE_LOG_KEY, "结束", "", "", "", "", "用户取消模板选择", CStr(Round(Timer - t0, 2))







        Exit Sub







    End If















    Set fd = Application.FileDialog(msoFileDialogFilePicker)







    With fd







        .Title = "请选择要进行表头路径比对的源工作簿，可多选"







        .AllowMultiSelect = True







        .Filters.Clear







        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"







        If .Show <> -1 Then







            RunLog_WriteRow PATH_COMPARE_LOG_KEY, "结束", "", "", "", "", "用户取消源文件选择", CStr(Round(Timer - t0, 2))







            Exit Sub







        End If







    End With















    Application.ScreenUpdating = False







    Application.DisplayAlerts = False







    On Error GoTo PathCompareErrHandler















    Set templateWb = Workbooks.Open(templatePath, ReadOnly:=True, UpdateLinks:=0)







    Set compareWb = CreateResultWorkbook(COMPARE_RESULT_SHEET_NAME, compareWs)







    InitCompareResultHeader compareWs







    resultRow = 2







    lastRuleRow = GetLastUsedRow(wsRule)















    For Each fileItem In fd.SelectedItems







        Set sourceWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True, UpdateLinks:=0)







        hitBooks = hitBooks + 1















        For ruleRow = 2 To lastRuleRow







            If ShouldSkipRuleRow(wsRule, ruleRow) Then







                skipRules = skipRules + 1







            Else







                CompareOneRuleByPath wsRule, ruleRow, templateWb, sourceWb, compareWs, resultRow, hitSheets, diffRows, sameRows, skipRules







            End If







        Next ruleRow















        SafeCloseWorkbook sourceWb







    Next fileItem















    On Error Resume Next







    compareWs.Columns("A:O").AutoFit







    Application.DisplayAlerts = True







    Application.ScreenUpdating = True







    SafeCloseWorkbook templateWb







    RunLog_WriteRow PATH_COMPARE_LOG_KEY, "结束", CStr(hitBooks), CStr(hitSheets), CStr(diffRows), "完成", "一致=" & sameRows & "，跳过规则=" & skipRules & "，跳过工作簿=" & skipBooks, CStr(Round(Timer - t0, 2))







    On Error GoTo 0







    MsgBox "表头路径比对完成。" & vbCrLf & "比对工作簿数：" & hitBooks & vbCrLf & "命中工作表数：" & hitSheets & vbCrLf & "一致条数：" & sameRows & vbCrLf & "差异条数：" & diffRows, vbInformation, "表头路径比对（按路径）"







    Exit Sub















PathCompareErrHandler:







    errNo = Err.Number







    errDesc = Err.Description







    Application.DisplayAlerts = True







    Application.ScreenUpdating = True







    SafeCloseWorkbook sourceWb







    SafeCloseWorkbook templateWb







    RunLog_WriteRow PATH_COMPARE_LOG_KEY, "结束", "", "", "", "失败", CStr(errNo) & " " & errDesc, CStr(Round(Timer - t0, 2))







    MsgBox "执行失败：" & CStr(errNo) & " " & errDesc, vbCritical, "表头路径比对（按路径）"







End Sub















Public Sub 执行表头混合比对()







    Dim t0 As Double







    Dim wsRule As Worksheet







    Dim compareWb As Workbook







    Dim compareWs As Worksheet







    Dim templatePath As String







    Dim templateWb As Workbook







    Dim fd As FileDialog







    Dim fileItem As Variant







    Dim sourceWb As Workbook







    Dim lastRuleRow As Long







    Dim ruleRow As Long







    Dim resultRow As Long







    Dim hitBooks As Long







    Dim hitSheets As Long







    Dim positionDiffRows As Long







    Dim positionSameRows As Long







    Dim pathDiffRows As Long







    Dim pathSameRows As Long







    Dim skipRules As Long







    Dim skipBooks As Long







    Dim errNo As Long







    Dim errDesc As String















    t0 = Timer







    RunLog_WriteRow HYBRID_COMPARE_LOG_KEY, "开始", "", "", "", "", "开始", ""















    Set wsRule = EnsureRuleSheet()







    InitRuleHeader wsRule


















    templatePath = PickOneWorkbookPath("请选择作为模板的工作簿")







    If templatePath = "" Then







        RunLog_WriteRow HYBRID_COMPARE_LOG_KEY, "结束", "", "", "", "", "用户取消模板选择", CStr(Round(Timer - t0, 2))







        Exit Sub







    End If















    Set fd = Application.FileDialog(msoFileDialogFilePicker)







    With fd







        .Title = "请选择要进行混合比对的源工作簿，可多选"







        .AllowMultiSelect = True







        .Filters.Clear







        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"







        If .Show <> -1 Then







            RunLog_WriteRow HYBRID_COMPARE_LOG_KEY, "结束", "", "", "", "", "用户取消源文件选择", CStr(Round(Timer - t0, 2))







            Exit Sub







        End If







    End With















    Application.ScreenUpdating = False







    Application.DisplayAlerts = False







    On Error GoTo HybridCompareErrHandler















    Set templateWb = Workbooks.Open(templatePath, ReadOnly:=True, UpdateLinks:=0)







    Set compareWb = CreateResultWorkbook(COMPARE_RESULT_SHEET_NAME, compareWs)







    InitCompareResultHeader compareWs







    resultRow = 2







    lastRuleRow = GetLastUsedRow(wsRule)















    For Each fileItem In fd.SelectedItems







        Set sourceWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True, UpdateLinks:=0)







        hitBooks = hitBooks + 1















        For ruleRow = 2 To lastRuleRow







            If ShouldSkipRuleRow(wsRule, ruleRow) Then







                skipRules = skipRules + 1







            Else







                CompareOneRuleHybrid wsRule, ruleRow, templateWb, sourceWb, compareWs, resultRow, hitSheets, positionDiffRows, positionSameRows, pathDiffRows, pathSameRows, skipRules







            End If







        Next ruleRow















        SafeCloseWorkbook sourceWb







    Next fileItem















    On Error Resume Next







    compareWs.Columns("A:O").AutoFit







    Application.DisplayAlerts = True







    Application.ScreenUpdating = True







    SafeCloseWorkbook templateWb







    RunLog_WriteRow HYBRID_COMPARE_LOG_KEY, "结束", CStr(hitBooks), CStr(hitSheets), CStr(positionDiffRows + pathDiffRows), "完成", "位置一致=" & positionSameRows & "，位置差异=" & positionDiffRows & "，路径一致=" & pathSameRows & "，路径差异=" & pathDiffRows & "，跳过规则=" & skipRules & "，跳过工作簿=" & skipBooks, CStr(Round(Timer - t0, 2))







    On Error GoTo 0







    MsgBox "混合比对完成。" & vbCrLf & "比对工作簿数：" & hitBooks & vbCrLf & "命中工作表数：" & hitSheets & vbCrLf & "位置一致条数：" & positionSameRows & vbCrLf & "位置差异条数：" & positionDiffRows & vbCrLf & "路径一致条数：" & pathSameRows & vbCrLf & "路径差异条数：" & pathDiffRows, vbInformation, "表头混合比对"







    Exit Sub















HybridCompareErrHandler:







    errNo = Err.Number







    errDesc = Err.Description







    Application.DisplayAlerts = True







    Application.ScreenUpdating = True







    SafeCloseWorkbook sourceWb







    SafeCloseWorkbook templateWb







    RunLog_WriteRow HYBRID_COMPARE_LOG_KEY, "结束", "", "", "", "失败", CStr(errNo) & " " & errDesc, CStr(Round(Timer - t0, 2))







    MsgBox "执行失败：" & CStr(errNo) & " " & errDesc, vbCritical, "表头混合比对"







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







    Dim pathMappings As Collection















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







    Set pathMappings = LoadPathMappings()







    resultRow = 2

    gTimelineFastMode = True
    gTimelineFastSlimMode = False
    Set gTimelineFastBuffer = New Collection
    gTimelineFastNextRow = resultRow







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







                ProcessOneExtractRule wsRule, ruleRow, wsResult, resultRow, targetWb, fileModified, dataDateText, dateSource, duplicateMap, pathMappings, hitSheets, outputRows, skipRules, duplicateRows







            End If







        Next ruleRow















        SafeCloseWorkbook targetWb







NextBook:







    Next fileItem















    FlushTimelineFastBuffer wsResult
    resultRow = gTimelineFastNextRow
    gTimelineFastMode = False
    Set gTimelineFastBuffer = Nothing

    wsResult.Columns("A:K").AutoFit







    Application.DisplayAlerts = True







    Application.ScreenUpdating = True















    RunLog_WriteRow LOG_KEY, "结束", CStr(hitBooks), CStr(hitSheets), CStr(outputRows), "完成", "跳过规则=" & skipRules & "，跳过工作簿=" & skipBooks & "，重复记录=" & duplicateRows, CStr(Round(Timer - t0, 2))







    MsgBox BuildSummaryMessage(hitBooks, hitSheets, outputRows, skipRules, skipBooks, duplicateRows), vbInformation, "批量转时序数据"







    Exit Sub















ErrHandler:







    Application.DisplayAlerts = True







    Application.ScreenUpdating = True







    gTimelineFastMode = False
    Set gTimelineFastBuffer = Nothing
    SafeCloseWorkbook targetWb







    RunLog_WriteRow LOG_KEY, "结束", "", "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))







    MsgBox "执行失败：" & Err.Number & " " & Err.Description, vbCritical, "批量转时序数据"







End Sub















Public Sub ExecuteTimelineExtractionFast()







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







    Dim pathMappings As Collection







    Dim oldCalculation As XlCalculation







    Dim oldEnableEvents As Boolean







    Dim oldScreenUpdating As Boolean







    Dim oldDisplayAlerts As Boolean















    t0 = Timer







    RunLog_WriteRow FAST_TIMELINE_LOG_KEY, "开始", "", "", "", "", "开始", ""















    Set wsRule = EnsureRuleSheet()







    InitRuleHeader wsRule


















    If MsgBox("请先确认目标工作簿已完成表格校验且无错误。" & vbCrLf & "是否继续执行时序提取（快速）？", vbQuestion + vbYesNo, "提取时序（快速）") <> vbYes Then







        RunLog_WriteRow FAST_TIMELINE_LOG_KEY, "结束", "", "", "", "", "用户取消，未确认表格校验", CStr(Round(Timer - t0, 2))







        Exit Sub







    End If















    Set fd = Application.FileDialog(msoFileDialogFilePicker)







    With fd







        .Title = "请选择要转时序的工作簿（可多选）"







        .AllowMultiSelect = True







        .Filters.Clear







        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"







        If .Show <> -1 Then







            RunLog_WriteRow FAST_TIMELINE_LOG_KEY, "结束", "", "", "", "", "用户取消", CStr(Round(Timer - t0, 2))







            Exit Sub







        End If







    End With















    oldScreenUpdating = Application.ScreenUpdating







    oldDisplayAlerts = Application.DisplayAlerts







    oldEnableEvents = Application.EnableEvents







    oldCalculation = Application.Calculation















    Application.ScreenUpdating = False







    Application.DisplayAlerts = False







    Application.EnableEvents = False







    On Error Resume Next







    Application.Calculation = xlCalculationManual







    On Error GoTo FastErrHandler







    On Error GoTo FastErrHandler















    Set resultWb = CreateResultWorkbook(RESULT_SHEET_NAME, wsResult)







    InitResultHeader wsResult







    Set duplicateMap = CreateObject("Scripting.Dictionary")







    duplicateMap.CompareMode = vbTextCompare







    Set pathMappings = LoadPathMappings()







    resultRow = 2







    lastRuleRow = GetLastUsedRow(wsRule)















    gTimelineFastSlimMode = False







    Set gTimelineFastSlimBuffer = Nothing







    gTimelineFastMode = True







    Set gTimelineFastBuffer = New Collection







    gTimelineFastNextRow = 2















    For Each fileItem In fd.SelectedItems







        fileModified = ""







        On Error Resume Next







        fileModified = Format(FileDateTime(CStr(fileItem)), "yyyy/mm/dd hh:nn:ss")







        On Error GoTo FastErrHandler















        Set targetWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True, UpdateLinks:=0)







        If Not ResolveWorkbookDataDate(targetWb, dataDateText, dateSource) Then







            skipBooks = skipBooks + 1







            SafeCloseWorkbook targetWb







            GoTo NextFastBook







        End If















        hitBooks = hitBooks + 1















        For ruleRow = 2 To lastRuleRow







            If ShouldSkipRuleRow(wsRule, ruleRow) Then







                skipRules = skipRules + 1







            Else







                ProcessOneExtractRule wsRule, ruleRow, wsResult, resultRow, targetWb, fileModified, dataDateText, dateSource, duplicateMap, pathMappings, hitSheets, outputRows, skipRules, duplicateRows







            End If







        Next ruleRow















        SafeCloseWorkbook targetWb







NextFastBook:







    Next fileItem















    FlushTimelineFastBuffer wsResult







    wsResult.Columns("A:K").AutoFit















    On Error Resume Next







    Application.Calculation = oldCalculation







    Application.EnableEvents = oldEnableEvents







    Application.DisplayAlerts = oldDisplayAlerts







    Application.ScreenUpdating = oldScreenUpdating







    On Error GoTo 0















    On Error Resume Next







    RunLog_WriteRow FAST_TIMELINE_LOG_KEY, "结束", CStr(hitBooks), CStr(hitSheets), CStr(outputRows), "完成", "跳过规则=" & skipRules & "跳过工作簿=" & skipBooks & "自动去重=" & duplicateRows, CStr(Round(Timer - t0, 2))







    On Error GoTo 0







    MsgBox BuildSummaryMessage(hitBooks, hitSheets, outputRows, skipRules, skipBooks, duplicateRows), vbInformation, "提取时序（快速）"







    gTimelineFastMode = False







    gTimelineFastSlimMode = False







    Set gTimelineFastBuffer = Nothing







    Set gTimelineFastSlimBuffer = Nothing







    Exit Sub















FastErrHandler:







    On Error Resume Next







    Application.Calculation = oldCalculation







    Application.EnableEvents = oldEnableEvents







    Application.DisplayAlerts = oldDisplayAlerts







    Application.ScreenUpdating = oldScreenUpdating







    On Error GoTo 0







    If Not wsResult Is Nothing Then







        FlushTimelineFastBuffer wsResult







    End If







    gTimelineFastMode = False







    gTimelineFastSlimMode = False







    Set gTimelineFastBuffer = Nothing







    Set gTimelineFastSlimBuffer = Nothing







    SafeCloseWorkbook targetWb







    RunLog_WriteRow FAST_TIMELINE_LOG_KEY, "结束", "", "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))







    MsgBox "执行失败：" & Err.Number & " " & Err.Description, vbCritical, "提取时序（快速）"







End Sub















Public Sub ExecuteTimelineExtractionFastSlim()







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







    Dim pathMappings As Collection







    Dim oldCalculation As XlCalculation







    Dim oldEnableEvents As Boolean







    Dim oldScreenUpdating As Boolean







    Dim oldDisplayAlerts As Boolean















    t0 = Timer







    RunLog_WriteRow FAST_TIMELINE_SLIM_LOG_KEY, "开始", "", "", "", "", "开始", ""















    Set wsRule = EnsureRuleSheet()







    InitRuleHeader wsRule


















    If MsgBox("请先确认目标工作簿已完成表格校验且无错误。" & vbCrLf & "是否继续执行时序提取（快速精简）？", vbQuestion + vbYesNo, "提取时序（快速精简）") <> vbYes Then







        RunLog_WriteRow FAST_TIMELINE_SLIM_LOG_KEY, "结束", "", "", "", "", "用户取消，未确认表格校验", CStr(Round(Timer - t0, 2))







        Exit Sub







    End If















    Set fd = Application.FileDialog(msoFileDialogFilePicker)







    With fd







        .Title = "请选择要转时序的工作簿（可多选）"







        .AllowMultiSelect = True







        .Filters.Clear







        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"







        If .Show <> -1 Then







            RunLog_WriteRow FAST_TIMELINE_SLIM_LOG_KEY, "结束", "", "", "", "", "用户取消", CStr(Round(Timer - t0, 2))







            Exit Sub







        End If







    End With















    oldScreenUpdating = Application.ScreenUpdating







    oldDisplayAlerts = Application.DisplayAlerts







    oldEnableEvents = Application.EnableEvents







    oldCalculation = Application.Calculation















    Application.ScreenUpdating = False







    Application.DisplayAlerts = False







    Application.EnableEvents = False







    On Error Resume Next







    Application.Calculation = xlCalculationManual







    On Error GoTo FastSlimErrHandler







    On Error GoTo FastSlimErrHandler















    Set resultWb = CreateResultWorkbook(RESULT_SHEET_NAME, wsResult)







    InitResultHeaderSlim wsResult







    Set duplicateMap = CreateObject("Scripting.Dictionary")







    duplicateMap.CompareMode = vbTextCompare







    Set pathMappings = LoadPathMappings()







    resultRow = 2







    lastRuleRow = GetLastUsedRow(wsRule)















    gTimelineFastMode = True







    gTimelineFastSlimMode = True







    Set gTimelineFastBuffer = Nothing







    Set gTimelineFastSlimBuffer = New Collection







    gTimelineFastNextRow = 2















    For Each fileItem In fd.SelectedItems







        fileModified = ""







        On Error Resume Next







        fileModified = Format(FileDateTime(CStr(fileItem)), "yyyy/mm/dd hh:nn:ss")







        On Error GoTo FastSlimErrHandler















        Set targetWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True, UpdateLinks:=0)







        If Not ResolveWorkbookDataDate(targetWb, dataDateText, dateSource) Then







            skipBooks = skipBooks + 1







            SafeCloseWorkbook targetWb







            GoTo NextFastSlimBook







        End If















        hitBooks = hitBooks + 1















        For ruleRow = 2 To lastRuleRow







            If ShouldSkipRuleRow(wsRule, ruleRow) Then







                skipRules = skipRules + 1







            Else







                ProcessOneExtractRule wsRule, ruleRow, wsResult, resultRow, targetWb, fileModified, dataDateText, dateSource, duplicateMap, pathMappings, hitSheets, outputRows, skipRules, duplicateRows







            End If







        Next ruleRow















        SafeCloseWorkbook targetWb







NextFastSlimBook:







    Next fileItem















    FlushTimelineFastSlimBuffer wsResult







    wsResult.Columns("A:G").AutoFit















    On Error Resume Next







    Application.Calculation = oldCalculation







    Application.EnableEvents = oldEnableEvents







    Application.DisplayAlerts = oldDisplayAlerts







    Application.ScreenUpdating = oldScreenUpdating







    On Error GoTo 0















    On Error Resume Next







    RunLog_WriteRow FAST_TIMELINE_SLIM_LOG_KEY, "结束", CStr(hitBooks), CStr(hitSheets), CStr(outputRows), "完成", "跳过规则=" & skipRules & "跳过工作簿=" & skipBooks & "自动去重=" & duplicateRows, CStr(Round(Timer - t0, 2))







    On Error GoTo 0







    MsgBox BuildSummaryMessage(hitBooks, hitSheets, outputRows, skipRules, skipBooks, duplicateRows), vbInformation, "提取时序（快速精简）"







    gTimelineFastMode = False







    gTimelineFastSlimMode = False







    Set gTimelineFastBuffer = Nothing







    Set gTimelineFastSlimBuffer = Nothing







    Exit Sub















FastSlimErrHandler:







    On Error Resume Next







    Application.Calculation = oldCalculation







    Application.EnableEvents = oldEnableEvents







    Application.DisplayAlerts = oldDisplayAlerts







    Application.ScreenUpdating = oldScreenUpdating







    On Error GoTo 0







    If Not wsResult Is Nothing Then







        FlushTimelineFastSlimBuffer wsResult







    End If







    gTimelineFastMode = False







    gTimelineFastSlimMode = False







    Set gTimelineFastBuffer = Nothing







    Set gTimelineFastSlimBuffer = Nothing







    SafeCloseWorkbook targetWb







    RunLog_WriteRow FAST_TIMELINE_SLIM_LOG_KEY, "结束", "", "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))







    MsgBox "执行失败：" & Err.Number & " " & Err.Description, vbCritical, "提取时序（快速精简）"







End Sub















Public Sub ExecuteRuleWideSummary()







    Dim t0 As Double







    Dim wsRule As Worksheet







    Dim resultWb As Workbook







    Dim fd As FileDialog







    Dim fileItem As Variant







    Dim targetWb As Workbook







    Dim lastRuleRow As Long







    Dim ruleRow As Long







    Dim hitBooks As Long







    Dim hitSheets As Long







    Dim outputRows As Long







    Dim skipRules As Long







    Dim skipBooks As Long







    Dim conflictRows As Long







    Dim dataDateText As String







    Dim dateSource As String







    Dim pathMappings As Collection







    Dim groupStore As Object







    Dim groupOrder As Collection







    Dim errNo As Long







    Dim errDesc As String







    Dim oneGroupKey As Variant







    Dim groupItem As Object







    Dim totalWideRows As Long







    Dim totalDynamicCols As Long







    Dim wsResult As Worksheet







    Dim resultSheetNameMap As Object















    t0 = Timer







    RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, "开始", "", "", "", "", "开始", ""















    Set wsRule = EnsureRuleSheet()







    InitRuleHeader wsRule


















    If MsgBox("请先确认目标工作簿已完成表格校验且无错误。" & vbCrLf & "是否继续执行规则汇总（宽表）？", vbQuestion + vbYesNo, "规则汇总（宽表）") <> vbYes Then







        RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, "取消", "", "", "", "", "用户取消，未确认表格校验", CStr(Round(Timer - t0, 2))







        Exit Sub







    End If















    Set fd = Application.FileDialog(msoFileDialogFilePicker)







    With fd







        .Title = "请选择要执行规则汇总（宽表）的工作簿，可多选"







        .AllowMultiSelect = True







        .Filters.Clear







        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"







        If .Show <> -1 Then







            RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, "取消", "", "", "", "", "用户取消文件选择", CStr(Round(Timer - t0, 2))







            Exit Sub







        End If







    End With















    Application.ScreenUpdating = False







    Application.DisplayAlerts = False







    On Error GoTo WideSummaryErrHandler















    Set groupStore = CreateObject("Scripting.Dictionary")





    groupStore.CompareMode = vbTextCompare







    Set groupOrder = New Collection

    gWideVerboseLogEnabled = IsWideVerboseLogEnabled()





    Set pathMappings = LoadPathMappings()







    lastRuleRow = GetLastUsedRow(wsRule)















    For Each fileItem In fd.SelectedItems







        Set targetWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True, UpdateLinks:=0)







        If Not ResolveWorkbookDataDate(targetWb, dataDateText, dateSource) Then







            skipBooks = skipBooks + 1







            RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, "跳过工作簿", targetWb.Name, "", "", "跳过", "未识别到数据日期", ""







            SafeCloseWorkbook targetWb







            GoTo NextWideBook







        End If















        hitBooks = hitBooks + 1















        For ruleRow = 2 To lastRuleRow







            If ShouldSkipRuleRow(wsRule, ruleRow) Then







                skipRules = skipRules + 1







            Else







                ProcessOneWideSummaryRule wsRule, ruleRow, targetWb, dataDateText, pathMappings, groupStore, groupOrder, hitSheets, skipRules, outputRows, conflictRows







            End If







        Next ruleRow















        SafeCloseWorkbook targetWb







NextWideBook:







    Next fileItem















    Set resultWb = CreateResultWorkbook(WIDE_RESULT_SHEET_NAME, wsResult)







    Set resultSheetNameMap = CreateObject("Scripting.Dictionary")







    resultSheetNameMap.CompareMode = vbTextCompare















    For Each oneGroupKey In groupOrder







        Set groupItem = groupStore(CStr(oneGroupKey))







        If CLng(groupItem("rowOrder").Count) = 0 Then







            GoTo NextWideGroup







        End If







        If StrComp(CStr(oneGroupKey), CStr(groupOrder(1)), vbTextCompare) = 0 Then







            Set wsResult = resultWb.Worksheets(1)







        Else







            Set wsResult = resultWb.Worksheets.Add(After:=resultWb.Worksheets(resultWb.Worksheets.Count))







        End If







        wsResult.Cells.Clear







        wsResult.Name = BuildWideSummarySheetName(CStr(oneGroupKey), resultSheetNameMap)







        WriteWideSummarySheet wsResult, groupItem("rowStore"), groupItem("rowOrder"), groupItem("colOrder")

        ApplyWideSummarySheetLayout wsResult, CLng(groupItem("colOrder").Count)





        totalWideRows = totalWideRows + CLng(groupItem("rowOrder").Count)







        totalDynamicCols = totalDynamicCols + CLng(groupItem("colOrder").Count)







NextWideGroup:







    Next oneGroupKey















    Application.DisplayAlerts = True







    Application.ScreenUpdating = True















    RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, "结果", CStr(hitBooks), CStr(hitSheets), CStr(outputRows), "完成", "结果Sheet数=" & groupOrder.Count & " 宽表总行数=" & totalWideRows & " 动态列总数=" & totalDynamicCols & " 跳过规则=" & skipRules & " 跳过工作簿=" & skipBooks & " 冲突数=" & conflictRows, CStr(Round(Timer - t0, 2))







    MsgBox BuildWideSummaryMessage(hitBooks, hitSheets, totalWideRows, totalDynamicCols, skipRules, skipBooks, conflictRows, CLng(groupOrder.Count)), vbInformation, "规则汇总（宽表）"







    Exit Sub















WideSummaryErrHandler:







    errNo = Err.Number







    errDesc = Err.Description







    Application.DisplayAlerts = True







    Application.ScreenUpdating = True







    SafeCloseWorkbook targetWb







    RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, "结果", "", "", "", "失败", CStr(errNo) & " " & errDesc, CStr(Round(Timer - t0, 2))







    MsgBox "执行失败：" & CStr(errNo) & " " & errDesc, vbCritical, "规则汇总（宽表）"







End Sub















Private Sub ProcessOneWideSummaryRule(ByVal wsRule As Worksheet, ByVal ruleRow As Long, ByVal targetWb As Workbook, ByVal dataDateText As String, ByVal pathMappings As Collection, ByVal groupStore As Object, ByVal groupOrder As Collection, ByRef hitSheets As Long, ByRef skipRules As Long, ByRef outputRows As Long, ByRef conflictRows As Long)







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







    Dim groupItem As Object















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







        RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, "跳过规则", targetWb.Name & "|" & ruleName, "", "", "跳过", "未匹配到工作表", ""







        Exit Sub







    End If















    Set groupItem = EnsureWideSummaryGroup(groupStore, groupOrder, ruleName)















    For Each item In matchedSheets







        ExtractSheetToWideSummary item, targetWb.Name, dataDateText, ruleName, rowHeaderCols, colHeaderRows, requiredColPaths, requiredRowPaths, dataStartRow, dataEndRow, dataStartCol, dataEndCol, skipKeywords, pathMappings, groupItem("rowStore"), groupItem("rowOrder"), groupItem("colIndexMap"), groupItem("colOrder"), hitSheets, skipRules, outputRows, conflictRows







    Next item







End Sub















Private Function EnsureWideSummaryGroup(ByVal groupStore As Object, ByVal groupOrder As Collection, ByVal ruleName As String) As Object







    Dim groupItem As Object







    Dim rowStore As Object







    Dim rowOrder As Collection







    Dim colIndexMap As Object







    Dim colOrder As Collection















    If groupStore.Exists(ruleName) Then







        Set EnsureWideSummaryGroup = groupStore(ruleName)







        Exit Function







    End If















    Set groupItem = CreateObject("Scripting.Dictionary")







    groupItem.CompareMode = vbTextCompare







    Set rowStore = CreateObject("Scripting.Dictionary")







    rowStore.CompareMode = vbTextCompare







    Set rowOrder = New Collection







    Set colIndexMap = CreateObject("Scripting.Dictionary")







    colIndexMap.CompareMode = vbTextCompare







    Set colOrder = New Collection















    groupItem.Add "rowStore", rowStore







    groupItem.Add "rowOrder", rowOrder







    groupItem.Add "colIndexMap", colIndexMap







    groupItem.Add "colOrder", colOrder







    groupStore.Add ruleName, groupItem







    groupOrder.Add ruleName







    Set EnsureWideSummaryGroup = groupItem







End Function















Private Function BuildWideSummarySheetName(ByVal rawRuleName As String, ByVal usedNames As Object) As String







    Dim baseName As String







    Dim tryName As String







    Dim suffixNo As Long







    Dim suffixText As String







    Dim maxBaseLen As Long















    baseName = SanitizeWideSummarySheetName(rawRuleName)







    tryName = baseName







    suffixNo = 1















    Do While usedNames.Exists(UCase$(tryName))







        suffixNo = suffixNo + 1







        suffixText = "_" & CStr(suffixNo)







        maxBaseLen = 31 - Len(suffixText)







        If maxBaseLen < 1 Then







            maxBaseLen = 1







        End If







        tryName = Left$(baseName, maxBaseLen) & suffixText







    Loop















    usedNames(UCase$(tryName)) = True







    BuildWideSummarySheetName = tryName







End Function















Private Function SanitizeWideSummarySheetName(ByVal rawRuleName As String) As String







    Dim result As String















    result = Trim$(CStr(rawRuleName))







    If result = "" Then







        result = WIDE_RESULT_SHEET_NAME







    End If















    result = Replace(result, "\", "_")







    result = Replace(result, "/", "_")







    result = Replace(result, ":", "_")







    result = Replace(result, "*", "_")







    result = Replace(result, "?", "_")







    result = Replace(result, "[", "_")







    result = Replace(result, "]", "_")















    Do While InStr(result, "__") > 0







        result = Replace(result, "__", "_")







    Loop















    result = Trim$(result)







    If result = "" Then







        result = WIDE_RESULT_SHEET_NAME







    End If















    If Len(result) > 31 Then







        result = Left$(result, 31)







    End If















    SanitizeWideSummarySheetName = result







End Function















Private Sub ExtractSheetToWideSummary(ByVal ws As Worksheet, ByVal sourceBookName As String, ByVal dataDateText As String, ByVal ruleName As String, ByVal rowHeaderCols As Collection, ByVal colHeaderRows As Collection, ByVal requiredColPaths As String, ByVal requiredRowPaths As String, ByVal dataStartRow As Long, ByVal dataEndRow As Long, ByVal dataStartCol As Long, ByVal dataEndCol As Long, ByVal skipKeywords As String, ByVal pathMappings As Collection, ByVal rowStore As Object, ByVal rowOrder As Collection, ByVal colIndexMap As Object, ByVal colOrder As Collection, ByRef hitSheets As Long, ByRef skipRules As Long, ByRef outputRows As Long, ByRef conflictRows As Long)







    Dim actualEndRow As Long







    Dim actualEndCol As Long







    Dim headerText() As String







    Dim rowPathMap As Object







    Dim colPathMap As Object







    Dim mappingLogMap As Object







    Dim rowNo As Long







    Dim colNo As Long







    Dim rowPath As String







    Dim colPath As String







    Dim cellValue As Variant







    Dim valuesWritten As Long



    Dim dataArr As Variant



    Dim rowOffset As Long



    Dim colOffset As Long



    Dim cellAddress As String















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







    If gWideVerboseLogEnabled Then

        Set mappingLogMap = CreateObject("Scripting.Dictionary")

    Else

        Set mappingLogMap = Nothing

    End If





    If Not mappingLogMap Is Nothing Then mappingLogMap.CompareMode = vbTextCompare





    Set rowPathMap = BuildMappedRowPathMap(sourceBookName, ws.Name, ruleName, ws, rowHeaderCols, dataStartRow, actualEndRow, skipKeywords, pathMappings, mappingLogMap)







    Set colPathMap = BuildMappedColPathMap(sourceBookName, ws.Name, ruleName, headerText, dataStartCol, actualEndCol, pathMappings, colHeaderRows, mappingLogMap)







    If Not ValidateRequiredAnchors(requiredColPaths, requiredRowPaths, rowPathMap, colPathMap) Then







        skipRules = skipRules + 1







        RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, "跳过规则", sourceBookName & "|" & ws.Name & "|" & ruleName, "", "", "跳过", "未命中必含列头或必含行头", ""







        Exit Sub







    End If















    dataArr = ws.Range(ws.Cells(dataStartRow, dataStartCol), ws.Cells(actualEndRow, actualEndCol)).Value



    hitSheets = hitSheets + 1







    For rowNo = dataStartRow To actualEndRow







        rowPath = GetUniquePathFromMap(rowPathMap, rowNo)







        If rowPath <> "" Then







            If Not MatchAnyKeyword(rowPath, skipKeywords) Then







                For colNo = dataStartCol To actualEndCol



                    rowOffset = rowNo - dataStartRow + 1



                    colOffset = colNo - dataStartCol + 1



                    If IsArray(dataArr) Then



                        cellValue = dataArr(rowOffset, colOffset)



                    Else



                        cellValue = dataArr



                    End If







                    If IsOutputValue(cellValue) Then







                        colPath = GetUniquePathFromMap(colPathMap, colNo)







                        If colPath <> "" Then







                            cellAddress = ws.Cells(rowNo, colNo).Address(False, False)



                            AddWideSummaryValue rowStore, rowOrder, colIndexMap, colOrder, sourceBookName, ws.Name, dataDateText, rowPath, colPath, cellValue, cellAddress, ruleName, conflictRows, valuesWritten







                        End If







                    End If







                Next colNo







            End If







        End If







    Next rowNo















    outputRows = outputRows + valuesWritten







    If gWideVerboseLogEnabled Then

        RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, "汇总工作表", sourceBookName & "|" & ws.Name & "|" & ruleName, "", CStr(valuesWritten), "完成", "OK", ""

    End If





End Sub















Private Function BuildMappedRowPathMap(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal ws As Worksheet, ByVal rowHeaderCols As Collection, ByVal startRow As Long, ByVal endRow As Long, ByVal skipKeywords As String, ByVal pathMappings As Collection, ByVal mappingLogMap As Object) As Object





    Dim countMap As Object







    Dim seqMap As Object







    Dim result As Object







    Dim rowNo As Long







    Dim rawPath As String

    Dim uniquePath As String

    Dim mappedPath As String

    Dim seq As Long

    Dim rawPaths() As String













    Set countMap = CreateObject("Scripting.Dictionary")







    countMap.CompareMode = vbTextCompare







    Set seqMap = CreateObject("Scripting.Dictionary")







    seqMap.CompareMode = vbTextCompare







    Set result = CreateObject("Scripting.Dictionary")







    result.CompareMode = vbTextCompare















    ReDim rawPaths(startRow To endRow)



    For rowNo = startRow To endRow

        rawPath = BuildRowPath(ws, rowNo, rowHeaderCols)

        rawPaths(rowNo) = rawPath

        If rawPath <> "" Then

            If countMap.Exists(rawPath) Then

                countMap(rawPath) = CLng(countMap(rawPath)) + 1

            Else

                countMap(rawPath) = 1

            End If

        End If

    Next rowNo













    For rowNo = startRow To endRow

        rawPath = rawPaths(rowNo)



        If rawPath <> "" Then





            If CLng(countMap(rawPath)) > 1 Then







                If seqMap.Exists(rawPath) Then







                    seq = CLng(seqMap(rawPath)) + 1







                Else







                    seq = 1







                End If







                seqMap(rawPath) = seq







                uniquePath = rawPath & "_" & CStr(seq)

                LogWidePathMapping mappingLogMap, "行头重命名", sourceBookName, sourceSheetName, ruleName, rawPath, uniquePath, "行" & CStr(rowNo)





            Else







                uniquePath = rawPath







            End If















            mappedPath = ApplyPathMapping(uniquePath, "行头", ruleName, sourceBookName, sourceSheetName, pathMappings)

            LogWidePathMapping mappingLogMap, "行头映射", sourceBookName, sourceSheetName, ruleName, uniquePath, mappedPath, "行" & CStr(rowNo)





            If mappedPath <> "" Then







                If Not MatchAnyKeyword(mappedPath, skipKeywords) Then







                    result(CStr(rowNo)) = mappedPath







                End If







            End If







        End If







    Next rowNo















    Set BuildMappedRowPathMap = result







End Function















Private Function BuildMappedColPathMap(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByRef headerText() As String, ByVal startCol As Long, ByVal endCol As Long, ByVal pathMappings As Collection, ByVal colHeaderRows As Collection, ByVal mappingLogMap As Object) As Object





    Dim countMap As Object







    Dim seqMap As Object







    Dim result As Object







    Dim colNo As Long







    Dim rawPath As String

    Dim uniquePath As String

    Dim mappedPath As String

    Dim seq As Long

    Dim rawPaths() As String













    Set countMap = CreateObject("Scripting.Dictionary")







    countMap.CompareMode = vbTextCompare







    Set seqMap = CreateObject("Scripting.Dictionary")







    seqMap.CompareMode = vbTextCompare







    Set result = CreateObject("Scripting.Dictionary")







    result.CompareMode = vbTextCompare















    ReDim rawPaths(startCol To endCol)



    For colNo = startCol To endCol

        rawPath = BuildColPath(headerText, colNo)

        rawPaths(colNo) = rawPath

        If rawPath <> "" Then

            If countMap.Exists(rawPath) Then

                countMap(rawPath) = CLng(countMap(rawPath)) + 1

            Else

                countMap(rawPath) = 1

            End If

        End If

    Next colNo













    For colNo = startCol To endCol

        rawPath = rawPaths(colNo)





        If rawPath <> "" Then







            If CLng(countMap(rawPath)) > 1 Then







                If seqMap.Exists(rawPath) Then







                    seq = CLng(seqMap(rawPath)) + 1







                Else







                    seq = 1







                End If







                seqMap(rawPath) = seq







                uniquePath = rawPath & "_" & CStr(seq)







                LogWidePathMapping mappingLogMap, "列头重命名", sourceBookName, sourceSheetName, ruleName, rawPath, uniquePath, "列" & ColumnNumberToLetter(colNo) & " 行" & BuildHeaderRowsLabel(colHeaderRows)







            Else







                uniquePath = rawPath







            End If















            mappedPath = ApplyPathMapping(uniquePath, "列头", ruleName, sourceBookName, sourceSheetName, pathMappings)







            LogWidePathMapping mappingLogMap, "列头映射", sourceBookName, sourceSheetName, ruleName, uniquePath, mappedPath, "列" & ColumnNumberToLetter(colNo) & " 行" & BuildHeaderRowsLabel(colHeaderRows)







            If mappedPath <> "" Then







                result(CStr(colNo)) = mappedPath







            End If







        End If







    Next colNo















    Set BuildMappedColPathMap = result







End Function















Private Sub AddWideSummaryValue(ByVal rowStore As Object, ByVal rowOrder As Collection, ByVal colIndexMap As Object, ByVal colOrder As Collection, ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal dataDateText As String, ByVal rowPath As String, ByVal colPath As String, ByVal cellValue As Variant, ByVal cellAddress As String, ByVal ruleName As String, ByRef conflictRows As Long, ByRef valuesWritten As Long)







    Dim rowKey As String







    Dim rowItem As Object







    Dim valueMap As Object











    Dim conflictMsg As String













    rowKey = sourceBookName & "|" & sourceSheetName & "|" & dataDateText & "|" & rowPath







    If rowStore.Exists(rowKey) Then







        Set rowItem = rowStore(rowKey)







    Else







        Set rowItem = CreateObject("Scripting.Dictionary")







        rowItem.CompareMode = vbTextCompare







        rowItem.Add "sourceBook", sourceBookName







        rowItem.Add "sourceSheet", sourceSheetName







        rowItem.Add "dataDate", dataDateText







        rowItem.Add "rowPath", rowPath







        Set valueMap = CreateObject("Scripting.Dictionary")







        valueMap.CompareMode = vbTextCompare















        rowItem.Add "values", valueMap











        rowStore.Add rowKey, rowItem







        rowOrder.Add rowKey







    End If















    Set valueMap = rowItem("values")



















    If Not colIndexMap.Exists(colPath) Then







        colIndexMap.Add colPath, CLng(colOrder.Count) + 1







        colOrder.Add colPath







    End If















    If valueMap.Exists(colPath) Then







        conflictRows = conflictRows + 1







        If gWideVerboseLogEnabled Then





            conflictMsg = "列头=" & colPath & "|当前=" & cellAddress

            RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, "宽表冲突", sourceBookName & "|" & sourceSheetName & "|" & ruleName, dataDateText, rowPath, "冲突", conflictMsg, ""

        End If





        Exit Sub







    End If















    valueMap.Add colPath, cellValue











    valuesWritten = valuesWritten + 1







End Sub















Private Sub LogWidePathMapping(ByVal mappingLogMap As Object, ByVal actionName As String, ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal rawPath As String, ByVal mappedPath As String, ByVal positionText As String)







    Dim logKey As String















    If Not gWideVerboseLogEnabled Then Exit Sub

    If Trim$(rawPath) = "" Then Exit Sub





    If StrComp(Trim$(rawPath), Trim$(mappedPath), vbTextCompare) = 0 Then Exit Sub















    logKey = actionName & "|" & sourceBookName & "|" & sourceSheetName & "|" & ruleName & "|" & rawPath & "|" & mappedPath







    If mappingLogMap Is Nothing Then Exit Sub

























    If mappingLogMap.Exists(logKey) Then Exit Sub







    mappingLogMap.Add logKey, True







    RunLog_WriteRow WIDE_SUMMARY_LOG_KEY, actionName, sourceBookName & "|" & sourceSheetName & "|" & ruleName, rawPath, mappedPath, "映射", "位置=" & positionText & "|原始=" & rawPath & "|标准=" & mappedPath, ""







End Sub















Private Sub WriteWideSummarySheet(ByVal ws As Worksheet, ByVal rowStore As Object, ByVal rowOrder As Collection, ByVal colOrder As Collection)







    Dim rowCount As Long







    Dim colCount As Long







    Dim outputArr() As Variant







    Dim rowIdx As Long







    Dim dynIdx As Long







    Dim rowKey As Variant







    Dim rowItem As Object







    Dim valueMap As Object







    Dim colPath As String















    rowCount = rowOrder.Count







    colCount = 4 + colOrder.Count















    ws.Cells(1, wrSourceBook).Value = "工作簿名"







    ws.Cells(1, wrSourceSheet).Value = "工作表名"







    ws.Cells(1, wrDataDate).Value = "数据日期"







    ws.Cells(1, wrRowPath).Value = "行头路径"







    For dynIdx = 1 To colOrder.Count







        ws.Cells(1, 4 + dynIdx).Value = CStr(colOrder(dynIdx))







    Next dynIdx







    ws.Rows(1).Font.Bold = True















    If rowCount = 0 Then







        Exit Sub







    End If















    ReDim outputArr(1 To rowCount, 1 To colCount)







    For rowIdx = 1 To rowCount







        rowKey = rowOrder(rowIdx)







        Set rowItem = rowStore(CStr(rowKey))







        outputArr(rowIdx, wrSourceBook) = rowItem("sourceBook")







        outputArr(rowIdx, wrSourceSheet) = rowItem("sourceSheet")







        outputArr(rowIdx, wrDataDate) = rowItem("dataDate")







        outputArr(rowIdx, wrRowPath) = rowItem("rowPath")















        Set valueMap = rowItem("values")







        For dynIdx = 1 To colOrder.Count







            colPath = CStr(colOrder(dynIdx))







            If valueMap.Exists(colPath) Then







                outputArr(rowIdx, 4 + dynIdx) = valueMap(colPath)







            End If







        Next dynIdx







    Next rowIdx















    ws.Range(ws.Cells(2, 1), ws.Cells(rowCount + 1, colCount)).Value = outputArr



End Sub







Private Function IsWideVerboseLogEnabled() As Boolean



    IsWideVerboseLogEnabled = False



End Function







Private Sub ApplyWideSummarySheetLayout(ByVal ws As Worksheet, ByVal dynamicColCount As Long)



    Dim lastCol As Long



    If ws Is Nothing Then Exit Sub



    ws.Columns("A:D").AutoFit



    If dynamicColCount <= 0 Then Exit Sub



    lastCol = 4 + dynamicColCount



    ws.Range(ws.Cells(1, 5), ws.Cells(1, lastCol)).EntireColumn.ColumnWidth = WIDE_DYNAMIC_COL_WIDTH



End Sub







Private Function BuildWideSummaryMessage(ByVal hitBooks As Long, ByVal hitSheets As Long, ByVal outputRows As Long, ByVal dynamicCols As Long, ByVal skipRules As Long, ByVal skipBooks As Long, ByVal conflictRows As Long, ByVal resultSheetCount As Long) As String





    BuildWideSummaryMessage = "处理文件数：" & hitBooks







    BuildWideSummaryMessage = BuildWideSummaryMessage & vbCrLf & "处理工作表数：" & hitSheets







    BuildWideSummaryMessage = BuildWideSummaryMessage & vbCrLf & "结果Sheet数：" & resultSheetCount







    BuildWideSummaryMessage = BuildWideSummaryMessage & vbCrLf & "宽表行数：" & outputRows







    BuildWideSummaryMessage = BuildWideSummaryMessage & vbCrLf & "动态列数：" & dynamicCols







    BuildWideSummaryMessage = BuildWideSummaryMessage & vbCrLf & "跳过规则数：" & skipRules







    BuildWideSummaryMessage = BuildWideSummaryMessage & vbCrLf & "跳过工作簿数：" & skipBooks







    BuildWideSummaryMessage = BuildWideSummaryMessage & vbCrLf & "冲突记录数：" & conflictRows







    BuildWideSummaryMessage = BuildWideSummaryMessage & vbCrLf & vbCrLf & "结果已按规则名称分Sheet写入新工作簿"



    BuildWideSummaryMessage = BuildWideSummaryMessage



End Function















Private Sub ProcessOneExtractRule(ByVal wsRule As Worksheet, ByVal ruleRow As Long, ByVal wsResult As Worksheet, ByRef resultRow As Long, ByVal targetWb As Workbook, ByVal fileModified As String, ByVal dataDateText As String, ByVal dateSource As String, ByVal duplicateMap As Object, ByVal pathMappings As Collection, ByRef hitSheets As Long, ByRef outputRows As Long, ByRef skipRules As Long, ByRef duplicateRows As Long)







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







        If Not gTimelineFastMode Then







            RunLog_WriteRow LOG_KEY, "跳过规则", targetWb.Name & "|" & ruleName, "", "", "跳过", "未匹配到工作表", ""







        End If







        Exit Sub







    End If















    For Each item In matchedSheets







        ExtractSheetToTimeline item, wsResult, resultRow, targetWb.Name, fileModified, dataDateText, dateSource, ruleName, rowHeaderCols, colHeaderRows, requiredColPaths, requiredRowPaths, dataStartRow, dataEndRow, dataStartCol, dataEndCol, skipKeywords, duplicateMap, pathMappings, hitSheets, skipRules, outputRows, duplicateRows







    Next item







End Sub















Private Sub ExtractSheetToTimeline(ByVal ws As Worksheet, ByVal wsResult As Worksheet, ByRef resultRow As Long, ByVal sourceBookName As String, ByVal fileModified As String, ByVal dataDateText As String, ByVal dateSource As String, ByVal ruleName As String, ByVal rowHeaderCols As Collection, ByVal colHeaderRows As Collection, ByVal requiredColPaths As String, ByVal requiredRowPaths As String, ByVal dataStartRow As Long, ByVal dataEndRow As Long, ByVal dataStartCol As Long, ByVal dataEndCol As Long, ByVal skipKeywords As String, ByVal duplicateMap As Object, ByVal pathMappings As Collection, ByRef hitSheets As Long, ByRef skipRules As Long, ByRef outputRows As Long, ByRef duplicateRows As Long)







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



    Dim dataArr As Variant



    Dim rowOffset As Long



    Dim colOffset As Long



    Dim cellAddress As String















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







    Set rowPathMap = BuildUniqueRowPathMap(sourceBookName, ws.Name, ruleName, ws, rowHeaderCols, dataStartRow, actualEndRow, skipKeywords, pathMappings)







    Set colPathMap = BuildUniqueColPathMap(sourceBookName, ws.Name, ruleName, colHeaderRows, headerText, dataStartCol, actualEndCol, pathMappings)







    If Not ValidateRequiredAnchors(requiredColPaths, requiredRowPaths, rowPathMap, colPathMap) Then







        skipRules = skipRules + 1







        If Not gTimelineFastMode Then







            RunLog_WriteRow LOG_KEY, "跳过规则", sourceBookName & "|" & ws.Name & "|" & ruleName, "", "", "跳过", "未命中必含列头或必含行头", ""







        End If







        Exit Sub







    End If















    dataArr = ws.Range(ws.Cells(dataStartRow, dataStartCol), ws.Cells(actualEndRow, actualEndCol)).Value



    hitSheets = hitSheets + 1







    For rowNo = dataStartRow To actualEndRow







        rowPath = GetUniquePathFromMap(rowPathMap, rowNo)







        If rowPath <> "" Then







            If Not MatchAnyKeyword(rowPath, skipKeywords) Then







                For colNo = dataStartCol To actualEndCol







                    rowOffset = rowNo - dataStartRow + 1



                    colOffset = colNo - dataStartCol + 1



                    If IsArray(dataArr) Then



                        cellValue = dataArr(rowOffset, colOffset)



                    Else



                        cellValue = dataArr



                    End If







                    If IsOutputValue(cellValue) Then







                        colPath = GetUniquePathFromMap(colPathMap, colNo)







                        If colPath <> "" Then







                            oneKey = BuildDuplicateKey(sourceBookName, ws.Name, dataDateText, rowPath, colPath, cellValue)







                            If duplicateMap.Exists(oneKey) Then







                                duplicateRows = duplicateRows + 1







                                If Not gTimelineFastMode Then







                                    cellAddress = ws.Cells(rowNo, colNo).Address(False, False)



                                    RunLog_WriteRow LOG_KEY, "重复记录", sourceBookName & "|" & ws.Name & "|" & ruleName, dataDateText, rowPath, "跳过", "当前=" & cellAddress & "|列头=" & colPath & "|首次=" & CStr(duplicateMap(oneKey)), ""







                                End If







                            Else







                                cellAddress = ws.Cells(rowNo, colNo).Address(False, False)



                                duplicateMap(oneKey) = "规则=" & ruleName & "|单元格=" & cellAddress & "|行头=" & rowPath & "|列头=" & colPath







                                WriteTimelineRow wsResult, resultRow, sourceBookName, ws.Name, ruleName, fileModified, dataDateText, dateSource, rowPath, colPath, cellValue, cellAddress







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







    If Not gTimelineFastMode Then







        RunLog_WriteRow LOG_KEY, "提取工作表", sourceBookName & "|" & ws.Name & "|" & ruleName, "", CStr(valuesWritten), "完成", "OK", ""







    End If







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















Private Function BuildUniqueRowPathMap(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal ws As Worksheet, ByVal rowHeaderCols As Collection, ByVal startRow As Long, ByVal endRow As Long, ByVal skipKeywords As String, ByVal pathMappings As Collection) As Object







    Set BuildUniqueRowPathMap = BuildUniquePathMapForRows(sourceBookName, sourceSheetName, ruleName, ws, rowHeaderCols, startRow, endRow, skipKeywords, pathMappings)







End Function















Private Function BuildUniqueColPathMap(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal colHeaderRows As Collection, ByRef headerText() As String, ByVal startCol As Long, ByVal endCol As Long, ByVal pathMappings As Collection) As Object







    Set BuildUniqueColPathMap = BuildUniquePathMapForCols(sourceBookName, sourceSheetName, ruleName, colHeaderRows, headerText, startCol, endCol, pathMappings)







End Function















Private Function BuildUniquePathMapForRows(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal ws As Worksheet, ByVal rowHeaderCols As Collection, ByVal startRow As Long, ByVal endRow As Long, ByVal skipKeywords As String, ByVal pathMappings As Collection) As Object







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







        basePath = ApplyPathMapping(basePath, "行头", ruleName, sourceBookName, sourceSheetName, pathMappings)







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







        basePath = ApplyPathMapping(basePath, "行头", ruleName, sourceBookName, sourceSheetName, pathMappings)







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







                    renamedPath = ApplyPathMappingForUniquePath(renamedPath, "行头", ruleName, sourceBookName, sourceSheetName, pathMappings)







                    result(CStr(rowNo)) = renamedPath







                    If Not gTimelineFastMode Then







                        RunLog_WriteRow LOG_KEY, "行头重命名", sourceBookName & "|" & sourceSheetName & "|" & ruleName, basePath, renamedPath, "完成", "位置=R" & CStr(rowNo) & "|从=" & basePath & "|到=" & renamedPath, ""







                    End If







                Else







                    result(CStr(rowNo)) = basePath







                End If







            End If







        End If







    Next rowNo















    Set BuildUniquePathMapForRows = result







End Function















Private Function BuildUniquePathMapForCols(ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal colHeaderRows As Collection, ByRef headerText() As String, ByVal startCol As Long, ByVal endCol As Long, ByVal pathMappings As Collection) As Object







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







        basePath = ApplyPathMapping(basePath, "列头", ruleName, sourceBookName, sourceSheetName, pathMappings)







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







        basePath = ApplyPathMapping(basePath, "列头", ruleName, sourceBookName, sourceSheetName, pathMappings)







        If basePath <> "" Then







            If CLng(countMap(basePath)) > 1 Then







                If seqMap.Exists(basePath) Then







                    seq = CLng(seqMap(basePath)) + 1







                Else







                    seq = 1







                End If







                seqMap(basePath) = seq







                renamedPath = basePath & "_" & CStr(seq)







                renamedPath = ApplyPathMappingForUniquePath(renamedPath, "列头", ruleName, sourceBookName, sourceSheetName, pathMappings)







                result(CStr(colNo)) = renamedPath







                If Not gTimelineFastMode Then







                    RunLog_WriteRow LOG_KEY, "列头重命名", sourceBookName & "|" & sourceSheetName & "|" & ruleName, basePath, renamedPath, "完成", "位置=列" & ColumnNumberToLetter(colNo) & " 行" & BuildHeaderRowsLabel(colHeaderRows) & "|从=" & basePath & "|到=" & renamedPath, ""







                End If







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















Private Function ApplyPathMappingForUniquePath(ByVal uniquePath As String, ByVal targetType As String, ByVal ruleName As String, ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal pathMappings As Collection) As String







    Dim item As Variant







    Dim matchMode As String







    Dim originalPath As String















    ApplyPathMappingForUniquePath = uniquePath







    If uniquePath = "" Then Exit Function







    If pathMappings Is Nothing Then Exit Function















    For Each item In pathMappings







        If StrComp(CStr(item("targetType")), targetType, vbTextCompare) = 0 Then







            If MappingScopeMatched(item, ruleName, sourceBookName, sourceSheetName) Then







                originalPath = Trim$(CStr(item("originalPath")))







                If originalPath = "" Then GoTo NextItem







                If Not HasAutoUniqueSuffix(originalPath) Then GoTo NextItem







                matchMode = Trim$(CStr(item("matchMode")))







                If matchMode = "" Then matchMode = "精确"







                If StrComp(matchMode, "精确", vbTextCompare) = 0 Then







                    If StrComp(uniquePath, originalPath, vbTextCompare) = 0 Then







                        ApplyPathMappingForUniquePath = CStr(item("standardPath"))







                        Exit Function







                    End If







                ElseIf StrComp(matchMode, "包含", vbTextCompare) = 0 Then







                    If InStr(1, uniquePath, originalPath, vbTextCompare) > 0 Then







                        ApplyPathMappingForUniquePath = CStr(item("standardPath"))







                        Exit Function







                    End If







                End If







            End If







        End If







NextItem:







    Next item







End Function















Private Function HasAutoUniqueSuffix(ByVal pathText As String) As Boolean







    Dim pos As Long







    Dim tailText As String















    pathText = Trim$(pathText)







    If pathText = "" Then Exit Function















    pos = InStrRev(pathText, "_")







    If pos <= 0 Then Exit Function







    If pos = Len(pathText) Then Exit Function















    tailText = Mid$(pathText, pos + 1)







    If tailText = "" Then Exit Function







    If Not IsNumeric(tailText) Then Exit Function















    HasAutoUniqueSuffix = True







End Function















Private Function LoadPathMappings() As Collection







    Dim ws As Worksheet







    Dim lastRow As Long







    Dim rowNo As Long







    Dim item As Object















    Set LoadPathMappings = New Collection







    Set ws = EnsurePathMapSheet()







    InitPathMapHeader ws















    lastRow = GetLastUsedRow(ws)







    For rowNo = 2 To lastRow







        If IsEnabledValue(ws.Cells(rowNo, mcEnabled).Value) Then







            If Trim$(CStr(ws.Cells(rowNo, mcOriginalPath).Value)) <> "" Then







                If Trim$(CStr(ws.Cells(rowNo, mcStandardPath).Value)) <> "" Then







                    Set item = CreateObject("Scripting.Dictionary")







                    item.CompareMode = vbTextCompare







                    item("ruleName") = Trim$(CStr(ws.Cells(rowNo, mcRuleName).Value))







                    item("bookKeywords") = Trim$(CStr(ws.Cells(rowNo, mcBookKeywords).Value))







                    item("sheetKeywords") = Trim$(CStr(ws.Cells(rowNo, mcSheetKeywords).Value))







                    item("targetType") = Trim$(CStr(ws.Cells(rowNo, mcTargetType).Value))







                    item("matchMode") = Trim$(CStr(ws.Cells(rowNo, mcMatchMode).Value))







                    item("originalPath") = Trim$(CStr(ws.Cells(rowNo, mcOriginalPath).Value))







                    item("standardPath") = Trim$(CStr(ws.Cells(rowNo, mcStandardPath).Value))







                    LoadPathMappings.Add item







                End If







            End If







        End If







    Next rowNo







End Function















Private Function ApplyPathMapping(ByVal basePath As String, ByVal targetType As String, ByVal ruleName As String, ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal pathMappings As Collection) As String







    Dim item As Variant







    Dim matchMode As String















    ApplyPathMapping = basePath







    If basePath = "" Then Exit Function







    If pathMappings Is Nothing Then Exit Function















    For Each item In pathMappings







        If StrComp(CStr(item("targetType")), targetType, vbTextCompare) = 0 Then







            If MappingScopeMatched(item, ruleName, sourceBookName, sourceSheetName) Then







                matchMode = Trim$(CStr(item("matchMode")))







                If matchMode = "" Then matchMode = "精确"







                If StrComp(matchMode, "精确", vbTextCompare) = 0 Then







                    If StrComp(basePath, CStr(item("originalPath")), vbTextCompare) = 0 Then







                        ApplyPathMapping = CStr(item("standardPath"))







                        Exit Function







                    End If







                ElseIf StrComp(matchMode, "包含", vbTextCompare) = 0 Then







                    If InStr(1, basePath, CStr(item("originalPath")), vbTextCompare) > 0 Then







                        ApplyPathMapping = CStr(item("standardPath"))







                        Exit Function







                    End If







                End If







            End If







        End If







    Next item







End Function















Private Function MappingScopeMatched(ByVal item As Object, ByVal ruleName As String, ByVal sourceBookName As String, ByVal sourceSheetName As String) As Boolean







    MappingScopeMatched = True















    If Trim$(CStr(item("ruleName"))) <> "" Then







        If StrComp(Trim$(CStr(item("ruleName"))), ruleName, vbTextCompare) <> 0 Then







            MappingScopeMatched = False







            Exit Function







        End If







    End If















    If Trim$(CStr(item("bookKeywords"))) <> "" Then







        If Not MatchAllKeywords(sourceBookName, CStr(item("bookKeywords"))) Then







            MappingScopeMatched = False







            Exit Function







        End If







    End If















    If Trim$(CStr(item("sheetKeywords"))) <> "" Then







        If Not MatchAllKeywords(sourceSheetName, CStr(item("sheetKeywords"))) Then







            MappingScopeMatched = False







            Exit Function







        End If







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







    If gTimelineFastSlimMode Then







        QueueTimelineRowSlim ws, sourceBookName, sourceSheetName, ruleName, dataDateText, rowPath, colPath, valueText







        Exit Sub







    End If







    If gTimelineFastMode Then







        QueueTimelineRow ws, sourceBookName, sourceSheetName, ruleName, fileModified, dataDateText, dateSource, rowPath, colPath, valueText, cellAddress







        Exit Sub







    End If















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















Private Sub QueueTimelineRowSlim(ByVal ws As Worksheet, ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal dataDateText As String, ByVal rowPath As String, ByVal colPath As String, ByVal valueText As Variant)







    Dim rowData(1 To 7) As Variant















    If gTimelineFastSlimBuffer Is Nothing Then







        Set gTimelineFastSlimBuffer = New Collection







    End If















    rowData(1) = sourceBookName







    rowData(2) = sourceSheetName







    rowData(3) = ruleName







    rowData(4) = dataDateText







    rowData(5) = rowPath







    rowData(6) = colPath







    rowData(7) = valueText















    gTimelineFastSlimBuffer.Add rowData







    If gTimelineFastSlimBuffer.Count >= TIMELINE_FAST_FLUSH_SIZE Then







        FlushTimelineFastSlimBuffer ws







    End If







End Sub















Private Sub QueueTimelineRow(ByVal ws As Worksheet, ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal ruleName As String, ByVal fileModified As String, ByVal dataDateText As String, ByVal dateSource As String, ByVal rowPath As String, ByVal colPath As String, ByVal valueText As Variant, ByVal cellAddress As String)







    Dim rowData(1 To 11) As Variant















    If gTimelineFastBuffer Is Nothing Then







        Set gTimelineFastBuffer = New Collection







    End If















    rowData(rsExecTime) = Format(Now, "yyyy/mm/dd hh:nn:ss")







    rowData(rsSourceBook) = sourceBookName







    rowData(rsSourceSheet) = sourceSheetName







    rowData(rsRuleName) = ruleName







    rowData(rsFileModified) = fileModified







    rowData(rsDataDate) = dataDateText







    rowData(rsDateSource) = dateSource







    rowData(rsRowPath) = rowPath







    rowData(rsColPath) = colPath







    rowData(rsValue) = valueText







    rowData(rsCellAddress) = cellAddress















    gTimelineFastBuffer.Add rowData







    If gTimelineFastBuffer.Count >= TIMELINE_FAST_FLUSH_SIZE Then







        FlushTimelineFastBuffer ws







    End If







End Sub















Private Sub FlushTimelineFastBuffer(ByVal ws As Worksheet)







    Dim bufferCount As Long







    Dim outputArr() As Variant







    Dim idx As Long







    Dim colNo As Long







    Dim oneRow As Variant







    Dim startRow As Long







    Dim endRow As Long















    If Not gTimelineFastMode Then Exit Sub







    If gTimelineFastBuffer Is Nothing Then Exit Sub







    bufferCount = gTimelineFastBuffer.Count







    If bufferCount = 0 Then Exit Sub















    ReDim outputArr(1 To bufferCount, 1 To 11)







    For idx = 1 To bufferCount







        oneRow = gTimelineFastBuffer(idx)







        For colNo = 1 To 11







            outputArr(idx, colNo) = oneRow(colNo)







        Next colNo







    Next idx















    startRow = gTimelineFastNextRow







    endRow = startRow + bufferCount - 1







    ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 11)).Value = outputArr







    gTimelineFastNextRow = endRow + 1







    Set gTimelineFastBuffer = New Collection







End Sub















Private Sub FlushTimelineFastSlimBuffer(ByVal ws As Worksheet)







    Dim bufferCount As Long







    Dim outputArr() As Variant







    Dim idx As Long







    Dim colNo As Long







    Dim oneRow As Variant







    Dim startRow As Long







    Dim endRow As Long















    If Not gTimelineFastSlimMode Then Exit Sub







    If gTimelineFastSlimBuffer Is Nothing Then Exit Sub







    bufferCount = gTimelineFastSlimBuffer.Count







    If bufferCount = 0 Then Exit Sub















    ReDim outputArr(1 To bufferCount, 1 To 7)







    For idx = 1 To bufferCount







        oneRow = gTimelineFastSlimBuffer(idx)







        For colNo = 1 To 7







            outputArr(idx, colNo) = oneRow(colNo)







        Next colNo







    Next idx















    startRow = gTimelineFastNextRow







    endRow = startRow + bufferCount - 1







    ws.Range(ws.Cells(startRow, 1), ws.Cells(endRow, 7)).Value = outputArr







    gTimelineFastNextRow = endRow + 1







    Set gTimelineFastSlimBuffer = New Collection







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















Private Function ValidateRequiredAnchors(ByVal requiredColPaths As String, ByVal requiredRowPaths As String, ByVal rowPathMap As Object, ByVal colPathMap As Object) As Boolean



    Dim requiredCols As Collection



    Dim requiredRows As Collection



    Dim idx As Long



    Dim valueSet As Object



    Dim oneToken As String







    Set requiredCols = SplitTokens(requiredColPaths)



    If requiredCols.Count > 0 Then



        Set valueSet = BuildPathValueSet(colPathMap)



        For idx = 1 To requiredCols.Count



            oneToken = CStr(requiredCols(idx))



            If oneToken <> "" Then



                If valueSet Is Nothing Then Exit Function



                If Not valueSet.Exists(oneToken) Then Exit Function



            End If



        Next idx



    End If







    Set requiredRows = SplitTokens(requiredRowPaths)



    If requiredRows.Count > 0 Then



        Set valueSet = BuildPathValueSet(rowPathMap)



        For idx = 1 To requiredRows.Count



            oneToken = CStr(requiredRows(idx))



            If oneToken <> "" Then



                If valueSet Is Nothing Then Exit Function



                If Not valueSet.Exists(oneToken) Then Exit Function



            End If



        Next idx



    End If







    ValidateRequiredAnchors = True



End Function







Private Function BuildPathValueSet(ByVal pathMap As Object) As Object



    Dim result As Object



    Dim oneKey As Variant



    Dim oneValue As String







    If pathMap Is Nothing Then Exit Function



    Set result = CreateObject("Scripting.Dictionary")



    result.CompareMode = vbTextCompare







    For Each oneKey In pathMap.Keys



        oneValue = Trim$(CStr(pathMap(oneKey)))



        If oneValue <> "" Then



            If Not result.Exists(oneValue) Then



                result.Add oneValue, True



            End If



        End If



    Next oneKey







    Set BuildPathValueSet = result



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







    If gRuleValidationReady Then



        If Not gRuleInvalidRows Is Nothing Then



            If gRuleInvalidRows.Exists(CStr(rowNo)) Then



                ShouldSkipRuleRow = True



                Exit Function



            End If



        End If



    End If







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















Private Function EnsurePathMapSheet() As Worksheet







    On Error Resume Next







    Set EnsurePathMapSheet = ThisWorkbook.Worksheets(MAP_SHEET_NAME)







    On Error GoTo 0















    If EnsurePathMapSheet Is Nothing Then







        Set EnsurePathMapSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))







        EnsurePathMapSheet.Name = MAP_SHEET_NAME







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















Private Sub InitPathMapHeader(ByVal ws As Worksheet)







    ws.Cells(1, mcEnabled).Value = "是否启用"







    ws.Cells(1, mcMapName).Value = "映射名称"







    ws.Cells(1, mcRuleName).Value = "适用规则名"







    ws.Cells(1, mcBookKeywords).Value = "工作簿关键字"







    ws.Cells(1, mcSheetKeywords).Value = "工作表关键字"







    ws.Cells(1, mcTargetType).Value = "作用对象"







    ws.Cells(1, mcMatchMode).Value = "匹配方式"







    ws.Cells(1, mcOriginalPath).Value = "原始路径"







    ws.Cells(1, mcStandardPath).Value = "标准路径"







    ws.Cells(1, mcRemark).Value = "备注"







    ws.Rows(1).Font.Bold = True







    SetHeaderComment ws.Cells(1, mcEnabled), "是否启用此条映射。填写 Y/1/TRUE/是 时生效。"







    SetHeaderComment ws.Cells(1, mcMapName), "给人工看的映射名称，可留空。"







    SetHeaderComment ws.Cells(1, mcRuleName), "优先限制到某一条时序提取规则；填写后仅该规则命中时生效。"







    SetHeaderComment ws.Cells(1, mcBookKeywords), "进一步限制工作簿名必须包含这些关键字，可用分号分隔多个关键字。"







    SetHeaderComment ws.Cells(1, mcSheetKeywords), "进一步限制工作表名必须包含这些关键字，可用分号分隔多个关键字。"







    SetHeaderComment ws.Cells(1, mcTargetType), "作用对象。只允许填写：行头 或 列头。"







    SetHeaderComment ws.Cells(1, mcMatchMode), "匹配方式。第一版建议填写 精确；包含 仅在少数场景使用。"







    SetHeaderComment ws.Cells(1, mcOriginalPath), "源表中提取到的原始路径。"







    SetHeaderComment ws.Cells(1, mcStandardPath), "标准化后的目标路径。"







    SetHeaderComment ws.Cells(1, mcRemark), "备注说明，不参与程序判断。"







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















Private Sub InitResultHeaderSlim(ByVal ws As Worksheet)







    ws.Cells(1, 1).Value = "源文件"







    ws.Cells(1, 2).Value = "工作表名"







    ws.Cells(1, 3).Value = "规则名称"







    ws.Cells(1, 4).Value = "数据日期"







    ws.Cells(1, 5).Value = "行头路径"







    ws.Cells(1, 6).Value = "列头路径"







    ws.Cells(1, 7).Value = "数值"







    ws.Rows(1).Font.Bold = True







End Sub















Private Sub InitCompareResultHeader(ByVal ws As Worksheet)







    ws.Cells(1, crcEnabled).Value = "是否启用"







    ws.Cells(1, crcMapName).Value = "映射名称"







    ws.Cells(1, crcRuleName).Value = "适用规则名"







    ws.Cells(1, crcBookKeywords).Value = "工作簿关键字"







    ws.Cells(1, crcSheetKeywords).Value = "工作表关键字"







    ws.Cells(1, crcTargetType).Value = "作用对象"







    ws.Cells(1, crcMatchMode).Value = "匹配方式"







    ws.Cells(1, crcOriginalPath).Value = "原始路径"







    ws.Cells(1, crcStandardPath).Value = "标准路径"







    ws.Cells(1, crcRemark).Value = "备注"







    ws.Cells(1, crcSourceBook).Value = "源文件"







    ws.Cells(1, crcSourceSheet).Value = "工作表名"







    ws.Cells(1, crcTemplatePosition).Value = "模板位置"







    ws.Cells(1, crcSourcePosition).Value = "源表位置"







    ws.Cells(1, crcCompareResult).Value = "比对结果"







    ws.Rows(1).Font.Bold = True







    SetHeaderComment ws.Cells(1, crcEnabled), "复制到路径标准化映射表时使用。这里默认留空，由人工确认后再填写 Y。"







    SetHeaderComment ws.Cells(1, crcMapName), "复制到路径标准化映射表时可填写一个便于识别的映射名称。"







    SetHeaderComment ws.Cells(1, crcRuleName), "已自动带出当前时序提取规则名，复制到映射表后可直接使用。"







    SetHeaderComment ws.Cells(1, crcBookKeywords), "已自动带出当前规则的工作簿关键字，可按需要保留或清空。"







    SetHeaderComment ws.Cells(1, crcSheetKeywords), "已自动带出当前规则的工作表关键字，可按需要保留或清空。"







    SetHeaderComment ws.Cells(1, crcTargetType), "作用对象。程序会自动填 行头 或 列头。"







    SetHeaderComment ws.Cells(1, crcMatchMode), "匹配方式。默认输出为 精确，复制到映射表后一般无需调整。"







    SetHeaderComment ws.Cells(1, crcOriginalPath), "源表当前识别到的路径，复制到映射表后作为原始路径。"







    SetHeaderComment ws.Cells(1, crcStandardPath), "模板中识别到的路径，复制到映射表后可直接作为标准路径。"







    SetHeaderComment ws.Cells(1, crcRemark), "备注列，复制到映射表后可补充人工说明。"







    SetHeaderComment ws.Cells(1, crcSourceBook), "发生差异的源文件名，仅用于人工核对。"







    SetHeaderComment ws.Cells(1, crcSourceSheet), "发生差异的工作表名，仅用于人工核对。"







    SetHeaderComment ws.Cells(1, crcTemplatePosition), "模板中该路径对应的位置，用于判断是文本变化还是结构位置变化。"







    SetHeaderComment ws.Cells(1, crcSourcePosition), "源表中该路径对应的位置，用于判断是文本变化还是结构位置变化。"







    SetHeaderComment ws.Cells(1, crcCompareResult), "本次比对结论，例如 文本变化、模板缺失、源表缺失。"







End Sub















Private Sub SetHeaderComment(ByVal targetCell As Range, ByVal commentText As String)







    On Error Resume Next







    If Not targetCell.Comment Is Nothing Then







        targetCell.Comment.Delete







    End If







    targetCell.AddComment commentText







    targetCell.Comment.Visible = False







    On Error GoTo 0







End Sub















Private Sub CompareOneRule(ByVal wsRule As Worksheet, ByVal ruleRow As Long, ByVal templateWb As Workbook, ByVal sourceWb As Workbook, ByVal compareWs As Worksheet, ByRef resultRow As Long, ByRef hitSheets As Long, ByRef diffRows As Long, ByRef sameRows As Long, ByRef skipRules As Long)







    Dim ruleName As String







    Dim bookKeywords As String







    Dim sheetKeywords As String







    Dim rowHeaderCols As Collection







    Dim colHeaderRows As Collection







    Dim dataStartRow As Long







    Dim dataEndRow As Long







    Dim dataStartCol As Long







    Dim dataEndCol As Long







    Dim sourceSheets As Collection







    Dim templateSheets As Collection







    Dim sourceItem As Variant







    Dim templateWs As Worksheet















    ruleName = Trim$(CStr(wsRule.Cells(ruleRow, rcRuleName).Value))







    bookKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcBookKeywords).Value))







    sheetKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcSheetKeywords).Value))







    Set rowHeaderCols = ParseColumnCollection(CStr(wsRule.Cells(ruleRow, rcRowHeaderCols).Value))







    Set colHeaderRows = ParseNumberCollection(CStr(wsRule.Cells(ruleRow, rcColHeaderRows).Value))







    dataStartRow = ParseLongValue(wsRule.Cells(ruleRow, rcStartRow).Value, 0)







    dataEndRow = ParseLongValue(wsRule.Cells(ruleRow, rcEndRow).Value, 0)







    dataStartCol = ParseColumnSpec(wsRule.Cells(ruleRow, rcStartCol).Value, 0)







    dataEndCol = ParseColumnSpec(wsRule.Cells(ruleRow, rcEndCol).Value, 0)















    If ruleName = "" Then







        ruleName = "规则" & CStr(ruleRow)







    End If















    If bookKeywords <> "" Then







        If Not MatchAllKeywords(sourceWb.Name, bookKeywords) Then Exit Sub







    End If















    If colHeaderRows.Count = 0 Then Exit Sub







    If dataStartRow <= 0 Or dataStartCol <= 0 Then Exit Sub















    Set sourceSheets = MatchWorksheets(sourceWb, sheetKeywords)







    If sourceSheets Is Nothing Or sourceSheets.Count = 0 Then Exit Sub















    Set templateSheets = MatchWorksheets(templateWb, sheetKeywords)







    If templateSheets Is Nothing Or templateSheets.Count = 0 Then







        skipRules = skipRules + 1







        RunLog_WriteRow COMPARE_LOG_KEY, "跳过规则", templateWb.Name & "|" & ruleName, "", "", "跳过", "模板未匹配到工作表", ""







        Exit Sub







    End If















    For Each sourceItem In sourceSheets







        Set templateWs = ResolveTemplateSheet(templateSheets, CStr(sourceItem.Name))







        If templateWs Is Nothing Then







            skipRules = skipRules + 1







            RunLog_WriteRow COMPARE_LOG_KEY, "跳过规则", sourceWb.Name & "|" & CStr(sourceItem.Name) & "|" & ruleName, "", "", "跳过", "模板工作表匹配不唯一", ""







        Else







            hitSheets = hitSheets + 1







            CompareSheetPaths templateWs, sourceItem, compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, rowHeaderCols, colHeaderRows, dataStartRow, dataEndRow, dataStartCol, dataEndCol, sourceWb.Name, diffRows, sameRows, True







        End If







    Next sourceItem







End Sub















Private Sub CompareOneRuleByPath(ByVal wsRule As Worksheet, ByVal ruleRow As Long, ByVal templateWb As Workbook, ByVal sourceWb As Workbook, ByVal compareWs As Worksheet, ByRef resultRow As Long, ByRef hitSheets As Long, ByRef diffRows As Long, ByRef sameRows As Long, ByRef skipRules As Long)







    Dim ruleName As String







    Dim bookKeywords As String







    Dim sheetKeywords As String







    Dim rowHeaderCols As Collection







    Dim colHeaderRows As Collection







    Dim dataStartRow As Long







    Dim dataEndRow As Long







    Dim dataStartCol As Long







    Dim dataEndCol As Long







    Dim skipKeywords As String







    Dim sourceSheets As Collection







    Dim templateSheets As Collection







    Dim sourceItem As Variant







    Dim templateWs As Worksheet















    ruleName = Trim$(CStr(wsRule.Cells(ruleRow, rcRuleName).Value))







    bookKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcBookKeywords).Value))







    sheetKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcSheetKeywords).Value))







    Set rowHeaderCols = ParseColumnCollection(CStr(wsRule.Cells(ruleRow, rcRowHeaderCols).Value))







    Set colHeaderRows = ParseNumberCollection(CStr(wsRule.Cells(ruleRow, rcColHeaderRows).Value))







    dataStartRow = ParseLongValue(wsRule.Cells(ruleRow, rcStartRow).Value, 0)







    dataEndRow = ParseLongValue(wsRule.Cells(ruleRow, rcEndRow).Value, 0)







    dataStartCol = ParseColumnSpec(wsRule.Cells(ruleRow, rcStartCol).Value, 0)







    dataEndCol = ParseColumnSpec(wsRule.Cells(ruleRow, rcEndCol).Value, 0)







    skipKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcSkipKeywords).Value))















    If ruleName = "" Then







        ruleName = "规则" & CStr(ruleRow)







    End If















    If bookKeywords <> "" Then







        If Not MatchAllKeywords(sourceWb.Name, bookKeywords) Then Exit Sub







    End If















    If colHeaderRows.Count = 0 Then Exit Sub







    If dataStartRow <= 0 Or dataStartCol <= 0 Then Exit Sub















    Set sourceSheets = MatchWorksheets(sourceWb, sheetKeywords)







    If sourceSheets Is Nothing Or sourceSheets.Count = 0 Then Exit Sub















    Set templateSheets = MatchWorksheets(templateWb, sheetKeywords)







    If templateSheets Is Nothing Or templateSheets.Count = 0 Then







        skipRules = skipRules + 1







        RunLog_WriteRow PATH_COMPARE_LOG_KEY, "跳过规则", templateWb.Name & "|" & ruleName, "", "", "跳过", "模板未匹配到工作表", ""







        Exit Sub







    End If















    For Each sourceItem In sourceSheets







        Set templateWs = ResolveTemplateSheet(templateSheets, CStr(sourceItem.Name))







        If templateWs Is Nothing Then







            skipRules = skipRules + 1







            RunLog_WriteRow PATH_COMPARE_LOG_KEY, "跳过规则", sourceWb.Name & "|" & CStr(sourceItem.Name) & "|" & ruleName, "", "", "跳过", "模板工作表匹配不唯一", ""







        Else







            hitSheets = hitSheets + 1







            CompareSheetPathsByPath templateWs, sourceItem, compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, rowHeaderCols, colHeaderRows, dataStartRow, dataEndRow, dataStartCol, dataEndCol, skipKeywords, sourceWb.Name, diffRows, sameRows







        End If







    Next sourceItem







End Sub















Private Sub CompareOneRuleHybrid(ByVal wsRule As Worksheet, ByVal ruleRow As Long, ByVal templateWb As Workbook, ByVal sourceWb As Workbook, ByVal compareWs As Worksheet, ByRef resultRow As Long, ByRef hitSheets As Long, ByRef positionDiffRows As Long, ByRef positionSameRows As Long, ByRef pathDiffRows As Long, ByRef pathSameRows As Long, ByRef skipRules As Long)







    Dim ruleName As String







    Dim bookKeywords As String







    Dim sheetKeywords As String







    Dim rowHeaderCols As Collection







    Dim colHeaderRows As Collection







    Dim dataStartRow As Long







    Dim dataEndRow As Long







    Dim dataStartCol As Long







    Dim dataEndCol As Long







    Dim skipKeywords As String







    Dim sourceSheets As Collection







    Dim templateSheets As Collection







    Dim sourceItem As Variant







    Dim templateWs As Worksheet







    Dim beforeDiff As Long















    ruleName = Trim$(CStr(wsRule.Cells(ruleRow, rcRuleName).Value))







    bookKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcBookKeywords).Value))







    sheetKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcSheetKeywords).Value))







    Set rowHeaderCols = ParseColumnCollection(CStr(wsRule.Cells(ruleRow, rcRowHeaderCols).Value))







    Set colHeaderRows = ParseNumberCollection(CStr(wsRule.Cells(ruleRow, rcColHeaderRows).Value))







    dataStartRow = ParseLongValue(wsRule.Cells(ruleRow, rcStartRow).Value, 0)







    dataEndRow = ParseLongValue(wsRule.Cells(ruleRow, rcEndRow).Value, 0)







    dataStartCol = ParseColumnSpec(wsRule.Cells(ruleRow, rcStartCol).Value, 0)







    dataEndCol = ParseColumnSpec(wsRule.Cells(ruleRow, rcEndCol).Value, 0)







    skipKeywords = Trim$(CStr(wsRule.Cells(ruleRow, rcSkipKeywords).Value))















    If ruleName = "" Then







        ruleName = "规则" & CStr(ruleRow)







    End If















    If bookKeywords <> "" Then







        If Not MatchAllKeywords(sourceWb.Name, bookKeywords) Then Exit Sub







    End If















    If colHeaderRows.Count = 0 Then Exit Sub







    If dataStartRow <= 0 Or dataStartCol <= 0 Then Exit Sub















    Set sourceSheets = MatchWorksheets(sourceWb, sheetKeywords)







    If sourceSheets Is Nothing Or sourceSheets.Count = 0 Then Exit Sub















    Set templateSheets = MatchWorksheets(templateWb, sheetKeywords)







    If templateSheets Is Nothing Or templateSheets.Count = 0 Then







        skipRules = skipRules + 1







        RunLog_WriteRow HYBRID_COMPARE_LOG_KEY, "跳过规则", templateWb.Name & "|" & ruleName, "", "", "跳过", "模板未匹配到工作表", ""







        Exit Sub







    End If















    For Each sourceItem In sourceSheets







        Set templateWs = ResolveTemplateSheet(templateSheets, CStr(sourceItem.Name))







        If templateWs Is Nothing Then







            skipRules = skipRules + 1







            RunLog_WriteRow HYBRID_COMPARE_LOG_KEY, "跳过规则", sourceWb.Name & "|" & CStr(sourceItem.Name) & "|" & ruleName, "", "", "跳过", "模板工作表匹配不唯一", ""







        Else







            hitSheets = hitSheets + 1







            beforeDiff = positionDiffRows







            CompareSheetPaths templateWs, sourceItem, compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, rowHeaderCols, colHeaderRows, dataStartRow, dataEndRow, dataStartCol, dataEndCol, sourceWb.Name, positionDiffRows, positionSameRows, False







            If positionDiffRows > beforeDiff Then







                RunLog_WriteRow HYBRID_COMPARE_LOG_KEY, "追加路径比对", sourceWb.Name & "|" & CStr(sourceItem.Name) & "|" & ruleName, "", "", "继续", "按位置存在差异，追加按路径比对", ""







                CompareSheetPathsByPath templateWs, sourceItem, compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, rowHeaderCols, colHeaderRows, dataStartRow, dataEndRow, dataStartCol, dataEndCol, skipKeywords, sourceWb.Name, pathDiffRows, pathSameRows, False







            End If







        End If







    Next sourceItem







End Sub















Private Sub CompareSheetPaths(ByVal templateWs As Worksheet, ByVal sourceWs As Worksheet, ByVal compareWs As Worksheet, ByRef resultRow As Long, ByVal ruleName As String, ByVal bookKeywords As String, ByVal sheetKeywords As String, ByVal rowHeaderCols As Collection, ByVal colHeaderRows As Collection, ByVal dataStartRow As Long, ByVal dataEndRow As Long, ByVal dataStartCol As Long, ByVal dataEndCol As Long, ByVal sourceBookName As String, ByRef diffRows As Long, ByRef sameRows As Long, Optional ByVal outputAllRows As Boolean = True)







    Dim templateEndRow As Long







    Dim sourceEndRow As Long







    Dim templateEndCol As Long







    Dim sourceEndCol As Long







    Dim maxEndRow As Long







    Dim maxEndCol As Long







    Dim templateHeaderText() As String







    Dim sourceHeaderText() As String







    Dim rowNo As Long







    Dim colNo As Long







    Dim templatePath As String







    Dim sourcePath As String







    Dim templatePos As String







    Dim sourcePos As String







    Dim sheetDiffRows As Long







    Dim sheetSameRows As Long















    templateEndRow = ResolveSheetEndRow(templateWs, dataEndRow)







    sourceEndRow = ResolveSheetEndRow(sourceWs, dataEndRow)







    templateEndCol = ResolveSheetEndCol(templateWs, dataEndCol)







    sourceEndCol = ResolveSheetEndCol(sourceWs, dataEndCol)







    If templateEndCol < dataStartCol Then templateEndCol = dataStartCol







    If sourceEndCol < dataStartCol Then sourceEndCol = dataStartCol







    maxEndRow = IIf(templateEndRow > sourceEndRow, templateEndRow, sourceEndRow)







    maxEndCol = IIf(templateEndCol > sourceEndCol, templateEndCol, sourceEndCol)















    ReDim templateHeaderText(1 To colHeaderRows.Count, dataStartCol To templateEndCol)







    ReDim sourceHeaderText(1 To colHeaderRows.Count, dataStartCol To sourceEndCol)







    BuildHeaderCache templateWs, colHeaderRows, dataStartCol, templateEndCol, templateHeaderText







    BuildHeaderCache sourceWs, colHeaderRows, dataStartCol, sourceEndCol, sourceHeaderText















    For colNo = dataStartCol To maxEndCol







        templatePath = ""







        sourcePath = ""







        If colNo <= templateEndCol Then templatePath = BuildColPath(templateHeaderText, colNo)







        If colNo <= sourceEndCol Then sourcePath = BuildColPath(sourceHeaderText, colNo)







        templatePos = BuildColPositionLabel(colNo, colHeaderRows)







        sourcePos = BuildColPositionLabel(colNo, colHeaderRows)







        If StrComp(templatePath, sourcePath, vbTextCompare) = 0 Then







            sameRows = sameRows + 1







            sheetSameRows = sheetSameRows + 1







            If outputAllRows Then







                WriteCompareResultRow compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, "列头", sourcePath, templatePath, sourceBookName, sourceWs.Name, templatePos, sourcePos, BuildCompareResultText(templatePath, sourcePath), "按位置"







                resultRow = resultRow + 1







            End If







        Else







            diffRows = diffRows + 1







            sheetDiffRows = sheetDiffRows + 1







            WriteCompareResultRow compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, "列头", sourcePath, templatePath, sourceBookName, sourceWs.Name, templatePos, sourcePos, BuildCompareResultText(templatePath, sourcePath), "按位置"







            resultRow = resultRow + 1







        End If







    Next colNo















    For rowNo = dataStartRow To maxEndRow







        templatePath = ""







        sourcePath = ""







        If rowNo <= templateEndRow Then templatePath = BuildRowPath(templateWs, rowNo, rowHeaderCols)







        If rowNo <= sourceEndRow Then sourcePath = BuildRowPath(sourceWs, rowNo, rowHeaderCols)







        templatePos = BuildRowPositionLabel(rowNo, rowHeaderCols)







        sourcePos = BuildRowPositionLabel(rowNo, rowHeaderCols)







        If StrComp(templatePath, sourcePath, vbTextCompare) = 0 Then







            sameRows = sameRows + 1







            sheetSameRows = sheetSameRows + 1







            If outputAllRows Then







                WriteCompareResultRow compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, "行头", sourcePath, templatePath, sourceBookName, sourceWs.Name, templatePos, sourcePos, BuildCompareResultText(templatePath, sourcePath), "按位置"







                resultRow = resultRow + 1







            End If







        Else







            diffRows = diffRows + 1







            sheetDiffRows = sheetDiffRows + 1







            WriteCompareResultRow compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, "行头", sourcePath, templatePath, sourceBookName, sourceWs.Name, templatePos, sourcePos, BuildCompareResultText(templatePath, sourcePath), "按位置"







            resultRow = resultRow + 1







        End If







    Next rowNo















    RunLog_WriteRow COMPARE_LOG_KEY, "比对工作表", sourceBookName & "|" & sourceWs.Name & "|" & ruleName, CStr(sheetSameRows), CStr(sheetDiffRows), "完成", "模板=" & templateWs.Name, ""







End Sub















Private Sub CompareSheetPathsByPath(ByVal templateWs As Worksheet, ByVal sourceWs As Worksheet, ByVal compareWs As Worksheet, ByRef resultRow As Long, ByVal ruleName As String, ByVal bookKeywords As String, ByVal sheetKeywords As String, ByVal rowHeaderCols As Collection, ByVal colHeaderRows As Collection, ByVal dataStartRow As Long, ByVal dataEndRow As Long, ByVal dataStartCol As Long, ByVal dataEndCol As Long, ByVal skipKeywords As String, ByVal sourceBookName As String, ByRef diffRows As Long, ByRef sameRows As Long, Optional ByVal outputAllRows As Boolean = True)







    Dim templateRowMap As Object







    Dim sourceRowMap As Object







    Dim templateColMap As Object







    Dim sourceColMap As Object







    Dim sheetDiffRows As Long







    Dim sheetSameRows As Long















    Set templateRowMap = BuildPathOccurrenceMap(templateWs, rowHeaderCols, colHeaderRows, dataStartRow, dataEndRow, dataStartCol, dataEndCol, skipKeywords, True)







    Set sourceRowMap = BuildPathOccurrenceMap(sourceWs, rowHeaderCols, colHeaderRows, dataStartRow, dataEndRow, dataStartCol, dataEndCol, skipKeywords, True)







    Set templateColMap = BuildPathOccurrenceMap(templateWs, rowHeaderCols, colHeaderRows, dataStartRow, dataEndRow, dataStartCol, dataEndCol, skipKeywords, False)







    Set sourceColMap = BuildPathOccurrenceMap(sourceWs, rowHeaderCols, colHeaderRows, dataStartRow, dataEndRow, dataStartCol, dataEndCol, skipKeywords, False)















    CompareOccurrenceMaps compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, "行头", sourceBookName, sourceWs.Name, templateRowMap, sourceRowMap, sheetSameRows, sheetDiffRows, outputAllRows







    CompareOccurrenceMaps compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, "列头", sourceBookName, sourceWs.Name, templateColMap, sourceColMap, sheetSameRows, sheetDiffRows, outputAllRows















    sameRows = sameRows + sheetSameRows







    diffRows = diffRows + sheetDiffRows







    RunLog_WriteRow PATH_COMPARE_LOG_KEY, "比对工作表", sourceBookName & "|" & sourceWs.Name & "|" & ruleName, CStr(sheetSameRows), CStr(sheetDiffRows), "完成", "模板=" & templateWs.Name, ""







End Sub















Private Function ResolveTemplateSheet(ByVal templateSheets As Collection, ByVal sourceSheetName As String) As Worksheet







    Dim item As Variant















    For Each item In templateSheets







        If StrComp(CStr(item.Name), sourceSheetName, vbTextCompare) = 0 Then







            Set ResolveTemplateSheet = item







            Exit Function







        End If







    Next item















    If templateSheets.Count = 1 Then







        Set ResolveTemplateSheet = templateSheets(1)







    End If







End Function















Private Function BuildPathOccurrenceMap(ByVal ws As Worksheet, ByVal rowHeaderCols As Collection, ByVal colHeaderRows As Collection, ByVal dataStartRow As Long, ByVal dataEndRow As Long, ByVal dataStartCol As Long, ByVal dataEndCol As Long, ByVal skipKeywords As String, ByVal isRowPath As Boolean) As Object







    Dim actualEndRow As Long







    Dim actualEndCol As Long







    Dim headerText() As String







    Dim rowNo As Long







    Dim colNo As Long







    Dim onePath As String







    Dim onePos As String















    Set BuildPathOccurrenceMap = CreateObject("Scripting.Dictionary")







    BuildPathOccurrenceMap.CompareMode = vbTextCompare















    actualEndRow = ResolveSheetEndRow(ws, dataEndRow)







    actualEndCol = ResolveSheetEndCol(ws, dataEndCol)







    If actualEndRow < dataStartRow Or actualEndCol < dataStartCol Then Exit Function















    ReDim headerText(1 To colHeaderRows.Count, dataStartCol To actualEndCol)







    BuildHeaderCache ws, colHeaderRows, dataStartCol, actualEndCol, headerText















    If isRowPath Then







        For rowNo = dataStartRow To actualEndRow







            onePath = BuildRowPath(ws, rowNo, rowHeaderCols)







            If onePath <> "" Then







                If Not MatchAnyKeyword(onePath, skipKeywords) Then







                    onePos = BuildRowPositionLabel(rowNo, rowHeaderCols)







                    AddPathOccurrence BuildPathOccurrenceMap, onePath, onePos







                End If







            End If







        Next rowNo







    Else







        For colNo = dataStartCol To actualEndCol







            onePath = BuildColPath(headerText, colNo)







            If onePath <> "" Then







                If Not MatchAnyKeyword(onePath, skipKeywords) Then







                    onePos = BuildColPositionLabel(colNo, colHeaderRows)







                    AddPathOccurrence BuildPathOccurrenceMap, onePath, onePos







                End If







            End If







        Next colNo







    End If







End Function















Private Sub AddPathOccurrence(ByVal pathMap As Object, ByVal onePath As String, ByVal onePosition As String)







    Dim item As Object















    If pathMap.Exists(onePath) Then







        Set item = pathMap(onePath)







    Else







        Set item = CreateObject("Scripting.Dictionary")







        item.CompareMode = vbTextCompare







        item("count") = 0







        item("positions") = ""







        pathMap.Add onePath, item







    End If















    item("count") = CLng(item("count")) + 1







    If Trim$(CStr(item("positions"))) <> "" Then







        item("positions") = CStr(item("positions")) & "；" & onePosition







    Else







        item("positions") = onePosition







    End If







End Sub















Private Sub CompareOccurrenceMaps(ByVal compareWs As Worksheet, ByRef resultRow As Long, ByVal ruleName As String, ByVal bookKeywords As String, ByVal sheetKeywords As String, ByVal targetType As String, ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal templateMap As Object, ByVal sourceMap As Object, ByRef sameRows As Long, ByRef diffRows As Long, Optional ByVal outputAllRows As Boolean = True)







    Dim allKeys As Object







    Dim oneKey As Variant







    Dim templatePath As String







    Dim sourcePath As String







    Dim templatePos As String







    Dim sourcePos As String







    Dim compareResult As String















    Set allKeys = CreateObject("Scripting.Dictionary")







    allKeys.CompareMode = vbTextCompare















    MergeKeys allKeys, templateMap







    MergeKeys allKeys, sourceMap















    For Each oneKey In allKeys.Keys







        templatePath = ""







        sourcePath = ""







        templatePos = ""







        sourcePos = ""















        If Not templateMap Is Nothing Then







            If templateMap.Exists(CStr(oneKey)) Then







                templatePath = CStr(oneKey)







                templatePos = CStr(templateMap(CStr(oneKey))("positions"))







            End If







        End If















        If Not sourceMap Is Nothing Then







            If sourceMap.Exists(CStr(oneKey)) Then







                sourcePath = CStr(oneKey)







                sourcePos = CStr(sourceMap(CStr(oneKey))("positions"))







            End If







        End If















        compareResult = BuildPathCompareResultText(templateMap, sourceMap, CStr(oneKey))







        If compareResult = "一致" Then







            sameRows = sameRows + 1







            If outputAllRows Then







                WriteCompareResultRow compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, targetType, sourcePath, templatePath, sourceBookName, sourceSheetName, templatePos, sourcePos, compareResult, "按路径"







                resultRow = resultRow + 1







            End If







        Else







            diffRows = diffRows + 1







            WriteCompareResultRow compareWs, resultRow, ruleName, bookKeywords, sheetKeywords, targetType, sourcePath, templatePath, sourceBookName, sourceSheetName, templatePos, sourcePos, compareResult, "按路径"







            resultRow = resultRow + 1







        End If







    Next oneKey







End Sub















Private Sub MergeKeys(ByVal allKeys As Object, ByVal sourceMap As Object)







    Dim oneKey As Variant















    If sourceMap Is Nothing Then Exit Sub







    For Each oneKey In sourceMap.Keys







        If Not allKeys.Exists(CStr(oneKey)) Then







            allKeys(CStr(oneKey)) = True







        End If







    Next oneKey







End Sub















Private Function BuildPathCompareResultText(ByVal templateMap As Object, ByVal sourceMap As Object, ByVal oneKey As String) As String







    Dim templateExists As Boolean







    Dim sourceExists As Boolean







    Dim templateCount As Long







    Dim sourceCount As Long















    If Not templateMap Is Nothing Then templateExists = templateMap.Exists(oneKey)







    If Not sourceMap Is Nothing Then sourceExists = sourceMap.Exists(oneKey)















    If templateExists Then







        templateCount = CLng(templateMap(oneKey)("count"))







    End If







    If sourceExists Then







        sourceCount = CLng(sourceMap(oneKey)("count"))







    End If















    If templateExists And sourceExists Then







        If templateCount = sourceCount Then







            BuildPathCompareResultText = "一致"







        Else







            BuildPathCompareResultText = "数量变化"







        End If







    ElseIf sourceExists Then







        BuildPathCompareResultText = "模板缺失"







    Else







        BuildPathCompareResultText = "源表缺失"







    End If







End Function















Private Function ResolveSheetEndRow(ByVal ws As Worksheet, ByVal configuredEndRow As Long) As Long







    ResolveSheetEndRow = configuredEndRow







    If ResolveSheetEndRow <= 0 Then







        ResolveSheetEndRow = GetLastUsedRow(ws)







    End If







End Function















Private Function ResolveSheetEndCol(ByVal ws As Worksheet, ByVal configuredEndCol As Long) As Long







    ResolveSheetEndCol = configuredEndCol







    If ResolveSheetEndCol <= 0 Then







        ResolveSheetEndCol = GetLastUsedCol(ws)







    End If







End Function















Private Sub WriteCompareResultRow(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal ruleName As String, ByVal bookKeywords As String, ByVal sheetKeywords As String, ByVal targetType As String, ByVal sourcePath As String, ByVal templatePath As String, ByVal sourceBookName As String, ByVal sourceSheetName As String, ByVal templatePosition As String, ByVal sourcePosition As String, ByVal compareResult As String, Optional ByVal remarkText As String = "")







    ws.Cells(rowNo, crcEnabled).Value = ""







    ws.Cells(rowNo, crcMapName).Value = ""







    ws.Cells(rowNo, crcRuleName).Value = ruleName







    ws.Cells(rowNo, crcBookKeywords).Value = bookKeywords







    ws.Cells(rowNo, crcSheetKeywords).Value = sheetKeywords







    ws.Cells(rowNo, crcTargetType).Value = targetType







    ws.Cells(rowNo, crcMatchMode).Value = "精确"







    ws.Cells(rowNo, crcOriginalPath).Value = sourcePath







    ws.Cells(rowNo, crcStandardPath).Value = templatePath







    ws.Cells(rowNo, crcRemark).Value = remarkText







    ws.Cells(rowNo, crcSourceBook).Value = sourceBookName







    ws.Cells(rowNo, crcSourceSheet).Value = sourceSheetName







    ws.Cells(rowNo, crcTemplatePosition).Value = templatePosition







    ws.Cells(rowNo, crcSourcePosition).Value = sourcePosition







    ws.Cells(rowNo, crcCompareResult).Value = compareResult







End Sub















Private Function BuildCompareResultText(ByVal templatePath As String, ByVal sourcePath As String) As String







    If StrComp(templatePath, sourcePath, vbTextCompare) = 0 Then







        BuildCompareResultText = "一致"







    ElseIf templatePath = "" And sourcePath <> "" Then







        BuildCompareResultText = "模板缺失"







    ElseIf templatePath <> "" And sourcePath = "" Then







        BuildCompareResultText = "源表缺失"







    Else







        BuildCompareResultText = "文本变化"







    End If







End Function















Private Function BuildColPositionLabel(ByVal colNo As Long, ByVal colHeaderRows As Collection) As String







    BuildColPositionLabel = "列" & ColumnNumberToLetter(colNo) & " 行" & BuildHeaderRowsLabel(colHeaderRows)







End Function















Private Function BuildRowPositionLabel(ByVal rowNo As Long, ByVal rowHeaderCols As Collection) As String







    BuildRowPositionLabel = "行" & CStr(rowNo) & " 列" & BuildColumnCollectionLabel(rowHeaderCols)







End Function















Private Function BuildColumnCollectionLabel(ByVal rowHeaderCols As Collection) As String







    Dim idx As Long















    If rowHeaderCols Is Nothing Then Exit Function







    For idx = 1 To rowHeaderCols.Count







        If BuildColumnCollectionLabel <> "" Then







            BuildColumnCollectionLabel = BuildColumnCollectionLabel & ","







        End If







        BuildColumnCollectionLabel = BuildColumnCollectionLabel & ColumnNumberToLetter(CLng(rowHeaderCols(idx)))







    Next idx







End Function















Private Function PickOneWorkbookPath(ByVal dialogTitle As String) As String







    Dim fd As FileDialog















    Set fd = Application.FileDialog(msoFileDialogFilePicker)







    With fd







        .Title = dialogTitle







        .AllowMultiSelect = False







        .Filters.Clear







        .Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"







        If .Show = -1 Then







            PickOneWorkbookPath = CStr(.SelectedItems(1))







        End If







    End With







End Function















Private Function ParseNumberCollection(ByVal rawText As String, Optional ByRef invalidCount As Long = 0) As Collection



    Dim result As New Collection



    Dim arr As Variant



    Dim i As Long



    Dim token As String



    Dim parsedValue As Long







    rawText = Replace(rawText, "，", ",")



    rawText = Replace(rawText, "；", ",")



    arr = Split(rawText, ",")







    For i = LBound(arr) To UBound(arr)



        token = Trim$(CStr(arr(i)))



        If token <> "" Then



            If TryParseLongStrict(token, parsedValue) Then



                result.Add parsedValue



            Else



                invalidCount = invalidCount + 1



            End If



        End If



    Next i







    Set ParseNumberCollection = result



End Function















Private Function ParseColumnCollection(ByVal rawText As String, Optional ByRef invalidCount As Long = 0) As Collection



    Dim result As New Collection



    Dim arr As Variant



    Dim i As Long



    Dim token As String



    Dim parsedValue As Long







    rawText = Replace(rawText, "，", ",")



    rawText = Replace(rawText, "；", ",")



    arr = Split(rawText, ",")







    For i = LBound(arr) To UBound(arr)



        token = Trim$(CStr(arr(i)))



        If token <> "" Then



            If TryParseColumnStrict(token, parsedValue) Then



                result.Add parsedValue



            Else



                invalidCount = invalidCount + 1



            End If



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















    If TryParseColumnStrict(textValue, ParseColumnSpec) Then



        Exit Function



    End If



    ParseColumnSpec = defaultValue







End Function















Private Function ColLetterToNumber(ByVal colText As String) As Long







    Dim i As Long















    colText = UCase$(Trim$(colText))



    If colText = "" Then Exit Function



    For i = 1 To Len(colText)



        If Mid$(colText, i, 1) < "A" Or Mid$(colText, i, 1) > "Z" Then



            ColLetterToNumber = 0



            Exit Function



        End If



        ColLetterToNumber = ColLetterToNumber * 26 + Asc(Mid$(colText, i, 1)) - 64



    Next i







End Function







Private Function TryParseLongStrict(ByVal rawText As String, ByRef outValue As Long) As Boolean



    Dim txt As String







    txt = Trim$(CStr(rawText))



    If txt = "" Then Exit Function



    If Not IsNumeric(txt) Then Exit Function



    On Error GoTo ParseFail



    outValue = CLng(txt)



    TryParseLongStrict = True



    Exit Function



ParseFail:



    TryParseLongStrict = False



End Function







Private Function TryParseColumnStrict(ByVal rawText As String, ByRef outValue As Long) As Boolean



    Dim txt As String



    Dim colNo As Long







    txt = Trim$(CStr(rawText))



    If txt = "" Then Exit Function



    If IsNumeric(txt) Then



        On Error GoTo ParseFail



        colNo = CLng(txt)



        If colNo <= 0 Then Exit Function



        outValue = colNo



        TryParseColumnStrict = True



        Exit Function



    End If







    colNo = ColLetterToNumber(txt)



    If colNo <= 0 Then Exit Function



    outValue = colNo



    TryParseColumnStrict = True



    Exit Function



ParseFail:



    TryParseColumnStrict = False



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



    BuildSummaryMessage = BuildSummaryMessage



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






