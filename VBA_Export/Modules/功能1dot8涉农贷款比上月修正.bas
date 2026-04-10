Attribute VB_Name = "功能1dot8涉农贷款比上月修正"



Option Explicit







Private Const LOG_KEY As String = "2.7 修正涉农贷款比上月"



Private Const INST_SHEET_KEY1 As String = "惠州市"



Private Const INST_SHEET_KEY2 As String = "涉农贷款分机构"



Private Const AREA_SHEET_KEY1 As String = "本外币"



Private Const AREA_SHEET_KEY2 As String = "涉农贷款分地区"







Private Const HEADER_SCAN_ROWS As Long = 10



Private Const NAME_SCAN_COLS As Long = 3



Private Const INST_DATA_START_ROW As Long = 3



Private Const INST_NAME_COL As Long = 2



Private Const INST_VALUE_COL As Long = 3



Private Const INST_DIFF_COL As Long = 4



Private Const COMPARE_TOLERANCE As Double = 0.01







Public Sub 修正涉农贷款比上月()



    Dim t0 As Double



    Dim currentPath As String



    Dim previousPath As String



    Dim currentWb As Workbook



    Dim previousWb As Workbook



    Dim currentInstWs As Worksheet



    Dim previousInstWs As Worksheet



    Dim currentAreaWs As Worksheet



    Dim currentRows As Object



    Dim currentValues As Object



    Dim previousValues As Object



    Dim previousOnly As Collection



    Dim unmatchedCurrent As Collection



    Dim duplicateCurrent As Collection



    Dim duplicatePrevious As Collection



    Dim matchedCount As Long



    Dim fillTotal As Double



    Dim areaTotal As Double



    Dim checkMsg As String



    Dim currentStep As String



    Dim errNo As Long



    Dim errDesc As String



    Dim errSource As String



    Dim logResult As String



    Dim logDetail As String







    t0 = Timer



    currentStep = "初始化"



    RunLog_WriteRow LOG_KEY, "开始", "", "", "", "", "开始", ""







    currentStep = "选择本期文件"



    currentPath = PickExcelFile("请选择本期 Excel 文件")



    If currentPath = "" Then



        RunLog_WriteRow LOG_KEY, "结束", "", "", "", "", "取消选择本期文件", CStr(Round(Timer - t0, 2))



        Exit Sub



    End If







    currentStep = "选择上期文件"



    previousPath = PickExcelFile("请选择上期 Excel 文件")



    If previousPath = "" Then



        RunLog_WriteRow LOG_KEY, "结束", "", "", "", "", "取消选择上期文件", CStr(Round(Timer - t0, 2))



        Exit Sub



    End If







    Application.ScreenUpdating = False



    Application.DisplayAlerts = False



    On Error GoTo ErrHandler







    currentStep = "打开工作簿"



    Set currentWb = Workbooks.Open(currentPath, ReadOnly:=False, UpdateLinks:=0)



    Set previousWb = Workbooks.Open(previousPath, ReadOnly:=True, UpdateLinks:=0)







    currentStep = "定位工作表"



    Set currentInstWs = FindSheetByKeywords(currentWb, INST_SHEET_KEY1, INST_SHEET_KEY2)



    Set previousInstWs = FindSheetByKeywords(previousWb, INST_SHEET_KEY1, INST_SHEET_KEY2)



    Set currentAreaWs = FindSheetByKeywords(currentWb, AREA_SHEET_KEY1, AREA_SHEET_KEY2)







    If currentInstWs Is Nothing Then



        Err.Raise vbObjectError + 2001, , "本期文件未找到包含“惠州市”且包含“涉农贷款分机构”的工作表。"



    End If



    If previousInstWs Is Nothing Then



        Err.Raise vbObjectError + 2002, , "上期文件未找到包含“惠州市”且包含“涉农贷款分机构”的工作表。"



    End If



    If currentAreaWs Is Nothing Then



        Err.Raise vbObjectError + 2003, , "本期文件未找到包含“本外币”且包含“涉农贷款分地区”的工作表。"



    End If







    currentStep = "构建机构映射"



    Set currentRows = CreateObject("Scripting.Dictionary")



    currentRows.CompareMode = vbTextCompare



    Set duplicateCurrent = New Collection



    Set currentValues = BuildInstitutionValueMap(currentInstWs, INST_VALUE_COL, currentRows, duplicateCurrent)







    Set duplicatePrevious = New Collection



    Set previousValues = BuildInstitutionValueMap(previousInstWs, INST_VALUE_COL, Nothing, duplicatePrevious)







    If duplicateCurrent.Count > 0 Then



        Err.Raise vbObjectError + 2004, , "本期分机构表存在重复机构名称：" & JoinCollection(duplicateCurrent, "、")



    End If



    If duplicatePrevious.Count > 0 Then



        Err.Raise vbObjectError + 2005, , "上期分机构表存在重复机构名称：" & JoinCollection(duplicatePrevious, "、")



    End If







    currentStep = "回填涉农贷款比上月"



    Set unmatchedCurrent = New Collection



    fillTotal = FillMonthDiffValues(currentInstWs, currentRows, currentValues, previousValues, unmatchedCurrent, matchedCount)



    Set previousOnly = CollectMissingKeys(previousValues, currentValues)







    currentStep = "读取分地区总数"



    areaTotal = GetAreaMonthDiffTotal(currentAreaWs)



    checkMsg = BuildResultMessage(matchedCount, fillTotal, areaTotal, unmatchedCurrent, previousOnly)







    currentStep = "保存本期文件"



    currentWb.Save







    If Abs(fillTotal - areaTotal) <= COMPARE_TOLERANCE Then



        logResult = "成功"



    Else



        logResult = "校验不一致"



    End If



    logDetail = "匹配机构数=" & matchedCount



    RunLog_WriteRow LOG_KEY, "回填", currentWb.Name & " | " & previousWb.Name, CStr(areaTotal), CStr(fillTotal), logResult, logDetail, ""







    SafeCloseWorkbook previousWb, False



    SafeCloseWorkbook currentWb, False







    Application.DisplayAlerts = True



    Application.ScreenUpdating = True







    If Abs(fillTotal - areaTotal) > COMPARE_TOLERANCE Then



        MsgBox checkMsg, vbCritical, "涉农贷款比上月修正失败"



    ElseIf unmatchedCurrent.Count > 0 Or previousOnly.Count > 0 Then



        MsgBox checkMsg, vbExclamation, "涉农贷款比上月修正完成"



    Else



        MsgBox checkMsg, vbInformation, "涉农贷款比上月修正完成"



    End If







    RunLog_WriteRow LOG_KEY, "结束", currentPath & " | " & previousPath, "", "", "完成", "", CStr(Round(Timer - t0, 2))



    Exit Sub







ErrHandler:



    errNo = Err.Number



    errDesc = Err.Description



    errSource = Err.Source



    If Trim$(errDesc) = "" Then errDesc = "未返回错误描述"







    Application.DisplayAlerts = True



    Application.ScreenUpdating = True







    On Error Resume Next



    SafeCloseWorkbook previousWb, False



    SafeCloseWorkbook currentWb, False



    On Error GoTo 0







    checkMsg = "步骤：" & currentStep

    checkMsg = checkMsg & vbCrLf & "错误号：" & CStr(errNo)

    checkMsg = checkMsg & vbCrLf & "错误来源：" & errSource

    checkMsg = checkMsg & vbCrLf & "错误信息：" & errDesc







    RunLog_WriteRow LOG_KEY, "结束", currentPath & " | " & previousPath, "", "", "失败", checkMsg, CStr(Round(Timer - t0, 2))



    MsgBox checkMsg, vbCritical, "涉农贷款比上月修正失败"



End Sub







Private Function PickExcelFile(ByVal dialogTitle As String) As String



    Dim fd As FileDialog







    Set fd = Application.FileDialog(msoFileDialogFilePicker)



    With fd



        .Title = dialogTitle



        .AllowMultiSelect = False



        .Filters.Clear



        .Filters.Add "Excel 文件", "*.xls;*.xlsx;*.xlsm"



        If .Show = -1 Then



            PickExcelFile = CStr(.SelectedItems(1))



        Else



            PickExcelFile = ""



        End If



    End With



End Function







Private Function FindSheetByKeywords(ByVal wb As Workbook, ByVal keyword1 As String, ByVal keyword2 As String, Optional ByVal keyword3 As String = "") As Worksheet



    Dim ws As Worksheet



    Dim hitCount As Long



    Dim hitWs As Worksheet







    For Each ws In wb.Worksheets



        If ContainsText(ws.Name, keyword1) And ContainsText(ws.Name, keyword2) Then



            If keyword3 = "" Or ContainsText(ws.Name, keyword3) Then



                hitCount = hitCount + 1



                Set hitWs = ws



            End If



        End If



    Next ws







    If hitCount > 1 Then



        Err.Raise vbObjectError + 2101, , "工作簿“" & wb.Name & "”匹配到多个目标工作表，请保留唯一目标表。"



    End If







    Set FindSheetByKeywords = hitWs



End Function







Private Function BuildInstitutionValueMap(ByVal ws As Worksheet, ByVal valueCol As Long, ByVal rowMap As Object, ByRef duplicateNames As Collection) As Object



    Dim dict As Object



    Dim lastRow As Long



    Dim rowNo As Long



    Dim instName As String







    Set dict = CreateObject("Scripting.Dictionary")



    dict.CompareMode = vbTextCompare



    lastRow = GetLastUsedRow(ws)







    For rowNo = INST_DATA_START_ROW To lastRow



        instName = NormalizeInstitutionName(CStr(ws.Cells(rowNo, INST_NAME_COL).Value))



        If instName <> "" Then



            If dict.Exists(instName) Then



                duplicateNames.Add instName



            Else



                dict(instName) = ToNumber(ws.Cells(rowNo, valueCol).Value)



                If Not rowMap Is Nothing Then rowMap(instName) = rowNo



            End If



        End If



    Next rowNo







    Set BuildInstitutionValueMap = dict



End Function







Private Function FillMonthDiffValues(ByVal ws As Worksheet, ByVal rowMap As Object, ByVal currentValues As Object, ByVal previousValues As Object, ByRef unmatchedCurrent As Collection, ByRef matchedCount As Long) As Double



    Dim key As Variant



    Dim rowNo As Long



    Dim diffValue As Double







    FillMonthDiffValues = 0



    matchedCount = 0







    For Each key In rowMap.Keys



        rowNo = CLng(rowMap(key))



        If previousValues.Exists(CStr(key)) Then



            diffValue = CDbl(currentValues(key)) - CDbl(previousValues(key))



            ws.Cells(rowNo, INST_DIFF_COL).Value = diffValue



            FillMonthDiffValues = FillMonthDiffValues + diffValue



            matchedCount = matchedCount + 1



        Else



            ws.Cells(rowNo, INST_DIFF_COL).ClearContents



            unmatchedCurrent.Add CStr(key)



        End If



    Next key



End Function







Private Function CollectMissingKeys(ByVal sourceDict As Object, ByVal targetDict As Object) As Collection



    Dim key As Variant



    Dim result As New Collection







    For Each key In sourceDict.Keys



        If Not targetDict.Exists(CStr(key)) Then



            result.Add CStr(key)



        End If



    Next key







    Set CollectMissingKeys = result



End Function







Private Function GetAreaMonthDiffTotal(ByVal ws As Worksheet) As Double



    Dim headerCol As Long



    Dim cityRow As Long







    headerCol = FindHeaderColumn(ws, "涉农贷款比上月")



    If headerCol = 0 Then



        Err.Raise vbObjectError + 2201, , "未在“本外币涉农贷款分地区”表中找到表头“涉农贷款比上月”。"



    End If







    cityRow = FindCityRow(ws, "惠州市")



    If cityRow = 0 Then



        Err.Raise vbObjectError + 2202, , "未在“本外币涉农贷款分地区”表中找到“惠州市”所在行。"



    End If







    GetAreaMonthDiffTotal = ToNumber(ws.Cells(cityRow, headerCol).Value)



End Function







Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long



    Dim lastCol As Long



    Dim rowNo As Long



    Dim colNo As Long



    Dim maxScanRow As Long







    lastCol = GetLastUsedCol(ws)



    maxScanRow = GetLastUsedRow(ws)



    If maxScanRow < 1 Then Exit Function



    If maxScanRow > HEADER_SCAN_ROWS Then maxScanRow = HEADER_SCAN_ROWS







    For rowNo = 1 To maxScanRow



        For colNo = 1 To lastCol



            If NormalizeInstitutionName(CStr(ws.Cells(rowNo, colNo).Value)) = headerText Then



                FindHeaderColumn = colNo



                Exit Function



            End If



        Next colNo



    Next rowNo



End Function







Private Function FindCityRow(ByVal ws As Worksheet, ByVal cityName As String) As Long



    Dim lastRow As Long



    Dim rowNo As Long



    Dim colNo As Long







    lastRow = GetLastUsedRow(ws)



    For rowNo = 1 To lastRow



        For colNo = 1 To NAME_SCAN_COLS



            If NormalizeInstitutionName(CStr(ws.Cells(rowNo, colNo).Value)) = cityName Then



                FindCityRow = rowNo



                Exit Function



            End If



        Next colNo



    Next rowNo



End Function







Private Function BuildResultMessage(ByVal matchedCount As Long, ByVal fillTotal As Double, ByVal areaTotal As Double, ByVal unmatchedCurrent As Collection, ByVal previousOnly As Collection) As String



    Dim msg As String







    msg = "处理完成。"
    msg = msg & vbCrLf & vbCrLf & "已回填机构数：" & matchedCount
    msg = msg & vbCrLf & "分机构合计：" & Format(fillTotal, "0.00############")
    msg = msg & vbCrLf & "分地区总数：" & Format(areaTotal, "0.00############")







    If unmatchedCurrent.Count > 0 Then



        msg = msg & vbCrLf & vbCrLf & "本期未匹配到上期的机构：" & vbCrLf & JoinCollection(unmatchedCurrent, vbCrLf)



    End If







    If previousOnly.Count > 0 Then



        msg = msg & vbCrLf & vbCrLf & "上期存在但本期未出现的机构：" & vbCrLf & JoinCollection(previousOnly, vbCrLf)



    End If







    If Abs(fillTotal - areaTotal) > COMPARE_TOLERANCE Then



        msg = msg & vbCrLf & vbCrLf & "校验失败：分机构合计与分地区“惠州市-涉农贷款比上月”不一致。"



    Else



        msg = msg & vbCrLf & vbCrLf & "校验通过：分机构合计与分地区总数一致。"



    End If







    BuildResultMessage = msg



End Function







Private Function NormalizeInstitutionName(ByVal rawText As String) As String



    Dim txt As String







    txt = CStr(rawText)



    txt = Replace(txt, vbCr, " ")



    txt = Replace(txt, vbLf, " ")



    txt = Replace(txt, ChrW(&H3000), " ")



    txt = Replace(txt, Chr(160), " ")



    txt = Trim$(txt)







    Do While InStr(txt, "  ") > 0



        txt = Replace(txt, "  ", " ")



    Loop







    NormalizeInstitutionName = txt



End Function







Private Function ContainsText(ByVal sourceText As String, ByVal keyword As String) As Boolean



    ContainsText = (InStr(1, CStr(sourceText), CStr(keyword), vbTextCompare) > 0)



End Function







Private Function ToNumber(ByVal rawValue As Variant) As Double



    Dim txt As String







    If IsError(rawValue) Or IsEmpty(rawValue) Then Exit Function



    If IsNumeric(rawValue) Then



        ToNumber = CDbl(rawValue)



        Exit Function



    End If







    txt = Trim$(CStr(rawValue))



    If txt = "" Or txt = "-" Then Exit Function



    txt = Replace(txt, ",", "")



    If IsNumeric(txt) Then ToNumber = CDbl(txt)



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







Private Function JoinCollection(ByVal items As Collection, ByVal delimiter As String) As String



    Dim i As Long







    For i = 1 To items.Count



        If JoinCollection <> "" Then JoinCollection = JoinCollection & delimiter



        JoinCollection = JoinCollection & CStr(items(i))



    Next i



End Function

Private Sub SafeCloseWorkbook(ByRef wb As Workbook, Optional ByVal saveChanges As Boolean = False)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=saveChanges
        Set wb = Nothing
    End If
    On Error GoTo 0
End Sub
