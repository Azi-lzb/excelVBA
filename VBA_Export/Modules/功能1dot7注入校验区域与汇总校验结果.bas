Attribute VB_Name = "FnInjectCheckResult"
Option Explicit

' 1.7 注入校验区域：按模板批注区域将公式/值/批注注入源文件
' 1.8 汇总校验结果：汇总源文件中校验区域内命中的错误结果
' 1.9 清空校验区域：按关键字清空源文件中指定批注区域
' 1.10 一键校验流程：依次执行 1.7 / 1.8 / 1.9

Private Const PANEL_SHEET_NAME As String = "执行面板"
Private Const PANEL_DATA_START_ROW As Long = 5
Private Const PANEL_COL_PATH As Long = 2
Private Const RESULT_SHEET_NAME As String = "校验结果"
Private Const TMPL_KEY_LOCAL As String = "本地校验区域"
Private Const TMPL_KEY_CROSS As String = "跨表校验区域"
Private Const TMPL_SHEET_NAME As String = "模板"

Private Enum ResultCols
    rcTmplWb = 1
    rcTmplWs = 2
    rcSrcWb = 3
    rcSrcWs = 4
    rcRow = 5
    rcCol = 6
    rcRowName = 7
    rcErrType = 8
    rcRsv1 = 9
    rcRsv2 = 10
    rcRsv3 = 11
    rcRsv4 = 12
    rcComment = 13
End Enum

Private Function GetPanelWs() As Worksheet
    On Error Resume Next
    Set GetPanelWs = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function GetTemplatePathFromPanel(ByVal wsPanel As Worksheet) As String
    If wsPanel Is Nothing Then
        GetTemplatePathFromPanel = ""
    Else
        GetTemplatePathFromPanel = Trim$(CStr(wsPanel.Cells(2, 1).Value))
    End If
End Function

Private Sub CollectSourcePathsFromPanel(ByVal wsPanel As Worksheet, ByRef outPaths As Collection)
    Dim lastRow As Long
    Dim r As Long
    Dim p As String

    Set outPaths = New Collection
    If wsPanel Is Nothing Then Exit Sub

    lastRow = wsPanel.Cells(wsPanel.Rows.Count, PANEL_COL_PATH).End(xlUp).Row
    If lastRow < PANEL_DATA_START_ROW Then Exit Sub

    For r = PANEL_DATA_START_ROW To lastRow
        p = Trim$(CStr(wsPanel.Cells(r, PANEL_COL_PATH).Value))
        If p <> "" Then outPaths.Add p
    Next r
End Sub

Private Function Col2Num(ByVal s As String) As Long
    Dim i As Long
    Dim n As Long

    n = 0
    For i = 1 To Len(s)
        n = n * 26 + Asc(UCase$(Mid$(s, i, 1))) - 64
    Next i
    Col2Num = n
End Function

Private Sub SplitAddr(ByVal a As String, ByRef cs As String, ByRef rn As Long)
    Dim i As Long

    cs = ""
    rn = 0
    a = Replace$(a, "$", "")
    For i = 1 To Len(a)
        If Mid$(a, i, 1) Like "[A-Za-z]" Then
            cs = cs & UCase$(Mid$(a, i, 1))
        Else
            rn = CLng(Mid$(a, i))
            Exit For
        End If
    Next i
End Sub

Private Function CellCommentText(ByVal c As Range) As String
    On Error Resume Next
    If c Is Nothing Then
        CellCommentText = ""
    ElseIf c.Comment Is Nothing Then
        CellCommentText = ""
    Else
        CellCommentText = c.Comment.Text
        If CellCommentText = "" Then CellCommentText = c.Comment.Comment.Text
    End If
    On Error GoTo 0
End Function

Private Sub ExtractComments(ByVal ws As Worksheet, ByRef atc As Object)
    Dim cm As Comment
    Dim addr As String

    Set atc = CreateObject("Scripting.Dictionary")
    atc.CompareMode = vbTextCompare
    If ws Is Nothing Then Exit Sub

    For Each cm In ws.Comments
        addr = cm.Parent.Address(False, False)
        atc(addr) = cm.Text
    Next cm
End Sub

Private Function ExtractRegions(ByRef atc As Object, ByVal kw As String) As Collection
    Dim nums As Object
    Dim k As Variant
    Dim txt As String
    Dim p As Long
    Dim ci As Long
    Dim sfx As String
    Dim ds As String
    Dim numArr() As Long
    Dim idx As Long
    Dim ii As Long
    Dim jj As Long
    Dim tmp As Long
    Dim sa As String
    Dim ea As String
    Dim sc As String
    Dim sr As Long
    Dim ec As String
    Dim er As Long
    Dim scn As Long
    Dim ecn As Long

    Set ExtractRegions = New Collection
    If atc Is Nothing Then Exit Function

    Set nums = CreateObject("Scripting.Dictionary")
    nums.CompareMode = vbTextCompare

    For Each k In atc.Keys
        txt = CStr(atc(k))
        p = InStr(1, txt, kw, vbTextCompare)
        If p > 0 Then
            sfx = Mid$(txt, p + Len(kw))
            If Left$(sfx, 1) = "#" Then sfx = Mid$(sfx, 2)
            ds = ""
            For ci = 1 To Len(sfx)
                If Mid$(sfx, ci, 1) Like "[0-9]" Then
                    ds = ds & Mid$(sfx, ci, 1)
                Else
                    Exit For
                End If
            Next ci
            If ds <> "" Then nums(CLng(ds)) = True
        End If
    Next k

    If nums.Count = 0 Then Exit Function

    ReDim numArr(1 To nums.Count)
    idx = 1
    For Each k In nums.Keys
        numArr(idx) = CLng(k)
        idx = idx + 1
    Next k

    For ii = 1 To UBound(numArr) - 1
        For jj = ii + 1 To UBound(numArr)
            If numArr(ii) > numArr(jj) Then
                tmp = numArr(ii)
                numArr(ii) = numArr(jj)
                numArr(jj) = tmp
            End If
        Next jj
    Next ii

    For ii = 1 To UBound(numArr)
        sa = ""
        ea = ""
        For Each k In atc.Keys
            txt = CStr(atc(k))
            If InStr(1, txt, kw & CStr(numArr(ii)), vbTextCompare) > 0 And _
               InStr(1, txt, kw & "#" & CStr(numArr(ii)), vbTextCompare) = 0 Then
                sa = CStr(k)
            End If
            If InStr(1, txt, kw & "#" & CStr(numArr(ii)), vbTextCompare) > 0 Then
                ea = CStr(k)
            End If
        Next k

        If sa <> "" And ea <> "" Then
            SplitAddr sa, sc, sr
            SplitAddr ea, ec, er
            If sr > er Then tmp = sr: sr = er: er = tmp
            scn = Col2Num(sc)
            ecn = Col2Num(ec)
            If scn > ecn Then tmp = scn: scn = ecn: ecn = tmp
            ExtractRegions.Add Array(sr, er, scn, ecn)
        End If
    Next ii
End Function

Private Function MergeRegions(ByVal a As Collection, ByVal b As Collection) As Collection
    Dim out As New Collection
    Dim i As Long

    If Not a Is Nothing Then
        For i = 1 To a.Count
            out.Add a(i)
        Next i
    End If
    If Not b Is Nothing Then
        For i = 1 To b.Count
            out.Add b(i)
        Next i
    End If
    Set MergeRegions = out
End Function

Private Function EnsureResultSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RESULT_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = RESULT_SHEET_NAME
    End If
    Set EnsureResultSheet = ws
End Function

Private Sub InitResultHeader(ByVal ws As Worksheet)
    If ws Is Nothing Then Exit Sub

    ws.Cells.Clear
    ws.Cells(1, rcTmplWb).Value = "模板工作簿名"
    ws.Cells(1, rcTmplWs).Value = "模板工作表名"
    ws.Cells(1, rcSrcWb).Value = "源文件工作簿名"
    ws.Cells(1, rcSrcWs).Value = "源文件工作表名"
    ws.Cells(1, rcRow).Value = "行序号"
    ws.Cells(1, rcCol).Value = "列序号"
    ws.Cells(1, rcRowName).Value = "行名称"
    ws.Cells(1, rcErrType).Value = "错误类型"
    ws.Cells(1, rcRsv1).Value = "预留字段1"
    ws.Cells(1, rcRsv2).Value = "预留字段2"
    ws.Cells(1, rcRsv3).Value = "预留字段3"
    ws.Cells(1, rcRsv4).Value = "预留字段4"
    ws.Cells(1, rcComment).Value = "错误单元格的批注描述"
    ws.Rows(1).Font.Bold = True
End Sub

Private Function ContainsAnyKey(ByVal s As String) As Boolean
    Dim t As String

    t = CStr(s)
    If t = "" Then
        ContainsAnyKey = False
    Else
        ContainsAnyKey = (InStr(1, t, "错", vbTextCompare) > 0 Or _
                          InStr(1, t, "校验失败", vbTextCompare) > 0 Or _
                          InStr(1, t, "硬性", vbTextCompare) > 0 Or _
                          InStr(1, t, "软性", vbTextCompare) > 0 Or _
                          InStr(1, t, "警告", vbTextCompare) > 0)
    End If
End Function

Private Sub ParsePipe5(ByVal s As String, ByRef errType As String, ByRef r1 As String, ByRef r2 As String, ByRef r3 As String, ByRef r4 As String)
    Dim arr As Variant

    errType = ""
    r1 = ""
    r2 = ""
    r3 = ""
    r4 = ""
    If s = "" Then Exit Sub

    arr = Split(CStr(s), "|")
    If UBound(arr) >= 0 Then errType = CStr(arr(0))
    If UBound(arr) >= 1 Then r1 = CStr(arr(1))
    If UBound(arr) >= 2 Then r2 = CStr(arr(2))
    If UBound(arr) >= 3 Then r3 = CStr(arr(3))
    If UBound(arr) >= 4 Then r4 = CStr(arr(4))
End Sub

Public Sub 注入校验区域()
    Dim t0 As Double
    Dim logKey As String
    Dim wsPanel As Worksheet
    Dim tmplPath As String
    Dim srcPaths As Collection
    Dim tmplWb As Workbook
    Dim srcWb As Workbook
    Dim tmplWs As Worksheet
    Dim srcWs As Worksheet
    Dim atc As Object
    Dim regsLocal As Collection
    Dim regsCross As Collection
    Dim regs As Collection
    Dim rg As Variant
    Dim r As Long
    Dim c As Long
    Dim fi As Long
    Dim okCnt As Long
    Dim failCnt As Long
    Dim cmTxt As String

    t0 = Timer
    logKey = "1.7 注入校验区域"
    On Error Resume Next
    RunLog_WriteRow logKey, "开始", "", "", "", "", "开始", ""
    On Error GoTo 0

    Set wsPanel = GetPanelWs()
    If wsPanel Is Nothing Then
        MsgBox "未找到执行面板，请先初始化执行面板。", vbExclamation
        RunLog_WriteRow logKey, "完成", "", "", "", "", "执行面板不存在", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    tmplPath = GetTemplatePathFromPanel(wsPanel)
    If tmplPath = "" Then
        MsgBox "执行面板 A2 未填写模板路径。", vbExclamation
        RunLog_WriteRow logKey, "完成", "", "", "", "", "模板路径为空", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    CollectSourcePathsFromPanel wsPanel, srcPaths
    If srcPaths Is Nothing Or srcPaths.Count = 0 Then
        MsgBox "执行面板 B5 起没有源文件路径。", vbExclamation
        RunLog_WriteRow logKey, "完成", "", "", "", "", "源文件为空", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrInject

    Set tmplWb = Workbooks.Open(tmplPath, ReadOnly:=True, UpdateLinks:=0)

    For fi = 1 To srcPaths.Count
        On Error Resume Next
        Set srcWb = Workbooks.Open(CStr(srcPaths(fi)), ReadOnly:=False, UpdateLinks:=0)
        If Err.Number <> 0 Or srcWb Is Nothing Then
            failCnt = failCnt + 1
            RunLog_WriteRow logKey, "打开源文件", CStr(srcPaths(fi)), "", "", "失败", Err.Number & " " & Err.Description, ""
            Err.Clear
            On Error GoTo 0
            GoTo NextFile
        End If
        On Error GoTo 0

        For Each srcWs In srcWb.Worksheets
            On Error Resume Next
            Set tmplWs = Nothing
            Set tmplWs = tmplWb.Worksheets(srcWs.Name)
            If tmplWs Is Nothing Then Set tmplWs = tmplWb.Worksheets(TMPL_SHEET_NAME)
            On Error GoTo 0
            If tmplWs Is Nothing Then GoTo NextSheet

            ExtractComments tmplWs, atc
            Set regsLocal = ExtractRegions(atc, TMPL_KEY_LOCAL)
            Set regsCross = ExtractRegions(atc, TMPL_KEY_CROSS)
            Set regs = MergeRegions(regsLocal, regsCross)
            If regs.Count = 0 Then GoTo NextSheet

            For Each rg In regs
                For r = CLng(rg(0)) To CLng(rg(1))
                    For c = CLng(rg(2)) To CLng(rg(3))
                        If tmplWs.Cells(r, c).HasFormula Then
                            srcWs.Cells(r, c).Formula = tmplWs.Cells(r, c).Formula
                        Else
                            srcWs.Cells(r, c).Value2 = tmplWs.Cells(r, c).Value2
                        End If

                        cmTxt = CellCommentText(tmplWs.Cells(r, c))
                        If cmTxt <> "" Then
                            On Error Resume Next
                            If Not srcWs.Cells(r, c).Comment Is Nothing Then srcWs.Cells(r, c).Comment.Delete
                            srcWs.Cells(r, c).AddComment cmTxt
                            On Error GoTo 0
                        End If
                    Next c
                Next r
            Next rg
NextSheet:
        Next srcWs

        srcWb.Save
        okCnt = okCnt + 1
        RunLog_WriteRow logKey, "保存源文件", srcWb.Name, "", "", "成功", "OK", ""
        srcWb.Close SaveChanges:=False
NextFile:
        Set srcWb = Nothing
    Next fi

    tmplWb.Close SaveChanges:=False
    Set tmplWb = Nothing
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    RunLog_WriteRow logKey, "完成", "", "", "", "", "成功 " & okCnt & " 个，失败 " & failCnt & " 个", CStr(Round(Timer - t0, 2))
    MsgBox "注入完成" & vbCrLf & "成功：" & okCnt & vbCrLf & "失败：" & failCnt, vbInformation
    Exit Sub

ErrInject:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If Not tmplWb Is Nothing Then tmplWb.Close SaveChanges:=False
    On Error GoTo 0
    RunLog_WriteRow logKey, "完成", "", "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "发生错误 " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Public Sub 一键校验流程()
    注入校验区域
    汇总校验结果
    清空校验区域
End Sub

Public Sub 汇总校验结果()
    Dim t0 As Double
    Dim logKey As String
    Dim wsPanel As Worksheet
    Dim tmplPath As String
    Dim srcPaths As Collection
    Dim wsRes As Worksheet
    Dim tmplWb As Workbook
    Dim srcWb As Workbook
    Dim tmplWs As Worksheet
    Dim srcWs As Worksheet
    Dim atc As Object
    Dim regsLocal As Collection
    Dim regsCross As Collection
    Dim regs As Collection
    Dim rg As Variant
    Dim r As Long
    Dim c As Long
    Dim fi As Long
    Dim cellVal As String
    Dim errType As String
    Dim r1 As String
    Dim r2 As String
    Dim r3 As String
    Dim r4 As String
    Dim cmTxt As String
    Dim rowName As String
    Dim outRow As Long
    Dim hitCnt As Long

    t0 = Timer
    logKey = "1.8 汇总校验结果"
    On Error Resume Next
    RunLog_WriteRow logKey, "开始", "", "", "", "", "开始", ""
    On Error GoTo 0

    Set wsPanel = GetPanelWs()
    If wsPanel Is Nothing Then
        MsgBox "未找到执行面板，请先初始化执行面板。", vbExclamation
        RunLog_WriteRow logKey, "完成", "", "", "", "", "执行面板不存在", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    tmplPath = GetTemplatePathFromPanel(wsPanel)
    If tmplPath = "" Then
        MsgBox "执行面板 A2 未填写模板路径。", vbExclamation
        RunLog_WriteRow logKey, "完成", "", "", "", "", "模板路径为空", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    CollectSourcePathsFromPanel wsPanel, srcPaths
    If srcPaths Is Nothing Or srcPaths.Count = 0 Then
        MsgBox "执行面板 B5 起没有源文件路径。", vbExclamation
        RunLog_WriteRow logKey, "完成", "", "", "", "", "源文件为空", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    Set wsRes = EnsureResultSheet()
    InitResultHeader wsRes
    outRow = 2

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrSum

    ' [优化F] 预分配结果数组，避免逐格写入工作表
    Const MAX_RESULT As Long = 50000
    Dim resArr() As Variant
    ReDim resArr(1 To MAX_RESULT, 1 To 13)
    Dim arrIdx As Long
    arrIdx = 0

    ' [优化G] 按模板表名缓存批注字典，避免重复提取
    Dim atcCache As Object
    Set atcCache = CreateObject("Scripting.Dictionary")
    atcCache.CompareMode = vbTextCompare

    Set tmplWb = Workbooks.Open(tmplPath, ReadOnly:=True, UpdateLinks:=0)

    For fi = 1 To srcPaths.Count
        On Error Resume Next
        Set srcWb = Workbooks.Open(CStr(srcPaths(fi)), ReadOnly:=True, UpdateLinks:=0)
        If Err.Number <> 0 Or srcWb Is Nothing Then
            RunLog_WriteRow logKey, "打开源文件", CStr(srcPaths(fi)), "", "", "失败", Err.Number & " " & Err.Description, ""
            Err.Clear
            On Error GoTo 0
            GoTo NextFile2
        End If
        On Error GoTo 0

        For Each srcWs In srcWb.Worksheets
            On Error Resume Next
            Set tmplWs = Nothing
            Set tmplWs = tmplWb.Worksheets(srcWs.Name)
            If tmplWs Is Nothing Then Set tmplWs = tmplWb.Worksheets(TMPL_SHEET_NAME)
            On Error GoTo 0
            If tmplWs Is Nothing Then GoTo NextSheet2

            ' [优化G] 使用缓存，同一模板表不重复提取批注
            If Not atcCache.Exists(tmplWs.Name) Then
                ExtractComments tmplWs, atc
                Set atcCache(tmplWs.Name) = atc
            Else
                Set atc = atcCache(tmplWs.Name)
            End If
            Set regsLocal = ExtractRegions(atc, TMPL_KEY_LOCAL)
            Set regsCross = ExtractRegions(atc, TMPL_KEY_CROSS)
            Set regs = MergeRegions(regsLocal, regsCross)
            If regs.Count = 0 Then GoTo NextSheet2

            For Each rg In regs
                For r = CLng(rg(0)) To CLng(rg(1))
                    rowName = CStr(srcWs.Cells(r, 1).Value)
                    For c = CLng(rg(2)) To CLng(rg(3))
                        cellVal = CStr(srcWs.Cells(r, c).Value)
                        If ContainsAnyKey(cellVal) Then
                            ParsePipe5 cellVal, errType, r1, r2, r3, r4
                            cmTxt = CellCommentText(srcWs.Cells(r, c))
                            ' [优化F] 写入数组而非逐格写工作表
                            arrIdx = arrIdx + 1
                            If arrIdx <= MAX_RESULT Then
                                resArr(arrIdx, rcTmplWb) = tmplWb.Name
                                resArr(arrIdx, rcTmplWs) = tmplWs.Name
                                resArr(arrIdx, rcSrcWb) = srcWb.Name
                                resArr(arrIdx, rcSrcWs) = srcWs.Name
                                resArr(arrIdx, rcRow) = r
                                resArr(arrIdx, rcCol) = c
                                resArr(arrIdx, rcRowName) = rowName
                                resArr(arrIdx, rcErrType) = errType
                                resArr(arrIdx, rcRsv1) = r1
                                resArr(arrIdx, rcRsv2) = r2
                                resArr(arrIdx, rcRsv3) = r3
                                resArr(arrIdx, rcRsv4) = r4
                                resArr(arrIdx, rcComment) = cmTxt
                            End If
                            hitCnt = hitCnt + 1
                        End If
                    Next c
                Next r
            Next rg
NextSheet2:
        Next srcWs

        RunLog_WriteRow logKey, "扫描源文件", srcWb.Name, "", "", "成功", "累计命中 " & hitCnt, ""
        srcWb.Close SaveChanges:=False
NextFile2:
        Set srcWb = Nothing
    Next fi

    tmplWb.Close SaveChanges:=False
    Set tmplWb = Nothing
    ' [优化F] 一次性写入结果数组
    If arrIdx > 0 Then
        Dim writeCount As Long
        writeCount = IIf(arrIdx > MAX_RESULT, MAX_RESULT, arrIdx)
        wsRes.Cells(2, 1).Resize(writeCount, 13).Value = resArr
    End If
    wsRes.Columns.AutoFit
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    RunLog_WriteRow logKey, "完成", "", "", "", "", "命中 " & hitCnt & " 条", CStr(Round(Timer - t0, 2))
    MsgBox "汇总完成" & vbCrLf & "命中：" & hitCnt & vbCrLf & "结果已写入工作表《" & RESULT_SHEET_NAME & "》", vbInformation
    Exit Sub

ErrSum:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If Not tmplWb Is Nothing Then tmplWb.Close SaveChanges:=False
    On Error GoTo 0
    RunLog_WriteRow logKey, "完成", "", "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "发生错误 " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Public Sub 清空校验区域()
    Dim t0 As Double
    Dim logKey As String
    Dim wsPanel As Worksheet
    Dim srcPaths As Collection
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim atc As Object
    Dim regs As Collection
    Dim rg As Variant
    Dim rg2 As Variant
    Dim r As Long
    Dim c As Long
    Dim fi As Long
    Dim keyInput As String
    Dim keyArr() As String
    Dim iKey As Long
    Dim oneRegs As Collection

    t0 = Timer
    logKey = "1.9 清空校验区域"
    On Error Resume Next
    RunLog_WriteRow logKey, "开始", "", "", "", "", "开始", ""
    On Error GoTo 0

    Set wsPanel = GetPanelWs()
    If wsPanel Is Nothing Then
        MsgBox "未找到执行面板，请先初始化执行面板。", vbExclamation
        RunLog_WriteRow logKey, "完成", "", "", "", "", "执行面板不存在", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    CollectSourcePathsFromPanel wsPanel, srcPaths
    If srcPaths Is Nothing Or srcPaths.Count = 0 Then
        MsgBox "执行面板 B5 起没有源文件路径。", vbExclamation
        RunLog_WriteRow logKey, "完成", "", "", "", "", "源文件为空", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    keyInput = InputBox("请输入需要清空的批注区域关键字，多个关键字用英文逗号分隔。", _
                        "清空校验区域", _
                        "本地校验区域,跨表校验区域")
    If Trim$(keyInput) = "" Then Exit Sub
    keyArr = Split(keyInput, ",")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    On Error GoTo ErrClear

    For fi = 1 To srcPaths.Count
        On Error Resume Next
        Set srcWb = Workbooks.Open(CStr(srcPaths(fi)), ReadOnly:=False, UpdateLinks:=0)
        If Err.Number <> 0 Or srcWb Is Nothing Then
            RunLog_WriteRow logKey, "打开源文件", CStr(srcPaths(fi)), "", "", "失败", Err.Number & " " & Err.Description, ""
            Err.Clear
            On Error GoTo 0
            GoTo NextFileClear
        End If
        On Error GoTo 0

        For Each srcWs In srcWb.Worksheets
            ExtractComments srcWs, atc
            Set regs = New Collection
            For iKey = LBound(keyArr) To UBound(keyArr)
                Set oneRegs = ExtractRegions(atc, Trim$(CStr(keyArr(iKey))))
                If Not oneRegs Is Nothing Then
                    For Each rg2 In oneRegs
                        regs.Add rg2
                    Next rg2
                End If
            Next iKey
            If regs.Count = 0 Then GoTo NextSheetClear

            ' [优化H] 先 Union 所有区域，再一次 ClearContents + ClearComments
            Dim clearRng As Range
            Dim rgBlock As Range
            For Each rg In regs
                Set rgBlock = srcWs.Range(srcWs.Cells(CLng(rg(0)), CLng(rg(2))), srcWs.Cells(CLng(rg(1)), CLng(rg(3))))
                If clearRng Is Nothing Then
                    Set clearRng = rgBlock
                Else
                    Set clearRng = Union(clearRng, rgBlock)
                End If
            Next rg
            If Not clearRng Is Nothing Then
                clearRng.ClearContents
                On Error Resume Next
                clearRng.ClearComments
                On Error GoTo 0
                Set clearRng = Nothing
            End If
NextSheetClear:
        Next srcWs

        srcWb.Save
        RunLog_WriteRow logKey, "保存源文件", srcWb.Name, "", "", "成功", "OK", ""
        srcWb.Close SaveChanges:=False
NextFileClear:
        Set srcWb = Nothing
    Next fi

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    RunLog_WriteRow logKey, "完成", "", "", "", "", "完成", CStr(Round(Timer - t0, 2))
    MsgBox "清空完成", vbInformation
    Exit Sub

ErrClear:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    On Error GoTo 0
    RunLog_WriteRow logKey, "完成", "", "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "发生错误 " & Err.Number & ": " & Err.Description, vbCritical
End Sub
