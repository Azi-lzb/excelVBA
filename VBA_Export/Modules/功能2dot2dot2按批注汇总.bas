Attribute VB_Name = "功能2dot2dot2按批注汇总"
Option Explicit

' 2.2.2 汇总：从执行面板读取模板与源文件列表，按批注配置汇总，并按 config 分列（支持按 Sheet 或强制按模板）

Private Const CFG_KEY_SPLIT_SHEET As String = "2.2.2 按批注汇总"
Private Const TMPL_SHT    As String = "模板"
Private Const PANEL_SHEET_NAME As String = "执行面板"
Private Const PANEL_HEADER_ROW As Long = 4
Private Const PANEL_DATA_START_ROW As Long = 5
Private Const PANEL_COL_PATH As Long = 2
Private Const PANEL_COL_SHORT As Long = 3

' ============================================================
'  从 config 读配置
' ============================================================
Private Function CfgVal(ByVal cfgKey As String, ByVal kn As String) As String
    Dim ws As Worksheet, r As Long
    CfgVal = ""
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If Trim(CStr(ws.Cells(r, 1).value)) = cfgKey And _
           LCase(Trim(CStr(ws.Cells(r, 2).value))) = LCase(kn) Then
            CfgVal = Trim(CStr(ws.Cells(r, 3).value))
            Exit Function
        End If
    Next r
End Function

Private Function CfgBool(ByVal cfgKey As String, ByVal kn As String, Optional ByVal def As Boolean = False) As Boolean
    Dim v As String: v = CfgVal(cfgKey, kn)
    If v = "" Then CfgBool = def: Exit Function
    CfgBool = (v = "1" Or LCase(v) = "是" Or LCase(v) = "true")
End Function

' ============================================================
'  按 config 分列配置对当前表指定列按分隔符拆成多列。
' ============================================================
Private Sub ApplySplitColumns(ByVal ws As Worksheet, ByVal cfgKey As String)
    Dim splitVal As String, lastRow As Long, lastCol As Long
    Dim rules As Collection, r As Long, c As Long
    Dim seg As Variant, part As String, pos As Long, colsStr As String, delim As String
    Dim colIdx As Long, i As Long, j As Long, maxParts As Long, nParts As Long
    Dim arr() As String, destCol As Long, headStr As String
    Dim k As Long

    splitVal = Trim(CfgVal(cfgKey, "分列"))
    If splitVal = "" Or splitVal = "0" Then Exit Sub
    On Error GoTo SplitErr
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Sub
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    If lastCol < 1 Then Exit Sub

    Set rules = New Collection
    part = splitVal
    Do While Len(part) > 0
        pos = InStr(part, ";")
        If pos > 0 Then
            seg = Trim(Left(part, pos - 1))
            part = Mid(part, pos + 1)
        Else
            seg = Trim(part)
            part = ""
        End If
        If Len(CStr(seg)) = 0 Then GoTo NextSeg
        pos = InStr(CStr(seg), ":")
        If pos <= 0 Then GoTo NextSeg
        colsStr = Trim(Left(seg, pos - 1))
        delim = Mid(seg, pos + 1)
        If Trim(delim) = "" Then delim = " "
        If Len(colsStr) = 0 Then GoTo NextSeg
        Dim colList As Collection
        Set colList = New Collection
        Dim colLetter As String, segs() As String, u As Long
        segs = Split(colsStr, ",")
        For u = LBound(segs) To UBound(segs)
            colLetter = Trim(segs(u))
            If Len(colLetter) > 0 Then
                colIdx = Col2Num(colLetter)
                If colIdx > 0 Then
                    On Error Resume Next
                    colList.Add colIdx, CStr(colIdx)
                    On Error GoTo 0
                End If
            End If
        Next u
        If colList.count > 0 Then
            For k = 1 To colList.count
                rules.Add Array(colList(k), delim)
            Next k
        End If
NextSeg:
    Loop

    If rules.count = 0 Then Exit Sub

    Dim idxList() As Long, delimList() As String
    ReDim idxList(1 To rules.count): ReDim delimList(1 To rules.count)
    For i = 1 To rules.count
        idxList(i) = rules(i)(0): delimList(i) = rules(i)(1)
    Next i
    For i = 1 To UBound(idxList) - 1
        For j = i + 1 To UBound(idxList)
            If idxList(i) < idxList(j) Then
                colIdx = idxList(i): idxList(i) = idxList(j): idxList(j) = colIdx
                part = delimList(i): delimList(i) = delimList(j): delimList(j) = part
            End If
        Next j
    Next i

    For i = 1 To UBound(idxList)
        colIdx = idxList(i)
        If colIdx > lastCol Then GoTo nextCol
        maxParts = 1
        For r = 2 To lastRow
            arr = Split(SafeCellStr(ws.Cells(r, colIdx)), delimList(i), -1, vbBinaryCompare)
            nParts = UBound(arr) - LBound(arr) + 1
            If nParts > maxParts Then maxParts = nParts
        Next r
        If maxParts <= 1 Then GoTo nextCol
        ws.Columns(colIdx + 1).Resize(, maxParts - 1).Insert Shift:=xlToRight
        lastCol = lastCol + maxParts - 1
        headStr = SafeCellStr(ws.Cells(1, colIdx))
        For j = 1 To maxParts - 1
            ws.Cells(1, colIdx + j).value = headStr
        Next j
        For r = 2 To lastRow
            arr = Split(SafeCellStr(ws.Cells(r, colIdx)), delimList(i), -1, vbBinaryCompare)
            For j = LBound(arr) To UBound(arr)
                destCol = colIdx + (j - LBound(arr))
                If destCol <= lastCol Then ws.Cells(r, destCol).value = Trim(arr(j))
            Next j
        Next r
nextCol:
    Next i
    Exit Sub
SplitErr:
End Sub

Private Function SafeCellStr(ByVal c As Range) As String
    If c Is Nothing Then SafeCellStr = "": Exit Function
    On Error Resume Next
    If IsNull(c.value) Then SafeCellStr = "": Exit Function
    If IsEmpty(c.value) Then SafeCellStr = "": Exit Function
    SafeCellStr = CStr(c.value)
    On Error GoTo 0
End Function

Private Function Col2Num(ByVal s As String) As Long
    Dim i As Long, n As Long: n = 0
    For i = 1 To Len(s): n = n * 26 + Asc(UCase(Mid(s, i, 1))) - 64: Next
    Col2Num = n
End Function

Private Function Num2Col(ByVal n As Long) As String
    Dim s As String: s = ""
    Do While n > 0: s = Chr(65 + (n - 1) Mod 26) & s: n = (n - 1) \ 26: Loop
    Num2Col = s
End Function

Private Sub SplitAddr(ByVal a As String, ByRef cs As String, ByRef rn As Long)
    Dim i As Long: cs = "": rn = 0: a = Replace(a, "$", "")
    For i = 1 To Len(a)
        If Mid(a, i, 1) Like "[A-Za-z]" Then
            cs = cs & UCase(Mid(a, i, 1))
        Else
            rn = CLng(Mid(a, i)): Exit For
        End If
    Next i
End Sub

Private Function GCV(ByVal c As Range) As Variant
    If c Is Nothing Then GCV = Empty: Exit Function
    On Error Resume Next
    If c.MergeCells Then GCV = c.MergeArea.Cells(1, 1).value Else GCV = c.value
    On Error GoTo 0
End Function

Private Sub ExtractComments(ByVal ws As Worksheet, _
                            ByRef atc As Object, ByRef atv As Object)
    Set atc = CreateObject("Scripting.Dictionary")
    Set atv = CreateObject("Scripting.Dictionary")
    Dim cm As Comment, addr As String
    For Each cm In ws.Comments
        addr = cm.Parent.Address(False, False)
        atc(addr) = cm.Text
        atv(addr) = GCV(cm.Parent)
    Next cm
End Sub

Private Function ExtractRegions(ByRef atc As Object, _
                                 ByVal kw As String) As Collection
    Set ExtractRegions = New Collection

    Dim nums As Object: Set nums = CreateObject("Scripting.Dictionary")
    Dim k As Variant, txt As String, p As Long
    Dim sfx As String, ds As String, ci As Long

    For Each k In atc.keys
        txt = CStr(atc(k)): p = InStr(txt, kw)
        If p > 0 Then
            sfx = Mid(txt, p + Len(kw))
            If Left(sfx, 1) = "#" Then sfx = Mid(sfx, 2)
            ds = ""
            For ci = 1 To Len(sfx)
                If Mid(sfx, ci, 1) Like "[0-9]" Then ds = ds & Mid(sfx, ci, 1) Else Exit For
            Next ci
            If ds <> "" Then nums(CLng(ds)) = True
        End If
    Next k

    If nums.count = 0 Then Exit Function

    Dim numArr() As Long, idx As Long
    ReDim numArr(1 To nums.count): idx = 1
    For Each k In nums.keys: numArr(idx) = CLng(k): idx = idx + 1: Next
    Dim ii As Long, jj As Long, tmp As Long
    For ii = 1 To UBound(numArr) - 1
        For jj = ii + 1 To UBound(numArr)
            If numArr(ii) > numArr(jj) Then tmp = numArr(ii): numArr(ii) = numArr(jj): numArr(jj) = tmp
        Next jj
    Next ii

    Dim sa As String, ea As String
    Dim sc As String, sr As Long, ec As String, er As Long
    Dim scn As Long, ecn As Long
    For ii = 1 To UBound(numArr)
        sa = "": ea = ""
        For Each k In atc.keys
            txt = CStr(atc(k))
            If InStr(txt, kw & CStr(numArr(ii))) > 0 And _
               InStr(txt, kw & "#" & CStr(numArr(ii))) = 0 Then sa = CStr(k)
            If InStr(txt, kw & "#" & CStr(numArr(ii))) > 0 Then ea = CStr(k)
        Next k
        If sa <> "" And ea <> "" Then
            SplitAddr sa, sc, sr: SplitAddr ea, ec, er
            If sr > er Then tmp = sr: sr = er: er = tmp
            scn = Col2Num(sc): ecn = Col2Num(ec)
            If scn > ecn Then tmp = scn: scn = ecn: ecn = tmp
            ExtractRegions.Add Array(sr, er, scn, ecn)
        End If
    Next ii
End Function

Private Sub ExtractSetInfo(ByRef atc As Object, _
                           ByRef names As Collection, ByRef addrs As Collection)
    Set names = New Collection: Set addrs = New Collection
    Dim k As Variant, txt As String, nm As String
    Dim p1 As Long, p2 As Long
    For Each k In atc.keys
        txt = CStr(atc(k))
        nm = ""
        p1 = InStr(txt, "set(")
        If p1 > 0 Then
            p2 = InStr(p1, txt, ")")
            If p2 > p1 Then nm = Mid(txt, p1 + 4, p2 - p1 - 4)
        End If
        If nm = "" Then
            p1 = InStr(txt, "set" & ChrW(&HFF08))
            If p1 > 0 Then
                p2 = InStr(p1, txt, ChrW(&HFF09))
                If p2 > p1 Then nm = Mid(txt, p1 + 4, p2 - p1 - 4)
            End If
        End If
        If nm <> "" Then names.Add nm: addrs.Add CStr(k)
    Next k
End Sub

Private Sub WriteHeaders(ByVal destWs As Worksheet, _
                         ByVal tmplWs As Worksheet, _
                         ByRef atc As Object, _
                         ByVal bWb As Boolean, ByVal bWs As Boolean, _
                         ByVal bSet As Boolean, ByVal bCol As Boolean, _
                         ByVal bRowN As Boolean)
    Dim col As Long: col = 1
    Dim si As Long, ri As Long, c As Long
    Dim rg As Variant, v1 As Variant, v2 As Variant, hv As Variant
    Dim sn As Collection, sa As Collection
    Dim cr As Collection

    If bWb Then destWs.Cells(1, col).value = "工作簿": col = col + 1
    If bWs Then destWs.Cells(1, col).value = "工作表": col = col + 1

    If bSet Then
        ExtractSetInfo atc, sn, sa
        For si = 1 To sn.count
            destWs.Cells(1, col).value = sn(si): col = col + 1
        Next si
    End If

    If bCol Then
        Set cr = ExtractRegions(atc, "列区域")
        For ri = 1 To cr.count
            rg = cr(ri)
            If rg(0) < rg(1) Then
                For c = rg(2) To rg(3)
                    v1 = GCV(tmplWs.Cells(rg(0), c))
                    v2 = GCV(tmplWs.Cells(rg(1), c))
                    If CStr(v1 & "") <> CStr(v2 & "") And Not IsEmpty(v2) Then
                        destWs.Cells(1, col).value = CStr(v1) & "_" & CStr(v2)
                    ElseIf Not IsEmpty(v1) Then
                        destWs.Cells(1, col).value = CStr(v1)
                    Else
                        destWs.Cells(1, col).value = Num2Col(c)
                    End If
                    col = col + 1
                Next c
            Else
                For c = rg(2) To rg(3)
                    hv = GCV(tmplWs.Cells(rg(0), c))
                    If Not IsEmpty(hv) Then
                        destWs.Cells(1, col).value = CStr(hv)
                    Else
                        destWs.Cells(1, col).value = Num2Col(c)
                    End If
                    col = col + 1
                Next c
            End If
        Next ri
    End If

    If bRowN Then destWs.Cells(1, col).value = "行号"
    destWs.Rows(1).Font.Bold = True
End Sub

Private Sub WriteDataRows(ByVal destWs As Worksheet, ByRef destRow As Long, _
                          ByVal srcWs As Worksheet, ByVal wbName As String, _
                          ByRef atc As Object, _
                          ByVal bWb As Boolean, ByVal bWs As Boolean, _
                          ByVal bSet As Boolean, ByVal bCol As Boolean, _
                          ByVal bRowN As Boolean)

    Dim rowRegs As Collection: Set rowRegs = ExtractRegions(atc, "行区域")
    Dim colRegs As Collection: Set colRegs = ExtractRegions(atc, "列区域")
    Dim sn As Collection, sa As Collection
    ExtractSetInfo atc, sn, sa

    If rowRegs.count = 0 Then Exit Sub

    Dim rri As Long, rrg As Variant, dataRow As Long, col As Long
    Dim sai As Long, sColStr As String, sRowNum As Long
    Dim cri As Long, crg As Variant, dc As Long

    For rri = 1 To rowRegs.count
        rrg = rowRegs(rri)
        For dataRow = rrg(0) To rrg(1)
            col = 1
            If bWb Then destWs.Cells(destRow, col).value = wbName: col = col + 1
            If bWs Then destWs.Cells(destRow, col).value = srcWs.Name: col = col + 1

            If bSet Then
                For sai = 1 To sa.count
                    SplitAddr CStr(sa(sai)), sColStr, sRowNum
                    destWs.Cells(destRow, col).value = GCV(srcWs.Cells(sRowNum, Col2Num(sColStr)))
                    col = col + 1
                Next sai
            End If

            If bCol Then
                For cri = 1 To colRegs.count
                    crg = colRegs(cri)
                    For dc = crg(2) To crg(3)
                        destWs.Cells(destRow, col).value = GCV(srcWs.Cells(dataRow, dc))
                        col = col + 1
                    Next dc
                Next cri
            End If

            If bRowN Then destWs.Cells(destRow, col).value = dataRow
            destRow = destRow + 1
        Next dataRow
    Next rri
End Sub

' 判断当前表名是否参与汇总（参与/不参与两栏都配时以参与为准）
Private Function SheetParticipate(ByVal sheetName As String, ByRef excludeList As Variant, ByRef includeList As Variant) As Boolean
    Dim i As Long, kw As String
    Dim useWhitelist As Boolean
    useWhitelist = False
    For i = LBound(includeList) To UBound(includeList)
        kw = Trim(CStr(includeList(i)))
        If kw <> "" Then useWhitelist = True: Exit For
    Next i
    If useWhitelist Then
        SheetParticipate = False
        For i = LBound(includeList) To UBound(includeList)
            kw = Trim(CStr(includeList(i)))
            If kw = "" Then GoTo NextInc
            If InStr(1, kw, sheetName, vbTextCompare) > 0 Or sheetName Like "*" & kw & "*" Then
                SheetParticipate = True: Exit Function
            End If
NextInc:
        Next i
    Else
        SheetParticipate = True
        For i = LBound(excludeList) To UBound(excludeList)
            kw = Trim(CStr(excludeList(i)))
            If kw = "" Then GoTo NextExc
            If InStr(1, kw, sheetName, vbTextCompare) > 0 Or sheetName Like "*" & kw & "*" Then
                SheetParticipate = False: Exit Function
            End If
NextExc:
        Next i
    End If
End Function

' ============================================================
'  入口：汇总（从执行面板读取）
' ============================================================
Public Sub SummarizeBySheetThenSplit()
    DoSummarizeSheetFromPanel
End Sub

' 主流程：从执行面板取模板与源文件，按参与/不参与过滤 Sheet 后汇总
Private Sub DoSummarizeSheetFromPanel()
    Dim wsPanel As Worksheet
    Dim tmplPath As String
    Dim filePaths As New Collection
    Dim r As Long
    Dim tmplWb As Workbook
    Dim newWb As Workbook
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim cfgKey As String
    Dim splitCfgKey As String
    Dim t0 As Double
    Dim logKey As String
    Dim bWb As Boolean, bWs As Boolean, bSet As Boolean, bCol As Boolean, bRowN As Boolean, bSkip As Boolean
    Dim sheetAtc As Object
    Dim atcTmp As Object, atvTmp As Object
    Dim sheetDestRow As Object, sheetDestWs As Object
    Dim firstSheet As Boolean, sKey As Variant
    Dim destWsR As Worksheet
    Dim curAtc As Object
    Dim curTmplWs As Worksheet
    Dim matchKey As String
    Dim dr As Long
    Dim wbName As String
    Dim totalFiles As Long, fi As Long
    Dim arrExclude As Variant, arrInclude As Variant
    Dim fd As FileDialog

    logKey = "2.2.2 按批注汇总"
    cfgKey = CFG_KEY_SPLIT_SHEET
    splitCfgKey = CFG_KEY_SPLIT_SHEET
    t0 = Timer
    RunLog_WriteRow logKey, "开始", "", "", "", "", "开始", ""

    On Error Resume Next
    Set wsPanel = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    On Error GoTo 0
    If wsPanel Is Nothing Then
        MsgBox "请先运行「4. VBA 同步 → 3.3 初始化配置」创建执行面板，并在执行面板中填写模板路径与源文件列表。", vbExclamation
        RunLog_WriteRow logKey, "取消", "", "", "", "", "无执行面板", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    tmplPath = Trim$(CStr(wsPanel.Cells(2, 1).value))  ' A2：模板文件路径
    If tmplPath = "" Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = "选择模板文件（执行面板未填模板路径）"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Excel 文件", "*.xls;*.xlsx;*.xlsm"
        End With
        If fd.Show <> True Then
            RunLog_WriteRow logKey, "取消", "", "", "", "", "未选模板", CStr(Round(Timer - t0, 2))
            Exit Sub
        End If
        tmplPath = fd.SelectedItems(1)
        wsPanel.Cells(2, 1).value = tmplPath
        wsPanel.Hyperlinks.Add Anchor:=wsPanel.Cells(2, 1), Address:=CStr(tmplPath), TextToDisplay:=CStr(tmplPath)
    End If

    For r = PANEL_DATA_START_ROW To wsPanel.Cells(wsPanel.Rows.count, PANEL_COL_PATH).End(xlUp).row
        If Trim$(CStr(wsPanel.Cells(r, PANEL_COL_PATH).value)) <> "" Then
            filePaths.Add Trim$(CStr(wsPanel.Cells(r, PANEL_COL_PATH).value))
        End If
    Next r
    If filePaths.count = 0 Then
        MsgBox "执行面板中无源文件列表，请用「1. 文件选择」添加。", vbExclamation
        RunLog_WriteRow logKey, "取消", "", "", "", "", "无源文件", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    arrExclude = Split(CfgVal(cfgKey, "不参与的工作表"), ";")
    arrInclude = Split(CfgVal(cfgKey, "参与的工作表"), ";")

    bWb = CfgBool(cfgKey, "工作簿", True)
    bWs = CfgBool(cfgKey, "工作表", True)
    bSet = CfgBool(cfgKey, "set区", True)
    bCol = CfgBool(cfgKey, "列区域", True)
    bRowN = CfgBool(cfgKey, "行号", False)
    bSkip = CfgBool(cfgKey, "跳过未匹配表", False)
    Dim bForceTemplate As Boolean
    bForceTemplate = CfgBool(cfgKey, "强制按模板", False)

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrPanel

    Set tmplWb = Workbooks.Open(tmplPath, ReadOnly:=True)
    Set sheetAtc = CreateObject("Scripting.Dictionary")
    sheetAtc.CompareMode = vbTextCompare
    For Each srcWs In tmplWb.Worksheets
        Set atcTmp = CreateObject("Scripting.Dictionary")
        Set atvTmp = CreateObject("Scripting.Dictionary")
        ExtractComments srcWs, atcTmp, atvTmp
        If atcTmp.count > 0 Then sheetAtc.Add srcWs.Name, atcTmp
    Next srcWs
    If sheetAtc.count = 0 Then
        MsgBox "模板无批注", vbExclamation
        SafeCloseWorkbookLocal tmplWb, False
        GoTo CleanupPanel
    End If

    ' 强制按模板：只保留「模板」表，单表 ALL、不参与过滤
    If bForceTemplate Then
        If Not sheetAtc.Exists(TMPL_SHT) Then
            MsgBox "模板工作簿中未找到工作表「" & TMPL_SHT & "」。强制按模板时必须有该表。", vbExclamation
            SafeCloseWorkbookLocal tmplWb, False
            GoTo CleanupPanel
        End If
        Set atcTmp = sheetAtc(TMPL_SHT)
        sheetAtc.RemoveAll
        sheetAtc.Add TMPL_SHT, atcTmp
    End If

    Set newWb = Workbooks.Add
    Do While newWb.Worksheets.count > 1
        newWb.Worksheets(newWb.Worksheets.count).Delete
    Loop

    Set sheetDestRow = CreateObject("Scripting.Dictionary")
    sheetDestRow.CompareMode = vbTextCompare
    Set sheetDestWs = CreateObject("Scripting.Dictionary")
    sheetDestWs.CompareMode = vbTextCompare
    firstSheet = True
    For Each sKey In sheetAtc.keys
        If firstSheet Then
            Set destWsR = newWb.Worksheets(1)
            firstSheet = False
        Else
            Set destWsR = newWb.Worksheets.Add(After:=newWb.Worksheets(newWb.Worksheets.count))
        End If
        If bForceTemplate Then destWsR.Name = "ALL" Else destWsR.Name = CStr(sKey)
        Set curAtc = sheetAtc(CStr(sKey))
        Set curTmplWs = tmplWb.Worksheets(CStr(sKey))
        WriteHeaders destWsR, curTmplWs, curAtc, bWb, bWs, bSet, bCol, bRowN
        sheetDestRow(CStr(sKey)) = 2
        sheetDestWs.Add CStr(sKey), destWsR
    Next sKey

    totalFiles = filePaths.count
    For fi = 1 To totalFiles
        Application.StatusBar = "处理 " & fi & "/" & totalFiles & " ..."
        Set srcWb = Workbooks.Open(CStr(filePaths(fi)), ReadOnly:=True)
        wbName = srcWb.Name
        If InStrRev(wbName, ".") > 0 Then wbName = Left(wbName, InStrRev(wbName, ".") - 1)
        For Each srcWs In srcWb.Worksheets
            If Not bForceTemplate And Not SheetParticipate(srcWs.Name, arrExclude, arrInclude) Then GoTo NextSrcWs

            If bForceTemplate And sheetAtc.Exists(TMPL_SHT) Then
                matchKey = TMPL_SHT
            ElseIf sheetAtc.Exists(srcWs.Name) Then
                matchKey = srcWs.Name
            ElseIf bSkip Then
                GoTo NextSrcWs
            ElseIf sheetAtc.Exists(TMPL_SHT) Then
                matchKey = TMPL_SHT
            Else
                GoTo NextSrcWs
            End If
            If Not sheetDestWs.Exists(matchKey) Then GoTo NextSrcWs
            Set curAtc = sheetAtc(matchKey)
            Set destWsR = sheetDestWs(matchKey)
            dr = sheetDestRow(matchKey)
            WriteDataRows destWsR, dr, srcWs, wbName, curAtc, bWb, bWs, bSet, bCol, bRowN
            sheetDestRow(matchKey) = dr
NextSrcWs:
        Next srcWs
        RunLog_WriteRow logKey, "汇总", srcWb.Name, "", "", "成功", "成功", ""
        SafeCloseWorkbookLocal srcWb, False
        Set srcWb = Nothing
    Next fi

    RunLog_WriteRow logKey, "完成", "", "", "", "", "Done", CStr(Round(Timer - t0, 2))
    SafeCloseWorkbookLocal tmplWb, False
    Set tmplWb = Nothing

    For Each srcWs In newWb.Worksheets
        srcWs.Columns.AutoFit
    Next srcWs
    For Each srcWs In newWb.Worksheets
        ApplySplitColumns srcWs, splitCfgKey
    Next srcWs

    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    Dim savePath As Variant
    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="汇总_" & Format(Now, "yyyy-mm-dd_hh-nn-ss"), _
        FileFilter:="Excel 工作簿 (*.xlsx), *.xlsx", _
        Title:="另存为")
    If VarType(savePath) = vbBoolean Then
        If savePath = False Then MsgBox "未保存", vbInformation
    Else
        newWb.SaveAs fileName:=CStr(savePath), FileFormat:=xlOpenXMLWorkbook
        Dim totalRows As Long: totalRows = 0
        For Each srcWs In newWb.Worksheets
            totalRows = totalRows + srcWs.Cells(srcWs.Rows.count, 1).End(xlUp).row - 1
        Next srcWs
        MsgBox "已汇总 " & totalRows & " 行" & vbCrLf & "已保存至" & vbCrLf & savePath, vbInformation
    End If
    Exit Sub

ErrPanel:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error Resume Next
    SafeCloseWorkbookLocal srcWb, False
    SafeCloseWorkbookLocal tmplWb, False
    On Error GoTo 0
    RunLog_WriteRow logKey, "失败", "", "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
    Exit Sub

CleanupPanel:
    Application.StatusBar = False
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    RunLog_WriteRow logKey, "取消", "", "", "", "", "取消", CStr(Round(Timer - t0, 2))
End Sub

Private Sub SafeCloseWorkbookLocal(ByRef wb As Workbook, Optional ByVal saveChanges As Boolean = False)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=saveChanges
        Set wb = Nothing
    End If
    On Error GoTo 0
End Sub
