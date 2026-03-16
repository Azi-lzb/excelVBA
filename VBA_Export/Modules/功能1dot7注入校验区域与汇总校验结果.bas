Attribute VB_Name = "FnInjectCheckResult"
Option Explicit

' 1.7 ??????????????/?/????????/????????????
' 1.8 ??????????????????? ??/??/??/??? ????????????
' 1.9 ??????????????????,?????????????????????????

Private Const PANEL_SHEET_NAME As String = "????"
Private Const PANEL_DATA_START_ROW As Long = 5
Private Const PANEL_COL_PATH As Long = 2

Private Const RESULT_SHEET_NAME As String = "????"

Private Const TMPL_KEY_LOCAL As String = "??????"
Private Const TMPL_KEY_CROSS As String = "??????"
Private Const TMPL_SHEET_NAME As String = "??"

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

' ---------- ???? ----------

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
    Dim lastRow As Long, r As Long
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
    Dim i As Long, n As Long
    n = 0
    For i = 1 To Len(s)
        n = n * 26 + Asc(UCase$(Mid$(s, i, 1))) - 64
    Next i
    Col2Num = n
End Function

Private Sub SplitAddr(ByVal a As String, ByRef cs As String, ByRef rn As Long)
    Dim i As Long
    cs = "": rn = 0
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
        If CellCommentText = "" Then
            CellCommentText = c.Comment.Comment.Text
        End If
    End If
    On Error GoTo 0
End Function

Private Sub ExtractComments(ByVal ws As Worksheet, ByRef atc As Object)
    Dim cm As Comment, addr As String
    Set atc = CreateObject("Scripting.Dictionary")
    atc.CompareMode = vbTextCompare
    If ws Is Nothing Then Exit Sub
    For Each cm In ws.Comments
        addr = cm.Parent.Address(False, False)
        atc(addr) = cm.Text
    Next cm
End Sub

' ?????? kw ?????????? Collection???? Array(sr, er, scn, ecn)
Private Function ExtractRegions(ByRef atc As Object, ByVal kw As String) As Collection
    Dim nums As Object
    Dim k As Variant, txt As String
    Dim p As Long, ci As Long
    Dim sfx As String, ds As String
    Dim numArr() As Long
    Dim idx As Long, ii As Long, jj As Long, tmp As Long
    Dim sa As String, ea As String
    Dim sc As String, sr As Long, ec As String, er As Long
    Dim scn As Long, ecn As Long

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
        sa = "": ea = ""
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
    ws.Cells(1, rcTmplWb).Value = "??????"
    ws.Cells(1, rcTmplWs).Value = "??????"
    ws.Cells(1, rcSrcWb).Value = "???????"
    ws.Cells(1, rcSrcWs).Value = "???????"
    ws.Cells(1, rcRow).Value = "???"
    ws.Cells(1, rcCol).Value = "???"
    ws.Cells(1, rcRowName).Value = "???"
    ws.Cells(1, rcErrType).Value = "????"
    ws.Cells(1, rcRsv1).Value = "????1"
    ws.Cells(1, rcRsv2).Value = "????2"
    ws.Cells(1, rcRsv3).Value = "????3"
    ws.Cells(1, rcRsv4).Value = "????4"
    ws.Cells(1, rcComment).Value = "??????????"
    ws.Rows(1).Font.Bold = True
End Sub

Private Function ContainsAnyKey(ByVal s As String) As Boolean
    Dim t As String
    t = CStr(s)
    If t = "" Then
        ContainsAnyKey = False
    Else
        ContainsAnyKey = (InStr(1, t, "?", vbTextCompare) > 0 Or _
                          InStr(1, t, "??", vbTextCompare) > 0 Or _
                          InStr(1, t, "??", vbTextCompare) > 0 Or _
                          InStr(1, t, "??", vbTextCompare) > 0)
    End If
End Function

Private Sub ParsePipe5(ByVal s As String, ByRef errType As String, ByRef r1 As String, ByRef r2 As String, ByRef r3 As String, ByRef r4 As String)
    Dim arr As Variant
    errType = "": r1 = "": r2 = "": r3 = "": r4 = ""
    If s = "" Then Exit Sub
    arr = Split(CStr(s), "|")
    If UBound(arr) >= 0 Then errType = CStr(arr(0))
    If UBound(arr) >= 1 Then r1 = CStr(arr(1))
    If UBound(arr) >= 2 Then r2 = CStr(arr(2))
    If UBound(arr) >= 3 Then r3 = CStr(arr(3))
    If UBound(arr) >= 4 Then r4 = CStr(arr(4))
End Sub

' ---------- 1.7 ?????? ----------

Public Sub ??????()
    Dim t0 As Double
    Dim logKey As String
    Dim wsPanel As Worksheet
    Dim tmplPath As String
    Dim srcPaths As Collection
    Dim tmplWb As Workbook, srcWb As Workbook
    Dim tmplWs As Worksheet, srcWs As Worksheet
    Dim atc As Object
    Dim regsLocal As Collection, regsCross As Collection, regs As Collection
    Dim rg As Variant
    Dim r As Long, c As Long, fi As Long
    Dim okCnt As Long, failCnt As Long

    t0 = Timer
    logKey = "1.7 ??????"
    On Error Resume Next
    RunLog_WriteRow logKey, "??", "", "", "", "", "??", ""
    On Error GoTo 0

    Set wsPanel = GetPanelWs()
    If wsPanel Is Nothing Then
        MsgBox "??????????????????", vbExclamation
        RunLog_WriteRow logKey, "??", "", "", "", "", "?????", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    tmplPath = GetTemplatePathFromPanel(wsPanel)
    If tmplPath = "" Then
        MsgBox "???? A2 ??????????", vbExclamation
        RunLog_WriteRow logKey, "??", "", "", "", "", "?????", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    CollectSourcePathsFromPanel wsPanel, srcPaths
    If srcPaths Is Nothing Or srcPaths.Count = 0 Then
        MsgBox "????????????B5 ???", vbExclamation
        RunLog_WriteRow logKey, "??", "", "", "", "", "????", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrInject

    Set tmplWb = Workbooks.Open(tmplPath, ReadOnly:=True, UpdateLinks:=0)

    For fi = 1 To srcPaths.Count
        On Error Resume Next
        Set srcWb = Workbooks.Open(CStr(srcPaths(fi)), ReadOnly:=False, UpdateLinks:=0)
        If Err.Number <> 0 Or srcWb Is Nothing Then
            failCnt = failCnt + 1
            RunLog_WriteRow logKey, "??", CStr(srcPaths(fi)), "", "", "??", Err.Number & " " & Err.Description, ""
            Err.Clear
            On Error GoTo 0
            GoTo NextFile
        End If
        On Error GoTo 0

        For Each srcWs In srcWb.Worksheets
            On Error Resume Next
            Set tmplWs = Nothing
            Set tmplWs = tmplWb.Worksheets(srcWs.Name)
            If tmplWs Is Nothing Then
                Set tmplWs = tmplWb.Worksheets(TMPL_SHEET_NAME)
            End If
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
                        Dim cmTxt As String
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
        RunLog_WriteRow logKey, "??", srcWb.Name, "", "", "??", "??", ""
        srcWb.Close SaveChanges:=False
NextFile:
        Set srcWb = Nothing
    Next fi

    tmplWb.Close SaveChanges:=False
    Set tmplWb = Nothing

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    RunLog_WriteRow logKey, "??", "", "", "", "", "?? " & okCnt & "??? " & failCnt, CStr(Round(Timer - t0, 2))
    MsgBox "?????" & vbCrLf & "???" & okCnt & vbCrLf & "???" & failCnt, vbInformation
    Exit Sub

ErrInject:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If Not tmplWb Is Nothing Then tmplWb.Close SaveChanges:=False
    On Error GoTo 0
    RunLog_WriteRow logKey, "??", "", "", "", "??", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "?? " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' ---------- 2. ????????? ? ?? ? ??? ----------

Public Sub ??????()
    ' ?????????????????????????????????
    ??????
    ??????
    ??????
End Sub

' ---------- 1.8 ?????? ----------

Public Sub ??????()
    Dim t0 As Double
    Dim logKey As String
    Dim wsPanel As Worksheet
    Dim tmplPath As String
    Dim srcPaths As Collection
    Dim wsRes As Worksheet
    Dim tmplWb As Workbook, srcWb As Workbook
    Dim tmplWs As Worksheet, srcWs As Worksheet
    Dim atc As Object
    Dim regsLocal As Collection, regsCross As Collection, regs As Collection
    Dim rg As Variant
    Dim r As Long, c As Long, fi As Long
    Dim cellVal As String
    Dim errType As String, r1 As String, r2 As String, r3 As String, r4 As String
    Dim cmTxt As String
    Dim rowName As String
    Dim outRow As Long
    Dim hitCnt As Long

    t0 = Timer
    logKey = "1.8 ??????"
    On Error Resume Next
    RunLog_WriteRow logKey, "??", "", "", "", "", "??", ""
    On Error GoTo 0

    Set wsPanel = GetPanelWs()
    If wsPanel Is Nothing Then
        MsgBox "??????????????????", vbExclamation
        RunLog_WriteRow logKey, "??", "", "", "", "", "?????", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    tmplPath = GetTemplatePathFromPanel(wsPanel)
    If tmplPath = "" Then
        MsgBox "???? A2 ??????????", vbExclamation
        RunLog_WriteRow logKey, "??", "", "", "", "", "?????", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    CollectSourcePathsFromPanel wsPanel, srcPaths
    If srcPaths Is Nothing Or srcPaths.Count = 0 Then
        MsgBox "????????????B5 ???", vbExclamation
        RunLog_WriteRow logKey, "??", "", "", "", "", "????", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    Set wsRes = EnsureResultSheet()
    InitResultHeader wsRes
    outRow = 2

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrSum

    Set tmplWb = Workbooks.Open(tmplPath, ReadOnly:=True, UpdateLinks:=0)

    For fi = 1 To srcPaths.Count
        On Error Resume Next
        Set srcWb = Workbooks.Open(CStr(srcPaths(fi)), ReadOnly:=True, UpdateLinks:=0)
        If Err.Number <> 0 Or srcWb Is Nothing Then
            RunLog_WriteRow logKey, "??", CStr(srcPaths(fi)), "", "", "??", Err.Number & " " & Err.Description, ""
            Err.Clear
            On Error GoTo 0
            GoTo NextFile2
        End If
        On Error GoTo 0

        For Each srcWs In srcWb.Worksheets
            On Error Resume Next
            Set tmplWs = Nothing
            Set tmplWs = tmplWb.Worksheets(srcWs.Name)
            If tmplWs Is Nothing Then
                Set tmplWs = tmplWb.Worksheets(TMPL_SHEET_NAME)
            End If
            On Error GoTo 0
            If tmplWs Is Nothing Then GoTo NextSheet2

            ExtractComments tmplWs, atc
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
                            wsRes.Cells(outRow, rcTmplWb).Value = tmplWb.Name
                            wsRes.Cells(outRow, rcTmplWs).Value = tmplWs.Name
                            wsRes.Cells(outRow, rcSrcWb).Value = srcWb.Name
                            wsRes.Cells(outRow, rcSrcWs).Value = srcWs.Name
                            wsRes.Cells(outRow, rcRow).Value = r
                            wsRes.Cells(outRow, rcCol).Value = c
                            wsRes.Cells(outRow, rcRowName).Value = rowName
                            wsRes.Cells(outRow, rcErrType).Value = errType
                            wsRes.Cells(outRow, rcRsv1).Value = r1
                            wsRes.Cells(outRow, rcRsv2).Value = r2
                            wsRes.Cells(outRow, rcRsv3).Value = r3
                            wsRes.Cells(outRow, rcRsv4).Value = r4
                            wsRes.Cells(outRow, rcComment).Value = cmTxt
                            outRow = outRow + 1
                            hitCnt = hitCnt + 1
                        End If
                    Next c
                Next r
            Next rg
NextSheet2:
        Next srcWs

        RunLog_WriteRow logKey, "??", srcWb.Name, "", "", "??", "?? " & hitCnt, ""
        srcWb.Close SaveChanges:=False
NextFile2:
        Set srcWb = Nothing
    Next fi

    tmplWb.Close SaveChanges:=False
    Set tmplWb = Nothing

    wsRes.Columns.AutoFit
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    RunLog_WriteRow logKey, "??", "", "", "", "", "?? " & hitCnt, CStr(Round(Timer - t0, 2))
    MsgBox "?????" & vbCrLf & "???" & hitCnt & vbCrLf & "?????????" & RESULT_SHEET_NAME & "??", vbInformation
    Exit Sub

ErrSum:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If Not tmplWb Is Nothing Then tmplWb.Close SaveChanges:=False
    On Error GoTo 0
    RunLog_WriteRow logKey, "??", "", "", "", "??", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "?? " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' ---------- 1.9 ?????? ----------

Public Sub ??????()
    Dim t0 As Double
    Dim logKey As String
    Dim wsPanel As Worksheet
    Dim srcPaths As Collection
    Dim srcWb As Workbook
    Dim srcWs As Worksheet
    Dim atc As Object
    Dim regs As Collection
    Dim rg As Variant
    Dim r As Long, c As Long, fi As Long
    Dim keyInput As String
    Dim keyArr() As String
    Dim iKey As Long

    t0 = Timer
    logKey = "1.9 ??????"
    On Error Resume Next
    RunLog_WriteRow logKey, "??", "", "", "", "", "??", ""
    On Error GoTo 0

    Set wsPanel = GetPanelWs()
    If wsPanel Is Nothing Then
        MsgBox "??????????????????", vbExclamation
        RunLog_WriteRow logKey, "??", "", "", "", "", "?????", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    CollectSourcePathsFromPanel wsPanel, srcPaths
    If srcPaths Is Nothing Or srcPaths.Count = 0 Then
        MsgBox "????????????B5 ???", vbExclamation
        RunLog_WriteRow logKey, "??", "", "", "", "", "????", CStr(Round(Timer - t0, 2))
        Exit Sub
    End If

    keyInput = InputBox("????????????????????? ??????,??????", _
                        "??????", _
                        "??????,??????")
    If Trim$(keyInput) = "" Then Exit Sub
    keyArr = Split(keyInput, ",")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrClear

    For fi = 1 To srcPaths.Count
        On Error Resume Next
        Set srcWb = Workbooks.Open(CStr(srcPaths(fi)), ReadOnly:=False, UpdateLinks:=0)
        If Err.Number <> 0 Or srcWb Is Nothing Then
            RunLog_WriteRow logKey, "??", CStr(srcPaths(fi)), "", "", "??", Err.Number & " " & Err.Description, ""
            Err.Clear
            On Error GoTo 0
            GoTo NextFileClear
        End If
        On Error GoTo 0

        For Each srcWs In srcWb.Worksheets
            ' ????????????????????????
            ExtractComments srcWs, atc
            Set regs = New Collection
            For iKey = LBound(keyArr) To UBound(keyArr)
                Dim oneRegs As Collection
                Set oneRegs = ExtractRegions(atc, Trim$(CStr(keyArr(iKey))))
                If Not oneRegs Is Nothing Then
                    Dim rg2 As Variant
                    For Each rg2 In oneRegs
                        regs.Add rg2
                    Next rg2
                End If
            Next iKey
            If regs.Count = 0 Then GoTo NextSheetClear

            For Each rg In regs
                For r = CLng(rg(0)) To CLng(rg(1))
                    For c = CLng(rg(2)) To CLng(rg(3))
                        srcWs.Cells(r, c).ClearContents
                        On Error Resume Next
                        If Not srcWs.Cells(r, c).Comment Is Nothing Then srcWs.Cells(r, c).Comment.Delete
                        On Error GoTo 0
                    Next c
                Next r
            Next rg
NextSheetClear:
        Next srcWs

        srcWb.Save
        RunLog_WriteRow logKey, "??", srcWb.Name, "", "", "??", "??", ""
        srcWb.Close SaveChanges:=False
NextFileClear:
        Set srcWb = Nothing
    Next fi

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    RunLog_WriteRow logKey, "??", "", "", "", "", "??", CStr(Round(Timer - t0, 2))
    MsgBox "?????", vbInformation
    Exit Sub

ErrClear:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error Resume Next
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    On Error GoTo 0
    RunLog_WriteRow logKey, "??", "", "", "", "??", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "?? " & Err.Number & ": " & Err.Description, vbCritical
End Sub