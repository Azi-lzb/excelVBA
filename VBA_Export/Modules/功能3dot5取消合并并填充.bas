Attribute VB_Name = "功能3dot5取消合并并填充"
Option Explicit

' 选中区域取消合并并填充左上角值：将选中范围内所有合并单元格取消合并，
' 取消合并后每个单元格均填入原合并区域左上角的值（如 A1:C3 合并，取消后 A1、A2、A3、B1..C3 均为原 A1 的值）。
' 运行日志写在本模块内，不依赖 vbaSync。

Private Const RUNLOG_SHEET As String = "运行日志"
Private Const RUNLOG_COL_SEQ As Long = 1
Private Const RUNLOG_COL_TIME As Long = 2
Private Const RUNLOG_COL_USER As Long = 3
Private Const RUNLOG_COL_MODULE As Long = 4
Private Const RUNLOG_COL_OP As Long = 5
Private Const RUNLOG_COL_OBJ As Long = 6
Private Const RUNLOG_COL_BEFORE As Long = 7
Private Const RUNLOG_COL_AFTER As Long = 8
Private Const RUNLOG_COL_RESULT As Long = 9
Private Const RUNLOG_COL_DETAIL As Long = 10
Private Const RUNLOG_COL_ELAPSED As Long = 11
Private Const RUNLOG_COL_PC As Long = 12

''' 对当前选中区域：取消所有合并单元格，并将原合并区域左上角的值填入该区域内每一个单元格。
''' 需先选中工作表上的一个区域，再执行本过程（菜单或 Alt+F8）。
Public Sub 取消合并并填充左上角值()
    Dim sel As Range
    Dim ws As Worksheet
    Dim c As Range
    Dim addr As String
    Dim processed As String
    Dim addrs() As String
    Dim values() As Variant
    Dim n As Long
    Dim i As Long
    Dim rng As Range
    Dim topLeftVal As Variant
    Dim t0 As Double

    On Error GoTo ErrHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "请先在工作表上选中一个单元格区域，再执行本功能。", vbExclamation
        Exit Sub
    End If

    Set sel = Selection
    If sel.Cells.count = 0 Then
        MsgBox "未选中有效区域，请先选中包含合并单元格的区域。", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrTrust
    Set ws = sel.Parent
    On Error GoTo ErrHandler
    If ws Is Nothing Or TypeName(ws) <> "Worksheet" Then
        MsgBox "当前选区不在工作表上，无法处理。", vbExclamation
        Exit Sub
    End If

    t0 = Timer
    On Error Resume Next
    写运行日志 "3.5 取消合并并填充", "开始", sel.Address(False, False), "", "", "", "选中区域 " & sel.Address(False, False), ""
    On Error GoTo ErrHandler

    ' 收集选中范围内所有不同的合并区域及其左上角值（每个 MergeArea 只记一次）
    processed = "|"
    n = 0
    For Each c In sel.Cells
        If c.MergeCells Then
            addr = c.MergeArea.Address(False, False)
            If InStr(processed, "|" & addr & "|") = 0 Then
                processed = processed & addr & "|"
                n = n + 1
                ReDim Preserve addrs(1 To n)
                ReDim Preserve values(1 To n)
                addrs(n) = addr
                values(n) = c.MergeArea.Cells(1, 1).value
            End If
        End If
    Next c

    If n = 0 Then
        MsgBox "选中区域内没有合并单元格。", vbInformation
        On Error Resume Next
        写运行日志 "3.5 取消合并并填充", "完成", sel.Address(False, False), "", "", "无合并单元格", "", CStr(Round(Timer - t0, 2))
        On Error GoTo ErrHandler
        Exit Sub
    End If

    Application.ScreenUpdating = False
    For i = 1 To n
        Set rng = ws.Range(addrs(i))
        topLeftVal = values(i)
        rng.UnMerge
        rng.value = topLeftVal
    Next i
    Application.ScreenUpdating = True

    On Error Resume Next
    写运行日志 "3.5 取消合并并填充", "完成", sel.Address(False, False), "", "", "已处理 " & n & " 个合并区域", "", CStr(Round(Timer - t0, 2))
    On Error GoTo ErrHandler
    MsgBox "已取消 " & n & " 个合并区域，并已将各区域左上角的值填入对应单元格。", vbInformation
    Exit Sub

ErrTrust:
    MsgBox "访问工作表时出错。若为信任设置问题，请在「文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置」中勾选「信任对 VBA 工程对象模型的访问」。", vbExclamation
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    MsgBox "执行出错：" & vbCrLf & Err.Description, vbCritical
End Sub

Private Sub 写运行日志(ByVal 功能模块 As String, ByVal 操作 As String, ByVal 记录ID对象 As String, _
    ByVal 操作前值 As String, ByVal 操作后值 As String, ByVal 结果 As String, ByVal 详细信息 As String, ByVal 耗时秒 As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim 序号 As Long
    On Error Resume Next
    确保运行日志表
    Set ws = ThisWorkbook.Worksheets(RUNLOG_SHEET)
    If ws Is Nothing Then Exit Sub
    nextRow = ws.Cells(ws.Rows.count, RUNLOG_COL_SEQ).End(xlUp).row + 1
    If nextRow < 2 Then Exit Sub
    序号 = nextRow - 1
    ws.Cells(nextRow, RUNLOG_COL_SEQ).value = 序号
    ws.Cells(nextRow, RUNLOG_COL_TIME).value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(nextRow, RUNLOG_COL_USER).value = ""
    ws.Cells(nextRow, RUNLOG_COL_MODULE).value = 功能模块
    ws.Cells(nextRow, RUNLOG_COL_OP).value = 操作
    ws.Cells(nextRow, RUNLOG_COL_OBJ).value = 记录ID对象
    ws.Cells(nextRow, RUNLOG_COL_BEFORE).value = 操作前值
    ws.Cells(nextRow, RUNLOG_COL_AFTER).value = 操作后值
    ws.Cells(nextRow, RUNLOG_COL_RESULT).value = 结果
    ws.Cells(nextRow, RUNLOG_COL_DETAIL).value = 详细信息
    ws.Cells(nextRow, RUNLOG_COL_ELAPSED).value = 耗时秒
    ws.Cells(nextRow, RUNLOG_COL_PC).value = ""
End Sub

Private Sub 确保运行日志表()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RUNLOG_SHEET)
    If Not ws Is Nothing Then Exit Sub
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    If ws Is Nothing Then Exit Sub
    ws.Name = RUNLOG_SHEET
    ws.Cells(1, RUNLOG_COL_SEQ).value = "序号"
    ws.Cells(1, RUNLOG_COL_TIME).value = "时间戳"
    ws.Cells(1, RUNLOG_COL_USER).value = "用户名"
    ws.Cells(1, RUNLOG_COL_MODULE).value = "功能模块"
    ws.Cells(1, RUNLOG_COL_OP).value = "操作"
    ws.Cells(1, RUNLOG_COL_OBJ).value = "记录ID/对象"
    ws.Cells(1, RUNLOG_COL_BEFORE).value = "操作前值"
    ws.Cells(1, RUNLOG_COL_AFTER).value = "操作后值"
    ws.Cells(1, RUNLOG_COL_RESULT).value = "结果"
    ws.Cells(1, RUNLOG_COL_DETAIL).value = "详细信息"
    ws.Cells(1, RUNLOG_COL_ELAPSED).value = "耗时(秒)"
    ws.Cells(1, RUNLOG_COL_PC).value = "电脑名"
    ws.Range(ws.Cells(1, 1), ws.Cells(1, RUNLOG_COL_PC)).Font.Bold = True
End Sub
