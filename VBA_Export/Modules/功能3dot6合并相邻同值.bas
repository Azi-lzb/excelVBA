Attribute VB_Name = "功能3dot6合并相邻同值"
Option Explicit

' 合并相邻同值单元格：按 config「合并方向」选择横向（按行合并左右同值）或竖向（按列合并上下同值）。
' 与 3.5 取消合并并填充 互为逆操作。运行日志写在本模块内，不依赖 vbaSync。

Private Const RUNLOG_SHEET As String = "运行日志"
Private Const CONFIG_SHEET As String = "config"
Private Const CFG_KEY_MERGE As String = "3.6 合并相邻同值"
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

''' 对当前选中区域：按 config「合并方向」横向或竖向合并相邻同值单元格。
''' 需先选中工作表上的一个区域，再执行本过程（菜单或 Alt+F8）。
Public Sub 合并相邻同值单元格()
    Dim sel As Range
    Dim ws As Worksheet
    Dim rowIdx As Long
    Dim colIdx As Long
    Dim startCol As Long
    Dim endCol As Long
    Dim startRow As Long
    Dim endRow As Long
    Dim c As Long
    Dim r As Long
    Dim runStart As Long
    Dim cVal As Variant
    Dim t0 As Double
    Dim mergeDir As String
    Dim resultMsg As String

    On Error GoTo ErrHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "请先在工作表上选中一个单元格区域，再执行本功能。", vbExclamation
        Exit Sub
    End If

    Set sel = Selection
    If sel.Cells.count = 0 Then
        MsgBox "未选中有效区域。", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrTrust
    Set ws = sel.Parent
    On Error GoTo ErrHandler
    If ws Is Nothing Or TypeName(ws) <> "Worksheet" Then
        MsgBox "当前选区不在工作表上，无法处理。", vbExclamation
        Exit Sub
    End If

    mergeDir = Trim(读取配置值(CFG_KEY_MERGE, "合并方向"))
    If mergeDir = "" Then mergeDir = "竖向"
    If LCase(mergeDir) <> "横向" Then mergeDir = "竖向"

    t0 = Timer
    On Error Resume Next
    写运行日志 "3.6 合并相邻同值", "开始", sel.Address(False, False), "", "", "", "选中区域 " & sel.Address(False, False) & " 合并方向=" & mergeDir, ""
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    If mergeDir = "竖向" Then
        ' 按列处理：每列从上到下找连续同值段并合并
        For colIdx = 1 To sel.Columns.count
            startRow = sel.Cells(1, colIdx).row
            endRow = sel.Cells(sel.Rows.count, colIdx).row
            r = startRow
            Do While r <= endRow
                cVal = ws.Cells(r, sel.Cells(1, colIdx).Column).value
                runStart = r
                r = r + 1
                Do While r <= endRow And CellsValueEqual(ws.Cells(r, sel.Cells(1, colIdx).Column).value, cVal)
                    r = r + 1
                Loop
                If r - 1 > runStart Then
                    ws.Range(ws.Cells(runStart, sel.Cells(1, colIdx).Column), ws.Cells(r - 1, sel.Cells(1, colIdx).Column)).Merge
                End If
            Loop
        Next colIdx
        resultMsg = "已按列合并竖向相邻同值"
    Else
        ' 按行处理：每行从左到右找横向连续同值段并合并
        For rowIdx = 1 To sel.Rows.count
            startCol = sel.Cells(rowIdx, 1).Column
            endCol = sel.Cells(rowIdx, sel.Columns.count).Column
            c = startCol
            Do While c <= endCol
                cVal = ws.Cells(sel.Cells(rowIdx, 1).row, c).value
                runStart = c
                c = c + 1
                Do While c <= endCol And CellsValueEqual(ws.Cells(sel.Cells(rowIdx, 1).row, c).value, cVal)
                    c = c + 1
                Loop
                If c - 1 > runStart Then
                    ws.Range(ws.Cells(sel.Cells(rowIdx, 1).row, runStart), ws.Cells(sel.Cells(rowIdx, 1).row, c - 1)).Merge
                End If
            Loop
        Next rowIdx
        resultMsg = "已按行合并横向相邻同值"
    End If

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    On Error Resume Next
    写运行日志 "3.6 合并相邻同值", "完成", sel.Address(False, False), "", "", resultMsg, "", CStr(Round(Timer - t0, 2))
    On Error GoTo ErrHandler
    MsgBox "已对选中区域" & resultMsg & "单元格。", vbInformation
    Exit Sub

ErrTrust:
    MsgBox "访问工作表时出错。若为信任设置问题，请在「文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置」中勾选「信任对 VBA 工程对象模型的访问」。", vbExclamation
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
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

' 从 config 表读取配置：键（如 "3.6 合并相邻同值"）、键名（如 "合并方向"）
Private Function 读取配置值(ByVal 键 As String, ByVal 键名 As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim aVal As String
    Dim bVal As String
    读取配置值 = ""
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Function
    For i = 2 To lastRow
        aVal = Trim(CStr(ws.Cells(i, 1).value))
        bVal = Trim(CStr(ws.Cells(i, 2).value))
        If (aVal = "" Or aVal = 键) And LCase(bVal) = LCase(键名) Then
            读取配置值 = Trim(CStr(ws.Cells(i, 3).value))
            Exit Function
        End If
    Next i
End Function

' 比较两格值是否视为相同（用于合并判断）
Private Function CellsValueEqual(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsEmpty(a) And IsEmpty(b) Then CellsValueEqual = True: Exit Function
    If IsNull(a) And IsNull(b) Then CellsValueEqual = True: Exit Function
    If IsEmpty(a) Or IsNull(a) Then
        CellsValueEqual = (IsEmpty(b) Or IsNull(b) Or Trim(CStr(b)) = "")
        Exit Function
    End If
    If IsEmpty(b) Or IsNull(b) Then
        CellsValueEqual = (Trim(CStr(a)) = "")
        Exit Function
    End If
    CellsValueEqual = (a = b)
End Function
