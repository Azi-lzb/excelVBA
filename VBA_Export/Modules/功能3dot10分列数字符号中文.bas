Attribute VB_Name = "功能3dot10分列数字符号中文"
Option Explicit

' 3.10 分列：数字 / 符号 / 中文（及其他）
' 规则：
'   数字段：0-9 以及 26 个英文字母（字母和数字视为一类）
'   符号段：出现在配置中「符号字符」里的任何字符（默认： *\-/+ , . ; : ! ? @ # $ % ^ & ( ) [ ] 等）
'   中文段：既不是数字字母，也不是符号字符的其它字符（通常是 CJK 中文）

Private Const RUNLOG_SHEET As String = "运行日志"
Private Const CONFIG_SHEET As String = "config"
Private Const CFG_KEY As String = "3.10 分列数字符号中文"

' 输出顺序默认： 数字;符号;中文
Private Const ORDER_DEFAULT As String = "数;符;中"

' 默认符号字符集合（可在 config 表中配置覆盖）
Private Const SYMBOLS_DEFAULT As String = " *\-/+,.;:!?@#$%^&()[]"

' 运行日志列号
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

'================= 主入口 =================
' 从选中列起，将单元格内容拆成：数字段 / 符号段 / 中文段 三列
' 输出列顺序可以在 config 中通过「数;符;中」配置
Public Sub 功能3dot10_分列数字符号中文()
    Dim ws As Worksheet
    Dim sel As Range
    Dim colIdx As Long
    Dim lastRow As Long
    Dim orderStr As String
    Dim orderParts() As String
    Dim segNum As String
    Dim segSym As String
    Dim segChn As String
    Dim outCols(1 To 3) As Long
    Dim r As Long
    Dim cellVal As String
    Dim i As Long
    Dim t0 As Double
    Dim symChars As String

    On Error GoTo ErrHandler

    If TypeName(Selection) <> "Range" Then
        MsgBox "请先选中需要分列的列（单列或多行单列）。", vbExclamation
        Exit Sub
    End If

    Set sel = Selection
    Set ws = sel.Parent
    If ws Is Nothing Or TypeName(ws) <> "Worksheet" Then
        MsgBox "当前选择不在工作表中。", vbExclamation
        Exit Sub
    End If

    colIdx = sel.Cells(1, 1).Column
    lastRow = ws.Cells(ws.Rows.count, colIdx).End(xlUp).row
    If lastRow < 1 Then
        MsgBox "当前列没有数据。", vbExclamation
        Exit Sub
    End If

    ' 1. 读取输出顺序配置：数;符;中
    orderStr = Trim(读取配置值(CFG_KEY, "输出顺序"))
    If orderStr = "" Then orderStr = ORDER_DEFAULT
    orderParts = Split(orderStr, ";")
    If UBound(orderParts) - LBound(orderParts) + 1 < 3 Then
        MsgBox "3.10 分列数字符号中文 的配置有误，请在 config 中设置为：数;符;中 形式。", vbExclamation
        Exit Sub
    End If

    ' 将「数/符/中」映射到 1/2/3
    For i = 0 To 2
        Select Case Trim(LCase(orderParts(i)))
            Case "数"
                outCols(i + 1) = 1
            Case "符"
                outCols(i + 1) = 2
            Case "中"
                outCols(i + 1) = 3
            Case Else
                ' 非法值则按默认顺序：第 1 个→1，第 2 个→2，第 3 个→3
                outCols(i + 1) = i + 1
        End Select
    Next i

    ' 2. 读取符号字符集合
    symChars = Trim(读取配置值(CFG_KEY, "符号字符"))
    If symChars = "" Then symChars = SYMBOLS_DEFAULT

    t0 = Timer
    On Error Resume Next
    写运行日志 CFG_KEY, "开始", "C" & colIdx, "", "", "", "输出顺序=" & orderStr, ""
    On Error GoTo ErrHandler

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' 在选中列右侧插入两个新列，用来放另外两段
    ws.Columns(colIdx + 1).Insert
    ws.Columns(colIdx + 1).Insert

    ' 3. 循环逐行分列
    For r = 1 To lastRow
        cellVal = Trim(CStr(ws.Cells(r, colIdx).value))
        Call 拆分数字符号中文(cellVal, symChars, segNum, segSym, segChn)

        ' outCols(1)~(3) 分别是 1/2/3，表示数字/符号/中文
        ws.Cells(r, colIdx).value = 取分段(outCols(1), segNum, segSym, segChn)
        ws.Cells(r, colIdx + 1).value = 取分段(outCols(2), segNum, segSym, segChn)
        ws.Cells(r, colIdx + 2).value = 取分段(outCols(3), segNum, segSym, segChn)
    Next r

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    On Error Resume Next
    写运行日志 CFG_KEY, "完成", "C" & colIdx, "", "", "成功", "", CStr(Round(Timer - t0, 2))
    On Error GoTo ErrHandler

    MsgBox "数字 / 符号 / 中文 分列完成，共处理 " & lastRow & " 行。", vbInformation
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "3.10 分列数字符号中文 执行出错：" & vbCrLf & Err.Description, vbCritical
End Sub

'================= 辅助函数 =================

' typ: 1=数字段 2=符号段 3=中文段
Private Function 取分段(ByVal typ As Long, _
                        ByVal segNum As String, _
                        ByVal segSym As String, _
                        ByVal segChn As String) As String
    Select Case typ
        Case 1: 取分段 = segNum
        Case 2: 取分段 = segSym
        Case 3: 取分段 = segChn
        Case Else: 取分段 = segNum
    End Select
End Function

' 将字符串 s 拆成：首个数字/字母段、首个符号段、首个中文/其他段
' symbolChars 为符号字符集合
Private Sub 拆分数字符号中文(ByVal s As String, ByVal symbolChars As String, _
                             ByRef segNum As String, ByRef segSym As String, ByRef segChn As String)
    Dim i As Long
    Dim c As String
    Dim seg As String
    Dim kind As Long
    Dim prevKind As Long
    Dim got(1 To 3) As Boolean

    segNum = ""
    segSym = ""
    segChn = ""
    If Len(s) = 0 Then Exit Sub

    prevKind = 0
    seg = ""

    For i = 1 To Len(s)
        c = Mid$(s, i, 1)

        ' 数字字母：0-9 + A-Z / a-z
        If (c >= "0" And c <= "9") Or (c >= "A" And c <= "Z") Or (c >= "a" And c <= "z") Then
            kind = 1
        ' 符号：出现在 symbolChars 中
        ElseIf InStr(1, symbolChars, c, vbBinaryCompare) > 0 Then
            kind = 2
        ' 其它（通常为中文）
        Else
            kind = 3
        End If

        If kind <> prevKind Then
            ' 把前一段落到对应分段里（只取第一次出现的那段）
            If prevKind = 1 And Not got(1) Then segNum = seg: got(1) = True
            If prevKind = 2 And Not got(2) Then segSym = seg: got(2) = True
            If prevKind = 3 And Not got(3) Then segChn = seg: got(3) = True
            seg = ""
        End If

        seg = seg & c
        prevKind = kind
    Next i

    ' 处理最后一段
    If prevKind = 1 And Not got(1) Then segNum = seg: got(1) = True
    If prevKind = 2 And Not got(2) Then segSym = seg: got(2) = True
    If prevKind = 3 And Not got(3) Then segChn = seg: got(3) = True
End Sub

'================= 运行日志 & 配置 =================

Private Sub 写运行日志(ByVal 模块名 As String, ByVal 操作 As String, ByVal 对象ID As String, _
                      ByVal 修改前 As String, ByVal 修改后 As String, ByVal 结果 As String, _
                      ByVal 详情 As String, ByVal 耗时秒 As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim seq As Long

    On Error Resume Next
    初始化运行日志表
    Set ws = ThisWorkbook.Worksheets(RUNLOG_SHEET)
    If ws Is Nothing Then Exit Sub

    nextRow = ws.Cells(ws.Rows.count, RUNLOG_COL_SEQ).End(xlUp).row + 1
    If nextRow < 2 Then Exit Sub

    seq = nextRow - 1

    ws.Cells(nextRow, RUNLOG_COL_SEQ).value = seq
    ws.Cells(nextRow, RUNLOG_COL_TIME).value = Format(Now, "yyyy/mm/dd hh:nn:ss")
    ws.Cells(nextRow, RUNLOG_COL_USER).value = Environ$("Username")
    ws.Cells(nextRow, RUNLOG_COL_MODULE).value = 模块名
    ws.Cells(nextRow, RUNLOG_COL_OP).value = 操作
    ws.Cells(nextRow, RUNLOG_COL_OBJ).value = 对象ID
    ws.Cells(nextRow, RUNLOG_COL_BEFORE).value = 修改前
    ws.Cells(nextRow, RUNLOG_COL_AFTER).value = 修改后
    ws.Cells(nextRow, RUNLOG_COL_RESULT).value = 结果
    ws.Cells(nextRow, RUNLOG_COL_DETAIL).value = 详情
    ws.Cells(nextRow, RUNLOG_COL_ELAPSED).value = 耗时秒
    ws.Cells(nextRow, RUNLOG_COL_PC).value = ""
End Sub

Private Sub 初始化运行日志表()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RUNLOG_SHEET)
    If Not ws Is Nothing Then Exit Sub

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
    If ws Is Nothing Then Exit Sub

    ws.Name = RUNLOG_SHEET

    ws.Cells(1, RUNLOG_COL_SEQ).value = "序号"
    ws.Cells(1, RUNLOG_COL_TIME).value = "时间"
    ws.Cells(1, RUNLOG_COL_USER).value = "用户"
    ws.Cells(1, RUNLOG_COL_MODULE).value = "模块"
    ws.Cells(1, RUNLOG_COL_OP).value = "操作"
    ws.Cells(1, RUNLOG_COL_OBJ).value = "对象ID/范围"
    ws.Cells(1, RUNLOG_COL_BEFORE).value = "修改前"
    ws.Cells(1, RUNLOG_COL_AFTER).value = "修改后"
    ws.Cells(1, RUNLOG_COL_RESULT).value = "结果"
    ws.Cells(1, RUNLOG_COL_DETAIL).value = "详情"
    ws.Cells(1, RUNLOG_COL_ELAPSED).value = "耗时(秒)"
    ws.Cells(1, RUNLOG_COL_PC).value = "机器"

    ws.Range(ws.Cells(1, 1), ws.Cells(1, RUNLOG_COL_PC)).Font.Bold = True
End Sub

' 从 config 表中按「键 + 子键」读取值：
' 第 1 列：功能键（如 3.10 分列数字符号中文，或留空表示通用）
' 第 2 列：子键（如 输出顺序 / 符号字符）
' 第 3 列：对应配置值
Private Function 读取配置值(ByVal 主键 As String, ByVal 子键 As String) As String
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

        ' 主键为空或等于传入主键，且子键匹配（不区分大小写）
        If (aVal = "" Or aVal = 主键) And LCase$(bVal) = LCase$(子键) Then
            读取配置值 = Trim(CStr(ws.Cells(i, 3).value))
            Exit Function
        End If
    Next i
End Function
