Attribute VB_Name = "功能1dot6模板与源数据表格校验"
Option Explicit

' 执行面板布局：A2=模板路径，B5起=源文件路径，D列=表格数量校验，E列=表格样式检验
' 表格样式校验与「2.2.2 按批注汇总」一致：强制按模板时仅用模板的「模板」表行列区域对源文件所有 sheet 比对；否则同名 sheet 比对，未匹配的源 sheet 与模板「模板」表比对
Private Const PANEL_SHEET_NAME As String = "执行面板"
Private Const PANEL_DATA_START_ROW As Long = 5
Private Const PANEL_COL_PATH As Long = 2
Private Const PANEL_COL_COUNT_CHECK As Long = 4
Private Const PANEL_COL_STYLE_CHECK As Long = 5
Private Const CONFIG_SHEET_NAME As String = "config"
Private Const CONFIG_KEY_STYLE As String = "1.6 模板与源数据表格校验"
Private Const CFG_KEY_BIZHU As String = "2.2.2 按批注汇总"
Private Const TMPL_SHT As String = "模板"

' 从 config 表按 键、键名 取值
Private Function 读取配置(ByVal 键 As String, ByVal 键名 As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim aVal As String, bVal As String
    读取配置 = ""
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 2 Then Exit Function
    For i = 2 To lastRow
        aVal = Trim(CStr(ws.Cells(i, 1).value))
        bVal = Trim(CStr(ws.Cells(i, 2).value))
        If (aVal = "" Or aVal = 键) And LCase(bVal) = LCase(键名) Then
            读取配置 = Trim(CStr(ws.Cells(i, 3).value))
            Exit Function
        End If
    Next i
End Function

Private Function GetPanelWs() As Worksheet
    On Error Resume Next
    Set GetPanelWs = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    On Error GoTo 0
End Function

Private Function 工作簿工作表名称串(ByVal wb As Workbook, ByVal sep As String) As String
    Dim i As Long
    Dim s As String
    If wb Is Nothing Then 工作簿工作表名称串 = "": Exit Function
    s = ""
    For i = 1 To wb.Worksheets.count
        If i > 1 Then s = s & sep
        s = s & wb.Worksheets(i).Name
    Next i
    工作簿工作表名称串 = s
End Function

' 列号转列字母
Private Function ColNumToLetter(ByVal colNum As Long) As String
    Dim n As Long
    ColNumToLetter = ""
    n = colNum
    Do While n > 0
        ColNumToLetter = Chr(65 + ((n - 1) Mod 26)) & ColNumToLetter
        n = (n - 1) \ 26
    Loop
End Function

' 从批注文本中解析区域编号。批注如「行区域1」「行区域#1」「列区域2」。关键字为「行区域」或「列区域」，返回数字 N，0 表示无效。
Private Function 解析批注区域号(ByVal commentText As String, ByVal 关键字 As String) As Long
    Dim pos As Long
    Dim s As String
    Dim i As Long
    Dim numStr As String
    解析批注区域号 = 0
    If commentText = "" Or 关键字 = "" Then Exit Function
    commentText = Replace(commentText, vbLf, " ")
    commentText = Replace(commentText, vbCr, " ")
    pos = InStr(1, commentText, 关键字, vbTextCompare)
    If pos <= 0 Then Exit Function
    pos = pos + Len(关键字)
    s = Mid(commentText, pos)
    ' 跳过 # 若有
    If Len(s) >= 1 And Mid(s, 1, 1) = "#" Then s = Mid(s, 2)
    s = Trim(s)
    numStr = ""
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            numStr = numStr & Mid(s, i, 1)
        Else
            Exit For
        End If
    Next i
    If numStr = "" Then Exit Function
    On Error Resume Next
    解析批注区域号 = CLng(numStr)
    If 解析批注区域号 <= 0 Then 解析批注区域号 = 0
    On Error GoTo 0
End Function

' 获取单元格批注文本（兼容不同 Excel 版本）
Private Function 单元格批注文本(ByVal rng As Range) As String
    On Error Resume Next
    If rng.Comment Is Nothing Then 单元格批注文本 = "": Exit Function
    单元格批注文本 = rng.Comment.Text
    If 单元格批注文本 = "" Then 单元格批注文本 = rng.Comment.Comment.Text
    On Error GoTo 0
End Function

' 是否执行该类型校验（配置值为 是/1/true 等视为执行）
Private Function 是否执行校验(ByVal 配置值 As String) As Boolean
    Dim v As String
    v = LCase(Trim(配置值))
    If v = "是" Or v = "1" Or v = "true" Or v = "y" Or v = "yes" Then 是否执行校验 = True: Exit Function
    是否执行校验 = False
End Function

' 读取 config 布尔（与按批注汇总一致：1/是/true 为 True）
Private Function 配置布尔(ByVal 键 As String, ByVal 键名 As String, Optional ByVal 默认 As Boolean = False) As Boolean
    Dim v As String
    v = LCase(Trim(读取配置(键, 键名)))
    If v = "" Then 配置布尔 = 默认: Exit Function
    配置布尔 = (v = "1" Or v = "是" Or v = "true")
End Function

' 模板工作簿是否包含指定名称的工作表
Private Function 模板工作簿是否有表(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    On Error Resume Next
    模板工作簿是否有表 = (wb Is Nothing = False And Len(sheetName) > 0 And Not wb.Worksheets(sheetName) Is Nothing)
    On Error GoTo 0
End Function

' 从模板工作簿中收集「关键字」批注围成的矩形区域。关键字为「行区域」或「列区域」。
' 仅模板表=True 时只扫描名为「模板」的工作表（与强制按模板汇总一致）。
' 返回：字典 key = 表名 & "|||" & 区域号，item = "minR|maxR|minC|maxC"
Private Sub 收集批注矩形区域(ByVal templateWb As Workbook, ByVal 关键字 As String, ByRef outRects As Object, Optional ByVal 仅模板表 As Boolean = False)
    Dim sh As Long
    Dim ws As Worksheet
    Dim usedRng As Range
    Dim cell As Range
    Dim commentText As String
    Dim regionId As Long
    Dim key As String
    Dim minR As Long, maxR As Long, minC As Long, maxC As Long
    Dim arr() As String
    Dim r As Long, c As Long

    If templateWb Is Nothing Then Exit Sub
    If outRects Is Nothing Then Set outRects = CreateObject("Scripting.Dictionary")

    For sh = 1 To templateWb.Worksheets.count
        Set ws = templateWb.Worksheets(sh)
        If 仅模板表 And ws.Name <> TMPL_SHT Then GoTo NextSheet
        On Error Resume Next
        Set usedRng = ws.UsedRange
        On Error GoTo 0
        If usedRng Is Nothing Then GoTo NextSheet

        For Each cell In usedRng.Cells
            commentText = 单元格批注文本(cell)
            If commentText <> "" Then
                regionId = 解析批注区域号(commentText, 关键字)
                If regionId > 0 Then
                    key = ws.Name & "|||" & CStr(regionId)
                    If outRects.Exists(key) Then
                        arr = Split(outRects(key), "|")
                        minR = CLng(arr(0))
                        maxR = CLng(arr(1))
                        minC = CLng(arr(2))
                        maxC = CLng(arr(3))
                        If cell.Row < minR Then minR = cell.Row
                        If cell.Row > maxR Then maxR = cell.Row
                        If cell.Column < minC Then minC = cell.Column
                        If cell.Column > maxC Then maxC = cell.Column
                        outRects(key) = minR & "|" & maxR & "|" & minC & "|" & maxC
                    Else
                        outRects(key) = cell.Row & "|" & cell.Row & "|" & cell.Column & "|" & cell.Column
                    End If
                End If
            End If
        Next cell
NextSheet:
    Next sh
End Sub

' 表格样式校验：与按批注汇总一致。强制按模板时用模板「模板」表行列区域对源文件所有 sheet 逐格比对；否则按 matchKey（同名表或兜底「模板」表）比对。返回格式：sheet1:A1与模板文件sheet1:A1不一致；sheet2:A2与模板文件 模板:A2不一致
Private Function 表格样式校验(ByVal templateWb As Workbook, ByVal sourceWb As Workbook, ByVal 执行行区域 As Boolean, ByVal 执行列区域 As Boolean, ByVal 强制按模板 As Boolean) As String
    Dim rects As Object
    Dim key As Variant
    Dim arr() As String
    Dim tmplSheetName As String
    Dim minR As Long, maxR As Long, minC As Long, maxC As Long
    Dim tWs As Worksheet
    Dim srcWs As Worksheet
    Dim r As Long, c As Long
    Dim vT As String, vS As String
    Dim diffList As String
    Dim matchKey As String

    Set rects = CreateObject("Scripting.Dictionary")
    diffList = ""

    If 强制按模板 Then
        If 执行行区域 Then 收集批注矩形区域 templateWb, "行区域", rects, True
        If 执行列区域 Then 收集批注矩形区域 templateWb, "列区域", rects, True
    Else
        If 执行行区域 Then 收集批注矩形区域 templateWb, "行区域", rects, False
        If 执行列区域 Then 收集批注矩形区域 templateWb, "列区域", rects, False
    End If

    If 强制按模板 Then
        ' 用模板「模板」表对所有源 sheet 的同一矩形区域比对
        For Each srcWs In sourceWb.Worksheets
            For Each key In rects.Keys
                arr = Split(rects(key), "|")
                If UBound(arr) >= 3 Then
                    minR = CLng(arr(0)): maxR = CLng(arr(1)): minC = CLng(arr(2)): maxC = CLng(arr(3))
                    Set tWs = templateWb.Worksheets(TMPL_SHT)
                    For r = minR To maxR
                        For c = minC To maxC
                            vT = Trim(CStr(tWs.Cells(r, c).value))
                            vS = Trim(CStr(srcWs.Cells(r, c).value))
                            If vT <> vS Then
                                If diffList <> "" Then diffList = diffList & "；"
                                diffList = diffList & srcWs.Name & ":" & ColNumToLetter(c) & r & "与模板文件" & TMPL_SHT & ":" & ColNumToLetter(c) & r & "不一致"
                            End If
                        Next c
                    Next r
                End If
            Next key
        Next srcWs
    Else
        ' 按 sheet：每个源 sheet 对应 matchKey（同名或兜底「模板」），只比对该 matchKey 的矩形
        For Each srcWs In sourceWb.Worksheets
            If 模板工作簿是否有表(templateWb, srcWs.Name) Then
                matchKey = srcWs.Name
            ElseIf 模板工作簿是否有表(templateWb, TMPL_SHT) Then
                matchKey = TMPL_SHT
            Else
                GoTo NextSrcSheet
            End If
            For Each key In rects.Keys
                tmplSheetName = Split(CStr(key), "|||")(0)
                If tmplSheetName <> matchKey Then GoTo NextRectKey
                arr = Split(rects(key), "|")
                If UBound(arr) >= 3 Then
                    minR = CLng(arr(0)): maxR = CLng(arr(1)): minC = CLng(arr(2)): maxC = CLng(arr(3))
                    Set tWs = templateWb.Worksheets(matchKey)
                    For r = minR To maxR
                        For c = minC To maxC
                            vT = Trim(CStr(tWs.Cells(r, c).value))
                            vS = Trim(CStr(srcWs.Cells(r, c).value))
                            If vT <> vS Then
                                If diffList <> "" Then diffList = diffList & "；"
                                diffList = diffList & srcWs.Name & ":" & ColNumToLetter(c) & r & "与模板文件 " & matchKey & ":" & ColNumToLetter(c) & r & "不一致"
                            End If
                        Next c
                    Next r
                End If
NextRectKey:
            Next key
NextSrcSheet:
        Next srcWs
    End If

    表格样式校验 = diffList
End Function

Public Sub 模板与源数据表格校验()
    Dim ws As Worksheet
    Dim templatePath As String
    Dim templateWb As Workbook
    Dim sourceWb As Workbook
    Dim lastRow As Long
    Dim r As Long
    Dim srcPath As String
    Dim templateSheetNames As String
    Dim sourceSheetNames As String
    Dim templateCount As Long
    Dim sourceCount As Long
    Dim countWarn As Long
    Dim styleWarn As Long
    Dim 行区域配置 As String, 列区域配置 As String
    Dim 执行行区域 As Boolean, 执行列区域 As Boolean
    Dim 强制按模板 As Boolean
    Dim styleDiff As String
    Dim t0 As Double
    Dim oldScreenUpdating As Boolean
    Dim oldDisplayAlerts As Boolean

    t0 = Timer
    On Error Resume Next
    RunLog_WriteRow "1.6 模板与源数据表格校验", "开始", "", "", "", "", "开始", CStr(Round(Timer - t0, 2))
    On Error GoTo 0

    oldScreenUpdating = Application.ScreenUpdating
    oldDisplayAlerts = Application.DisplayAlerts
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set ws = GetPanelWs()
    If ws Is Nothing Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.DisplayAlerts = oldDisplayAlerts
        MsgBox "未找到执行面板，请先运行「6.2 初始化执行面板」或「5.3 初始化执行面板」。", vbExclamation
        Exit Sub
    End If

    templatePath = Trim$(CStr(ws.Cells(2, 1).value))
    If templatePath = "" Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.DisplayAlerts = oldDisplayAlerts
        MsgBox "执行面板 A2 未填写模板文件路径，请先用「1.1 选择模板文件」选择模板。", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set templateWb = Workbooks.Open(templatePath, ReadOnly:=True, UpdateLinks:=0)
    If Err.Number <> 0 Or templateWb Is Nothing Then
        Application.ScreenUpdating = oldScreenUpdating
        Application.DisplayAlerts = oldDisplayAlerts
        MsgBox "无法打开模板文件：" & templatePath & vbCrLf & Err.Description, vbCritical
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    templateCount = templateWb.Worksheets.count
    templateSheetNames = 工作簿工作表名称串(templateWb, "/")
    行区域配置 = 读取配置(CONFIG_KEY_STYLE, "行区域")
    列区域配置 = 读取配置(CONFIG_KEY_STYLE, "列区域")
    执行行区域 = 是否执行校验(行区域配置)
    执行列区域 = 是否执行校验(列区域配置)
    强制按模板 = 配置布尔(CFG_KEY_BIZHU, "强制按模板", False)

    lastRow = ws.Cells(ws.Rows.count, PANEL_COL_PATH).End(xlUp).row
    If lastRow < PANEL_DATA_START_ROW Then lastRow = PANEL_DATA_START_ROW - 1

    countWarn = 0
    styleWarn = 0
    For r = PANEL_DATA_START_ROW To lastRow
        srcPath = Trim$(CStr(ws.Cells(r, PANEL_COL_PATH).value))
        ws.Cells(r, PANEL_COL_COUNT_CHECK).value = ""
        ws.Cells(r, PANEL_COL_STYLE_CHECK).value = ""
        If srcPath = "" Then GoTo NextRow

        On Error Resume Next
        Set sourceWb = Workbooks.Open(srcPath, ReadOnly:=True, UpdateLinks:=0)
        If Err.Number <> 0 Or sourceWb Is Nothing Then
            ws.Cells(r, PANEL_COL_COUNT_CHECK).value = "无法打开：" & Err.Description
            countWarn = countWarn + 1
            On Error GoTo 0
            GoTo NextRow
        End If
        On Error GoTo 0

        sourceCount = sourceWb.Worksheets.count
        sourceSheetNames = 工作簿工作表名称串(sourceWb, "\")

        If sourceCount <> templateCount Then
            ws.Cells(r, PANEL_COL_COUNT_CHECK).value = "警告！与模板文件表格数量不一致，源文件工作表有" & sourceSheetNames & ";模板工作表有" & templateSheetNames
            countWarn = countWarn + 1
        Else
            ws.Cells(r, PANEL_COL_COUNT_CHECK).value = "校验通过"
        End If

        If (执行行区域 Or 执行列区域) Then
            styleDiff = 表格样式校验(templateWb, sourceWb, 执行行区域, 执行列区域, 强制按模板)
            If styleDiff <> "" Then
                ws.Cells(r, PANEL_COL_STYLE_CHECK).value = "表格样式不一致：" & styleDiff
                styleWarn = styleWarn + 1
            Else
                ws.Cells(r, PANEL_COL_STYLE_CHECK).value = "校验通过"
            End If
        End If

        sourceWb.Close SaveChanges:=False
NextRow:
    Next r

    templateWb.Close SaveChanges:=False

    Application.ScreenUpdating = oldScreenUpdating
    Application.DisplayAlerts = oldDisplayAlerts

    On Error Resume Next
    RunLog_WriteRow "1.6 模板与源数据表格校验", "完成", "", "", "", "", "数量不一致 " & countWarn & "，样式不一致 " & styleWarn, CStr(Round(Timer - t0, 2))
    On Error GoTo 0

    MsgBox "校验完成。" & vbCrLf & "表格数量不一致： " & countWarn & "（D 列）" & vbCrLf & "表格样式不一致： " & styleWarn & "（E 列）", vbInformation
End Sub
