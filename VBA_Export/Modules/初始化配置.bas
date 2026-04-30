Attribute VB_Name = "初始化配置"
Option Explicit

Private Const CONFIG_SHEET_NAME As String = "config"
Public Const RUNLOG_SHEET_NAME As String = "运行日志"
Public Const PANEL_SHEET_NAME As String = "执行面板"
Public Const MAPPING_SHEET_NAME As String = "机构映射表"
Public Const COMPARE_SHEET_NAME As String = "表格比对"
Public Const EXTRACT_SHEET_NAME As String = "工作表提取"
Public Const CONFIG_RENAME_SHEET_NAME As String = "config_rename"
Public Const HOME_SHEET_NAME As String = "统计工具"
Private Const INITDATA_DIR As String = "VBA_Export\InitData\"
Private Const TIMELINE_RULE_SHEET_NAME As String = "时序提取规则"
Private Const PATH_STD_MAP_SHEET_NAME As String = "路径标准化映射"
' 执行面板布局：
' A1=模板文件（标签），A2=模板文件路径
' B1=外部文件（标签），B2=外部文件路径
' 源文件区域表头行为第4行：A4=源文件，B4=路径，C4=文件名，D4=表格数量校验，E4=表格样式检验，F4=执行结果
' 从第5行起为源文件数据行：B列=完整路径，C列=文件名
Private Const PANEL_HEADER_ROW As Long = 4          ' 源文件表头行
Private Const PANEL_DATA_START_ROW As Long = 5      ' 源文件数据起始行
Private Const PANEL_COL_PATH As Long = 2            ' 源文件路径列（B）
Private Const PANEL_COL_SHORT As Long = 3           ' 源文件文件名列（C）

' 从 config 表按 A列=键 B列=键名 取 C列=值
Public Function 读取配置(ByVal 键 As String, ByVal 键名 As String) As String
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

' ============================================================
'  默认配置项定义
'  返回二维数组 (1..N, 1..4)：键、键名、值、备注
'  汇总功能使用 config 键「2.2.2 按批注汇总」
' ============================================================
Private Function GetDefaultConfigs() As Variant
    Dim d(1 To 22, 1 To 4) As Variant

    d(1, 1) = "2.2.1 按使用区域汇总": d(1, 2) = "跳过表头"
    d(1, 3) = "否":                    d(1, 4) = "填 是/1/true 开启；留空或否则弹窗询问"

    ' 2.2.2 按批注汇总
    d(2, 1) = "2.2.2 按批注汇总": d(2, 2) = "工作簿"
    d(2, 3) = "1":                   d(2, 4) = "结果中是否包含工作簿列（默认是）"

    d(3, 1) = "2.2.2 按批注汇总": d(3, 2) = "工作表"
    d(3, 3) = "1":                   d(3, 4) = "结果中是否包含工作表列（默认是）"

    d(4, 1) = "2.2.2 按批注汇总": d(4, 2) = "set区"
    d(4, 3) = "1":                   d(4, 4) = "是否提取 set() 字段（默认是）"

    d(5, 1) = "2.2.2 按批注汇总": d(5, 2) = "列区域"
    d(5, 3) = "1":                   d(5, 4) = "是否提取列区域数据（默认是）"

    d(6, 1) = "2.2.2 按批注汇总": d(6, 2) = "行号"
    d(6, 3) = "0":                   d(6, 4) = "是否追加行号列（默认否）"

    d(7, 1) = "2.2.2 按批注汇总": d(7, 2) = "跳过未匹配表"
    d(7, 3) = "0":                   d(7, 4) = "源表无匹配模板时是否跳过（默认否）"

    d(8, 1) = "2.2.2 按批注汇总": d(8, 2) = "分列"
    d(8, 3) = "0":                   d(8, 4) = "0=不分列；A: =A列按空格；A,B:_=A B按下划线；多规则用;分隔"

    d(9, 1) = "2.2.2 按批注汇总": d(9, 2) = "不参与的工作表"
    d(9, 3) = "":                    d(9, 4) = "分号分隔，匹配到的表不参与汇总"

    d(10, 1) = "2.2.2 按批注汇总": d(10, 2) = "参与的工作表"
    d(10, 3) = "":                   d(10, 4) = "分号分隔，仅匹配到的表参与；与不参与同时配置时以此为准"

    d(11, 1) = "2.2.2 按批注汇总": d(11, 2) = "强制按模板"
    d(11, 3) = "0":                  d(11, 4) = "1/是/true=仅用模板表「模板」、只生成一张表 ALL、不参与过滤；0=按 Sheet 名与参与/不参与规则"

    ' 2.4 批量Excel格式转换
    d(12, 1) = "2.4 批量Excel格式转换": d(12, 2) = "目的格式"
    d(12, 3) = "xlsx":               d(12, 4) = "另存为格式：xls,xlsx,xlsm,csv,xlt,xltx,xltm,xlsb；结果存源目录下同名子文件夹"

    ' 3.7 注入VBA到源文件
    d(13, 1) = "3.7 注入VBA到源文件": d(13, 2) = "模块"
    d(13, 3) = "all":                d(13, 4) = "要注入的 .bas 模块名（与文件名不含扩展名一致），分号分隔；填 all 表示注入全部"
    d(14, 1) = "3.7 注入VBA到源文件": d(14, 2) = "跳过模块"
    d(14, 3) = "vbaSync":            d(14, 4) = "不注入的 .bas 模块名，分号分隔；默认跳过 vbaSync"
    d(15, 1) = "3.7 注入VBA到源文件": d(15, 2) = "复制ThisWorkbook"
    d(15, 3) = "是":                 d(15, 4) = "是否将本工作簿的 ThisWorkbook 代码复制到目标；填 是/1/true 开启"

    ' 3.8 清除目标工作簿VBA
    d(16, 1) = "3.8 清除目标工作簿VBA": d(16, 2) = "清除ThisWorkbook"
    d(16, 3) = "是":                  d(16, 4) = "是否清空目标工作簿 ThisWorkbook 内代码；填 是/1/true 则清空，否则仅删除 .bas 模块"

    ' 3.9 追加列（按批注）
    d(17, 1) = "3.9 追加列（按批注）": d(17, 2) = "关键字"
    d(17, 3) = "sheet1;sheet2":      d(17, 4) = "分号分隔；源文件名包含某关键字则取该关键字对应工作表指定列追加到外部文件"
    d(18, 1) = "3.9 追加列（按批注）": d(18, 2) = "追加列批注"
    d(18, 3) = "追加列":             d(18, 4) = "模板表头批注包含此文本的列参与字典，默认「追加列」"

    ' 3.6 合并相邻同值
    d(19, 1) = "3.6 合并相邻同值": d(19, 2) = "合并方向"
    d(19, 3) = "竖向":              d(19, 4) = "横向=按行合并左右相邻同值；竖向=按列合并上下相邻同值"

    ' 2.5 批量Word格式转换
    d(20, 1) = "2.5 批量Word格式转换": d(20, 2) = "目的格式"
    d(20, 3) = "docx":                  d(20, 4) = "另存为格式：doc 或 docx；结果存源目录下同名子文件夹"

    ' 1.6 模板与源数据表格校验（是否执行行区域/列区域批注矩形校验）
    d(21, 1) = "1.6 模板与源数据表格校验": d(21, 2) = "行区域"
    d(21, 3) = "是":                      d(21, 4) = "是否执行行区域校验：是=在模板批注中找「行区域N」与「行区域#N」围成的矩形逐格比对"
    d(22, 1) = "1.6 模板与源数据表格校验": d(22, 2) = "列区域"
    d(22, 3) = "否":                      d(22, 4) = "是否执行列区域校验：是=在模板批注中找「列区域N」与「列区域#N」围成的矩形逐格比对"

    GetDefaultConfigs = d
End Function

' ============================================================
'  检查 A+B 组合是否已存在（不区分大小写）
' ============================================================
Private Function ConfigExists(ByVal ws As Worksheet, ByVal lastRow As Long, _
                               ByVal keyA As String, ByVal keyB As String) As Boolean
    Dim i As Long
    ConfigExists = False
    For i = 2 To lastRow
        If LCase(Trim(CStr(ws.Cells(i, 1).value))) = LCase(keyA) And _
           LCase(Trim(CStr(ws.Cells(i, 2).value))) = LCase(keyB) Then
            ConfigExists = True
            Exit Function
        End If
    Next i
End Function

' ============================================================
'  初始化 config 表
'  · 不存在 → 新建表 + 写表头 + 写全部默认配置
'  · 已存在 → 逐条检查 A+B 是否重复，缺失则追加，已有则跳过
' ============================================================
Public Sub 初始化config()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim created As Boolean
    Dim defs As Variant
    Dim i As Long, addCount As Long
    Dim t0 As Double

    t0 = Timer

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET_NAME)
    On Error GoTo 0

    created = False
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = CONFIG_SHEET_NAME
        created = True
    End If

    If created Or Trim(CStr(ws.Cells(1, 1).value)) = "" Then
        ws.Cells(1, 1).value = "键"
        ws.Cells(1, 2).value = "键名"
        ws.Cells(1, 3).value = "值"
        ws.Cells(1, 4).value = "备注"
        ws.Rows(1).Font.Bold = True
    End If

    defs = GetDefaultConfigs()
    addCount = 0

    For i = 1 To UBound(defs, 1)
        lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If lastRow < 1 Then lastRow = 1

        If Not ConfigExists(ws, lastRow, CStr(defs(i, 1)), CStr(defs(i, 2))) Then
            lastRow = lastRow + 1
            ws.Cells(lastRow, 1).value = defs(i, 1)
            ws.Cells(lastRow, 2).value = defs(i, 2)
            ws.Cells(lastRow, 3).value = defs(i, 3)
            ws.Cells(lastRow, 4).value = defs(i, 4)
            addCount = addCount + 1
        End If
    Next i

    ws.Columns("A:D").AutoFit
    ws.Activate

    Call 初始化执行面板
    If created Then
        MsgBox "已创建 config 表并写入 " & addCount & " 条默认配置。", vbInformation
    ElseIf addCount > 0 Then
        MsgBox "config 表已存在，新增 " & addCount & " 条缺失配置。", vbInformation
    Else
        MsgBox "config 表已是最新，无需更新。", vbInformation
    End If
End Sub

' ============================================================
'  初始化执行面板表（供 2.2.3 按Sheet页汇总 使用）
'  第1行：A1=模板文件（标签），B1=外部文件（标签）
'  第2行：A2=模板文件路径，B2=外部文件路径
'  第4行：源文件表头 A4=源文件，B4=路径，C4=文件名，D4=校验结果，F4=执行结果
'  第5行起：源文件数据行，B列=完整路径，C列=文件名，D/F 预留给后续功能
' ============================================================
Public Sub 初始化执行面板()
    Dim ws As Worksheet
    Dim found As Boolean

    found = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    If Not ws Is Nothing Then found = True
    On Error GoTo 0

    If Not found Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = PANEL_SHEET_NAME
    End If

    With ws
        ' 顶部模板 / 外部文件区域
        .Cells(1, 1).value = "模板文件"
        .Cells(1, 2).value = "外部文件"
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 2).Font.Bold = True
        .Cells(2, 1).value = vbNullString   ' 模板文件路径（由“选择模板文件”过程写入）
        .Cells(2, 2).value = vbNullString   ' 外部文件路径（预留，后续功能使用）

        ' 源文件表头与数据区域
        .Cells(PANEL_HEADER_ROW, 1).value = "源文件"
        .Cells(PANEL_HEADER_ROW, PANEL_COL_PATH).value = "路径"
        .Cells(PANEL_HEADER_ROW, PANEL_COL_SHORT).value = "文件名"
        .Cells(PANEL_HEADER_ROW, 4).value = "表格数量校验"
        .Cells(PANEL_HEADER_ROW, 5).value = "表格样式检验"
        .Cells(PANEL_HEADER_ROW, 6).value = "执行结果"
        .Rows(PANEL_HEADER_ROW).Font.Bold = True

        .Columns("A:F").AutoFit
    End With
End Sub

' ============================================================
'  初始化运行日志表
' ============================================================
Public Sub 初始化运行日志()
    Dim ws As Worksheet
    Dim found As Boolean

    found = False
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(RUNLOG_SHEET_NAME)
    If Not ws Is Nothing Then found = True
    On Error GoTo 0

    If Not found Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = RUNLOG_SHEET_NAME
    End If

    With ws
        .Cells(1, 1).value = "序号"
        .Cells(1, 2).value = "时间戳"
        .Cells(1, 3).value = "用户名"
        .Cells(1, 4).value = "功能模块"
        .Cells(1, 5).value = "操作"
        .Cells(1, 6).value = "记录ID/对象"
        .Cells(1, 7).value = "操作前值"
        .Cells(1, 8).value = "操作后值"
        .Cells(1, 9).value = "结果"
        .Cells(1, 10).value = "详细信息"
        .Cells(1, 11).value = "耗时(秒)"
        .Cells(1, 12).value = "电脑名"
        .Range("A1:L1").Font.Bold = True
        .Columns("A:L").AutoFit
    End With
End Sub

' ============================================================
'  初始化机构映射表、表格比对、config_rename（整表从 InitData TSV 回填）
'  导出初始化数据：将当前工作簿三张表整表导出到 InitData，方便后续初始化录入
' ============================================================
Private Function EnsureSheetByName(ByVal sheetName As String, ByRef created As Boolean) As Worksheet
    created = False
    On Error Resume Next
    Set EnsureSheetByName = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If EnsureSheetByName Is Nothing Then
        Set EnsureSheetByName = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        If Not EnsureSheetByName Is Nothing Then
            EnsureSheetByName.Name = sheetName
            created = True
        End If
    End If
End Function

' 从 TSV 文件整表回填到工作表。expectCols>0 时写入该列数；=0 时按文件每行最大列数写入。
Private Function 从TSV回填工作表(ByVal ws As Worksheet, ByVal relativePath As String, ByVal expectCols As Long) As Boolean
    Dim fso As Object, ts As Object
    Dim p As String
    Dim lines As Variant, cells As Variant
    Dim i As Long, j As Long
    Dim lineText As String
    Dim maxCol As Long

    从TSV回填工作表 = False
    If ws Is Nothing Then Exit Function

    p = ThisWorkbook.path
    If Right$(p, 1) <> "\" Then p = p & "\"
    p = p & relativePath

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(p) Then Exit Function

    Set ts = fso.OpenTextFile(p, 1, False, -2)
    If ts Is Nothing Then Exit Function
    lineText = ts.ReadAll
    ts.Close

    ws.Cells.Clear
    lines = Split(Replace(lineText, vbCrLf, vbLf), vbLf)
    If expectCols <= 0 Then
        maxCol = 0
        For i = LBound(lines) To UBound(lines)
            If Len(Trim(CStr(lines(i)))) > 0 Then
                cells = Split(CStr(lines(i)), vbTab)
                If UBound(cells) - LBound(cells) + 1 > maxCol Then maxCol = UBound(cells) - LBound(cells) + 1
            End If
        Next i
        If maxCol < 1 Then maxCol = 1
        expectCols = maxCol
    End If

    For i = LBound(lines) To UBound(lines)
        If i = UBound(lines) And Trim(CStr(lines(i))) = "" Then Exit For
        cells = Split(CStr(lines(i)), vbTab)
        For j = 0 To expectCols - 1
            If j <= UBound(cells) Then
                ws.Cells(i + 1, j + 1).value = CStr(cells(j))
            Else
                ws.Cells(i + 1, j + 1).value = ""
            End If
        Next j
    Next i
    ws.Rows(1).Font.Bold = True
    从TSV回填工作表 = True
End Function

Public Sub 初始化机构映射表()
    Dim ws As Worksheet
    Dim created As Boolean
    Set ws = EnsureSheetByName(MAPPING_SHEET_NAME, created)
    If ws Is Nothing Then Exit Sub

    If Not 从TSV回填工作表(ws, INITDATA_DIR & "机构映射表.tsv", 8) Then
        With ws
            .Cells.Clear
            .Cells(1, 1).value = "报表机构名"
            .Cells(1, 2).value = "映射机构名"
            .Cells(1, 3).value = "是否外资行"
            .Range("A1:C1").Font.Bold = True
            .Columns("A:H").AutoFit
        End With
        MsgBox "未找到基准数据文件，已仅初始化表头：" & MAPPING_SHEET_NAME, vbExclamation
        Exit Sub
    End If

    ws.Columns("A:H").AutoFit
    MsgBox "已按基准数据整表初始化：" & MAPPING_SHEET_NAME, vbInformation
End Sub

Public Sub 初始化表格比对()
    Dim ws As Worksheet
    Dim created As Boolean
    Set ws = EnsureSheetByName(COMPARE_SHEET_NAME, created)
    If ws Is Nothing Then Exit Sub

    If 从TSV回填工作表(ws, INITDATA_DIR & "表格比对.tsv", 0) Then
        ws.Columns("A:O").AutoFit
        MsgBox "已按基准数据整表初始化：" & COMPARE_SHEET_NAME, vbInformation
        Exit Sub
    End If

    With ws
        .Cells.Clear
        .Cells(1, 1).value = "A-币种关键字"
        .Cells(1, 2).value = "A-地区关键字"
        .Cells(1, 3).value = "A-机构关键字"
        .Cells(1, 4).value = "A-类型关键字"
        .Cells(1, 5).value = "A-名称关键字"
        .Cells(1, 6).value = "B-币种关键字"
        .Cells(1, 7).value = "B-地区关键字"
        .Cells(1, 8).value = "B-机构关键字"
        .Cells(1, 9).value = "B-类型关键字"
        .Cells(1, 10).value = "B-名称关键字"
        .Cells(1, 11).value = "是否执行(1=是)"
        .Cells(1, 12).value = "是否全量核对(1=是)"
        .Cells(1, 13).value = "表A-位置"
        .Cells(1, 14).value = "表B-位置"
        .Cells(1, 15).value = "备注"
        .Range("A1:O1").Font.Bold = True
        .Columns("A:O").AutoFit
    End With
    MsgBox "未找到表格比对.tsv，已仅初始化表头：" & COMPARE_SHEET_NAME, vbExclamation
End Sub

Public Sub 初始化工作表提取()
    Dim ws As Worksheet
    Dim created As Boolean
    Set ws = EnsureSheetByName(EXTRACT_SHEET_NAME, created)
    If ws Is Nothing Then Exit Sub

    With ws
        If created Or Trim$(CStr(.Cells(1, 1).value)) = "" Then
            .Cells(1, 1).value = "币种关键字"
            .Cells(1, 2).value = "地区关键字"
            .Cells(1, 3).value = "机构关键字"
            .Cells(1, 4).value = "类型关键字"
            .Cells(1, 5).value = "名称关键字"
            .Cells(1, 6).value = "是否禁用(1=是)"
            .Cells(1, 7).value = "是否整表提取(1=是)"
            .Cells(1, 8).value = "提取行(如2:100)"
            .Cells(1, 9).value = "提取列(如A:Z)"
            .Cells(1, 10).value = "输出文件名"
            .Cells(1, 11).value = "备注"
            .Range("A1:K1").Font.Bold = True
        End If
        .Columns("A:K").AutoFit
    End With
    MsgBox "已初始化工作表：" & EXTRACT_SHEET_NAME, vbInformation
End Sub

Public Sub 初始化config_rename()
    Dim ws As Worksheet
    Dim created As Boolean
    Set ws = EnsureSheetByName(CONFIG_RENAME_SHEET_NAME, created)
    If ws Is Nothing Then Exit Sub

    If Not 从TSV回填工作表(ws, INITDATA_DIR & "config_rename.tsv", 11) Then
        With ws
            .Cells.Clear
            .Cells(1, 1).value = "简称"
            .Cells(1, 2).value = "全称"
            .Cells(1, 4).value = "代码"
            .Cells(1, 5).value = "全称(用于全称->代码映射)"
            .Cells(1, 7).value = "键"
            .Cells(1, 8).value = "值"
            .Cells(1, 10).value = "原表名(J)"
            .Cells(1, 11).value = "新表名(K)"
            .Range("A1:K1").Font.Bold = True
            .Columns("A:K").AutoFit
        End With
        MsgBox "未找到基准数据文件，已仅初始化表头：" & CONFIG_RENAME_SHEET_NAME, vbExclamation
        Exit Sub
    End If

    ws.Columns("A:K").AutoFit
    MsgBox "已按基准数据整表初始化：" & CONFIG_RENAME_SHEET_NAME, vbInformation
End Sub

Public Sub 初始化统计工具()
    Dim ws As Worksheet
    Dim created As Boolean
    Set ws = EnsureSheetByName(HOME_SHEET_NAME, created)
    If ws Is Nothing Then Exit Sub

    With ws
        If created Or Trim$(CStr(.Cells(1, 1).value)) = "" Then
            .Cells(1, 1).value = "报表工具工作簿首页"
            .Cells(1, 1).Font.Bold = True
            .Cells(3, 1).value = "配置表/功能表："
            .Cells(4, 1).value = "config"
            .Cells(5, 1).value = "执行面板"
            .Cells(6, 1).value = "运行日志"
            .Cells(7, 1).value = "机构映射表"
            .Cells(8, 1).value = "表格比对"
            .Cells(9, 1).value = "工作表提取"
            .Cells(10, 1).value = "config_rename"
            .Columns("A:B").AutoFit
        End If
    End With
    MsgBox "已初始化工作表：" & HOME_SHEET_NAME, vbInformation
End Sub

Public Sub InitTimelineRuleSheet()
    Dim ws As Worksheet
    Dim created As Boolean
    Set ws = EnsureSheetByName(TIMELINE_RULE_SHEET_NAME, created)
    If ws Is Nothing Then Exit Sub

    With ws
        .Cells(1, 1).value = "是否启用"
        .Cells(1, 2).value = "规则名称"
        .Cells(1, 3).value = "工作簿关键字"
        .Cells(1, 4).value = "工作表关键字"
        .Cells(1, 5).value = "行头列"
        .Cells(1, 6).value = "列表头行"
        .Cells(1, 7).value = "必含列头"
        .Cells(1, 8).value = "必含行头"
        .Cells(1, 9).value = "数据起始行"
        .Cells(1, 10).value = "数据结束行"
        .Cells(1, 11).value = "数据起始列"
        .Cells(1, 12).value = "数据结束列"
        .Cells(1, 13).value = "跳过关键字"
        .Cells(1, 14).value = "目标工作簿路径"
        .Cells(1, 15).value = "目标工作表"
        .Cells(1, 16).value = "启用目标写入"
        .Range("A1:P1").Font.Bold = True
        .Columns("A:P").AutoFit
    End With

    SetHeaderCommentLocal ws.Cells(1, 1), "是否启用该规则。填写 Y/1/TRUE/是 时生效。"
    SetHeaderCommentLocal ws.Cells(1, 2), "规则名称。建议同类规则命名一致。"
    SetHeaderCommentLocal ws.Cells(1, 3), "工作簿关键字。支持多个关键字，用分号分隔。"
    SetHeaderCommentLocal ws.Cells(1, 4), "工作表关键字。支持多个关键字，用分号分隔。"
    SetHeaderCommentLocal ws.Cells(1, 5), "行头所在列，如 B 或 2。"
    SetHeaderCommentLocal ws.Cells(1, 6), "列表头所在行，支持多行，如 38,39。"
    SetHeaderCommentLocal ws.Cells(1, 7), "必含列头路径，支持多个，用分号分隔。"
    SetHeaderCommentLocal ws.Cells(1, 8), "必含行头路径，支持多个，用分号分隔。"
    SetHeaderCommentLocal ws.Cells(1, 9), "数据起始行。"
    SetHeaderCommentLocal ws.Cells(1, 10), "数据结束行。可为空表示到尾部。"
    SetHeaderCommentLocal ws.Cells(1, 11), "数据起始列，如 C 或 3。"
    SetHeaderCommentLocal ws.Cells(1, 12), "数据结束列。可为空表示到尾部。"
    SetHeaderCommentLocal ws.Cells(1, 13), "跳过关键字。命中则跳过该表。"
    SetHeaderCommentLocal ws.Cells(1, 14), "可选。目标工作簿路径。为空时回退到临时结果工作簿。"
    SetHeaderCommentLocal ws.Cells(1, 15), "可选。目标工作表。不存在会自动新建。"
    SetHeaderCommentLocal ws.Cells(1, 16), "可选。填写 Y/1/TRUE/是 时，允许写入目标工作簿。留空默认启用。"

    MsgBox "已初始化工作表：" & TIMELINE_RULE_SHEET_NAME, vbInformation
End Sub

Public Sub InitPathStandardMapSheet()
    Dim ws As Worksheet
    Dim created As Boolean
    Set ws = EnsureSheetByName(PATH_STD_MAP_SHEET_NAME, created)
    If ws Is Nothing Then Exit Sub

    With ws
        .Cells(1, 1).value = "是否启用"
        .Cells(1, 2).value = "规则名称"
        .Cells(1, 3).value = "工作簿关键字"
        .Cells(1, 4).value = "工作表关键字"
        .Cells(1, 5).value = "目标类型"
        .Cells(1, 6).value = "匹配方式"
        .Cells(1, 7).value = "原始路径"
        .Cells(1, 8).value = "标准路径"
        .Cells(1, 9).value = "备注"
        .Range("A1:I1").Font.Bold = True
        .Columns("A:I").AutoFit
    End With

    MsgBox "已初始化工作表：" & PATH_STD_MAP_SHEET_NAME, vbInformation
End Sub

Public Sub 初始化时序提取规则()
    InitTimelineRuleSheet
End Sub

Public Sub 初始化路径标准化映射_旧入口()
    InitPathStandardMapSheet
End Sub

Private Sub SetHeaderCommentLocal(ByVal targetCell As Range, ByVal commentText As String)
    On Error Resume Next
    If Not targetCell.Comment Is Nothing Then targetCell.Comment.Delete
    targetCell.AddComment commentText
    On Error GoTo 0
End Sub

' 将工作表整表（含表头与数据）导出为 TSV，用于保存到 InitData 供后续初始化使用。
Private Function 工作表导出为TSV(ByVal ws As Worksheet, ByVal fullPath As String) As Boolean
    Dim fso As Object, ts As Object
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim arr() As String
    Dim line As String

    工作表导出为TSV = False
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If ws.Cells(ws.Rows.count, 2).End(xlUp).row > lastRow Then lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    On Error GoTo 0
    If lastRow < 1 Or lastCol < 1 Then Exit Function

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(fullPath, True, False)
    If ts Is Nothing Then Exit Function

    For i = 1 To lastRow
        ReDim arr(1 To lastCol)
        For j = 1 To lastCol
            arr(j) = CStr(ws.Cells(i, j).value)
        Next j
        line = Join(arr, vbTab)
        ts.WriteLine line
    Next i
    ts.Close
    工作表导出为TSV = True
End Function

' 导出初始化数据：将当前工作簿的机构映射表、表格比对、config_rename 整表导出到 VBA_Export\InitData\，方便后续初始化录入。
Public Sub 导出初始化数据()
    Dim fso As Object
    Dim baseDir As String
    Dim ws As Worksheet
    Dim cnt As Long

    baseDir = ThisWorkbook.path
    If Right$(baseDir, 1) <> "\" Then baseDir = baseDir & "\"
    baseDir = baseDir & INITDATA_DIR

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(baseDir) Then
        Dim parentDir As String
        parentDir = fso.GetParentFolderName(Left$(baseDir, Len(baseDir) - 1))
        If Len(parentDir) > 0 And Not fso.FolderExists(parentDir) Then fso.CreateFolder parentDir
        fso.CreateFolder Left$(baseDir, Len(baseDir) - 1)
    End If

    cnt = 0
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(MAPPING_SHEET_NAME)
    If Not ws Is Nothing Then
        If 工作表导出为TSV(ws, baseDir & "机构映射表.tsv") Then cnt = cnt + 1
    End If
    Set ws = ThisWorkbook.Worksheets(COMPARE_SHEET_NAME)
    If Not ws Is Nothing Then
        If 工作表导出为TSV(ws, baseDir & "表格比对.tsv") Then cnt = cnt + 1
    End If
    Set ws = ThisWorkbook.Worksheets(CONFIG_RENAME_SHEET_NAME)
    If Not ws Is Nothing Then
        If 工作表导出为TSV(ws, baseDir & "config_rename.tsv") Then cnt = cnt + 1
    End If
    On Error GoTo 0

    MsgBox "已导出 " & cnt & " 张表到 " & baseDir & "（整表，含表头与数据）。", vbInformation
End Sub

' ============================================================
'  执行面板与文件选择（与上面执行面板布局约定一致）
' ============================================================
Private Function GetPanelWs() As Worksheet
    On Error Resume Next
    Set GetPanelWs = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    On Error GoTo 0
End Function

' 执行面板源文件数据区 B 列是否已存在同路径
Private Function PanelHasSourcePath(ByVal fullPath As String) As Boolean
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    fullPath = Trim$(fullPath)
    If fullPath = "" Then PanelHasSourcePath = False: Exit Function
    Set ws = GetPanelWs()
    If ws Is Nothing Then PanelHasSourcePath = False: Exit Function
    lastRow = ws.Cells(ws.Rows.count, PANEL_COL_PATH).End(xlUp).row
    If lastRow < PANEL_DATA_START_ROW Then PanelHasSourcePath = False: Exit Function
    For r = PANEL_DATA_START_ROW To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, PANEL_COL_PATH).value)), fullPath, vbTextCompare) = 0 Then
            PanelHasSourcePath = True: Exit Function
        End If
    Next r
    PanelHasSourcePath = False
End Function

' 向执行面板追加一行源文件（去重 + 超链接），写入 B/C 列
Private Sub PanelAddSourcePath(ByVal fullPath As String)
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    fullPath = Trim$(fullPath)
    If fullPath = "" Then Exit Sub
    If PanelHasSourcePath(fullPath) Then Exit Sub
    Set ws = GetPanelWs()
    If ws Is Nothing Then Exit Sub
    lastRow = ws.Cells(ws.Rows.count, PANEL_COL_PATH).End(xlUp).row
    If lastRow < PANEL_DATA_START_ROW Then lastRow = PANEL_HEADER_ROW
    r = lastRow + 1
    If r < PANEL_DATA_START_ROW Then r = PANEL_DATA_START_ROW
    ws.Cells(r, PANEL_COL_PATH).value = fullPath
    ws.Cells(r, PANEL_COL_SHORT).value = ShortFileName(fullPath)
    ws.Hyperlinks.Add Anchor:=ws.Cells(r, PANEL_COL_PATH), Address:=CStr(fullPath), TextToDisplay:=CStr(fullPath)
End Sub

Private Function ShortFileName(ByVal fullPath As String) As String
    Dim p As Long
    fullPath = Trim$(fullPath)
    If fullPath = "" Then ShortFileName = "": Exit Function
    p = InStrRev(fullPath, "\")
    If p > 0 Then ShortFileName = Mid$(fullPath, p + 1) Else ShortFileName = fullPath
End Function

' 1.1 选择模板文件：写入 A2=模板文件路径，带超链接
Public Sub 选择模板文件()
    Dim fd As FileDialog
    Dim ws As Worksheet
    Set ws = GetPanelWs()
    If ws Is Nothing Then
        MsgBox "请先运行「3.3 初始化配置」创建执行面板。", vbExclamation
        Exit Sub
    End If
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择模板文件"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel 文件", "*.xls;*.xlsx;*.xlsm"
    End With
    If fd.Show <> True Then Exit Sub
    With ws
        .Cells(2, 1).value = fd.SelectedItems(1)
        .Hyperlinks.Add Anchor:=.Cells(2, 1), Address:=CStr(fd.SelectedItems(1)), TextToDisplay:=CStr(fd.SelectedItems(1))
    End With
    MsgBox "已写入模板路径。", vbInformation
End Sub

' 1.2 选择外部文件：写入 B2=外部文件路径，带超链接
Public Sub 选择外部文件()
    Dim fd As FileDialog
    Dim ws As Worksheet
    Set ws = GetPanelWs()
    If ws Is Nothing Then
        MsgBox "请先运行「3.3 初始化配置」创建执行面板。", vbExclamation
        Exit Sub
    End If
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择外部文件"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel 文件", "*.xls;*.xlsx;*.xlsm"
    End With
    If fd.Show <> True Then Exit Sub
    With ws
        .Cells(2, 2).value = fd.SelectedItems(1)
        .Hyperlinks.Add Anchor:=.Cells(2, 2), Address:=CStr(fd.SelectedItems(1)), TextToDisplay:=CStr(fd.SelectedItems(1))
    End With
    MsgBox "已写入外部文件路径。", vbInformation
End Sub

' 1.3 选择源文件：多选，去重追加到执行面板
Public Sub 选择源文件()
    Dim fd As FileDialog
    Dim ws As Worksheet
    Dim i As Long, added As Long
    Set ws = GetPanelWs()
    If ws Is Nothing Then
        MsgBox "请先运行「3.3 初始化配置」创建执行面板。", vbExclamation
        Exit Sub
    End If
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择要汇总的源文件"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel 文件", "*.xls;*.xlsx;*.xlsm"
    End With
    If fd.Show <> True Then Exit Sub
    added = 0
    For i = 1 To fd.SelectedItems.count
        If Not PanelHasSourcePath(fd.SelectedItems(i)) Then
            PanelAddSourcePath fd.SelectedItems(i)
            added = added + 1
        End If
    Next i
    MsgBox "已追加 " & added & " 个文件（去重后）。", vbInformation
End Sub

' 1.3 批量选择源文件：选文件夹，仅当前文件夹（不递归），去重
Public Sub 批量选择源文件()
    Dim fd As FileDialog
    Dim ws As Worksheet
    Dim folder As String, fName As String
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "选择源文件所在文件夹"
    If fd.Show <> True Then Exit Sub
    folder = fd.SelectedItems(1)
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    Set ws = GetPanelWs()
    If ws Is Nothing Then
        MsgBox "请先运行「3.3 初始化配置」创建执行面板。", vbExclamation
        Exit Sub
    End If
    fName = Dir$(folder & "*.*")
    Do While fName <> ""
        If LCase$(Right$(fName, 4)) = ".xls" Or LCase$(Right$(fName, 5)) = ".xlsx" Or LCase$(Right$(fName, 5)) = ".xlsm" Then
            If Not PanelHasSourcePath(folder & fName) Then PanelAddSourcePath folder & fName
        End If
        fName = Dir$()
    Loop
    MsgBox "已从文件夹追加文件（去重后）。", vbInformation
End Sub

' 1.4 批量选择源文件(含子文件夹)：递归，去重
Public Sub 批量选择源文件含子文件夹()
    Dim fd As FileDialog
    Dim ws As Worksheet
    Dim folder As String
    Dim col As New Collection
    Dim i As Long, added As Long
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "选择源文件所在文件夹（将包含子文件夹）"
    If fd.Show <> True Then Exit Sub
    folder = fd.SelectedItems(1)
    If Right$(folder, 1) <> "\" Then folder = folder & "\"
    Set ws = GetPanelWs()
    If ws Is Nothing Then
        MsgBox "请先运行「3.3 初始化配置」创建执行面板。", vbExclamation
        Exit Sub
    End If
    CollectExcelFiles folder, col
    added = 0
    For i = 1 To col.count
        If Not PanelHasSourcePath(col(i)) Then
            PanelAddSourcePath col(i)
            added = added + 1
        End If
    Next i
    MsgBox "已追加 " & added & " 个文件（含子文件夹，去重后）。", vbInformation
End Sub

Private Sub CollectExcelFiles(ByVal folder As String, ByRef col As Collection)
    Dim fso As Object
    Dim f As Object
    Dim sf As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folder) Then Exit Sub
    On Error Resume Next
    For Each f In fso.GetFolder(folder).Files
        If LCase$(fso.GetExtensionName(f.Name)) = "xls" Or LCase$(fso.GetExtensionName(f.Name)) = "xlsx" Or LCase$(fso.GetExtensionName(f.Name)) = "xlsm" Then
            col.Add f.path
        End If
    Next f
    For Each sf In fso.GetFolder(folder).SubFolders
        CollectExcelFiles sf.path & "\", col
    Next sf
    On Error GoTo 0
End Sub
