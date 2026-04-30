# -*- coding: utf-8 -*-
"""
1dot6 提取工作表数据 + 1dot7 注入/汇总/清空 性能优化补丁

优化清单：
  [A] 1dot6  循环顺序倒置：源文件外层、配置行内层（减少 N_config 倍文件 Open/Close）
  [B] 1dot6  ExtractPartialData 逐格 Copy→数组整块读写
  [C] 1dot6  删除空白行列  逐行/列 Delete→Union 后一次删除
  [D] 1dot6  查找所有匹配工作表 去重 O(n²)→Dictionary O(1)
  [E] 1dot6  补加 Calculation=Manual / EnableEvents=False
  [F] 1dot7  汇总校验结果 逐格写结果→数组收集、循环后整块写入
  [G] 1dot7  汇总校验结果 ExtractComments 按 tmplWs.Name 缓存
  [H] 1dot7  清空校验区域 逐格 ClearContents→Union+ClearComments
  [I] 1dot7  注入/汇总/清空 三个公开 Sub 补加 Calculation=Manual / EnableEvents=False
"""
import os, sys
sys.stdout.reconfigure(encoding='utf-8')

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MOD  = os.path.join(BASE, 'VBA_Export', 'Modules')

def rp(path): return os.path.join(MOD, path)

def read_gbk(path):
    with open(path, 'r', encoding='gbk', newline='') as f:
        return f.read()

def write_utf8(path, content):
    with open(path, 'w', encoding='utf-8', newline='') as f:
        f.write(content)

def ok(label):  print(f'  [OK] {label}')
def warn(label): print(f'  [WARN] 未找到替换目标: {label}')

def rep(content, old, new, label):
    if old not in content:
        warn(label)
        return content
    ok(label)
    return content.replace(old, new, 1)


# ══════════════════════════════════════════════════════════
#  1dot6  提取工作表数据
# ══════════════════════════════════════════════════════════
def patch_1dot6_extract():
    path = rp('功能1dot6提取工作表数据.bas')
    c = read_gbk(path)

    # ── [E] 补 Calculation/EnableEvents ──────────────────
    c = rep(c,
        '        Application.ScreenUpdating = False\r\n'
        '        Application.DisplayAlerts = False\r\n'
        '        \r\n'
        '        \' 创建字典来管理输出文件',
        '        Application.ScreenUpdating = False\r\n'
        '        Application.DisplayAlerts = False\r\n'
        '        Application.Calculation = xlCalculationManual\r\n'
        '        Application.EnableEvents = False\r\n'
        '        \r\n'
        '        \' 创建字典来管理输出文件',
        '[E] 1dot6 补 Calculation/EnableEvents 关闭')

    # ── [A] 循环顺序倒置 ──────────────────────────────────
    OLD_LOOP = (
        '        \' 处理每个配置行\r\n'
        '        For i = 2 To lastRow\r\n'
        '            Dim shouldSkip As Boolean\r\n'
        '            Dim shouldFullExtract As Boolean\r\n'
        '            Dim sheetInfo(1 To 5) As String\r\n'
        '            Dim outputFileName As String\r\n'
        '            Dim extractRows As String\r\n'
        '            Dim extractCols As String\r\n'
        '            \r\n'
        '            \' 读取配置\r\n'
        '            With configSheet\r\n'
        '                \' 是否禁用 - F列\r\n'
        '                shouldSkip = (.Cells(i, "F").value = "是" Or .Cells(i, "F").value = "1" Or .Cells(i, "F").value = True)\r\n'
        '                \r\n'
        '                \' 如果禁用则跳过此配置\r\n'
        '                If Not shouldSkip Then\r\n'
        '                    \' 前五列确定sheet名\r\n'
        '                    sheetInfo(1) = CStr(.Cells(i, "A").value) \' 币种\r\n'
        '                    sheetInfo(2) = CStr(.Cells(i, "B").value) \' 地区\r\n'
        '                    sheetInfo(3) = CStr(.Cells(i, "C").value) \' 机构\r\n'
        '                    sheetInfo(4) = CStr(.Cells(i, "D").value) \' 类型\r\n'
        '                    sheetInfo(5) = CStr(.Cells(i, "E").value) \' 名称\r\n'
        '                    \r\n'
        '                    \' 是否整表提取 - G列\r\n'
        '                    shouldFullExtract = (.Cells(i, "G").value = "是" Or .Cells(i, "G").value = "1" Or .Cells(i, "G").value = True)\r\n'
        '                    \r\n'
        '                    \' 提取行列配置\r\n'
        '                    extractRows = CStr(.Cells(i, "H").value) \' 提取行\r\n'
        '                    extractCols = CStr(.Cells(i, "I").value) \' 提取列\r\n'
        '                    \r\n'
        '                    \' 输出文件名 - J列\r\n'
        '                    outputFileName = CStr(.Cells(i, "J").value)\r\n'
        '                    If outputFileName = "" Then\r\n'
        '                        outputFileName = "提取结果_" & Format(Now, "yyyymmdd_hhmmss")\r\n'
        '                    End If\r\n'
        '                    \r\n'
        '                    \' 在所有源文件中查找匹配的sheet\r\n'
        '                    For Each fileItem In fd.SelectedItems\r\n'
        '                        Dim sourceWb As Workbook\r\n'
        '                        Dim sourceWs As Worksheet\r\n'
        '                        Dim targetWs As Worksheet\r\n'
        '                        Dim targetWb As Workbook\r\n'
        '                        \r\n'
        '                        On Error Resume Next\r\n'
        '                        Set sourceWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True)\r\n'
        '                        If Err.Number <> 0 Then\r\n'
        '                            Debug.Print "无法打开文件: " & fileItem\r\n'
        '                            On Error GoTo 0\r\n'
        '                        Else\r\n'
        '                            On Error GoTo 0\r\n'
        '                            \r\n'
        '                            \' 查找所有匹配的工作表（修复：不只匹配第一个）\r\n'
        '                            Dim matchedSheets As Collection\r\n'
        '                            Set matchedSheets = 查找所有匹配工作表(sourceWb, sheetInfo)\r\n'
        '                            \r\n'
        '                            \' 处理每一个匹配的工作表\r\n'
        '                            If matchedSheets.count > 0 Then\r\n'
        '                                Dim wsItem As Variant\r\n'
        '                                \r\n'
        '                                For Each wsItem In matchedSheets\r\n'
        '                                    Set sourceWs = wsItem\r\n'
        '                                    \r\n'
        '                                    \' 获取或创建目标工作簿\r\n'
        '                                    If Not outputDict.Exists(outputFileName) Then\r\n'
        '                                        Set targetWb = Workbooks.Add\r\n'
        '                                        targetWb.SaveAs sourceWb.path & "\\" & outputFileName & ".xlsx"\r\n'
        '                                        outputDict.Add outputFileName, targetWb\r\n'
        '                                    Else\r\n'
        '                                        Set targetWb = outputDict(outputFileName)\r\n'
        '                                    End If\r\n'
        '                                    \r\n'
        '                                    \' 在目标工作簿中创建新工作表（使用源工作表的原名）\r\n'
        '                                    On Error Resume Next\r\n'
        '                                    Set targetWs = targetWb.Worksheets.Add(After:=targetWb.Worksheets(targetWb.Worksheets.count))\r\n'
        '                                    targetWs.Name = 获取唯一工作表名称(targetWb, sourceWs.Name)\r\n'
        '                                    On Error GoTo 0\r\n'
        '                                    \r\n'
        '                                    \' 提取数据\r\n'
        '                                    If shouldFullExtract Then\r\n'
        '                                        \' 整表提取\r\n'
        '                                        整表提取 sourceWs, targetWs\r\n'
        '                                    Else\r\n'
        '                                        \' 按行列提取\r\n'
        '                                        If Len(Trim$(extractRows)) > 0 Or Len(Trim$(extractCols)) > 0 Then\r\n'
        '                                            ExtractPartialData sourceWs, targetWs, extractRows, extractCols\r\n'
        '                                        Else\r\n'
        '                                            \' 如果没有指定行列，则整表提取\r\n'
        '                                            整表提取 sourceWs, targetWs\r\n'
        '                                        End If\r\n'
        '                                    End If\r\n'
        '                                    \r\n'
        '                                    \' 删除空白行列，使表格紧凑\r\n'
        '                                    删除空白行列 targetWs\r\n'
        '                                    \r\n'
        '                                    processedCount = processedCount + 1\r\n'
        '                                    Debug.Print "已提取: " & sourceWb.Name & " - " & sourceWs.Name & " -> " & outputFileName\r\n'
        '                                Next wsItem\r\n'
        '                            End If\r\n'
        '                            \r\n'
        '                            sourceWb.Close SaveChanges:=False\r\n'
        '                        End If\r\n'
        '                    Next fileItem\r\n'
        '                End If\r\n'
        '            End With\r\n'
        '        Next i'
    )
    NEW_LOOP = (
        '        \' [优化A] 预读所有配置行，避免每条配置都重复打开同一文件\r\n'
        '        Dim iCfg As Long\r\n'
        '        Dim cfgSkip() As Boolean\r\n'
        '        Dim cfgFullExtract() As Boolean\r\n'
        '        Dim cfgSheetInfo() As Variant\r\n'
        '        Dim cfgOutputFile() As String\r\n'
        '        Dim cfgRows() As String\r\n'
        '        Dim cfgCols() As String\r\n'
        '        ReDim cfgSkip(2 To lastRow)\r\n'
        '        ReDim cfgFullExtract(2 To lastRow)\r\n'
        '        ReDim cfgSheetInfo(2 To lastRow, 1 To 5)\r\n'
        '        ReDim cfgOutputFile(2 To lastRow)\r\n'
        '        ReDim cfgRows(2 To lastRow)\r\n'
        '        ReDim cfgCols(2 To lastRow)\r\n'
        '        For iCfg = 2 To lastRow\r\n'
        '            With configSheet\r\n'
        '                cfgSkip(iCfg) = (.Cells(iCfg, "F").value = "是" Or .Cells(iCfg, "F").value = "1" Or .Cells(iCfg, "F").value = True)\r\n'
        '                If Not cfgSkip(iCfg) Then\r\n'
        '                    cfgSheetInfo(iCfg, 1) = CStr(.Cells(iCfg, "A").value)\r\n'
        '                    cfgSheetInfo(iCfg, 2) = CStr(.Cells(iCfg, "B").value)\r\n'
        '                    cfgSheetInfo(iCfg, 3) = CStr(.Cells(iCfg, "C").value)\r\n'
        '                    cfgSheetInfo(iCfg, 4) = CStr(.Cells(iCfg, "D").value)\r\n'
        '                    cfgSheetInfo(iCfg, 5) = CStr(.Cells(iCfg, "E").value)\r\n'
        '                    cfgFullExtract(iCfg) = (.Cells(iCfg, "G").value = "是" Or .Cells(iCfg, "G").value = "1" Or .Cells(iCfg, "G").value = True)\r\n'
        '                    cfgRows(iCfg) = CStr(.Cells(iCfg, "H").value)\r\n'
        '                    cfgCols(iCfg) = CStr(.Cells(iCfg, "I").value)\r\n'
        '                    cfgOutputFile(iCfg) = CStr(.Cells(iCfg, "J").value)\r\n'
        '                    If cfgOutputFile(iCfg) = "" Then\r\n'
        '                        cfgOutputFile(iCfg) = "提取结果_" & Format(Now, "yyyymmdd_hhmmss")\r\n'
        '                    End If\r\n'
        '                End If\r\n'
        '            End With\r\n'
        '        Next iCfg\r\n'
        '        \r\n'
        '        \' 外层：每个源文件只打开一次\r\n'
        '        Dim sourceWb As Workbook\r\n'
        '        Dim sourceWs As Worksheet\r\n'
        '        Dim targetWs As Worksheet\r\n'
        '        Dim targetWb As Workbook\r\n'
        '        Dim matchedSheets As Collection\r\n'
        '        Dim wsItem As Variant\r\n'
        '        Dim sheetInfoArr(1 To 5) As String\r\n'
        '        \r\n'
        '        For Each fileItem In fd.SelectedItems\r\n'
        '            On Error Resume Next\r\n'
        '            Set sourceWb = Workbooks.Open(CStr(fileItem), ReadOnly:=True)\r\n'
        '            If Err.Number <> 0 Then\r\n'
        '                Debug.Print "无法打开文件: " & fileItem\r\n'
        '                On Error GoTo 0\r\n'
        '                GoTo NextFile\r\n'
        '            End If\r\n'
        '            On Error GoTo 0\r\n'
        '            \r\n'
        '            \' 内层：对该文件应用所有配置行\r\n'
        '            For iCfg = 2 To lastRow\r\n'
        '                If Not cfgSkip(iCfg) Then\r\n'
        '                    sheetInfoArr(1) = cfgSheetInfo(iCfg, 1)\r\n'
        '                    sheetInfoArr(2) = cfgSheetInfo(iCfg, 2)\r\n'
        '                    sheetInfoArr(3) = cfgSheetInfo(iCfg, 3)\r\n'
        '                    sheetInfoArr(4) = cfgSheetInfo(iCfg, 4)\r\n'
        '                    sheetInfoArr(5) = cfgSheetInfo(iCfg, 5)\r\n'
        '                    Set matchedSheets = 查找所有匹配工作表(sourceWb, sheetInfoArr)\r\n'
        '                    If matchedSheets.count > 0 Then\r\n'
        '                        For Each wsItem In matchedSheets\r\n'
        '                            Set sourceWs = wsItem\r\n'
        '                            If Not outputDict.Exists(cfgOutputFile(iCfg)) Then\r\n'
        '                                Set targetWb = Workbooks.Add\r\n'
        '                                targetWb.SaveAs sourceWb.path & "\\" & cfgOutputFile(iCfg) & ".xlsx"\r\n'
        '                                outputDict.Add cfgOutputFile(iCfg), targetWb\r\n'
        '                            Else\r\n'
        '                                Set targetWb = outputDict(cfgOutputFile(iCfg))\r\n'
        '                            End If\r\n'
        '                            On Error Resume Next\r\n'
        '                            Set targetWs = targetWb.Worksheets.Add(After:=targetWb.Worksheets(targetWb.Worksheets.count))\r\n'
        '                            targetWs.Name = 获取唯一工作表名称(targetWb, sourceWs.Name)\r\n'
        '                            On Error GoTo 0\r\n'
        '                            If cfgFullExtract(iCfg) Then\r\n'
        '                                整表提取 sourceWs, targetWs\r\n'
        '                            Else\r\n'
        '                                If Len(Trim$(cfgRows(iCfg))) > 0 Or Len(Trim$(cfgCols(iCfg))) > 0 Then\r\n'
        '                                    ExtractPartialData sourceWs, targetWs, cfgRows(iCfg), cfgCols(iCfg)\r\n'
        '                                Else\r\n'
        '                                    整表提取 sourceWs, targetWs\r\n'
        '                                End If\r\n'
        '                            End If\r\n'
        '                            删除空白行列 targetWs\r\n'
        '                            processedCount = processedCount + 1\r\n'
        '                            Debug.Print "已提取: " & sourceWb.Name & " - " & sourceWs.Name & " -> " & cfgOutputFile(iCfg)\r\n'
        '                        Next wsItem\r\n'
        '                    End If\r\n'
        '                End If\r\n'
        '            Next iCfg\r\n'
        '            \r\n'
        '            sourceWb.Close SaveChanges:=False\r\n'
        'NextFile:\r\n'
        '            Set sourceWb = Nothing\r\n'
        '        Next fileItem'
    )
    c = rep(c, OLD_LOOP, NEW_LOOP, '[A] 1dot6 循环顺序倒置')

    # ── [E] 恢复 Calculation/EnableEvents ─────────────────
    c = rep(c,
        '        Application.ScreenUpdating = True\r\n'
        '        Application.DisplayAlerts = True\r\n'
        '        \r\n'
        '        RunLog_WriteRow "1.6 提取工作表", "完成"',
        '        Application.Calculation = xlCalculationAutomatic\r\n'
        '        Application.EnableEvents = True\r\n'
        '        Application.ScreenUpdating = True\r\n'
        '        Application.DisplayAlerts = True\r\n'
        '        \r\n'
        '        RunLog_WriteRow "1.6 提取工作表", "完成"',
        '[E] 1dot6 恢复 Calculation/EnableEvents')

    # ── [B] ExtractPartialData 数组化 ─────────────────────
    OLD_PARTIAL = (
        '    targetRow = 1\r\n'
        '    For Each sourceRow In rowIndexes\r\n'
        '        targetCol = 1\r\n'
        '        For Each sourceCol In colIndexes\r\n'
        '            sourceWs.Cells(CLng(sourceRow), CLng(sourceCol)).Copy\r\n'
        '            targetWs.Cells(targetRow, targetCol).PasteSpecial Paste:=xlPasteAll\r\n'
        '            targetCol = targetCol + 1\r\n'
        '        Next sourceCol\r\n'
        '        targetRow = targetRow + 1\r\n'
        '    Next sourceRow\r\n'
        '\r\n'
        '    Application.CutCopyMode = False\r\n'
        'End Sub'
    )
    NEW_PARTIAL = (
        '    \' [优化B] 读入二维数组，整块写出，消除逐格剪贴板操作\r\n'
        '    Dim nR As Long, nC As Long, ri As Long, ci As Long\r\n'
        '    Dim dataArr() As Variant\r\n'
        '    nR = rowIndexes.Count\r\n'
        '    nC = colIndexes.Count\r\n'
        '    If nR = 0 Or nC = 0 Then Exit Sub\r\n'
        '    ReDim dataArr(1 To nR, 1 To nC)\r\n'
        '    ri = 0\r\n'
        '    For Each sourceRow In rowIndexes\r\n'
        '        ri = ri + 1\r\n'
        '        ci = 0\r\n'
        '        For Each sourceCol In colIndexes\r\n'
        '            ci = ci + 1\r\n'
        '            dataArr(ri, ci) = sourceWs.Cells(CLng(sourceRow), CLng(sourceCol)).Value\r\n'
        '        Next sourceCol\r\n'
        '    Next sourceRow\r\n'
        '    targetWs.Cells(1, 1).Resize(nR, nC).Value = dataArr\r\n'
        'End Sub'
    )
    c = rep(c, OLD_PARTIAL, NEW_PARTIAL, '[B] ExtractPartialData 数组化')

    # ── [C] 删除空白行列 Union 化 ─────────────────────────
    OLD_DELROWS = (
        '    \' 删除空行\r\n'
        '    For row = lastRow To 1 Step -1\r\n'
        '        hasData = False\r\n'
        '        For col = 1 To lastCol\r\n'
        '            If Len(Trim(CStr(ws.Cells(row, col).value))) > 0 Then\r\n'
        '                hasData = True\r\n'
        '                Exit For\r\n'
        '            End If\r\n'
        '        Next col\r\n'
        '        \r\n'
        '        If Not hasData Then\r\n'
        '            ws.Rows(row).Delete\r\n'
        '        End If\r\n'
        '    Next row'
    )
    NEW_DELROWS = (
        '    \' [优化C] 先收集空行，最后 Union 一次删除\r\n'
        '    Dim delRows As Range\r\n'
        '    For row = lastRow To 1 Step -1\r\n'
        '        hasData = False\r\n'
        '        For col = 1 To lastCol\r\n'
        '            If Len(Trim(CStr(ws.Cells(row, col).Value))) > 0 Then\r\n'
        '                hasData = True\r\n'
        '                Exit For\r\n'
        '            End If\r\n'
        '        Next col\r\n'
        '        If Not hasData Then\r\n'
        '            If delRows Is Nothing Then\r\n'
        '                Set delRows = ws.Rows(row)\r\n'
        '            Else\r\n'
        '                Set delRows = Union(delRows, ws.Rows(row))\r\n'
        '            End If\r\n'
        '        End If\r\n'
        '    Next row\r\n'
        '    If Not delRows Is Nothing Then delRows.Delete Shift:=xlUp\r\n'
        '    Set delRows = Nothing'
    )
    c = rep(c, OLD_DELROWS, NEW_DELROWS, '[C] 删除空行 Union 化')

    OLD_DELCOLS = (
        '    \' 删除空列\r\n'
        '    For col = lastCol To 1 Step -1\r\n'
        '        hasData = False\r\n'
        '        For row = 1 To lastRow\r\n'
        '            If Len(Trim(CStr(ws.Cells(row, col).value))) > 0 Then\r\n'
        '                hasData = True\r\n'
        '                Exit For\r\n'
        '            End If\r\n'
        '        Next row\r\n'
        '        \r\n'
        '        If Not hasData Then\r\n'
        '            ws.Columns(col).Delete\r\n'
        '        End If\r\n'
        '    Next col'
    )
    NEW_DELCOLS = (
        '    \' [优化C] 先收集空列，最后 Union 一次删除\r\n'
        '    Dim delCols As Range\r\n'
        '    For col = lastCol To 1 Step -1\r\n'
        '        hasData = False\r\n'
        '        For row = 1 To lastRow\r\n'
        '            If Len(Trim(CStr(ws.Cells(row, col).Value))) > 0 Then\r\n'
        '                hasData = True\r\n'
        '                Exit For\r\n'
        '            End If\r\n'
        '        Next row\r\n'
        '        If Not hasData Then\r\n'
        '            If delCols Is Nothing Then\r\n'
        '                Set delCols = ws.Columns(col)\r\n'
        '            Else\r\n'
        '                Set delCols = Union(delCols, ws.Columns(col))\r\n'
        '            End If\r\n'
        '        End If\r\n'
        '    Next col\r\n'
        '    If Not delCols Is Nothing Then delCols.Delete Shift:=xlToLeft\r\n'
        '    Set delCols = Nothing'
    )
    c = rep(c, OLD_DELCOLS, NEW_DELCOLS, '[C] 删除空列 Union 化')

    # ── [D] 查找所有匹配工作表 去重 O(n²)→Dict ────────────
    OLD_DEDUP = (
        'Function 查找所有匹配工作表(wb As Workbook, sheetInfo As Variant) As Collection\r\n'
        '    Dim ws As Worksheet\r\n'
        '    Dim targetName As String\r\n'
        '    Dim matchedSheets As Collection\r\n'
        '    Set matchedSheets = New Collection\r\n'
        '    \r\n'
        '    \' 构建目标工作表名称（用于精确匹配）\r\n'
        '    targetName = 构建工作表名称(sheetInfo)\r\n'
        '    \r\n'
        '    \' 先尝试精确匹配\r\n'
        '    On Error Resume Next\r\n'
        '    Set ws = wb.Worksheets(targetName)\r\n'
        '    If Not ws Is Nothing Then\r\n'
        '        matchedSheets.Add ws\r\n'
        '    End If\r\n'
        '    On Error GoTo 0\r\n'
        '    \r\n'
        '    \' 然后进行模糊匹配\r\n'
        '    For Each ws In wb.Worksheets\r\n'
        '        \' 检查是否包含所有关键词，并且不是已经添加的工作表\r\n'
        '        If 工作表包含所有关键词(ws, sheetInfo) Then\r\n'
        '            \' 检查是否已经添加过这个工作表\r\n'
        '            Dim alreadyAdded As Boolean\r\n'
        '            alreadyAdded = False\r\n'
        '            \r\n'
        '            Dim existingWs As Variant\r\n'
        '            For Each existingWs In matchedSheets\r\n'
        '                If existingWs.Name = ws.Name Then\r\n'
        '                    alreadyAdded = True\r\n'
        '                    Exit For\r\n'
        '                End If\r\n'
        '            Next existingWs\r\n'
        '            \r\n'
        '            If Not alreadyAdded Then\r\n'
        '                matchedSheets.Add ws\r\n'
        '            End If\r\n'
        '        End If\r\n'
        '    Next ws\r\n'
        '    \r\n'
        '    Set 查找所有匹配工作表 = matchedSheets\r\n'
        'End Function'
    )
    NEW_DEDUP = (
        'Function 查找所有匹配工作表(wb As Workbook, sheetInfo As Variant) As Collection\r\n'
        '    Dim ws As Worksheet\r\n'
        '    Dim targetName As String\r\n'
        '    Dim matchedSheets As Collection\r\n'
        '    Dim addedNames As Object\r\n'
        '    Set matchedSheets = New Collection\r\n'
        '    Set addedNames = CreateObject("Scripting.Dictionary")\r\n'
        '    addedNames.CompareMode = vbTextCompare\r\n'
        '    \r\n'
        '    \' 构建目标工作表名称（用于精确匹配）\r\n'
        '    targetName = 构建工作表名称(sheetInfo)\r\n'
        '    \r\n'
        '    \' 先尝试精确匹配\r\n'
        '    On Error Resume Next\r\n'
        '    Set ws = wb.Worksheets(targetName)\r\n'
        '    If Not ws Is Nothing Then\r\n'
        '        matchedSheets.Add ws\r\n'
        '        addedNames(ws.Name) = True\r\n'
        '    End If\r\n'
        '    On Error GoTo 0\r\n'
        '    \r\n'
        '    \' 模糊匹配，用字典去重（O(1) 替代 O(n) 遍历）\r\n'
        '    For Each ws In wb.Worksheets\r\n'
        '        If 工作表包含所有关键词(ws, sheetInfo) Then\r\n'
        '            If Not addedNames.Exists(ws.Name) Then\r\n'
        '                matchedSheets.Add ws\r\n'
        '                addedNames(ws.Name) = True\r\n'
        '            End If\r\n'
        '        End If\r\n'
        '    Next ws\r\n'
        '    \r\n'
        '    Set 查找所有匹配工作表 = matchedSheets\r\n'
        'End Function'
    )
    c = rep(c, OLD_DEDUP, NEW_DEDUP, '[D] 查找所有匹配工作表 去重 Dict 化')

    write_utf8(path, c)
    print('  => 1dot6提取工作表数据 已写回')


# ══════════════════════════════════════════════════════════
#  1dot7  注入/汇总/清空
# ══════════════════════════════════════════════════════════
def patch_1dot7():
    path = rp('功能1dot7注入校验区域与汇总校验结果.bas')
    c = read_gbk(path)

    # ── [I] 注入校验区域 补 Calculation/EnableEvents ───────
    c = rep(c,
        '    Application.ScreenUpdating = False\r\n'
        '    Application.DisplayAlerts = False\r\n'
        '    On Error GoTo ErrInject',
        '    Application.ScreenUpdating = False\r\n'
        '    Application.DisplayAlerts = False\r\n'
        '    Application.Calculation = xlCalculationManual\r\n'
        '    Application.EnableEvents = False\r\n'
        '    On Error GoTo ErrInject',
        '[I] 注入校验区域 关闭 Calculation/EnableEvents')

    # 注入正常结束恢复
    c = rep(c,
        '    tmplWb.Close SaveChanges:=False\r\n'
        '    Set tmplWb = Nothing\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '\r\n'
        '    RunLog_WriteRow logKey, "完成", "", "", "", "", "成功 " & okCnt',
        '    tmplWb.Close SaveChanges:=False\r\n'
        '    Set tmplWb = Nothing\r\n'
        '    Application.Calculation = xlCalculationAutomatic\r\n'
        '    Application.EnableEvents = True\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '\r\n'
        '    RunLog_WriteRow logKey, "完成", "", "", "", "", "成功 " & okCnt',
        '[I] 注入校验区域 恢复 Calculation/EnableEvents（正常路径）')

    # 注入错误处理恢复
    c = rep(c,
        'ErrInject:\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '    On Error Resume Next\r\n'
        '    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False\r\n'
        '    If Not tmplWb Is Nothing Then tmplWb.Close SaveChanges:=False',
        'ErrInject:\r\n'
        '    Application.Calculation = xlCalculationAutomatic\r\n'
        '    Application.EnableEvents = True\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '    On Error Resume Next\r\n'
        '    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False\r\n'
        '    If Not tmplWb Is Nothing Then tmplWb.Close SaveChanges:=False',
        '[I] 注入校验区域 恢复 Calculation/EnableEvents（错误路径）')

    # ── [I][F][G] 汇总校验结果 整体重写 ───────────────────
    OLD_SUM_SETUP = (
        '    Set wsRes = EnsureResultSheet()\r\n'
        '    InitResultHeader wsRes\r\n'
        '    outRow = 2\r\n'
        '\r\n'
        '    Application.ScreenUpdating = False\r\n'
        '    Application.DisplayAlerts = False\r\n'
        '    On Error GoTo ErrSum\r\n'
        '\r\n'
        '    Set tmplWb = Workbooks.Open(tmplPath, ReadOnly:=True, UpdateLinks:=0)'
    )
    NEW_SUM_SETUP = (
        '    Set wsRes = EnsureResultSheet()\r\n'
        '    InitResultHeader wsRes\r\n'
        '    outRow = 2\r\n'
        '\r\n'
        '    Application.ScreenUpdating = False\r\n'
        '    Application.DisplayAlerts = False\r\n'
        '    Application.Calculation = xlCalculationManual\r\n'
        '    Application.EnableEvents = False\r\n'
        '    On Error GoTo ErrSum\r\n'
        '\r\n'
        '    \' [优化F] 预分配结果数组，避免逐格写入工作表\r\n'
        '    Const MAX_RESULT As Long = 50000\r\n'
        '    Dim resArr() As Variant\r\n'
        '    ReDim resArr(1 To MAX_RESULT, 1 To 13)\r\n'
        '    Dim arrIdx As Long\r\n'
        '    arrIdx = 0\r\n'
        '\r\n'
        '    \' [优化G] 按模板表名缓存批注字典，避免重复提取\r\n'
        '    Dim atcCache As Object\r\n'
        '    Set atcCache = CreateObject("Scripting.Dictionary")\r\n'
        '    atcCache.CompareMode = vbTextCompare\r\n'
        '\r\n'
        '    Set tmplWb = Workbooks.Open(tmplPath, ReadOnly:=True, UpdateLinks:=0)'
    )
    c = rep(c, OLD_SUM_SETUP, NEW_SUM_SETUP, '[I][F][G] 汇总校验结果 变量初始化')

    # ExtractComments 缓存化 + 内层写入改为写数组
    OLD_SUM_INNER = (
        '            ExtractComments tmplWs, atc\r\n'
        '            Set regsLocal = ExtractRegions(atc, TMPL_KEY_LOCAL)\r\n'
        '            Set regsCross = ExtractRegions(atc, TMPL_KEY_CROSS)\r\n'
        '            Set regs = MergeRegions(regsLocal, regsCross)\r\n'
        '            If regs.Count = 0 Then GoTo NextSheet2\r\n'
        '\r\n'
        '            For Each rg In regs\r\n'
        '                For r = CLng(rg(0)) To CLng(rg(1))\r\n'
        '                    rowName = CStr(srcWs.Cells(r, 1).Value)\r\n'
        '                    For c = CLng(rg(2)) To CLng(rg(3))\r\n'
        '                        cellVal = CStr(srcWs.Cells(r, c).Value)\r\n'
        '                        If ContainsAnyKey(cellVal) Then\r\n'
        '                            ParsePipe5 cellVal, errType, r1, r2, r3, r4\r\n'
        '                            cmTxt = CellCommentText(srcWs.Cells(r, c))\r\n'
        '                            wsRes.Cells(outRow, rcTmplWb).Value = tmplWb.Name\r\n'
        '                            wsRes.Cells(outRow, rcTmplWs).Value = tmplWs.Name\r\n'
        '                            wsRes.Cells(outRow, rcSrcWb).Value = srcWb.Name\r\n'
        '                            wsRes.Cells(outRow, rcSrcWs).Value = srcWs.Name\r\n'
        '                            wsRes.Cells(outRow, rcRow).Value = r\r\n'
        '                            wsRes.Cells(outRow, rcCol).Value = c\r\n'
        '                            wsRes.Cells(outRow, rcRowName).Value = rowName\r\n'
        '                            wsRes.Cells(outRow, rcErrType).Value = errType\r\n'
        '                            wsRes.Cells(outRow, rcRsv1).Value = r1\r\n'
        '                            wsRes.Cells(outRow, rcRsv2).Value = r2\r\n'
        '                            wsRes.Cells(outRow, rcRsv3).Value = r3\r\n'
        '                            wsRes.Cells(outRow, rcRsv4).Value = r4\r\n'
        '                            wsRes.Cells(outRow, rcComment).Value = cmTxt\r\n'
        '                            outRow = outRow + 1\r\n'
        '                            hitCnt = hitCnt + 1\r\n'
        '                        End If\r\n'
        '                    Next c\r\n'
        '                Next r\r\n'
        '            Next rg'
    )
    NEW_SUM_INNER = (
        '            \' [优化G] 使用缓存，同一模板表不重复提取批注\r\n'
        '            If Not atcCache.Exists(tmplWs.Name) Then\r\n'
        '                ExtractComments tmplWs, atc\r\n'
        '                Set atcCache(tmplWs.Name) = atc\r\n'
        '            Else\r\n'
        '                Set atc = atcCache(tmplWs.Name)\r\n'
        '            End If\r\n'
        '            Set regsLocal = ExtractRegions(atc, TMPL_KEY_LOCAL)\r\n'
        '            Set regsCross = ExtractRegions(atc, TMPL_KEY_CROSS)\r\n'
        '            Set regs = MergeRegions(regsLocal, regsCross)\r\n'
        '            If regs.Count = 0 Then GoTo NextSheet2\r\n'
        '\r\n'
        '            For Each rg In regs\r\n'
        '                For r = CLng(rg(0)) To CLng(rg(1))\r\n'
        '                    rowName = CStr(srcWs.Cells(r, 1).Value)\r\n'
        '                    For c = CLng(rg(2)) To CLng(rg(3))\r\n'
        '                        cellVal = CStr(srcWs.Cells(r, c).Value)\r\n'
        '                        If ContainsAnyKey(cellVal) Then\r\n'
        '                            ParsePipe5 cellVal, errType, r1, r2, r3, r4\r\n'
        '                            cmTxt = CellCommentText(srcWs.Cells(r, c))\r\n'
        '                            \' [优化F] 写入数组而非逐格写工作表\r\n'
        '                            arrIdx = arrIdx + 1\r\n'
        '                            If arrIdx <= MAX_RESULT Then\r\n'
        '                                resArr(arrIdx, rcTmplWb) = tmplWb.Name\r\n'
        '                                resArr(arrIdx, rcTmplWs) = tmplWs.Name\r\n'
        '                                resArr(arrIdx, rcSrcWb) = srcWb.Name\r\n'
        '                                resArr(arrIdx, rcSrcWs) = srcWs.Name\r\n'
        '                                resArr(arrIdx, rcRow) = r\r\n'
        '                                resArr(arrIdx, rcCol) = c\r\n'
        '                                resArr(arrIdx, rcRowName) = rowName\r\n'
        '                                resArr(arrIdx, rcErrType) = errType\r\n'
        '                                resArr(arrIdx, rcRsv1) = r1\r\n'
        '                                resArr(arrIdx, rcRsv2) = r2\r\n'
        '                                resArr(arrIdx, rcRsv3) = r3\r\n'
        '                                resArr(arrIdx, rcRsv4) = r4\r\n'
        '                                resArr(arrIdx, rcComment) = cmTxt\r\n'
        '                            End If\r\n'
        '                            hitCnt = hitCnt + 1\r\n'
        '                        End If\r\n'
        '                    Next c\r\n'
        '                Next r\r\n'
        '            Next rg'
    )
    c = rep(c, OLD_SUM_INNER, NEW_SUM_INNER, '[F][G] 汇总 ExtractComments 缓存 + 数组写入')

    # 汇总 循环后整块写入 + 恢复设置
    OLD_SUM_WRITE = (
        '    tmplWb.Close SaveChanges:=False\r\n'
        '    Set tmplWb = Nothing\r\n'
        '    wsRes.Columns.AutoFit\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '\r\n'
        '    RunLog_WriteRow logKey, "完成", "", "", "", "", "命中 " & hitCnt & " 条"'
    )
    NEW_SUM_WRITE = (
        '    tmplWb.Close SaveChanges:=False\r\n'
        '    Set tmplWb = Nothing\r\n'
        '    \' [优化F] 一次性写入结果数组\r\n'
        '    If arrIdx > 0 Then\r\n'
        '        Dim writeCount As Long\r\n'
        '        writeCount = IIf(arrIdx > MAX_RESULT, MAX_RESULT, arrIdx)\r\n'
        '        wsRes.Cells(2, 1).Resize(writeCount, 13).Value = resArr\r\n'
        '    End If\r\n'
        '    wsRes.Columns.AutoFit\r\n'
        '    Application.Calculation = xlCalculationAutomatic\r\n'
        '    Application.EnableEvents = True\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '\r\n'
        '    RunLog_WriteRow logKey, "完成", "", "", "", "", "命中 " & hitCnt & " 条"'
    )
    c = rep(c, OLD_SUM_WRITE, NEW_SUM_WRITE, '[F] 汇总 整块写入 + 恢复设置')

    # 汇总错误处理恢复
    c = rep(c,
        'ErrSum:\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '    On Error Resume Next\r\n'
        '    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False\r\n'
        '    If Not tmplWb Is Nothing Then tmplWb.Close SaveChanges:=False\r\n'
        '    On Error GoTo 0\r\n'
        '    RunLog_WriteRow logKey, "完成", "", "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))\r\n'
        '    MsgBox "发生错误 " & Err.Number & ": " & Err.Description, vbCritical\r\n'
        'End Sub\r\n'
        '\r\n'
        'Public Sub 清空校验区域()',
        'ErrSum:\r\n'
        '    Application.Calculation = xlCalculationAutomatic\r\n'
        '    Application.EnableEvents = True\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '    On Error Resume Next\r\n'
        '    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False\r\n'
        '    If Not tmplWb Is Nothing Then tmplWb.Close SaveChanges:=False\r\n'
        '    On Error GoTo 0\r\n'
        '    RunLog_WriteRow logKey, "完成", "", "", "", "失败", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))\r\n'
        '    MsgBox "发生错误 " & Err.Number & ": " & Err.Description, vbCritical\r\n'
        'End Sub\r\n'
        '\r\n'
        'Public Sub 清空校验区域()',
        '[I] 汇总 错误路径恢复 Calculation/EnableEvents')

    # ── [H][I] 清空校验区域 ────────────────────────────────
    c = rep(c,
        '    Application.ScreenUpdating = False\r\n'
        '    Application.DisplayAlerts = False\r\n'
        '    On Error GoTo ErrClear',
        '    Application.ScreenUpdating = False\r\n'
        '    Application.DisplayAlerts = False\r\n'
        '    Application.Calculation = xlCalculationManual\r\n'
        '    Application.EnableEvents = False\r\n'
        '    On Error GoTo ErrClear',
        '[I] 清空校验区域 关闭 Calculation/EnableEvents')

    OLD_CLEAR_INNER = (
        '            For Each rg In regs\r\n'
        '                For r = CLng(rg(0)) To CLng(rg(1))\r\n'
        '                    For c = CLng(rg(2)) To CLng(rg(3))\r\n'
        '                        srcWs.Cells(r, c).ClearContents\r\n'
        '                        On Error Resume Next\r\n'
        '                        If Not srcWs.Cells(r, c).Comment Is Nothing Then srcWs.Cells(r, c).Comment.Delete\r\n'
        '                        On Error GoTo 0\r\n'
        '                    Next c\r\n'
        '                Next r\r\n'
        '            Next rg'
    )
    NEW_CLEAR_INNER = (
        '            \' [优化H] 先 Union 所有区域，再一次 ClearContents + ClearComments\r\n'
        '            Dim clearRng As Range\r\n'
        '            Dim rgBlock As Range\r\n'
        '            For Each rg In regs\r\n'
        '                Set rgBlock = srcWs.Range(srcWs.Cells(CLng(rg(0)), CLng(rg(2))), srcWs.Cells(CLng(rg(1)), CLng(rg(3))))\r\n'
        '                If clearRng Is Nothing Then\r\n'
        '                    Set clearRng = rgBlock\r\n'
        '                Else\r\n'
        '                    Set clearRng = Union(clearRng, rgBlock)\r\n'
        '                End If\r\n'
        '            Next rg\r\n'
        '            If Not clearRng Is Nothing Then\r\n'
        '                clearRng.ClearContents\r\n'
        '                On Error Resume Next\r\n'
        '                clearRng.ClearComments\r\n'
        '                On Error GoTo 0\r\n'
        '                Set clearRng = Nothing\r\n'
        '            End If'
    )
    c = rep(c, OLD_CLEAR_INNER, NEW_CLEAR_INNER, '[H] 清空校验区域 Union+ClearComments')

    c = rep(c,
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '    RunLog_WriteRow logKey, "完成", "", "", "", "", "完成"',
        '    Application.Calculation = xlCalculationAutomatic\r\n'
        '    Application.EnableEvents = True\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '    RunLog_WriteRow logKey, "完成", "", "", "", "", "完成"',
        '[I] 清空校验区域 恢复 Calculation/EnableEvents（正常路径）')

    c = rep(c,
        'ErrClear:\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '    On Error Resume Next\r\n'
        '    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False',
        'ErrClear:\r\n'
        '    Application.Calculation = xlCalculationAutomatic\r\n'
        '    Application.EnableEvents = True\r\n'
        '    Application.DisplayAlerts = True\r\n'
        '    Application.ScreenUpdating = True\r\n'
        '    On Error Resume Next\r\n'
        '    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False',
        '[I] 清空校验区域 恢复 Calculation/EnableEvents（错误路径）')

    write_utf8(path, c)
    print('  => 1dot7注入校验区域与汇总校验结果 已写回')


if __name__ == '__main__':
    print('=== patch 1dot6 提取工作表数据 ===')
    patch_1dot6_extract()
    print()
    print('=== patch 1dot7 注入/汇总/清空 ===')
    patch_1dot7()
    print()
    print('全部完成，请运行 python convert.py 转换为 GBK。')
