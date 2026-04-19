Attribute VB_Name = "功能3dot9追加列按批注"
Option Explicit

' 追加列（按批注）：根据模板各 Sheet 第一行批注含「追加列」的单元格形成 {工作表名: 列字母} 字典；
' 关键字仅匹配工作簿名；源工作簿内遍历所有工作表，若工作表名在字典中则取对应列；在外部文件中按配置的关键字找到对应工作表并往该表最右列追加，表头为 工作簿名_工作表名_列字母。
' 执行面板：A2=模板路径，B2=外部文件路径，B5 起=源文件路径。

Private Const CFG_KEY As String = "3.9 追加列（按批注）"
Private Const PANEL_SHEET As String = "执行面板"
Private Const PANEL_ROW_TMPL As Long = 2
Private Const PANEL_ROW_EXT As Long = 2
Private Const PANEL_COL_TMPL As Long = 1
Private Const PANEL_COL_EXT As Long = 2
Private Const PANEL_DATA_START As Long = 5
Private Const PANEL_COL_PATH As Long = 2
Private Const CONFIG_SHEET As String = "config"
Private Const COMMENT_KW_DEFAULT As String = "追加列"

Public Sub 追加列按批注()
    Dim wsPanel As Worksheet
    Dim tmplPath As String, extPath As String
    Dim pathList As Collection
    Dim r As Long
    Dim tmplWb As Workbook, extWb As Workbook, srcWb As Workbook
    Dim extWs As Worksheet, srcWs As Worksheet
    Dim dictSheetCol As Object
    Dim keywordsStr As String, keywords() As String
    Dim commentKw As String
    Dim wbName As String, fileNameNoExt As String
    Dim nextCol As Long
    Dim kw As Variant
    Dim matchedKeyword As String
    Dim colLetter As String
    Dim colNum As Long
    Dim lastRowSrc As Long
    Dim destRow As Long
    Dim headerText As String
    Dim fd As FileDialog
    Dim t0 As Double
    Dim i As Long

    t0 = Timer
    On Error GoTo ErrHandler

    Set wsPanel = ThisWorkbook.Worksheets(PANEL_SHEET)
    If wsPanel Is Nothing Then
        MsgBox "未找到「执行面板」。请先运行「4.4 初始化执行面板」。", vbExclamation
        Exit Sub
    End If

    tmplPath = Trim(CStr(wsPanel.Cells(PANEL_ROW_TMPL, PANEL_COL_TMPL).value))
    If tmplPath = "" Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        With fd
            .Title = "选择模板文件（表头第一行批注含「追加列」的列将参与字典）"
            .AllowMultiSelect = False
            .Filters.Clear
            .Filters.Add "Excel 文件", "*.xls;*.xlsx;*.xlsm"
        End With
        If fd.Show <> -1 Then Exit Sub
        tmplPath = fd.SelectedItems(1)
        wsPanel.Cells(PANEL_ROW_TMPL, PANEL_COL_TMPL).value = tmplPath
    End If

    extPath = Trim(CStr(wsPanel.Cells(PANEL_ROW_EXT, PANEL_COL_EXT).value))
    If extPath = "" Then
        Set fd = Application.FileDialog(msoFileDialogFilePicker)
        fd.Title = "选择外部文件（列将追加到该文件最右侧）"
        fd.AllowMultiSelect = False
        fd.Filters.Clear
        fd.Filters.Add "Excel 文件", "*.xls;*.xlsx;*.xlsm"
        If fd.Show <> -1 Then Exit Sub
        extPath = fd.SelectedItems(1)
        wsPanel.Cells(PANEL_ROW_EXT, PANEL_COL_EXT).value = extPath
    End If

    Set pathList = New Collection
    For r = PANEL_DATA_START To wsPanel.Cells(wsPanel.Rows.count, PANEL_COL_PATH).End(xlUp).row
        If Len(Trim(CStr(wsPanel.Cells(r, PANEL_COL_PATH).value))) > 0 Then
            pathList.Add Trim(CStr(wsPanel.Cells(r, PANEL_COL_PATH).value))
        End If
    Next r
    If pathList.count = 0 Then
        MsgBox "执行面板 B5 起无有效路径。请先填写源文件路径或运行 1. 选择文件。", vbExclamation
        Exit Sub
    End If

    keywordsStr = 读取配置值(CFG_KEY, "关键字")
    If Trim(keywordsStr) = "" Then
        MsgBox "请在 config 表配置「" & CFG_KEY & "」的「关键字」（如 sheet1;sheet2），源文件名包含某关键字则追加到外部文件对应工作表。", vbExclamation
        Exit Sub
    End If
    keywords = Split(keywordsStr, ";")

    commentKw = 读取配置值(CFG_KEY, "追加列批注")
    If Trim(commentKw) = "" Then commentKw = COMMENT_KW_DEFAULT

    Set dictSheetCol = BuildSheetColDictFromTemplate(tmplPath, commentKw)
    If dictSheetCol Is Nothing Or dictSheetCol.count = 0 Then
        MsgBox "模板中未找到任何表头第一行批注含「" & commentKw & "」的列。请检查模板文件。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo OpenExtErr
    Set extWb = Workbooks.Open(extPath, ReadOnly:=False)
    On Error GoTo ErrHandler

    For i = 1 To pathList.count
        On Error GoTo OpenSrcErr
        Set srcWb = Workbooks.Open(CStr(pathList(i)), ReadOnly:=True)
        On Error GoTo ErrHandler

        wbName = srcWb.Name
        fileNameNoExt = wbName
        If InStrRev(fileNameNoExt, ".") > 0 Then fileNameNoExt = Left(fileNameNoExt, InStrRev(fileNameNoExt, ".") - 1)

        matchedKeyword = ""
        For Each kw In keywords
            If Len(Trim(CStr(kw))) > 0 And InStr(1, fileNameNoExt, Trim(CStr(kw)), vbTextCompare) > 0 Then
                matchedKeyword = Trim(CStr(kw))
                Exit For
            End If
        Next kw

        If matchedKeyword = "" Then
            srcWb.Close SaveChanges:=False
            Set srcWb = Nothing
            GoTo NextFile
        End If

        Set extWs = 确保工作表(extWb, matchedKeyword)
        nextCol = 最后数据列(extWs) + 1
        If nextCol <= 1 Then nextCol = 1

        For Each srcWs In srcWb.Worksheets
            If Not dictSheetCol.Exists(srcWs.Name) Then GoTo NextSrcSheet
            colLetter = dictSheetCol(srcWs.Name)
            colNum = 列字母转号(colLetter)
            If colNum <= 0 Then GoTo NextSrcSheet

            lastRowSrc = 最后数据行(srcWs, colNum)
            If lastRowSrc < 1 Then GoTo NextSrcSheet

            headerText = wbName & "_" & srcWs.Name & "_" & colLetter & "列"
            extWs.Cells(1, nextCol).value = headerText
            extWs.Cells(1, nextCol).Font.Bold = True

            CopySourceColumnToTarget srcWs, colNum, lastRowSrc, extWs, nextCol

            nextCol = nextCol + 1
NextSrcSheet:
        Next srcWs

        srcWb.Close SaveChanges:=False
        Set srcWb = Nothing
NextFile:
    Next i

    extWb.Save
    extWb.Close SaveChanges:=True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "追加列完成。", vbInformation
    Exit Sub

OpenExtErr:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "打开外部文件失败：" & vbCrLf & extPath & vbCrLf & Err.Description, vbExclamation
    Exit Sub
OpenSrcErr:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    MsgBox "打开源文件失败：" & vbCrLf & pathList(i) & vbCrLf & Err.Description, vbExclamation
    Exit Sub
ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    If Not extWb Is Nothing Then extWb.Close SaveChanges:=False
    MsgBox "执行出错：" & vbCrLf & Err.Description, vbCritical
End Sub

' 从模板工作簿各 Sheet 第一行中，找出批注文本包含 commentKeyword 的单元格，形成 Dictionary(sheetName -> columnLetter)。
' 字典键必须为工作表名 ws.Name，不可选用其他键；commentKeyword 为必填参数。
Private Function BuildSheetColDictFromTemplate(ByVal tmplPath As String, ByVal commentKeyword As String) As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim c As Range
    Dim colLetter As String
    Dim commentText As String

    If Len(Trim(commentKeyword)) = 0 Then
        Set BuildSheetColDictFromTemplate = Nothing
        Exit Function
    End If
    Set BuildSheetColDictFromTemplate = CreateObject("Scripting.Dictionary")
    BuildSheetColDictFromTemplate.CompareMode = 1

    On Error GoTo ErrBuild
    Set wb = Workbooks.Open(tmplPath, ReadOnly:=True)
    On Error GoTo 0

    For Each ws In wb.Worksheets
        On Error Resume Next
        For Each c In ws.Range(ws.Cells(1, 1), ws.Cells(1, 256))
            If Not c.Comment Is Nothing Then
                commentText = c.Comment.Text
                If InStr(1, commentText, commentKeyword, vbTextCompare) > 0 Then
                    If Not BuildSheetColDictFromTemplate.Exists(ws.Name) Then
                        colLetter = 列号转字母(c.Column)
                        BuildSheetColDictFromTemplate(ws.Name) = colLetter
                    End If
                    Exit For
                End If
            End If
        Next c
        On Error GoTo 0
    Next ws

    wb.Close SaveChanges:=False
    Set wb = Nothing
    Exit Function
ErrBuild:
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Set BuildSheetColDictFromTemplate = Nothing
End Function

Private Function 确保工作表(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = sheetName
    End If
    Set 确保工作表 = ws
End Function

Private Function 列号转字母(ByVal colNum As Long) As String
    Dim n As Long
    列号转字母 = ""
    If colNum < 1 Then Exit Function
    Do While colNum > 0
        n = (colNum - 1) Mod 26
        列号转字母 = Chr(65 + n) & 列号转字母
        colNum = (colNum - 1) \ 26
    Loop
End Function

Private Function 列字母转号(ByVal colLetter As String) As Long
    Dim i As Long
    Dim ch As String
    列字母转号 = 0
    colLetter = UCase(Trim(colLetter))
    For i = 1 To Len(colLetter)
        ch = Mid(colLetter, i, 1)
        If ch >= "A" And ch <= "Z" Then
            列字母转号 = 列字母转号 * 26 + (Asc(ch) - 64)
        End If
    Next i
End Function

Private Function 最后数据列(ByVal ws As Worksheet) As Long
    Dim r As Long
    最后数据列 = 0
    On Error Resume Next
    r = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    If r >= 1 Then 最后数据列 = r
    On Error GoTo 0
End Function

Private Function 最后数据行(ByVal ws As Worksheet, ByVal colNum As Long) As Long
    Dim r As Long
    最后数据行 = 0
    If colNum < 1 Then Exit Function
    On Error Resume Next
    r = ws.Cells(ws.Rows.count, colNum).End(xlUp).row
    If r >= 1 Then 最后数据行 = r
    On Error GoTo 0
End Function

Private Function 单元格值(ByVal c As Range) As Variant
    If c Is Nothing Then 单元格值 = Empty: Exit Function
    On Error Resume Next
    If c.MergeCells Then 单元格值 = c.MergeArea.Cells(1, 1).value Else 单元格值 = c.value
    On Error GoTo 0
End Function

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

Private Sub CopySourceColumnToTarget(ByVal srcWs As Worksheet, ByVal srcCol As Long, ByVal srcLastRow As Long, ByVal tgtWs As Worksheet, ByVal tgtCol As Long)
    Dim srcRange As Range
    Dim tgtRange As Range
    Dim srcArr As Variant
    Dim r As Long

    If srcLastRow < 2 Then Exit Sub

    Set srcRange = srcWs.Range(srcWs.Cells(2, srcCol), srcWs.Cells(srcLastRow, srcCol))
    Set tgtRange = tgtWs.Range(tgtWs.Cells(2, tgtCol), tgtWs.Cells(srcLastRow, tgtCol))

    If Not RangeHasAnyMerge(srcRange) Then
        tgtRange.Value2 = srcRange.Value2
        Exit Sub
    End If

    If srcRange.Cells.CountLarge = 1 Then
        ReDim srcArr(1 To 1, 1 To 1)
        srcArr(1, 1) = srcRange.Value2
    Else
        srcArr = srcRange.Value2
    End If

    For r = 1 To UBound(srcArr, 1)
        tgtRange.Cells(r, 1).Value2 = 单元格值(srcWs.Cells(r + 1, srcCol))
    Next r
End Sub

Private Function RangeHasAnyMerge(ByVal rg As Range) As Boolean
    Dim v As Variant

    On Error Resume Next
    v = rg.MergeCells
    If IsNull(v) Then
        RangeHasAnyMerge = True
    Else
        RangeHasAnyMerge = CBool(v)
    End If
    On Error GoTo 0
End Function
