Attribute VB_Name = "功能3dot13批量修改Sheet名"
Option Explicit

' 批量修改工作表名称：读取 config_rename 的 J 列（原表名）和 K 列（新表名）映射，
' 对执行面板中已登记的源文件逐个打开并重命名命中的工作表。

Private Const CONFIG_RENAME_SHEET As String = "config_rename"
Private Const PANEL_SHEET_NAME As String = "执行面板"
Private Const PANEL_DATA_START_ROW As Long = 5
Private Const PANEL_COL_PATH As Long = 2
Private Const COL_原表名 As Long = 10
Private Const COL_新表名 As Long = 11
Private Const LOG_KEY As String = "3.13 批量修改Sheet名"

Private Function 读取Sheet名映射() As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim oldName As String
    Dim newName As String
    Dim dict As Object

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_RENAME_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, COL_原表名).End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare

    For i = 2 To lastRow
        oldName = Trim$(CStr(ws.Cells(i, COL_原表名).Value))
        newName = Trim$(CStr(ws.Cells(i, COL_新表名).Value))
        If oldName <> "" And newName <> "" Then
            dict(oldName) = newName
        End If
    Next i

    Set 读取Sheet名映射 = dict
End Function

Private Function 读取源文件路径列表() As Collection
    Dim ws As Worksheet
    Dim result As New Collection
    Dim keys As Object
    Dim lastRow As Long
    Dim rowNo As Long
    Dim onePath As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    On Error GoTo 0
    If ws Is Nothing Then
        Set 读取源文件路径列表 = result
        Exit Function
    End If

    Set keys = CreateObject("Scripting.Dictionary")
    keys.CompareMode = vbTextCompare
    lastRow = ws.Cells(ws.Rows.Count, PANEL_COL_PATH).End(xlUp).Row

    For rowNo = PANEL_DATA_START_ROW To lastRow
        onePath = Trim$(CStr(ws.Cells(rowNo, PANEL_COL_PATH).Value))
        If onePath <> "" Then
            If Not keys.Exists(onePath) Then
                keys(onePath) = True
                result.Add onePath
            End If
        End If
    Next rowNo

    Set 读取源文件路径列表 = result
End Function

Public Sub 批量修改Sheet名()
    Dim dict As Object
    Dim sourcePaths As Collection
    Dim onePath As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim oldName As String
    Dim newName As String
    Dim countWorkbook As Long
    Dim countRename As Long
    Dim countSkip As Long
    Dim t0 As Double

    t0 = Timer
    RunLog_WriteRow LOG_KEY, "开始", "", "", "", "", "读取 config_rename 与执行面板源文件", ""

    Set dict = 读取Sheet名映射()
    If dict Is Nothing Or dict.Count = 0 Then
        RunLog_WriteRow LOG_KEY, "失败", "", "", "", "", "config_rename 缺失或 J/K 列为空", CStr(Round(Timer - t0, 2))
        MsgBox "请先在 config_rename 中维护 J 列原表名、K 列新表名（从第 2 行开始）。", vbExclamation
        Exit Sub
    End If

    Set sourcePaths = 读取源文件路径列表()
    If sourcePaths Is Nothing Or sourcePaths.Count = 0 Then
        RunLog_WriteRow LOG_KEY, "失败", "", "", "", "", "执行面板未登记源文件", CStr(Round(Timer - t0, 2))
        MsgBox "执行面板中未找到源文件。请先通过“1.3/1.4/1.5 选择源文件”登记要处理的工作簿。", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrHandler

    For Each onePath In sourcePaths
        If Dir$(CStr(onePath)) = "" Then
            RunLog_WriteRow LOG_KEY, "跳过文件", CStr(onePath), "", "", "文件不存在", "", ""
            countSkip = countSkip + 1
        Else
            Set wb = Workbooks.Open(CStr(onePath), ReadOnly:=False, UpdateLinks:=0)
            countWorkbook = countWorkbook + 1

            For Each ws In wb.Worksheets
                oldName = ws.Name
                If dict.Exists(oldName) Then
                    newName = CStr(dict(oldName))
                    If StrComp(oldName, newName, vbBinaryCompare) = 0 Then
                        RunLog_WriteRow LOG_KEY, "跳过Sheet", wb.Name & "|" & oldName, newName, "", "已是目标名", "", ""
                        countSkip = countSkip + 1
                    Else
                        On Error GoTo ErrRename
                        ws.Name = newName
                        On Error GoTo ErrHandler
                        RunLog_WriteRow LOG_KEY, "重命名", wb.Name & "|" & oldName, newName, "", "OK", "", ""
                        countRename = countRename + 1
                    End If
                End If
            Next ws

            wb.Close SaveChanges:=True
            Set wb = Nothing
        End If
    Next onePath

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    RunLog_WriteRow LOG_KEY, "结束", CStr(countWorkbook), CStr(countRename), CStr(countSkip), "完成", "批量处理完成", CStr(Round(Timer - t0, 2))
    MsgBox "批量修改工作表名完成。" & vbCrLf & "处理工作簿: " & countWorkbook & vbCrLf & "成功重命名: " & countRename & vbCrLf & "跳过: " & countSkip, vbInformation
    Exit Sub

ErrRename:
    RunLog_WriteRow LOG_KEY, "失败", wb.Name & "|" & oldName, newName, "", "重命名失败", Err.Number & " " & Err.Description, ""
    MsgBox "工作簿《" & wb.Name & "》中工作表《" & oldName & "》重命名为《" & newName & "》时失败。" & vbCrLf & Err.Description & vbCrLf & "请检查名称是否重复、包含非法字符或超过 31 个字符。", vbExclamation
    countSkip = countSkip + 1
    Err.Clear
    On Error GoTo ErrHandler
    Resume Next

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    SafeCloseWorkbook wb, False
    RunLog_WriteRow LOG_KEY, "失败", "", "", "", "异常", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
End Sub

Private Sub SafeCloseWorkbook(ByRef wb As Workbook, Optional ByVal saveChanges As Boolean = False)
    On Error Resume Next
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=saveChanges
        Set wb = Nothing
    End If
    On Error GoTo 0
End Sub
