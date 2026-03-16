Attribute VB_Name = "功能3dot13批量修改Sheet名"
Option Explicit

' 批量修改工作表名：从 config_rename 表 J 列（原表名）、K 列（新表名）读取映射，
' 对当前工作簿中「表名与 J 列全字段匹配」的工作表改名为对应 K 列。不修改未在 J 列出现的表。

Private Const CONFIG_RENAME_SHEET As String = "config_rename"
Private Const COL_原表名 As Long = 10   ' J 列
Private Const COL_新表名 As Long = 11   ' K 列

' 从 config_rename 第 2 行起读取 J、K 列，返回 Dictionary：原表名 -> 新表名（全字段匹配，键区分大小写）
Private Function 读取Sheet名映射() As Object
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim jVal As String, kVal As String
    Dim dict As Object

    Set 读取Sheet名映射 = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_RENAME_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.count, COL_原表名).End(xlUp).row
    If lastRow < 2 Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbBinaryCompare   ' 全字段匹配，区分大小写
    For i = 2 To lastRow
        jVal = Trim(CStr(ws.Cells(i, COL_原表名).value))
        kVal = Trim(CStr(ws.Cells(i, COL_新表名).value))
        If Len(jVal) > 0 And Len(kVal) > 0 Then dict(jVal) = kVal
    Next i
    Set 读取Sheet名映射 = dict
End Function

Public Sub 批量修改Sheet名()
    Dim dict As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim oldName As String
    Dim newName As String
    Dim countOk As Long
    Dim countSkip As Long
    Dim t0 As Double

    t0 = Timer
    RunLog_WriteRow "3.13 批量修改Sheet名", "开始", "", "", "", "", "读取 config_rename J/K 映射", ""

    Set dict = 读取Sheet名映射()
    If dict Is Nothing Or dict.count = 0 Then
        RunLog_WriteRow "3.13 批量修改Sheet名", "失败", "", "", "", "", "config_rename 缺失或 J、K 列为空", CStr(Round(Timer - t0, 2))
        MsgBox "请在本工作簿中维护 config_rename 表，并在 J 列填原表名、K 列填新表名（从第 2 行起）。", vbExclamation
        Exit Sub
    End If

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        RunLog_WriteRow "3.13 批量修改Sheet名", "失败", "", "", "", "", "无活动工作簿", CStr(Round(Timer - t0, 2))
        MsgBox "请先打开要修改的工作簿并保持为活动状态。", vbExclamation
        Exit Sub
    End If

    countOk = 0
    countSkip = 0
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler
    For Each ws In wb.Worksheets
        oldName = ws.Name
        If dict.Exists(oldName) Then
            newName = CStr(dict(oldName))
            If StrComp(oldName, newName, vbBinaryCompare) = 0 Then
                RunLog_WriteRow "3.13 批量修改Sheet名", "跳过", oldName, newName, "", "已是目标名", "", ""
                countSkip = countSkip + 1
            Else
                On Error GoTo ErrRename
                ws.Name = newName
                On Error GoTo ErrHandler
                RunLog_WriteRow "3.13 批量修改Sheet名", "重命名", oldName, newName, "", "OK", "", ""
                countOk = countOk + 1
            End If
        End If
NextSheet:
    Next ws

    Application.ScreenUpdating = True
    RunLog_WriteRow "3.13 批量修改Sheet名", "完成", wb.Name, "", "", "", "成功 " & countOk & "，跳过 " & countSkip, CStr(Round(Timer - t0, 2))
    MsgBox "批量修改工作表名完成。" & vbCrLf & "成功: " & countOk & vbCrLf & "跳过（已是目标名）: " & countSkip, vbInformation
    Exit Sub

ErrRename:
    RunLog_WriteRow "3.13 批量修改Sheet名", "失败", oldName, newName, "", "重命名失败", Err.Number & " " & Err.Description, ""
    MsgBox "将表「" & oldName & "」改为「" & newName & "」时出错：" & vbCrLf & Err.Description & vbCrLf & "（表名不可含 : \ / ? * [ ]，且不超过 31 个字符）", vbExclamation
    countSkip = countSkip + 1
    Application.ScreenUpdating = False
    Resume NextSheet
ErrHandler:
    Application.ScreenUpdating = True
    RunLog_WriteRow "3.13 批量修改Sheet名", "失败", "", "", "", "错误", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
End Sub
