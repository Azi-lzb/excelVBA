Attribute VB_Name = "功能2dot3批量重命名文件"
Option Explicit

' 批量重命名文件：从 config_rename 表读取「仲恺报表银行命名」配置，按 简称→全称、全称→代码、G→H 键值 生成新文件名。
' 配置表结构（与 QtProgram 仲恺报表银行命名一致）：
'   A列=简称, B列=全称, D列=代码, E列=全称(用于建 全称→代码 映射), G列=键, H列=值（如 数据日期、报表名称）
' 新文件名格式：数据日期 & " " & 代码 & " " & 全称 & " " & 报表名称 & 原扩展名

Private Const CONFIG_RENAME_SHEET As String = "config_rename"

Public Sub 批量重命名文件()
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim 简称到全称 As Object
    Dim 全称到代码 As Object
    Dim 键值 As Object
    Dim 数据日期 As String, 报表名称 As String
    Dim filePath As String, fileDir As String, fileName As String, FileExt As String
    Dim newFileName As String, newFilePath As String
    Dim shortName As Variant, fullName As String, code As String
    Dim matched As Boolean
    Dim t0 As Double
    Dim countOk As Long, countSkip As Long

    t0 = Timer
    RunLog_WriteRow "2.3 批量重命名文件", "开始", "", "", "", "", "读取 config_rename", ""

    Set 简称到全称 = 读取配置表键值(CONFIG_RENAME_SHEET, 1, 2)   ' A→B 简称→全称
    Set 全称到代码 = 读取配置表键值(CONFIG_RENAME_SHEET, 5, 4)   ' E→D 全称→代码
    Set 键值 = 读取配置表键值(CONFIG_RENAME_SHEET, 7, 8)         ' G→H 键→值

    If 简称到全称 Is Nothing Or 全称到代码 Is Nothing Or 键值 Is Nothing Then
        RunLog_WriteRow "2.3 批量重命名文件", "失败", "", "", "", "", "config_rename 表缺失或列为空", CStr(Round(Timer - t0, 2))
        MsgBox "请先在本工作簿中维护好 config_rename 表（结构见模块注释）。", vbExclamation
        Exit Sub
    End If

    数据日期 = 取键值(键值, "数据日期")
    报表名称 = 取键值(键值, "报表名称")
    If 数据日期 = "" Or 报表名称 = "" Then
        RunLog_WriteRow "2.3 批量重命名文件", "失败", "", "", "", "", "G-H 中缺少 数据日期 或 报表名称", CStr(Round(Timer - t0, 2))
        MsgBox "config_rename 表 G 列需有「数据日期」「报表名称」键，H 列为对应值。", vbExclamation
        Exit Sub
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择要重命名的文件"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "所有文件", "*.*"
        If .Show <> -1 Then
            RunLog_WriteRow "2.3 批量重命名文件", "取消", "", "", "", "", "用户取消选择", CStr(Round(Timer - t0, 2))
            Exit Sub
        End If
    End With

    countOk = 0
    countSkip = 0
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler
    For Each fileItem In fd.SelectedItems
        filePath = CStr(fileItem)
        If Dir(filePath) = "" Then
            RunLog_WriteRow "2.3 批量重命名文件", "跳过", filePath, "", "", "文件不存在", "", ""
            countSkip = countSkip + 1
            GoTo NextFile
        End If

        fileDir = Left(filePath, InStrRev(filePath, "\"))
        fileName = Mid(filePath, InStrRev(filePath, "\") + 1)
        FileExt = ""
        If InStrRev(fileName, ".") > 0 Then FileExt = Mid(fileName, InStrRev(fileName, "."))

        matched = False
        For Each shortName In 简称到全称.keys
            If InStr(1, fileName, CStr(shortName), vbTextCompare) > 0 Then
                fullName = CStr(简称到全称(shortName))
                If 全称到代码.Exists(fullName) Then
                    code = CStr(全称到代码(fullName))
                    newFileName = 数据日期 & " " & code & " " & fullName & " " & 报表名称 & FileExt
                    newFilePath = fileDir & newFileName
                    If StrComp(filePath, newFilePath, vbTextCompare) = 0 Then
                        RunLog_WriteRow "2.3 批量重命名文件", "跳过", fileName, newFileName, "", "名称已为目标", "", ""
                        countSkip = countSkip + 1
                    ElseIf ContainsInvalidFileNameChars(newFileName) Then
                        RunLog_WriteRow "2.3 批量重命名文件", "跳过", fileName, newFileName, "", "目标文件名含非法字符，已跳过", "", ""
                        countSkip = countSkip + 1
                    ElseIf Dir(newFilePath) <> "" Then
                        RunLog_WriteRow "2.3 批量重命名文件", "跳过", fileName, newFileName, "", "目标文件已存在", "", ""
                        countSkip = countSkip + 1
                    Else
                        Name filePath As newFilePath
                        RunLog_WriteRow "2.3 批量重命名文件", "重命名", fileName, newFileName, "", "OK", "", ""
                        countOk = countOk + 1
                    End If
                    matched = True
                End If
                Exit For
            End If
        Next shortName

        If Not matched Then
            RunLog_WriteRow "2.3 批量重命名文件", "跳过", fileName, "", "", "未匹配任何简称", "", ""
            countSkip = countSkip + 1
        End If
NextFile:
    Next fileItem

    Application.ScreenUpdating = True
    RunLog_WriteRow "2.3 批量重命名文件", "完成", "", "", "", "", "成功 " & countOk & "，跳过 " & countSkip, CStr(Round(Timer - t0, 2))
    MsgBox "批量重命名完成。" & vbCrLf & "成功: " & countOk & "，跳过: " & countSkip, vbInformation
    Set fd = Nothing
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    RunLog_WriteRow "2.3 批量重命名文件", "错误", filePath, "", "", Err.Description, "", CStr(Round(Timer - t0, 2))
    MsgBox "重命名时出错: " & Err.Description, vbCritical
End Sub

' 从指定表按 键列→值列 建 Dictionary（从第2行起，键/值非空才加入）
Private Function 读取配置表键值(ByVal sheetName As String, ByVal keyCol As Long, ByVal valueCol As Long) As Object
    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim k As String, v As String
    Dim dict As Object

    Set 读取配置表键值 = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.count, keyCol).End(xlUp).row
    If lastRow < 2 Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    For i = 2 To lastRow
        k = Trim(CStr(ws.Cells(i, keyCol).value))
        v = Trim(CStr(ws.Cells(i, valueCol).value))
        If Len(k) > 0 And Len(v) > 0 Then dict(k) = v
    Next i
    Set 读取配置表键值 = dict
End Function

Private Function 取键值(ByVal d As Object, ByVal key As String) As String
    If d Is Nothing Then Exit Function
    If d.Exists(key) Then 取键值 = CStr(d(key))
End Function

Private Function ContainsInvalidFileNameChars(ByVal fileNameOnly As String) As Boolean
    Dim invalidChars As Variant
    Dim ch As Variant

    invalidChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For Each ch In invalidChars
        If InStr(1, fileNameOnly, CStr(ch), vbBinaryCompare) > 0 Then
            ContainsInvalidFileNameChars = True
            Exit Function
        End If
    Next ch
End Function
