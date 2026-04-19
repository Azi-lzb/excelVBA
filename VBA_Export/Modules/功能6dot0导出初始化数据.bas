Attribute VB_Name = "功能6dot0导出初始化数据"
Option Explicit

Public Sub 导出初始化数据()
    Dim t0 As Double
    Dim okCount As Long
    Dim failCount As Long
    Dim failList As String
    Dim logKey As String

    t0 = Timer
    logKey = "6.0 导出初始化数据"

    On Error Resume Next
    RunLog_WriteRow logKey, "开始", "", "", "", "", "开始", ""
    On Error GoTo 0

    执行单项初始化 "初始化config", okCount, failCount, failList, logKey
    执行单项初始化 "初始化执行面板", okCount, failCount, failList, logKey
    执行单项初始化 "初始化运行日志", okCount, failCount, failList, logKey
    执行单项初始化 "初始化机构映射表", okCount, failCount, failList, logKey
    执行单项初始化 "初始化表格比对", okCount, failCount, failList, logKey
    执行单项初始化 "初始化工作表提取", okCount, failCount, failList, logKey
    执行单项初始化 "初始化config_rename", okCount, failCount, failList, logKey
    执行单项初始化 "初始化统计工具", okCount, failCount, failList, logKey
    执行单项初始化 "初始化路径标准化映射", okCount, failCount, failList, logKey
    执行单项初始化 "功能3dot12_初始化按配置查重", okCount, failCount, failList, logKey
    执行单项初始化 "初始化打印配置", okCount, failCount, failList, logKey

    RunLog_WriteRow logKey, "完成", "", "", "", IIf(failCount = 0, "成功", "部分失败"), _
                    "成功=" & okCount & " 失败=" & failCount, CStr(Round(Timer - t0, 2))

    If failCount = 0 Then
        MsgBox "初始化导出完成。" & vbCrLf & _
               "成功：" & okCount & " 项" & vbCrLf & _
               "失败：" & failCount & " 项", vbInformation, "导出初始化数据"
    Else
        MsgBox "初始化导出完成（部分失败）。" & vbCrLf & _
               "成功：" & okCount & " 项" & vbCrLf & _
               "失败：" & failCount & " 项" & vbCrLf & vbCrLf & _
               failList, vbExclamation, "导出初始化数据"
    End If
End Sub

Public Sub Run_InitExport_All()
    导出初始化数据
End Sub

Private Sub 执行单项初始化(ByVal procName As String, ByRef okCount As Long, ByRef failCount As Long, ByRef failList As String, ByVal logKey As String)
    On Error GoTo RunFail

    Application.Run procName
    okCount = okCount + 1
    RunLog_WriteRow logKey, "初始化", procName, "", "", "成功", "OK", ""
    Exit Sub

RunFail:
    failCount = failCount + 1
    If Len(failList) > 0 Then failList = failList & vbCrLf
    failList = failList & procName & " -> " & Err.Number & " " & Err.Description
    RunLog_WriteRow logKey, "初始化", procName, "", "", "失败", Err.Number & " " & Err.Description, ""
    Err.Clear
End Sub
