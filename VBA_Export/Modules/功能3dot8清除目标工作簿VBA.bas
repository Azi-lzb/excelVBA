Attribute VB_Name = "功能3dot8清除目标工作簿VBA"
Option Explicit

' 清除「执行面板」所列目标工作簿中的 VBA 代码：删除所有标准模块(.bas)，并清空 ThisWorkbook 中的代码。
' 为 3.7 注入VBA 的反向操作。目标须为 .xlsm/.xls/.xlsb；需信任 VBA 工程对象模型；引用 Extensibility 5.3。

Private Const PANEL_SHEET_NAME As String = "执行面板"
Private Const PANEL_DATA_START_ROW As Long = 5
Private Const PANEL_COL_PATH As Long = 2
Private Const CFG_KEY As String = "3.8 清除目标工作簿VBA"
Private Const CONFIG_SHEET_NAME As String = "config"
Private Const XL_OPENXML_WORKBOOK_MACRO As Long = 52
Private Const XL_EXCEL12 As Long = 50
Private Const XL_EXCEL8 As Long = 56
Private Const vbext_ct_StdModule As Long = 1
Private Const vbext_ct_Document As Long = 100

Public Sub 清除目标工作簿VBA()
    Dim wsPanel As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim pathList As Collection
    Dim p As Variant
    Dim wb As Workbook
    Dim countOk As Long
    Dim countFail As Long
    Dim t0 As Double
    Dim onePath As String
    Dim errMsg As String
    Dim cfgClearThisWorkbook As String

    On Error GoTo ErrHandler

    On Error GoTo ErrTrust
    Set wsPanel = ThisWorkbook.Worksheets(PANEL_SHEET_NAME)
    On Error GoTo ErrHandler
    If wsPanel Is Nothing Then
        MsgBox "未找到「执行面板」工作表。请先运行「4.4 初始化执行面板」或「4.3 初始化配置」。", vbExclamation
        Exit Sub
    End If

    lastRow = wsPanel.Cells(wsPanel.Rows.count, PANEL_COL_PATH).End(xlUp).row
    If lastRow < PANEL_DATA_START_ROW Then
        MsgBox "执行面板中暂无源文件路径（请从 B" & PANEL_DATA_START_ROW & " 起填写完整路径）。", vbExclamation
        Exit Sub
    End If

    Set pathList = New Collection
    For r = PANEL_DATA_START_ROW To lastRow
        onePath = Trim(CStr(wsPanel.Cells(r, PANEL_COL_PATH).value))
        If onePath <> "" Then pathList.Add onePath
    Next r
    If pathList.count = 0 Then
        MsgBox "执行面板 B 列未找到有效路径。", vbExclamation
        Exit Sub
    End If

    cfgClearThisWorkbook = Trim(CStr(读取配置值(CFG_KEY, "清除ThisWorkbook")))
    If cfgClearThisWorkbook = "" Then cfgClearThisWorkbook = "是"

    If MsgBox("将清除 " & pathList.count & " 个目标工作簿中的全部 .bas 模块及 ThisWorkbook 内代码，是否继续？", vbYesNo + vbExclamation, "清除VBA") <> vbYes Then
        Exit Sub
    End If

    t0 = Timer
    On Error Resume Next
    vbaSync.RunLog_WriteRow CFG_KEY, "开始", "", "", "", "", "目标数 " & pathList.count, ""
    On Error GoTo ErrHandler

    countOk = 0
    countFail = 0
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For Each p In pathList
        onePath = CStr(p)
        Set wb = Nothing
        On Error GoTo OpenErr
        Set wb = Workbooks.Open(onePath, UpdateLinks:=0, ReadOnly:=False)
        On Error GoTo ErrHandler

        If Not 工作簿支持宏(wb) Then
            errMsg = "该文件为 .xlsx 或无宏格式，无法处理 VBA。"
            On Error Resume Next
            vbaSync.RunLog_WriteRow CFG_KEY, "跳过", onePath, "", "", "格式不支持", errMsg, ""
            On Error GoTo ErrHandler
            wb.Close SaveChanges:=False
            countFail = countFail + 1
            GoTo NextPath
        End If

        On Error GoTo ClearErr
        ClearVBAInProject wb.VBProject, 配置为是(cfgClearThisWorkbook)
        wb.Save
        errMsg = wb.Name
        wb.Close SaveChanges:=True
        countOk = countOk + 1
        On Error Resume Next
        vbaSync.RunLog_WriteRow CFG_KEY, "清除成功", errMsg, "", "", "成功", "", ""
        On Error GoTo ErrHandler
        GoTo NextPath

ClearErr:
        errMsg = Err.Number & " " & Err.Description
        On Error Resume Next
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
        vbaSync.RunLog_WriteRow CFG_KEY, "清除失败", onePath, "", "", "失败", errMsg, ""
        countFail = countFail + 1
        MsgBox "清除失败：" & vbCrLf & onePath & vbCrLf & errMsg, vbExclamation
        On Error GoTo ErrHandler
        GoTo NextPath

OpenErr:
        errMsg = Err.Number & " " & Err.Description
        On Error Resume Next
        vbaSync.RunLog_WriteRow CFG_KEY, "打开失败", onePath, "", "", "失败", errMsg, ""
        countFail = countFail + 1
        MsgBox "打开失败：" & vbCrLf & onePath & vbCrLf & errMsg, vbExclamation
NextPath:
    Next p

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    On Error Resume Next
    vbaSync.RunLog_WriteRow CFG_KEY, "完成", "", "", "", "成功 " & countOk & "，失败 " & countFail, "", CStr(Round(Timer - t0, 2))
    On Error GoTo ErrHandler
    MsgBox "清除完成。" & vbCrLf & "成功： " & countOk & vbCrLf & "失败： " & countFail, vbInformation
    Exit Sub

ErrTrust:
    MsgBox "访问 VBA 工程时出错。请勾选「信任对 VBA 工程对象模型的访问」，并在 VBA 编辑器「工具-引用」中勾选 Microsoft Visual Basic for Applications Extensibility 5.3。", vbExclamation
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub
ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "执行出错：" & vbCrLf & Err.Description, vbCritical
End Sub

Private Function 配置为是(ByVal v As String) As Boolean
    Dim s As String
    s = LCase(Trim(CStr(v)))
    配置为是 = (s = "是" Or s = "1" Or s = "true" Or s = "y" Or s = "yes")
End Function

Private Function 工作簿支持宏(ByVal wb As Workbook) As Boolean
    On Error Resume Next
    Select Case wb.FileFormat
        Case XL_OPENXML_WORKBOOK_MACRO, XL_EXCEL12, XL_EXCEL8
            工作簿支持宏 = True
        Case Else
            工作簿支持宏 = False
    End Select
    On Error GoTo 0
End Function

' 清除 vbProj 中所有标准模块，并可选清空 ThisWorkbook 代码
Private Sub ClearVBAInProject(ByVal vbProj As VBIDE.VBProject, ByVal clearThisWorkbook As Boolean)
    Dim vbComp As VBIDE.VBComponent
    Dim namesToRemove() As String
    Dim n As Long
    Dim i As Long

    If vbProj.Protection = 1 Then
        Err.Raise 1000, "ClearVBAInProject", "目标工作簿的 VBA 工程已锁定，无法清除。请在目标文件中取消 VBA 工程保护。"
    End If

    n = 0
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = vbext_ct_StdModule Then
            n = n + 1
            ReDim Preserve namesToRemove(1 To n)
            namesToRemove(n) = vbComp.Name
        End If
    Next vbComp

    For i = 1 To n
        Set vbComp = FindComponent(vbProj, namesToRemove(i))
        If Not vbComp Is Nothing Then vbProj.VBComponents.Remove vbComp
    Next i

    If clearThisWorkbook Then ClearThisWorkbookCode vbProj
End Sub

' 清空 ThisWorkbook（工作簿文档模块）内的全部代码
Private Sub ClearThisWorkbookCode(ByVal vbProj As VBIDE.VBProject)
    Dim c As VBIDE.VBComponent
    For Each c In vbProj.VBComponents
        If c.Type = vbext_ct_Document And (StrComp(c.Name, "ThisWorkbook", vbTextCompare) = 0 Or StrComp(c.Name, "此工作簿", vbTextCompare) = 0) Then
            With c.CodeModule
                If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
            End With
            Exit Sub
        End If
    Next c
End Sub

Private Function FindComponent(ByVal vbProj As VBIDE.VBProject, ByVal compName As String) As VBIDE.VBComponent
    Dim c As VBIDE.VBComponent
    Set FindComponent = Nothing
    For Each c In vbProj.VBComponents
        If StrComp(c.Name, compName, vbTextCompare) = 0 Then
            Set FindComponent = c
            Exit Function
        End If
    Next c
End Function

Private Function 读取配置值(ByVal 键 As String, ByVal 键名 As String) As String
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim aVal As String
    Dim bVal As String
    读取配置值 = ""
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
            读取配置值 = Trim(CStr(ws.Cells(i, 3).value))
            Exit Function
        End If
    Next i
End Function
