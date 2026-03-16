Attribute VB_Name = "功能3dot7注入VBA到源文件"
Option Explicit

' 将本工作簿 VBA_Export\Modules 下指定的 .bas 注入到「执行面板」所选源文件；可选复制 ThisWorkbook 代码。
' 配置：3.7 注入VBA到源文件 - 模块(all/名称;分号)、跳过模块(默认vbaSync)、复制ThisWorkbook(是/否)。
' 依赖：信任 VBA 工程对象模型；引用 Microsoft Visual Basic for Applications Extensibility 5.3。目标须 .xlsm/.xls/.xlsb。

Private Const PANEL_SHEET_NAME As String = "执行面板"
Private Const PANEL_DATA_START_ROW As Long = 5
Private Const PANEL_COL_PATH As Long = 2
Private Const CFG_KEY As String = "3.7 注入VBA到源文件"
Private Const CFG_KN_MODULES As String = "模块"
Private Const CFG_KN_SKIP As String = "跳过模块"
Private Const CFG_KN_COPY_THISWORKBOOK As String = "复制ThisWorkbook"
Private Const VBA_MODULES_FOLDER As String = "VBA_Export\Modules\"
Private Const CONFIG_SHEET_NAME As String = "config"
Private Const XL_OPENXML_WORKBOOK_MACRO As Long = 52
Private Const XL_EXCEL12 As Long = 50
Private Const XL_EXCEL8 As Long = 56
Private Const vbext_ct_Document As Long = 100

Public Sub 注入VBA到所选源文件()
    Dim wsPanel As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim pathList As Collection
    Dim cfgModules As String
    Dim cfgSkip As String
    Dim cfgCopyTW As String
    Dim basPaths As Collection
    Dim srcThisWorkbookCode As String
    Dim p As Variant
    Dim wb As Workbook
    Dim countOk As Long
    Dim countFail As Long
    Dim t0 As Double
    Dim exportRoot As String
    Dim onePath As String
    Dim errMsg As String

    On Error GoTo ErrHandler

    exportRoot = ThisWorkbook.path
    If Right(exportRoot, 1) <> "\" Then exportRoot = exportRoot & "\"
    If Dir(exportRoot & VBA_MODULES_FOLDER, vbDirectory) = "" Then
        MsgBox "未找到 VBA_Export\Modules 目录，请确认当前工作簿所在路径下存在该目录。", vbExclamation
        Exit Sub
    End If

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

    cfgModules = Trim(CStr(读取配置值(CFG_KEY, CFG_KN_MODULES)))
    If cfgModules = "" Then
        MsgBox "请在 config 表中配置「" & CFG_KEY & "」-「" & CFG_KN_MODULES & "」：填模块名（分号分隔）或 all。", vbExclamation
        Exit Sub
    End If

    cfgSkip = Trim(CStr(读取配置值(CFG_KEY, CFG_KN_SKIP)))
    cfgCopyTW = Trim(CStr(读取配置值(CFG_KEY, CFG_KN_COPY_THISWORKBOOK)))
    Set basPaths = BuildBasList(exportRoot & VBA_MODULES_FOLDER, cfgModules, cfgSkip)
    If basPaths.count = 0 Then
        MsgBox "未找到要注入的 .bas 文件（配置：[" & cfgModules & "]）。请检查 config 与 VBA_Export\Modules 目录。", vbExclamation
        Exit Sub
    End If

    ' 若需复制 ThisWorkbook，先取出本工作簿的 ThisWorkbook 代码
    srcThisWorkbookCode = ""
    If 配置为是(cfgCopyTW) Then
        On Error GoTo ErrTrust
        srcThisWorkbookCode = GetThisWorkbookCode(ThisWorkbook.VBProject)
        On Error GoTo ErrHandler
    End If

    t0 = Timer
    On Error Resume Next
    vbaSync.RunLog_WriteRow CFG_KEY, "开始", "", "", "", "", "源文件数 " & pathList.count & "，注入模块数 " & basPaths.count, ""
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
            errMsg = "该文件为 .xlsx 或无宏格式，无法保存 VBA。请先将工作簿另存为 .xlsm 后再注入。"
            On Error Resume Next
            vbaSync.RunLog_WriteRow CFG_KEY, "跳过", onePath, "", "", "格式不支持", errMsg, ""
            On Error GoTo ErrHandler
            wb.Close SaveChanges:=False
            countFail = countFail + 1
            GoTo NextPath
        End If

        On Error GoTo InjectErr
        InjectBasIntoProject wb.VBProject, basPaths
        If Len(srcThisWorkbookCode) > 0 Then CopyThisWorkbookCode wb.VBProject, srcThisWorkbookCode
        wb.Save
        errMsg = wb.Name
        wb.Close SaveChanges:=True
        countOk = countOk + 1
        On Error Resume Next
        vbaSync.RunLog_WriteRow CFG_KEY, "注入成功", errMsg, "", "", "成功", "", ""
        On Error GoTo ErrHandler
        GoTo NextPath

InjectErr:
        errMsg = Err.Number & " " & Err.Description
        On Error Resume Next
        If Not wb Is Nothing Then wb.Close SaveChanges:=False
        vbaSync.RunLog_WriteRow CFG_KEY, "注入失败", onePath, "", "", "失败", errMsg, ""
        countFail = countFail + 1
        MsgBox "注入失败：" & vbCrLf & onePath & vbCrLf & errMsg, vbExclamation
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
    MsgBox "注入完成。" & vbCrLf & "成功： " & countOk & vbCrLf & "失败： " & countFail, vbInformation
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

' 从 vbProj 中取 ThisWorkbook（工作簿文档模块，非 Sheet）的完整代码
Private Function GetThisWorkbookCode(ByVal vbProj As VBIDE.VBProject) As String
    Dim c As VBIDE.VBComponent
    GetThisWorkbookCode = ""
    On Error GoTo 0
    For Each c In vbProj.VBComponents
        If c.Type = vbext_ct_Document And (StrComp(c.Name, "ThisWorkbook", vbTextCompare) = 0 Or StrComp(c.Name, "此工作簿", vbTextCompare) = 0) Then
            With c.CodeModule
                If .CountOfLines > 0 Then GetThisWorkbookCode = .lines(1, .CountOfLines)
            End With
            Exit Function
        End If
    Next c
End Function

' 将 codeContent 写入目标 vbProj 的 ThisWorkbook（工作簿文档模块）
Private Sub CopyThisWorkbookCode(ByVal vbProj As VBIDE.VBProject, ByVal codeContent As String)
    Dim c As VBIDE.VBComponent
    On Error GoTo 0
    For Each c In vbProj.VBComponents
        If c.Type = vbext_ct_Document And (StrComp(c.Name, "ThisWorkbook", vbTextCompare) = 0 Or StrComp(c.Name, "此工作簿", vbTextCompare) = 0) Then
            With c.CodeModule
                If .CountOfLines > 0 Then .DeleteLines 1, .CountOfLines
                If Len(Trim(codeContent)) > 0 Then .InsertLines 1, codeContent
            End With
            Exit Sub
        End If
    Next c
End Sub

' cfgSkip：跳过模块名，分号分隔（如 vbaSync 或 vbaSync;其他）
Private Function BuildBasList(ByVal modulesDir As String, ByVal cfgModules As String, ByVal cfgSkip As String) As Collection
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim baseName As String
    Dim wantAll As Boolean
    Dim names() As String
    Dim skipNames() As String
    Dim i As Long
    Dim j As Long
    Dim n As String
    Dim skip As Boolean

    Set BuildBasList = New Collection
    wantAll = (LCase(Trim(cfgModules)) = "all")
    names = Split(cfgModules, ";")
    skipNames = Split(cfgSkip, ";")

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(modulesDir)
    For Each file In folder.Files
        If LCase(fso.GetExtensionName(file.Name)) = "bas" Then
            baseName = fso.GetBaseName(file.Name)
            skip = False
            For j = 0 To UBound(skipNames)
                If Trim(CStr(skipNames(j))) <> "" And StrComp(baseName, Trim(CStr(skipNames(j))), vbTextCompare) = 0 Then
                    skip = True
                    Exit For
                End If
            Next j
            If skip Then GoTo NextFile
            If wantAll Then
                BuildBasList.Add file.path
            Else
                For i = 0 To UBound(names)
                    n = Trim(CStr(names(i)))
                    If n <> "" And StrComp(baseName, n, vbTextCompare) = 0 Then
                        BuildBasList.Add file.path
                        Exit For
                    End If
                Next i
            End If
NextFile:
        End If
    Next file
    Set folder = Nothing
    Set fso = Nothing
End Function

Private Sub InjectBasIntoProject(ByVal vbProj As VBIDE.VBProject, ByVal basPaths As Collection)
    Dim compName As String
    Dim vbComp As VBIDE.VBComponent
    Dim fso As Object
    Dim p As Variant

    If vbProj.Protection = 1 Then
        Err.Raise 1000, "InjectBasIntoProject", "目标工作簿的 VBA 工程已锁定，无法注入。请在目标文件中取消 VBA 工程保护。"
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each p In basPaths
        compName = fso.GetBaseName(CStr(p))
        Set vbComp = FindComponent(vbProj, compName)
        If Not vbComp Is Nothing Then vbProj.VBComponents.Remove vbComp
        vbProj.VBComponents.Import CStr(p)
    Next p
    Set fso = Nothing
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
