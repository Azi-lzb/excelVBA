Attribute VB_Name = "功能3dot11批量Word格式转换"
Option Explicit

' 批量 Word 格式转换：按 config「2.5 批量Word格式转换」-「目的格式」将选中 Word 文档另存为 doc 或 docx，
' 不覆盖原件，在源文件所在目录下新建以目的格式命名的文件夹（如 doc、docx）存放转换结果。
' 支持格式：doc（Word 97-2003）、docx（Word 2007+）
' 依赖：需安装 Microsoft Word；使用 Late Binding，无需勾选 Word 引用。

Private Const CFG_KEY As String = "2.5 批量Word格式转换"
Private Const CFG_KN_TARGET As String = "目的格式"

' Word WdSaveFormat: 0=doc, 12=docx
Private Const WD_FORMAT_DOC As Long = 0
Private Const WD_FORMAT_DOCX As Long = 12

' 返回目的格式对应的 Word 保存格式码，不支持返回 -1
Private Function GetWordSaveFormat(ByVal ext As String) As Long
    Dim e As String
    e = LCase(Trim(ext))
    If Left(e, 1) = "." Then e = Mid(e, 2)
    Select Case e
        Case "doc"
            GetWordSaveFormat = WD_FORMAT_DOC
        Case "docx"
            GetWordSaveFormat = WD_FORMAT_DOCX
        Case Else
            GetWordSaveFormat = -1
    End Select
End Function

' 返回规范扩展名（带点、小写），不支持则返回空
Private Function NormalizeExt(ByVal ext As String) As String
    Dim e As String
    e = LCase(Trim(ext))
    If Left(e, 1) = "." Then e = Mid(e, 2)
    If GetWordSaveFormat(e) >= 0 Then NormalizeExt = "." & e Else NormalizeExt = ""
End Function

Public Sub 批量Word格式转换()
    Dim fd As FileDialog
    Dim targetFormat As String
    Dim targetExt As String
    Dim fmt As Long
    Dim fileItem As Variant
    Dim srcPath As String, srcDir As String, srcName As String, baseName As String
    Dim outDir As String, outPath As String
    Dim wdApp As Object
    Dim doc As Object
    Dim t0 As Double
    Dim countOk As Long, countFail As Long
    Dim fso As Object

    t0 = Timer
    RunLog_WriteRow "3.11 批量Word格式转换", "开始", "", "", "", "", "读取配置", ""

    targetFormat = Trim(初始化配置.读取配置(CFG_KEY, CFG_KN_TARGET))
    If targetFormat = "" Then targetFormat = "docx"
    targetExt = NormalizeExt(targetFormat)
    If targetExt = "" Then
        RunLog_WriteRow "3.11 批量Word格式转换", "失败", "", "", "", "", "目的格式不支持或未配置，支持: doc,docx", CStr(Round(Timer - t0, 2))
        MsgBox "config 中「2.5 批量Word格式转换」-「目的格式」需为 doc 或 docx。" & vbCrLf & "当前为空或不支持，已按 docx 处理。", vbExclamation
        targetExt = ".docx"
        fmt = WD_FORMAT_DOCX
    Else
        fmt = GetWordSaveFormat(Replace(targetExt, ".", ""))
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择要转换格式的 Word 文件"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Word 文档", "*.doc;*.docx"
        .Filters.Add "所有文件", "*.*"
        If .Show <> -1 Then
            RunLog_WriteRow "3.11 批量Word格式转换", "取消", "", "", "", "", "用户取消", CStr(Round(Timer - t0, 2))
            Exit Sub
        End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    countOk = 0
    countFail = 0

    On Error GoTo ErrHandler
    Set wdApp = CreateObject("Word.Application")
    wdApp.Visible = False
    wdApp.DisplayAlerts = 0  ' wdAlertsNone

    For Each fileItem In fd.SelectedItems
        srcPath = CStr(fileItem)
        If Dir(srcPath) = "" Then
            RunLog_WriteRow "3.11 批量Word格式转换", "跳过", srcPath, "", "", "文件不存在", "", ""
            countFail = countFail + 1
            GoTo NextFile
        End If

        srcDir = fso.GetParentFolderName(srcPath)
        srcName = fso.GetFileName(srcPath)
        baseName = srcName
        If InStrRev(baseName, ".") > 0 Then baseName = Left(baseName, InStrRev(baseName, ".") - 1)

        outDir = srcDir & "\" & Replace(targetExt, ".", "")
        If Not fso.FolderExists(outDir) Then
            fso.CreateFolder outDir
        End If

        outPath = outDir & "\" & baseName & targetExt
        If LCase(srcPath) = LCase(outPath) Then
            RunLog_WriteRow "3.11 批量Word格式转换", "跳过", srcPath, "", "", "源与目标相同", "", ""
            countFail = countFail + 1
            GoTo NextFile
        End If

        Set doc = Nothing
        On Error GoTo ErrOne
        Set doc = wdApp.Documents.Open(srcPath, ConfirmConversions:=False, ReadOnly:=True, AddToRecentFiles:=False)
        doc.SaveAs2 fileName:=outPath, FileFormat:=fmt
        doc.Close SaveChanges:=False
        Set doc = Nothing
        RunLog_WriteRow "3.11 批量Word格式转换", "转换", srcPath, "", outPath, "成功", "", ""
        countOk = countOk + 1
        On Error GoTo ErrHandler
        GoTo NextFile
ErrOne:
        If Not doc Is Nothing Then
            On Error Resume Next
            doc.Close SaveChanges:=False
            On Error GoTo ErrHandler
        End If
        RunLog_WriteRow "3.11 批量Word格式转换", "失败", srcPath, "", "", "失败", Err.Number & " " & Err.Description, ""
        countFail = countFail + 1
        On Error GoTo ErrHandler
NextFile:
    Next fileItem

    On Error Resume Next
    wdApp.Quit
    On Error GoTo ErrHandler

Done:
    RunLog_WriteRow "3.11 批量Word格式转换", "完成", "", "", "", "", "成功 " & countOk & " 个，失败 " & countFail & " 个", CStr(Round(Timer - t0, 2))
    MsgBox "转换完成。" & vbCrLf & "成功: " & countOk & vbCrLf & "失败/跳过: " & countFail & vbCrLf & "结果保存在源文件同目录下「" & Replace(targetExt, ".", "") & "」文件夹中。", vbInformation
    Exit Sub
ErrHandler:
    On Error Resume Next
    If Not wdApp Is Nothing Then wdApp.Quit
    On Error GoTo 0
    RunLog_WriteRow "3.11 批量Word格式转换", "失败", "", "", "", "错误", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
End Sub
