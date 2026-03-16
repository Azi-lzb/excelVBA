Attribute VB_Name = "功能2dot4批量Excel格式转换"
Option Explicit

' 批量Excel格式转换：按 config「2.4 批量Excel格式转换」-「目的格式」将选中工作簿另存为指定格式，
' 不覆盖原件，在源文件所在目录下新建以目的格式命名的文件夹（如 xlsx）存放转换结果。
' 支持格式：xls, xlsx, xlsm, csv, xlt, xltx, xltm, xlsb（et/ett 为 WPS 格式，Excel 无法另存为）

Private Const CFG_KEY As String = "2.4 批量Excel格式转换"
Private Const CFG_KN_TARGET As String = "目的格式"

' 扩展名（小写无点） -> Excel FileFormat
Private Function GetFileFormat(ByVal ext As String) As Long
    Dim e As String
    e = LCase(Trim(ext))
    If Left(e, 1) = "." Then e = Mid(e, 2)
    Select Case e
        Case "xls"
            GetFileFormat = 56   ' xlExcel8
        Case "xlsx"
            GetFileFormat = 51  ' xlOpenXMLWorkbook
        Case "xlsm"
            GetFileFormat = 52  ' xlOpenXMLWorkbookMacroEnabled
        Case "csv"
            GetFileFormat = 6   ' xlCSV
        Case "xlt"
            GetFileFormat = 17  ' xlTemplate
        Case "xltx"
            GetFileFormat = 54  ' xlOpenXMLTemplate
        Case "xltm"
            GetFileFormat = 53  ' xlOpenXMLTemplateMacroEnabled
        Case "xlsb"
            GetFileFormat = 50  ' xlExcel12
        Case Else
            GetFileFormat = 0
    End Select
End Function

' 返回规范扩展名（带点、小写），不支持则返回空
Private Function NormalizeExt(ByVal ext As String) As String
    Dim e As String
    e = LCase(Trim(ext))
    If Left(e, 1) = "." Then e = Mid(e, 2)
    If GetFileFormat(e) <> 0 Then NormalizeExt = "." & e Else NormalizeExt = ""
End Function

Public Sub 批量Excel格式转换()
    Dim fd As FileDialog
    Dim targetFormat As String
    Dim targetExt As String
    Dim fmt As Long
    Dim fileItem As Variant
    Dim srcPath As String, srcDir As String, srcName As String, baseName As String
    Dim outDir As String, outPath As String
    Dim wb As Workbook
    Dim t0 As Double
    Dim countOk As Long, countFail As Long
    Dim fso As Object

    t0 = Timer
    RunLog_WriteRow "2.4 批量Excel格式转换", "开始", "", "", "", "", "读取配置", ""

    targetFormat = Trim(初始化配置.读取配置(CFG_KEY, CFG_KN_TARGET))
    If targetFormat = "" Then targetFormat = "xlsx"
    targetExt = NormalizeExt(targetFormat)
    If targetExt = "" Then
        RunLog_WriteRow "2.4 批量Excel格式转换", "失败", "", "", "", "", "目的格式不支持或未配置，支持: xls,xlsx,xlsm,csv,xlt,xltx,xltm,xlsb", CStr(Round(Timer - t0, 2))
        MsgBox "config 中「2.4 批量Excel格式转换」-「目的格式」需为以下之一：" & vbCrLf & "xls, xlsx, xlsm, csv, xlt, xltx, xltm, xlsb" & vbCrLf & "当前为空或不支持，已按 xlsx 处理。", vbExclamation
        targetExt = ".xlsx"
        fmt = 51
    Else
        fmt = GetFileFormat(Replace(targetExt, ".", ""))
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择要转换格式的 Excel 文件"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel 文件", "*.xls;*.xlsx;*.xlsm;*.xlt;*.xltx;*.xltm;*.xlsb"
        .Filters.Add "CSV", "*.csv"
        .Filters.Add "所有文件", "*.*"
        If .Show <> -1 Then
            RunLog_WriteRow "2.4 批量Excel格式转换", "取消", "", "", "", "", "用户取消", CStr(Round(Timer - t0, 2))
            Exit Sub
        End If
    End With

    Set fso = CreateObject("Scripting.FileSystemObject")
    countOk = 0
    countFail = 0
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo ErrHandler
    For Each fileItem In fd.SelectedItems
        srcPath = CStr(fileItem)
        If Dir(srcPath) = "" Then
            RunLog_WriteRow "2.4 批量Excel格式转换", "跳过", srcPath, "", "", "文件不存在", "", ""
            countFail = countFail + 1
            GoTo NextFile
        End If

        srcDir = fso.GetParentFolderName(srcPath)
        srcName = fso.GetFileName(srcPath)
        baseName = srcName
        If InStrRev(baseName, ".") > 0 Then baseName = Left(baseName, InStrRev(baseName, ".") - 1)

        ' 目标文件夹：源目录下以格式命名的子文件夹（无点，如 xlsx）
        outDir = srcDir & "\" & Replace(targetExt, ".", "")
        If Not fso.FolderExists(outDir) Then
            fso.CreateFolder outDir
        End If

        outPath = outDir & "\" & baseName & targetExt
        If LCase(srcPath) = LCase(outPath) Then
            RunLog_WriteRow "2.4 批量Excel格式转换", "跳过", srcPath, "", "", "源与目标相同", "", ""
            countFail = countFail + 1
            GoTo NextFile
        End If

        On Error GoTo ErrOne
        Set wb = Workbooks.Open(srcPath, ReadOnly:=False, Password:="", UpdateLinks:=0)
        wb.SaveAs fileName:=outPath, FileFormat:=fmt, CreateBackup:=False
        wb.Close SaveChanges:=False
        Set wb = Nothing
        RunLog_WriteRow "2.4 批量Excel格式转换", "转换", srcPath, "", outPath, "成功", "", ""
        countOk = countOk + 1
        On Error GoTo ErrHandler
        GoTo NextFile
ErrOne:
        If Not wb Is Nothing Then
            On Error Resume Next
            wb.Close SaveChanges:=False
            On Error GoTo ErrHandler
        End If
        RunLog_WriteRow "2.4 批量Excel格式转换", "失败", srcPath, "", "", "失败", Err.Number & " " & Err.Description, ""
        countFail = countFail + 1
        On Error GoTo ErrHandler
NextFile:
    Next fileItem

Done:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    RunLog_WriteRow "2.4 批量Excel格式转换", "完成", "", "", "", "", "成功 " & countOk & " 个，失败 " & countFail & " 个", CStr(Round(Timer - t0, 2))
    MsgBox "转换完成。" & vbCrLf & "成功: " & countOk & vbCrLf & "失败/跳过: " & countFail & vbCrLf & "结果保存在源文件同目录下「" & Replace(targetExt, ".", "") & "」文件夹中。", vbInformation
    Exit Sub
ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    RunLog_WriteRow "2.4 批量Excel格式转换", "失败", "", "", "", "错误", Err.Number & " " & Err.Description, CStr(Round(Timer - t0, 2))
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
End Sub


