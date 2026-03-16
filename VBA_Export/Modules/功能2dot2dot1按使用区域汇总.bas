Attribute VB_Name = "功能2dot2dot1按使用区域汇总"
Option Explicit

' 本模块：按各工作表的 UsedRange 汇总数据到一个新工作簿。
' 可从 config 表读取配置（如是否跳过表头）。
' 可选：跳过每个文件的表头行，仅保留第一个文件的表头。

Private Const MAX_DATA_COLS As Long = 256  ' 仅作上限，表头列数=第一个 UsedRange 的列数

Private Function CfgVal(ByVal cfgKey As String, ByVal kn As String) As String
    Dim ws As Worksheet, r As Long
    CfgVal = ""
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("config")
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    For r = 2 To ws.Cells(ws.Rows.count, 1).End(xlUp).row
        If Trim(CStr(ws.Cells(r, 1).value)) = cfgKey And _
           LCase(Trim(CStr(ws.Cells(r, 2).value))) = LCase(kn) Then
            CfgVal = Trim(CStr(ws.Cells(r, 3).value))
            Exit Function
        End If
    Next r
End Function

''' <summary>
''' 取得合并单元格或普通单元格的值。
''' </summary>
Private Function GetCellValue(ByVal rng As Range, ByVal rowIdx As Long, ByVal colIdx As Long) As Variant
    Dim c As Range
    On Error Resume Next
    Set c = rng.Cells(rowIdx, colIdx)
    If c Is Nothing Then
        GetCellValue = Empty
        Exit Function
    End If
    If c.MergeCells Then
        GetCellValue = c.MergeArea.Cells(1, 1).value
    Else
        GetCellValue = c.value
    End If
    On Error GoTo 0
End Function

''' <summary>
''' 按使用区域汇总：可从 config 读取配置或弹窗询问，汇总所选文件的 UsedRange 到新工作簿；
''' 表头列数 = 第一个遇到的 UsedRange 的列数。
''' </summary>
Public Sub SummarizeByUsedRange()
    Dim fd As FileDialog
    Dim selectedFile As Variant
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim newWb As Workbook
    Dim sumWs As Worksheet
    Dim usedRng As Range
    Dim destRow As Long
    Dim skipHeader As Boolean
    Dim headerOutputOnce As Boolean
    Dim wbName As String
    Dim i As Long
    Dim j As Long
    Dim dataCols As Long
    Dim savePath As Variant
    Dim headerCols As Long
    Dim hasConfig As Boolean
    Dim configVal As String

    configVal = CfgVal("2.2.1 按使用区域汇总", "跳过表头")
    If configVal = "" Then configVal = CfgVal("2.2.1 按使用区域汇总", "是否跳过表头")
    hasConfig = (configVal <> "")
    If hasConfig Then
        skipHeader = (LCase(configVal) = "是" Or configVal = "1" Or LCase(configVal) = "true" Or LCase(configVal) = "y" Or LCase(configVal) = "yes")
    End If

    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "请选择要汇总的 Excel/CSV/WPS 文档"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel/CSV/WPS 文档", "*.xls;*.xlsx;*.xlsm;*.csv;*.et"
        .Filters.Add "Excel 文档", "*.xls;*.xlsx;*.xlsm"
        .Filters.Add "CSV", "*.csv"
        .Filters.Add "WPS 表格(ET)", "*.et"
    End With

    If fd.Show <> True Then
        MsgBox "未选择文件。", vbInformation
        Exit Sub
    End If

    If Not hasConfig Then
        skipHeader = (MsgBox("是否跳过各文件的表头行？" & vbCrLf & vbCrLf & _
                             "选择“是”：仅保留第一个文件的第 1 行作为表头，其余文件从第 2 行开始汇总。" & vbCrLf & _
                             "选择“否”：每个工作表的所有行都参与汇总。", _
                             vbYesNo + vbQuestion, "表头设置") = vbYes)
    End If

    Dim t0 As Double
    t0 = Timer
    RunLog_WriteRow "2.2.1 按使用区域汇总", "开始", "", "", "", "", "运行", ""

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    On Error GoTo ErrHandle

    headerCols = 0
    Set newWb = Workbooks.Add
    Set sumWs = newWb.Sheets(1)
    sumWs.Name = "汇总"

    destRow = 2
    headerOutputOnce = False

    For Each selectedFile In fd.SelectedItems
        On Error Resume Next
        Set sourceWb = Workbooks.Open(CStr(selectedFile), ReadOnly:=True)
        If Err.Number <> 0 Then
            RunLog_WriteRow "2.2.1 按使用区域汇总", "打开文件失败", CStr(selectedFile), "", "", "错误", Err.Number & " " & Err.Description, ""
            On Error GoTo ErrHandle
        End If
        On Error GoTo ErrHandle

        wbName = sourceWb.Name
        If InStrRev(wbName, ".") > 0 Then wbName = Left(wbName, InStrRev(wbName, ".") - 1)

        For Each sourceWs In sourceWb.Worksheets
            On Error Resume Next
            Set usedRng = sourceWs.UsedRange
            On Error GoTo ErrHandle

            If Not usedRng Is Nothing Then
                If usedRng.Rows.count >= 1 And usedRng.Columns.count >= 1 Then
                    dataCols = usedRng.Columns.count
                    If dataCols > MAX_DATA_COLS Then dataCols = MAX_DATA_COLS

                    If headerCols = 0 Then
                        headerCols = dataCols
                        sumWs.Cells(1, 1).value = "工作簿"
                        sumWs.Cells(1, 2).value = "工作表"
                        sumWs.Cells(1, 3).value = "行"
                        For j = 1 To headerCols
                            sumWs.Cells(1, 3 + j).value = "列" & j
                        Next j
                        sumWs.Rows(1).Font.Bold = True
                    End If

                    For i = 1 To usedRng.Rows.count
                        If skipHeader And i = 1 Then
                            If headerOutputOnce Then GoTo nextRow
                            headerOutputOnce = True
                        End If

                        sumWs.Cells(destRow, 1).value = wbName
                        sumWs.Cells(destRow, 2).value = sourceWs.Name
                        sumWs.Cells(destRow, 3).value = i

                        For j = 1 To headerCols
                            If j <= dataCols Then
                                sumWs.Cells(destRow, 3 + j).value = GetCellValue(usedRng, i, j)
                            Else
                                sumWs.Cells(destRow, 3 + j).value = Empty
                            End If
                        Next j

                        destRow = destRow + 1
nextRow:
                    Next i
                End If
            End If
        Next sourceWs

        RunLog_WriteRow "2.2.1 按使用区域汇总", "打开文件", sourceWb.Name, "", "", "成功", "已处理", ""
        sourceWb.Close SaveChanges:=False
    Next selectedFile

    If destRow <= 2 Then
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        newWb.Close SaveChanges:=False
        RunLog_WriteRow "2.2.1 按使用区域汇总", "结束", "", "", "", "", "无数据", CStr(Round(Timer - t0, 2))
        MsgBox "所选文件中没有可汇总的数据。", vbInformation
        Exit Sub
    End If

    sumWs.Columns.AutoFit
    sumWs.UsedRange.Columns.AutoFit

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="汇总_" & Format(Now, "yyyy-mm-dd_hh-nn-ss"), _
        FileFilter:="Excel 工作簿(*.xlsx), *.xlsx", _
        Title:="另存为")

    RunLog_WriteRow "2.2.1 按使用区域汇总", "结束", "", "", "", "", "Done 共 " & (destRow - 2) & " 行", CStr(Round(Timer - t0, 2))
    If VarType(savePath) = vbBoolean Then
        If savePath = False Then
            MsgBox "已取消保存，未写入文件。", vbInformation
        End If
    Else
        newWb.SaveAs fileName:=CStr(savePath)
        MsgBox "共汇总 " & (destRow - 2) & " 行。" & vbCrLf & "已保存至：" & vbCrLf & CStr(savePath), vbInformation
    End If

    Exit Sub

ErrHandle:
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    If Not sourceWb Is Nothing Then
        On Error Resume Next
        sourceWb.Close SaveChanges:=False
    End If
    MsgBox "错误 " & Err.Number & ": " & Err.Description, vbCritical
End Sub
