Attribute VB_Name = "功能1dot3剔除外资行"


'从本地机构映射表加载外资行信息
Function 加载外资行信息() As Object
    Dim mappingWs As Worksheet
    Dim foreignBankDict As Object
    Dim lastRow As Long
    Dim i As Long
    Dim bankName As String
    Dim isForeign As Integer
    
    Set foreignBankDict = CreateObject("Scripting.Dictionary")
    foreignBankDict.CompareMode = vbTextCompare '不区分大小写
    
    On Error GoTo ErrorHandler
    
    '获取机构映射表工作表
    Set mappingWs = ThisWorkbook.Sheets("机构映射表")
    
    '获取最后一行数据
    lastRow = mappingWs.Cells(mappingWs.Rows.count, 1).End(xlUp).row
    
    '从第2行开始读取（假设第1行是标题）
    For i = 2 To lastRow
        bankName = Trim(mappingWs.Cells(i, 2).value) 'B列：标准化机构名称
        isForeign = val(mappingWs.Cells(i, 3).value)  'C列：是否为外资行（1=是，0=否）
        
        If bankName <> "" Then
            '将机构名称添加到字典，值为是否为外资行
            foreignBankDict(bankName) = isForeign
        End If
    Next i
    
    Debug.Print "加载外资行配置: " & foreignBankDict.count & " 个机构"
    
    Set 加载外资行信息 = foreignBankDict
    Exit Function
    
ErrorHandler:
    MsgBox "加载机构映射表时发生错误: " & Err.Description, vbCritical
    Set 加载外资行信息 = CreateObject("Scripting.Dictionary")
End Function

'在工作表中删除外资行
Function 删除外资行(ws As Worksheet, foreignBankDict As Object) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim deletedRows As Long
    Dim bankName As String
    Dim colA As Variant
    Dim marked() As Boolean
    Dim startRow As Long, endRow As Long

    deletedRows = 0
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 1 Then Exit Function

    colA = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)).Value2
    ReDim marked(1 To lastRow)

    For i = 1 To lastRow
        bankName = Trim(CStr(colA(i, 1)))
        If bankName <> "" Then
            If foreignBankDict.Exists(bankName) Then
                If CLng(Val(foreignBankDict(bankName))) = 1 Then
                    marked(i) = True
                End If
            End If
        End If
    Next i

    i = lastRow
    Do While i >= 1
        If marked(i) Then
            endRow = i
            Do While i >= 1 And marked(i)
                i = i - 1
            Loop
            startRow = i + 1
            ws.Rows(startRow & ":" & endRow).Delete Shift:=xlUp
            deletedRows = deletedRows + (endRow - startRow + 1)
        Else
            i = i - 1
        End If
    Loop
    
    删除外资行 = deletedRows
End Function

'增强版本：支持模糊匹配机构名称
Function 删除外资行增强版(ws As Worksheet, foreignBankDict As Object) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim deletedRows As Long
    Dim bankName As String
    Dim dictKey As Variant
    Dim dictKeys As Variant
    Dim colA As Variant
    Dim marked() As Boolean
    Dim startRow As Long, endRow As Long

    deletedRows = 0
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If lastRow < 1 Then Exit Function

    dictKeys = foreignBankDict.keys
    colA = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, 1)).Value2
    ReDim marked(1 To lastRow)

    For i = 1 To lastRow
        bankName = Trim(CStr(colA(i, 1)))
        If bankName <> "" Then
            If foreignBankDict.Exists(bankName) Then
                If CLng(Val(foreignBankDict(bankName))) = 1 Then
                    marked(i) = True
                End If
            Else
                For Each dictKey In dictKeys
                    If InStr(1, bankName, CStr(dictKey), vbTextCompare) > 0 Then
                        If CLng(Val(foreignBankDict(dictKey))) = 1 Then
                            marked(i) = True
                            Exit For
                        End If
                    End If
                Next dictKey
            End If
        End If
    Next i

    i = lastRow
    Do While i >= 1
        If marked(i) Then
            endRow = i
            Do While i >= 1 And marked(i)
                i = i - 1
            Loop
            startRow = i + 1
            ws.Rows(startRow & ":" & endRow).Delete Shift:=xlUp
            deletedRows = deletedRows + (endRow - startRow + 1)
        Else
            i = i - 1
        End If
    Loop
    
    删除外资行增强版 = deletedRows
End Function

'测试函数：显示外资行配置
Sub 显示外资行配置()
    Dim foreignBankDict As Object
    Dim bankName As Variant
    Dim foreignCount As Integer, totalCount As Integer
    
    Set foreignBankDict = 加载外资行信息()
    
    If foreignBankDict.count = 0 Then
        MsgBox "未找到外资行配置信息", vbInformation
        Exit Sub
    End If
    
    foreignCount = 0
    totalCount = foreignBankDict.count
    
    Debug.Print "外资行配置信息（共 " & totalCount & " 个机构）："
    Debug.Print "机构名称 | 是否为外资行（1=是，0=否）"
    Debug.Print String(50, "-")
    
    For Each bankName In foreignBankDict.keys
        If foreignBankDict(bankName) = 1 Then
            foreignCount = foreignCount + 1
            Debug.Print bankName & " | 是（外资行）"
        Else
            Debug.Print bankName & " | 否"
        End If
    Next bankName
    
    Debug.Print String(50, "-")
    Debug.Print "外资行数量: " & foreignCount & " / " & totalCount
    
    MsgBox "外资行配置信息已显示在立即窗口（按Ctrl+G查看）" & vbCrLf & _
           "外资行数量: " & foreignCount & " / " & totalCount, vbInformation
End Sub

'简化版本：快速处理单个文件
Sub 删除外资行_单个()
    Dim sourceWb As Workbook, tempWb As Workbook
    Dim sourceWs As Worksheet, tempWs As Worksheet
    Dim foreignBankDict As Object
    Dim deletedRows As Long
    Dim savePath As String, originalName As String, newName As String
    
    '从本地机构映射表加载外资行信息
    Set foreignBankDict = 加载外资行信息()
    If foreignBankDict.count = 0 Then Exit Sub
    
    '选择单个文件
    Dim selectedFile As Variant
    selectedFile = Application.GetOpenFilename( _
        FileFilter:="Excel Files (*.xls;*.xlsx;*.xlsm), *.xls;*.xlsx;*.xlsm", _
        Title:="选择要处理的Excel文件" _
    )
    
    If selectedFile = False Then Exit Sub
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '打开源工作簿
    Set sourceWb = Workbooks.Open(selectedFile, ReadOnly:=True)
    
    '创建临时工作簿
    Set tempWb = Workbooks.Add
    deletedRows = 0
    
    '复制所有工作表
    For Each sourceWs In sourceWb.Worksheets
        sourceWs.Copy After:=tempWb.Sheets(tempWb.Sheets.count)
        Set tempWs = tempWb.Sheets(tempWb.Sheets.count)
        
        '删除分机构表中的外资行
        If InStr(1, tempWs.Name, "分机构", vbTextCompare) > 0 Then
            deletedRows = deletedRows + 删除外资行(tempWs, foreignBankDict)
        End If
    Next sourceWs
    
    '删除默认空白工作表
    Do While tempWb.Sheets.count > sourceWb.Sheets.count
        tempWb.Sheets(1).Delete
    Loop
    
    '生成新文件名
    originalName = Mid(sourceWb.Name, 1, InStrRev(sourceWb.Name, ".") - 1)
    newName = originalName & "(金融局)" & Mid(sourceWb.Name, InStrRev(sourceWb.Name, "."))
    savePath = Left(selectedFile, InStrRev(selectedFile, "\")) & newName
    
    '保存副本
    tempWb.SaveAs fileName:=savePath, FileFormat:=sourceWb.FileFormat
    
    '关闭工作簿
    sourceWb.Close False
    tempWb.Close True
    
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
    MsgBox "处理完成！" & vbCrLf & _
           "删除外资行数: " & deletedRows & vbCrLf & _
           "副本已保存为: " & newName, vbInformation
End Sub


'批量版本：快速处理多个文件
Sub 删除外资行_批量()
    Dim fd As FileDialog
    Dim selectedFile As Variant
    Dim sourceWb As Workbook, tempWb As Workbook
    Dim sourceWs As Worksheet, tempWs As Worksheet
    Dim foreignBankDict As Object
    Dim deletedRows As Long, totalDeletedRows As Long
    Dim savePath As String, originalName As String, newName As String
    Dim fileCount As Integer, 分机构工作表数 As Integer
    Dim 保存成功文件数 As Integer
    Dim t0 As Double

    t0 = Timer
    RunLog_WriteRow "1.3 删除分机构表的外资行", "开始", "", "", "", "", "开始", ""

    '从本地机构映射表加载外资行信息
    Set foreignBankDict = 加载外资行信息()
    If foreignBankDict.count = 0 Then
        RunLog_WriteRow "1.3 删除分机构表的外资行", "完成", "", "", "", "失败", "未找到外资行配置", CStr(Round(Timer - t0, 2))
        MsgBox "未找到外资行配置信息，请检查'机构映射表'工作表", vbExclamation
        Exit Sub
    End If
    
    '选择多个Excel文件
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "选择要处理的Excel文件（可多选）"
    fd.AllowMultiSelect = True
    fd.Filters.Add "Excel Files", "*.xls; *.xlsx; *.xlsm"
    
    If fd.Show Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        
        fileCount = 0
        分机构工作表数 = 0
        totalDeletedRows = 0
        保存成功文件数 = 0
        
        For Each selectedFile In fd.SelectedItems
            fileCount = fileCount + 1
            deletedRows = 0
            
            '打开源工作簿
            On Error Resume Next
            Set sourceWb = Workbooks.Open(selectedFile, ReadOnly:=True)
            If Err.Number <> 0 Then
                RunLog_WriteRow "1.3 删除分机构表的外资行", "处理文件", CStr(selectedFile), "", "", "失败", Err.Number & " " & Err.Description, ""
                Debug.Print "无法打开文件: " & selectedFile & " - " & Err.Description
                On Error GoTo 0
                GoTo NextFile
            End If
            On Error GoTo 0
            
            '创建临时工作簿
            Set tempWb = Workbooks.Add
            deletedRows = 0
            
            '复制所有工作表
            For Each sourceWs In sourceWb.Worksheets
                sourceWs.Copy After:=tempWb.Sheets(tempWb.Sheets.count)
                Set tempWs = tempWb.Sheets(tempWb.Sheets.count)
                
                '删除分机构表中的外资行
                If InStr(1, tempWs.Name, "分机构", vbTextCompare) > 0 Then
                    分机构工作表数 = 分机构工作表数 + 1
                    deletedRows = deletedRows + 删除外资行(tempWs, foreignBankDict)
                End If
            Next sourceWs
            
            '删除默认空白工作表
            Do While tempWb.Sheets.count > sourceWb.Sheets.count
                Application.DisplayAlerts = False
                tempWb.Sheets(1).Delete
                Application.DisplayAlerts = True
            Loop
            
            totalDeletedRows = totalDeletedRows + deletedRows
            
            '生成新文件名
            originalName = Mid(sourceWb.Name, 1, InStrRev(sourceWb.Name, ".") - 1)
            newName = originalName & "(金融局)" & Mid(sourceWb.Name, InStrRev(sourceWb.Name, "."))
            savePath = Left(selectedFile, InStrRev(selectedFile, "\")) & newName
            
            '保存副本
            Dim saveOk As Boolean
            Dim saveErrNum As Long
            saveOk = False
            saveErrNum = 0
            On Error Resume Next
            Err.Clear
            tempWb.SaveAs fileName:=savePath, FileFormat:=sourceWb.FileFormat
            If Err.Number = 0 Then
                saveOk = True
            Else
                Err.Clear
                tempWb.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook
                saveOk = (Err.Number = 0)
                saveErrNum = Err.Number
            End If
            On Error GoTo 0
            
            If saveOk Then
                保存成功文件数 = 保存成功文件数 + 1
                RunLog_WriteRow "1.3 删除分机构表的外资行", "处理文件", newName, "", "", "成功", "删除 " & deletedRows & " 行外资行", ""
                Debug.Print "文件 " & fileCount & ": " & newName & " (删除 " & deletedRows & " 行外资行)"
            Else
                RunLog_WriteRow "1.3 删除分机构表的外资行", "处理文件", newName, "", "", "失败", "保存失败 " & saveErrNum, ""
                Debug.Print "文件 " & fileCount & ": 保存失败 - " & newName
            End If
            
            '关闭工作簿
            sourceWb.Close False
            tempWb.Close True
            Set sourceWb = Nothing
            Set tempWb = Nothing
            
            
NextFile:
        Next selectedFile
        
        '恢复Excel设置
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.CutCopyMode = False
        
        RunLog_WriteRow "1.3 删除分机构表的外资行", "完成", "", "", "", "", "Done", CStr(Round(Timer - t0, 2))
        '显示处理结果
        Dim resultMsg As String
        resultMsg = "批量处理完成！" & vbCrLf & vbCrLf
        resultMsg = resultMsg & "选择文件数: " & fd.SelectedItems.count & vbCrLf
        resultMsg = resultMsg & "成功处理文件数: " & 保存成功文件数 & vbCrLf
        resultMsg = resultMsg & "处理分机构工作表数: " & 分机构工作表数 & vbCrLf
        resultMsg = resultMsg & "总删除外资行数: " & totalDeletedRows & vbCrLf & vbCrLf
        resultMsg = resultMsg & "副本命名格式: 原文件名(金融局).xlsx"
        
        MsgBox resultMsg, vbInformation, "批量处理完成"
    Else
        RunLog_WriteRow "1.3 删除分机构表的外资行", "完成", "", "", "", "", "未选择文件", CStr(Round(Timer - t0, 2))
        MsgBox "未选择任何文件", vbInformation
    End If
    
    Set foreignBankDict = Nothing
End Sub
