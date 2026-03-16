Attribute VB_Name = "功能1dot4外汇页眉修改"
Function ProcessSingleFile(filePath As String, ByRef sheetsProcessed As Integer) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileSheetsProcessed As Integer
    
    On Error GoTo ErrorHandler
    
    ' 初始化计数器
    fileSheetsProcessed = 0
    
    ' 打开工作簿
    Set wb = Workbooks.Open(filePath, ReadOnly:=False)
    
    ' 遍历所有工作表
    For Each ws In wb.Worksheets
        ' 检查工作表名称是否包含"外汇"
        If WorksheetContainsKeyword(ws, "外汇") Then
            ' 调整该工作表的页眉
            AdjustWorksheetHeader ws
            fileSheetsProcessed = fileSheetsProcessed + 1
        End If
    Next ws
    
    ' 设置返回的计数
    sheetsProcessed = fileSheetsProcessed
    
    ' 如果有工作表被处理，保存工作簿
    If fileSheetsProcessed > 0 Then
        wb.Save
        Debug.Print "? 文件: " & GetFileName(filePath) & " - 处理了 " & fileSheetsProcessed & " 个工作表"
    Else
        Debug.Print "○ 文件: " & GetFileName(filePath) & " - 未找到包含'外汇'的工作表"
    End If
    
    ' 关闭工作簿
    wb.Close SaveChanges:=False
    
    ProcessSingleFile = True
    Exit Function
    
ErrorHandler:
    ' 错误处理
    Debug.Print "? 错误: " & GetFileName(filePath) & " - " & Err.Description
    sheetsProcessed = 0
    
    ' 尝试关闭工作簿
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    
    ProcessSingleFile = False
End Function

' 这是原来的代码，我已经修复了ByRef参数问题，下面是完整的正确版本：

Sub 批量调整外汇Sheet页眉_优化版()
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim processedFiles As Integer
    Dim processedSheets As Integer
    Dim filePath As String
    Dim localSheetsProcessed As Integer
    Dim t0 As Double
    Dim fName As String
    
    t0 = Timer
    RunLog_WriteRow "1.4 修改外汇表页眉", "开始", "", "", "", "", "开始", ""
    
    ' 初始化计数器
    processedFiles = 0
    processedSheets = 0
    
    ' 创建文件选择对话框
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "请选择要处理的Excel文件（可多选）"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xls; *.xlsx; *.xlsm; *.xlsb"
        .ButtonName = "开始处理"
        
        If .Show = -1 Then
            ' 开始处理
            Application.ScreenUpdating = False
            Application.DisplayAlerts = False
            
            For Each fileItem In .SelectedItems
                filePath = CStr(fileItem)
                localSheetsProcessed = 0  ' 局部变量记录当前文件处理的工作表数
                fName = GetFileName(filePath)
                If fName = "" Then fName = filePath
                
                ' 调用函数处理文件
                If ProcessSingleFileSimple(filePath, localSheetsProcessed) Then
                    processedFiles = processedFiles + 1
                    processedSheets = processedSheets + localSheetsProcessed
                    RunLog_WriteRow "1.4 修改外汇表页眉", "处理文件", fName, "", "", "成功", "处理 " & localSheetsProcessed & " 个外汇表", ""
                Else
                    RunLog_WriteRow "1.4 修改外汇表页眉", "处理文件", fName, "", "", "失败", "打开或处理失败", ""
                End If
            Next fileItem
            
            ' 恢复设置
            Application.ScreenUpdating = True
            Application.DisplayAlerts = True
            
            RunLog_WriteRow "1.4 修改外汇表页眉", "完成", "", "", "", "", "Done", CStr(Round(Timer - t0, 2))
            ' 显示结果
            MsgBox "处理完成！" & vbCrLf & _
                   "已处理文件数: " & processedFiles & vbCrLf & _
                   "已处理工作表数: " & processedSheets, _
                   vbInformation, "处理结果"
        Else
            RunLog_WriteRow "1.4 修改外汇表页眉", "完成", "", "", "", "", "已取消", CStr(Round(Timer - t0, 2))
            MsgBox "操作已取消", vbInformation
        End If
    End With
    
    Set fd = Nothing
End Sub

Function ProcessSingleFileSimple(filePath As String, ByRef sheetsProcessed As Integer) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim count As Integer
    
    On Error GoTo ErrorHandler
    
    ' 初始化
    count = 0
    Set wb = Workbooks.Open(filePath, ReadOnly:=False)
    
    ' 遍历工作表
    For Each ws In wb.Worksheets
        ' 检查是否包含"外汇"
        If InStr(1, ws.Name, "外汇", vbTextCompare) > 0 Then
            ' 处理页眉
            Call ProcessHeader(ws)
            count = count + 1
        End If
    Next ws
    
    ' 设置返回的工作表计数
    sheetsProcessed = count
    
    ' 保存并关闭
    If count > 0 Then
        wb.Save
    End If
    
    wb.Close SaveChanges:=False
    ProcessSingleFileSimple = True
    
    Exit Function
    
ErrorHandler:
    sheetsProcessed = 0
    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If
    ProcessSingleFileSimple = False
End Function

Sub ProcessHeader(ws As Worksheet)
    ' 处理左页眉
    If Len(ws.PageSetup.leftHeader) > 0 Then
        ws.PageSetup.leftHeader = ReplaceHeader(ws.PageSetup.leftHeader)
    End If
    
    ' 处理中页眉
    If Len(ws.PageSetup.centerHeader) > 0 Then
        ws.PageSetup.centerHeader = ReplaceHeader(ws.PageSetup.centerHeader)
    End If
    
    ' 处理右页眉
    If Len(ws.PageSetup.rightHeader) > 0 Then
        ws.PageSetup.rightHeader = ReplaceHeader(ws.PageSetup.rightHeader)
    End If
End Sub

Function ReplaceHeader(headerText As String) As String
    Dim result As String
    result = headerText
    
    ' 主要替换规则
    result = Replace(result, "万元", "万美元")
    
    ' 其他可能的形式
    result = Replace(result, "（万元）", "（万美元）")
    result = Replace(result, "(万元)", "(万美元)")
    result = Replace(result, "[万元]", "[万美元]")
    
    ' 单位说明
    result = Replace(result, "单位:万元", "单位:万美元")
    result = Replace(result, "单位：万元", "单位：万美元")
    
    ReplaceHeader = result
End Function

' 最简单直接的版本（推荐使用这个）
Sub 批量调整外汇Sheet页眉_直接版()
    Dim fd As FileDialog
    Dim fileItem As Variant
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim fileCount As Integer
    Dim sheetCount As Integer
    
    ' 初始化
    fileCount = 0
    sheetCount = 0
    
    ' 创建文件对话框
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "请选择要处理的Excel文件"
        .AllowMultiSelect = True
        .Filters.Clear
        .Filters.Add "Excel文件", "*.xls; *.xlsx; *.xlsm"
        
        If .Show = -1 Then
            ' 开始处理
            Application.ScreenUpdating = False
            
            For Each fileItem In .SelectedItems
                ' 打开文件
                Set wb = Workbooks.Open(CStr(fileItem))
                
                ' 遍历工作表
                For Each ws In wb.Worksheets
                    If InStr(1, ws.Name, "外汇", vbTextCompare) > 0 Then
                        ' 处理页眉
                        处理工作表页眉 ws
                        sheetCount = sheetCount + 1
                    End If
                Next ws
                
                ' 保存并关闭
                wb.Save
                wb.Close
                fileCount = fileCount + 1
                
                ' 显示进度
                Debug.Print "已处理: " & GetFileName(CStr(fileItem))
            Next fileItem
            
            Application.ScreenUpdating = True
            
            ' 显示结果
            MsgBox "成功处理 " & fileCount & " 个文件，" & sheetCount & " 个工作表", vbInformation
        End If
    End With
    
    Set fd = Nothing
End Sub

Sub 处理工作表页眉(ws As Worksheet)
    ' 这个函数直接处理工作表的页眉，不需要ByRef参数
    Dim leftHeader As String, centerHeader As String, rightHeader As String
    
    ' 获取当前页眉
    leftHeader = ws.PageSetup.leftHeader
    centerHeader = ws.PageSetup.centerHeader
    rightHeader = ws.PageSetup.rightHeader
    
    ' 替换"万元"为"万美元"
    leftHeader = Replace(leftHeader, "万元", "万美元")
    centerHeader = Replace(centerHeader, "万元", "万美元")
    rightHeader = Replace(rightHeader, "万元", "万美元")
    
    ' 额外的替换规则（可选）
    leftHeader = Replace(leftHeader, "（万元）", "（万美元）")
    centerHeader = Replace(centerHeader, "（万元）", "（万美元）")
    rightHeader = Replace(rightHeader, "（万元）", "（万美元）")
    
    ' 设置回工作表
    ws.PageSetup.leftHeader = leftHeader
    ws.PageSetup.centerHeader = centerHeader
    ws.PageSetup.rightHeader = rightHeader
End Sub

Function GetFileName(filePath As String) As String
    ' 提取文件名
    Dim pos As Integer
    pos = InStrRev(filePath, "\")
    If pos > 0 Then
        GetFileName = Mid(filePath, pos + 1)
    Else
        GetFileName = filePath
    End If
End Function

