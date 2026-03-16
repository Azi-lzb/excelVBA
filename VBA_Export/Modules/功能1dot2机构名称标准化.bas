Attribute VB_Name = "功能1dot2机构名称标准化"


Function 获取机构映射表字典() As Object
    '功能：从当前工作簿的"机构映射表"中读取A、B列数据，返回字典对象
    '返回值：Scripting.Dictionary，键为A列原始机构名，值为B列映射后机构名
    
    Dim mappingDict As Object
    Dim mappingWs As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim key As String, value As String
    
    '创建字典对象
    Set mappingDict = CreateObject("Scripting.Dictionary")
    
    '设置字典比较模式为文本模式（不区分大小写）
    mappingDict.CompareMode = vbTextCompare
    
    On Error GoTo ErrorHandler
    
    '获取机构映射表工作表
    Set mappingWs = ThisWorkbook.Sheets("机构映射表")
    
    '获取最后一行数据
    lastRow = mappingWs.Cells(mappingWs.Rows.count, 1).End(xlUp).row
    
    '从第2行开始读取（假设第1行是标题）
    For i = 2 To lastRow
        key = Trim(mappingWs.Cells(i, 1).value)     'A列：原始机构名
        value = Trim(mappingWs.Cells(i, 2).value)   'B列：映射后机构名
        
        '确保键值都不为空
        If key <> "" And value <> "" Then
            '添加到字典，如果键已存在则覆盖
            mappingDict(key) = value
        End If
    Next i
    
    '返回字典
    Set 获取机构映射表字典 = mappingDict
    
    Exit Function
    
ErrorHandler:
    '错误处理
    Select Case Err.Number
        Case 9 '工作表不存在
            MsgBox "错误：未找到名为'机构映射表'的工作表！" & vbCrLf & _
                   "请在当前工作簿中创建该工作表，并确保格式为：" & vbCrLf & _
                   "A列：原始机构名称" & vbCrLf & _
                   "B列：映射后机构名称", vbCritical
        Case Else
            MsgBox "读取机构映射表时发生错误：" & vbCrLf & Err.Description, vbCritical
    End Select
    
    '返回空字典
    Set 获取机构映射表字典 = CreateObject("Scripting.Dictionary")
End Function





'辅助函数：根据字典获取映射后的名称
Function 获取映射名称(originalName As String, mappingDict As Object) As String
    Dim key As Variant
    
    '先在字典中查找精确匹配
    If mappingDict.Exists(originalName) Then
        获取映射名称 = mappingDict(originalName)
        Exit Function
    End If
    
    '查找包含关系（不区分大小写）
    For Each key In mappingDict.keys
        If InStr(1, originalName, key, vbTextCompare) > 0 Then
            获取映射名称 = mappingDict(key)
            Exit Function
        End If
    Next key
    
    '如果没有找到映射，返回原始名称
    获取映射名称 = originalName
End Function
'简单测试一下功能
Sub 测试映射表有没有问题()
    Dim mappingDict As Object
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim totalMapped As Long
    Dim processedSheets As Integer
    Dim str As String
    
    '获取机构映射表字典
    Set mappingDict = 获取机构映射表字典()
    str = 获取映射名称("广州市商业银行", mappingDict)
    Debug.Print str
End Sub



Sub 批量修改分机构文件()
    Dim mappingDict As Object
    Dim fd As FileDialog
    Dim selectedFile As Variant
    Dim sourceWb As Workbook
    Dim sourceWs As Worksheet
    Dim totalMapped As Long
    Dim processedFiles As Integer
    Dim t0 As Double
    Dim 本文件映射数 As Long
    Dim mappedCount As Long

    t0 = Timer
    RunLog_WriteRow "1.2 机构名称标准化", "开始", "", "", "", "", "开始", ""

    '获取机构映射表字典
    Set mappingDict = 获取机构映射表字典()
    If mappingDict.count = 0 Then
        RunLog_WriteRow "1.2 机构名称标准化", "完成", "", "", "", "失败", "机构映射表为空", CStr(Round(Timer - t0, 2))
        MsgBox "机构映射表为空，无法继续处理", vbExclamation
        Exit Sub
    End If
    
    '选择要处理的Excel文件
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "选择需要将机构名称标准化的Excel文件"
    fd.AllowMultiSelect = True
    fd.Filters.Add "Excel文件", "*.xls;*.xlsx;*.xlsm"
    
    If fd.Show Then
        Application.ScreenUpdating = False
        Application.DisplayAlerts = False
        
        totalMapped = 0
        processedFiles = 0
        
        For Each selectedFile In fd.SelectedItems
            本文件映射数 = 0
            '打开源工作簿
            On Error Resume Next
            Set sourceWb = Workbooks.Open(selectedFile)
            If Err.Number <> 0 Then
                RunLog_WriteRow "1.2 机构名称标准化", "处理文件", CStr(selectedFile), "", "", "失败", Err.Number & " " & Err.Description, ""
                Debug.Print "无法打开文件: " & selectedFile & " - " & Err.Description
                On Error GoTo 0
                GoTo NextFile
            End If
            On Error GoTo 0
            
            processedFiles = processedFiles + 1
            
            '遍历所有工作表
            For Each sourceWs In sourceWb.Worksheets
                '确保sourceWs变量已正确设置
                If Not sourceWs Is Nothing Then
                    '只处理包含"分机构"的工作表
                    If InStr(1, sourceWs.Name, "分机构", vbTextCompare) > 0 Then
                        mappedCount = 映射分机构工作表(sourceWs, mappingDict)
                        totalMapped = totalMapped + mappedCount
                        本文件映射数 = 本文件映射数 + mappedCount
                        
                        If mappedCount > 0 Then
                            Debug.Print "文件: " & sourceWb.Name & " | 工作表: " & sourceWs.Name & " | 修改: " & mappedCount & " 行"
                        End If
                    End If
                End If
            Next sourceWs
            
            RunLog_WriteRow "1.2 机构名称标准化", "处理文件", sourceWb.Name, "", "", "成功", "映射 " & 本文件映射数 & " 行", ""
            '保存修改并关闭工作簿
            sourceWb.Close SaveChanges:=True
            
NextFile:
        Next selectedFile
        
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        
        RunLog_WriteRow "1.2 机构名称标准化", "完成", "", "", "", "", "Done", CStr(Round(Timer - t0, 2))
        '显示结果
        MsgBox "分机构数据修改完成！" & vbCrLf & _
               "处理文件数: " & processedFiles & vbCrLf & _
               "总修改数量: " & totalMapped & vbCrLf & _
               "机构名称已根据映射表直接更新", vbInformation
    Else
        RunLog_WriteRow "1.2 机构名称标准化", "完成", "", "", "", "", "未选择文件", CStr(Round(Timer - t0, 2))
        MsgBox "未选择任何文件", vbInformation
    End If
    
    Set mappingDict = Nothing
End Sub

'映射函数（确保变量正确设置）
Function 映射分机构工作表(ws As Worksheet, mappingDict As Object) As Long
    '功能：对包含"分机构"的工作表进行机构名称映射，直接修改源数据
    '参数：
    '   ws: 要处理的工作表对象
    '   mappingDict: 机构名称映射字典
    '   列A标题: 可选的A列标题名称（用于定位数据区域）
    '   列B标题: 可选的B列标题名称（用于定位数据区域）
    '返回值：成功映射的行数（实际修改的数据行数）
    
    '检查参数是否有效
    If ws Is Nothing Then
        映射分机构工作表 = 0
        Exit Function
    End If
    
    If mappingDict Is Nothing Then
        映射分机构工作表 = 0
        Exit Function
    End If
    
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim mappedCount As Long
    Dim dataStartRow As Long
    Dim 需要映射A列 As Boolean, 需要映射B列 As Boolean
    
    '初始化
    mappedCount = 0
    
    '检查工作表是否包含"分机构"字样
    If InStr(1, ws.Name, "分机构", vbTextCompare) = 0 Then
        映射分机构工作表 = 0
        Exit Function
    End If
    
    '获取工作表数据范围（机构名可能在A列或B列，取两列最大行数避免B列有数据而A列为空时被跳过）
    On Error Resume Next
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row
    If ws.Cells(ws.Rows.count, 2).End(xlUp).row > lastRow Then lastRow = ws.Cells(ws.Rows.count, 2).End(xlUp).row
    lastCol = ws.Cells(1, ws.Columns.count).End(xlToLeft).Column
    On Error GoTo 0
    
    '如果数据不足2行，直接返回
    If lastRow < 2 Then
        映射分机构工作表 = 0
        Exit Function
    End If
    
    '确定数据起始行（跳过标题行）
    dataStartRow = 2 '默认从第2行开始
    
    '查找A列和B列
    列A索引 = 1
    列B索引 = 2
    
    '检查A列是否为空
    需要映射A列 = False
    需要映射B列 = False
    
    '优先检查A列是否有数据
    For i = dataStartRow To lastRow
        If Trim(ws.Cells(i, 列A索引).value) <> "" Then
            需要映射A列 = True
            Exit For
        End If
    Next i
    
    '如果A列没有数据，则检查B列
    If Not 需要映射A列 Then
        For i = dataStartRow To lastRow
            If Trim(ws.Cells(i, 列B索引).value) <> "" Then
                需要映射B列 = True
                Exit For
            End If
        Next i
    End If
    
    '执行映射（直接修改工作表数据）
    If 需要映射A列 Then
        '映射A列
        For i = dataStartRow To lastRow
            Dim originalName As String
            Dim mappedName As String
            
            originalName = Trim(ws.Cells(i, 列A索引).value)
            If originalName <> "" Then
                '获取映射后的名称
                mappedName = 获取映射名称(originalName, mappingDict)
                
                '如果映射后的名称与原始名称不同，则修改单元格
                If mappedName <> originalName Then
                    ws.Cells(i, 列A索引).value = mappedName
                    mappedCount = mappedCount + 1
                End If
            End If
        Next i
    ElseIf 需要映射B列 Then
        '映射B列
        For i = dataStartRow To lastRow
            originalName = Trim(ws.Cells(i, 列B索引).value)
            If originalName <> "" Then
                '获取映射后的名称
                mappedName = 获取映射名称(originalName, mappingDict)
                
                '如果映射后的名称与原始名称不同，则修改单元格
                If mappedName <> originalName Then
                    ws.Cells(i, 列B索引).value = mappedName
                    mappedCount = mappedCount + 1
                End If
            End If
        Next i
    End If
    
    映射分机构工作表 = mappedCount
End Function

