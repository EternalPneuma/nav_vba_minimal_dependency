Option Explicit

'==============================================================
' 模块: 删除当前工作表空行
' 功能: 按用户指定的一列或多列判断空值, 删除当前活动工作表中这些列均为空的整行
'==============================================================

Public Sub Tool06_DeleteEmptyRows()
    
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    If ws Is Nothing Then
        MsgBox "删除当前工作表空行无法继续" & vbCrLf & vbCrLf & _
               "错误信息：未找到当前活动工作表。", vbExclamation, "删除当前工作表空行"
        Exit Sub
    End If
    
    Dim headerRowInput As String
    headerRowInput = InputBox("请输入表头所在行号：" & vbCrLf & vbCrLf & _
                              "默认值：1" & vbCrLf & _
                              "说明：如果按表头名指定列，将用这一行匹配字段名。", _
                              "删除当前工作表空行", "1")
    
    If Len(headerRowInput) = 0 Then Exit Sub
    
    Dim headerRow As Long
    If Not IsNumeric(headerRowInput) Then
        MsgBox "删除当前工作表空行无法继续" & vbCrLf & vbCrLf & _
               "错误信息：表头行号必须是数字。", vbExclamation, "删除当前工作表空行"
        Exit Sub
    End If
    
    headerRow = CLng(headerRowInput)
    If headerRow < 1 Or headerRow > ws.Rows.Count Then
        MsgBox "删除当前工作表空行无法继续" & vbCrLf & vbCrLf & _
               "错误信息：表头行号超出有效范围。", vbExclamation, "删除当前工作表空行"
        Exit Sub
    End If
    
    Dim criteriaInput As String
    criteriaInput = InputBox("请输入用于判断空行的列：" & vbCrLf & vbCrLf & _
                             "可输入列字母、列号或表头名，多个条件用逗号分隔。" & vbCrLf & _
                             "示例：A" & vbCrLf & _
                             "示例：A,C,F" & vbCrLf & _
                             "示例：产品编号,净值日期" & vbCrLf & vbCrLf & _
                             "删除规则：指定的列全部为空时，删除该行。", _
                             "删除当前工作表空行")
    
    If Len(criteriaInput) = 0 Then Exit Sub
    
    Dim criteriaCols() As Long
    Dim criteriaNames() As String
    If Not ParseColumnCriteria(ws, headerRow, criteriaInput, criteriaCols, criteriaNames) Then
        Exit Sub
    End If
    
    Dim lastRow As Long
    Dim lastCol As Long
    lastRow = GetLastUsedRow(ws)
    lastCol = GetLastUsedCol(ws)
    
    If lastRow <= headerRow Then
        MsgBox "删除当前工作表空行无需处理" & vbCrLf & vbCrLf & _
               "处理结果：" & vbCrLf & _
               "当前工作表没有可处理的数据行。", vbInformation, "删除当前工作表空行"
        Exit Sub
    End If
    
    If lastCol = 0 Then
        MsgBox "删除当前工作表空行无需处理" & vbCrLf & vbCrLf & _
               "处理结果：" & vbCrLf & _
               "当前工作表没有已使用的单元格。", vbInformation, "删除当前工作表空行"
        Exit Sub
    End If
    
    Dim dataStartRow As Long
    dataStartRow = headerRow + 1
    
    Dim delRows() As Long
    ReDim delRows(1 To lastRow - dataStartRow + 1)
    
    Dim r As Long
    Dim c As Long
    Dim deleteCount As Long
    Dim allEmpty As Boolean
    
    For r = dataStartRow To lastRow
        allEmpty = True
        For c = LBound(criteriaCols) To UBound(criteriaCols)
            If Len(Trim(CStr(ws.Cells(r, criteriaCols(c)).Value))) > 0 Then
                allEmpty = False
                Exit For
            End If
        Next c
        
        If allEmpty Then
            deleteCount = deleteCount + 1
            delRows(deleteCount) = r
        End If
    Next r
    
    If deleteCount = 0 Then
        MsgBox "删除当前工作表空行无需处理" & vbCrLf & vbCrLf & _
               "处理结果：" & vbCrLf & _
               "工作表：" & ws.Name & vbCrLf & _
               "检查列：" & Join(criteriaNames, "、") & vbCrLf & _
               "未找到符合条件的空行。", vbInformation, "删除当前工作表空行"
        Exit Sub
    End If
    
    Dim msg As String
    msg = "确认删除当前工作表空行？" & vbCrLf & vbCrLf & _
          "处理范围：" & vbCrLf & _
          "工作表：" & ws.Name & vbCrLf & _
          "表头行：" & headerRow & vbCrLf & _
          "数据行：" & dataStartRow & " 到 " & lastRow & vbCrLf & _
          "检查列：" & Join(criteriaNames, "、") & vbCrLf & vbCrLf & _
          "处理结果：" & vbCrLf & _
          "拟删除行数：" & deleteCount & vbCrLf & vbCrLf & _
          "注意事项：" & vbCrLf & _
          "删除规则：指定列全部为空时删除整行。" & vbCrLf & _
          "此操作不可撤销，是否继续？"
    
    Dim resp As VbMsgBoxResult
    resp = MsgBox(msg, vbYesNo + vbExclamation + vbDefaultButton2, "删除当前工作表空行")
    
    If resp <> vbYes Then
        MsgBox "删除当前工作表空行已取消" & vbCrLf & vbCrLf & _
               "处理结果：" & vbCrLf & _
               "未删除任何行。", vbInformation, "删除当前工作表空行"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim t0 As Double
    t0 = Timer
    
    On Error GoTo ErrHandler
    
    For r = deleteCount To 1 Step -1
        ws.Rows(delRows(r)).Delete Shift:=xlShiftUp
    Next r
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "删除当前工作表空行完成" & vbCrLf & vbCrLf & _
           "耗时：" & Format(Timer - t0, "0.00") & " 秒" & vbCrLf & vbCrLf & _
           "处理结果：" & vbCrLf & _
           "工作表：" & ws.Name & vbCrLf & _
           "检查列：" & Join(criteriaNames, "、") & vbCrLf & _
           "已删除行数：" & deleteCount, _
           vbInformation, "删除当前工作表空行"
    Exit Sub
    
ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
    MsgBox "删除当前工作表空行失败" & vbCrLf & vbCrLf & _
           "错误信息：" & Err.Description, vbCritical, "删除当前工作表空行"
End Sub

Private Function ParseColumnCriteria(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal criteriaInput As String, _
                                     ByRef criteriaCols() As Long, ByRef criteriaNames() As String) As Boolean
    
    Dim tokens() As String
    tokens = Split(Replace(criteriaInput, "，", ","), ",")
    
    Dim maxCol As Long
    maxCol = ws.Columns.Count
    
    Dim tempCols() As Long
    Dim tempNames() As String
    ReDim tempCols(1 To UBound(tokens) - LBound(tokens) + 1)
    ReDim tempNames(1 To UBound(tokens) - LBound(tokens) + 1)
    
    Dim count As Long
    Dim i As Long
    Dim token As String
    Dim colNum As Long
    
    For i = LBound(tokens) To UBound(tokens)
        token = Trim(tokens(i))
        If Len(token) > 0 Then
            colNum = ResolveColumnNumber(ws, headerRow, token)
            If colNum < 1 Or colNum > maxCol Then
                MsgBox "删除当前工作表空行无法继续" & vbCrLf & vbCrLf & _
                       "错误信息：无法识别列条件：" & token & vbCrLf & vbCrLf & _
                       "请使用列字母、列号或表头名。", vbExclamation, "删除当前工作表空行"
                ParseColumnCriteria = False
                Exit Function
            End If
            
            If Not ColumnExists(tempCols, count, colNum) Then
                count = count + 1
                tempCols(count) = colNum
                tempNames(count) = ColumnLabel(ws, colNum) & "列"
            End If
        End If
    Next i
    
    If count = 0 Then
        MsgBox "删除当前工作表空行无法继续" & vbCrLf & vbCrLf & _
               "错误信息：至少需要指定一列。", vbExclamation, "删除当前工作表空行"
        ParseColumnCriteria = False
        Exit Function
    End If
    
    ReDim criteriaCols(1 To count)
    ReDim criteriaNames(1 To count)
    
    For i = 1 To count
        criteriaCols(i) = tempCols(i)
        criteriaNames(i) = tempNames(i)
    Next i
    
    ParseColumnCriteria = True
End Function

Private Function ResolveColumnNumber(ByVal ws As Worksheet, ByVal headerRow As Long, ByVal token As String) As Long
    
    If IsNumeric(token) Then
        ResolveColumnNumber = CLng(token)
        Exit Function
    End If
    
    If IsColumnLetters(token) Then
        On Error Resume Next
        ResolveColumnNumber = ws.Range(UCase$(token) & "1").Column
        On Error GoTo 0
        Exit Function
    End If
    
    Dim lastCol As Long
    lastCol = GetLastUsedCol(ws)
    
    Dim c As Long
    For c = 1 To lastCol
        If Trim(CStr(ws.Cells(headerRow, c).Value)) = token Then
            ResolveColumnNumber = c
            Exit Function
        End If
    Next c
    
    ResolveColumnNumber = 0
End Function

Private Function IsColumnLetters(ByVal valueText As String) As Boolean
    
    Dim i As Long
    Dim ch As String
    
    If Len(valueText) = 0 Then Exit Function
    
    For i = 1 To Len(valueText)
        ch = Mid$(valueText, i, 1)
        If ch < "A" Or ch > "Z" Then
            ch = UCase$(ch)
            If ch < "A" Or ch > "Z" Then
                IsColumnLetters = False
                Exit Function
            End If
        End If
    Next i
    
    IsColumnLetters = True
End Function

Private Function ColumnExists(ByRef cols() As Long, ByVal count As Long, ByVal colNum As Long) As Boolean
    
    Dim i As Long
    For i = 1 To count
        If cols(i) = colNum Then
            ColumnExists = True
            Exit Function
        End If
    Next i
    
    ColumnExists = False
End Function

Private Function ColumnLabel(ByVal ws As Worksheet, ByVal colNum As Long) As String
    ColumnLabel = Split(ws.Cells(1, colNum).Address(True, False), "$")(0)
End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long
    
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    
    If lastCell Is Nothing Then
        GetLastUsedRow = 0
    Else
        GetLastUsedRow = lastCell.row
    End If
End Function

Private Function GetLastUsedCol(ByVal ws As Worksheet) As Long
    
    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    
    If lastCell Is Nothing Then
        GetLastUsedCol = 0
    Else
        GetLastUsedCol = lastCell.Column
    End If
End Function
