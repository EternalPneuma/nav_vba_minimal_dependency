Option Explicit

'==============================================================
' 模块: 核对净值数据
' 功能: 按“日期 + 信托计划代码”核对两个数据表的净值。
'       上层产品净值数据(181): B=日期, D=净值, I=信托计划代码
'       绘图净值数据:           A=日期, D=净值, B=信托计划代码
'       仅将净值绝对差 >= 0.0001 的记录输出到“差异详情”。
'==============================================================

Private Const SHEET_SOURCE As String = "上层产品净值数据(181)"
Private Const SHEET_TARGET As String = "绘图净值数据"
Private Const TOLERANCE As Double = 0.0001

Public Sub Tool04_CheckNavData()

    Dim oldCalculation As XlCalculation
    oldCalculation = Application.Calculation

    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim wsSource As Worksheet, wsTarget As Worksheet
    Set wsSource = GetWorksheet(ThisWorkbook, SHEET_SOURCE)
    Set wsTarget = GetWorksheet(ThisWorkbook, SHEET_TARGET)

    If wsSource Is Nothing Or wsTarget Is Nothing Then
        MsgBox "核对净值数据无法继续" & vbCrLf & vbCrLf & _
               "错误信息：未找到核对所需工作表。" & vbCrLf & _
               "请确认存在：" & SHEET_SOURCE & "、" & SHEET_TARGET, _
               vbExclamation, "核对净值数据"
        GoTo CleanUp
    End If

    Dim sourceLastRow As Long, targetLastRow As Long
    sourceLastRow = LastUsedRow(wsSource, "I")
    targetLastRow = LastUsedRow(wsTarget, "B")

    If sourceLastRow < 2 Or targetLastRow < 2 Then
        MsgBox "核对净值数据无法继续" & vbCrLf & vbCrLf & _
               "错误信息：至少有一个工作表没有可核对的数据。", vbExclamation, "核对净值数据"
        GoTo CleanUp
    End If

    Dim sourceData As Variant, targetData As Variant
    sourceData = wsSource.Range("B2:I" & sourceLastRow).Value2
    targetData = wsTarget.Range("A2:D" & targetLastRow).Value2

    ' 以绘图净值数据建立索引。重复主键不能可靠一一匹配，保留首条并在简报中提示。
    Dim targetIndex As Object
    Set targetIndex = CreateObject("Scripting.Dictionary")

    Dim targetDuplicateCount As Long, targetInvalidKeyCount As Long
    Dim i As Long, key As String, dateKey As String, code As String
    For i = 1 To UBound(targetData, 1)
        code = NormalizeCode(targetData(i, 2))
        dateKey = NormalizeDateKey(targetData(i, 1))

        If Len(code) = 0 Or Len(dateKey) = 0 Then
            targetInvalidKeyCount = targetInvalidKeyCount + 1
        Else
            key = dateKey & "|" & code
            If targetIndex.Exists(key) Then
                targetDuplicateCount = targetDuplicateCount + 1
            Else
                targetIndex.Add key, i
            End If
        End If
    Next i

    Dim detailRows() As Variant, detailCount As Long
    ReDim detailRows(1 To UBound(sourceData, 1), 1 To 9)

    Dim sourceInvalidKeyCount As Long, sourceUnmatchedCount As Long
    Dim sourceNonNumericNavCount As Long, targetNonNumericNavCount As Long
    Dim matchedCount As Long, withinToleranceCount As Long, differenceCount As Long
    Dim targetRowIndex As Long, sourceNav As Double, targetNav As Double, navDifference As Double

    For i = 1 To UBound(sourceData, 1)
        code = NormalizeCode(sourceData(i, 8)) ' I列
        dateKey = NormalizeDateKey(sourceData(i, 1)) ' B列

        If Len(code) = 0 Or Len(dateKey) = 0 Then
            sourceInvalidKeyCount = sourceInvalidKeyCount + 1
        Else
            key = dateKey & "|" & code
            If Not targetIndex.Exists(key) Then
                sourceUnmatchedCount = sourceUnmatchedCount + 1
            Else
                targetRowIndex = CLng(targetIndex(key))
                matchedCount = matchedCount + 1

                If IsError(sourceData(i, 3)) Then ' D列
                    sourceNonNumericNavCount = sourceNonNumericNavCount + 1
                ElseIf Not IsNumeric(sourceData(i, 3)) Then
                    sourceNonNumericNavCount = sourceNonNumericNavCount + 1
                ElseIf IsError(targetData(targetRowIndex, 4)) Then ' D列
                    targetNonNumericNavCount = targetNonNumericNavCount + 1
                ElseIf Not IsNumeric(targetData(targetRowIndex, 4)) Then
                    targetNonNumericNavCount = targetNonNumericNavCount + 1
                Else
                    sourceNav = CDbl(sourceData(i, 3))
                    targetNav = CDbl(targetData(targetRowIndex, 4))
                    navDifference = Abs(sourceNav - targetNav)

                    If navDifference < TOLERANCE Then
                        withinToleranceCount = withinToleranceCount + 1
                    Else
                        differenceCount = differenceCount + 1
                        detailCount = detailCount + 1
                        detailRows(detailCount, 1) = i + 1
                        detailRows(detailCount, 2) = targetRowIndex + 1
                        detailRows(detailCount, 3) = DateFromKey(dateKey)
                        detailRows(detailCount, 4) = code
                        detailRows(detailCount, 5) = sourceNav
                        detailRows(detailCount, 6) = targetNav
                        detailRows(detailCount, 7) = sourceNav - targetNav
                        detailRows(detailCount, 8) = navDifference
                        detailRows(detailCount, 9) = "绝对差不小于 " & Format(TOLERANCE, "0.0000")
                    End If
                End If
            End If
        End If
    Next i

    Dim wbResult As Workbook, wsSummary As Worksheet, wsDetail As Worksheet
    Set wbResult = Workbooks.Add(xlWBATWorksheet)
    Set wsSummary = wbResult.Worksheets(1)
    wsSummary.Name = "执行简报"
    Set wsDetail = wbResult.Worksheets.Add(After:=wsSummary)
    wsDetail.Name = "差异详情"

    WriteSummary wsSummary, sourceLastRow - 1, targetLastRow - 1, matchedCount, _
                 withinToleranceCount, differenceCount, sourceUnmatchedCount, _
                 sourceInvalidKeyCount, targetInvalidKeyCount, targetDuplicateCount, _
                 sourceNonNumericNavCount, targetNonNumericNavCount
    WriteDetails wsDetail, detailRows, detailCount

    Dim outputPath As String
    outputPath = BuildOutputPath(ThisWorkbook)
    wbResult.SaveAs Filename:=outputPath, FileFormat:=xlOpenXMLWorkbook
    wbResult.Close SaveChanges:=True

    MsgBox "核对净值数据完成" & vbCrLf & vbCrLf & _
           "处理结果：" & vbCrLf & _
           "匹配记录数：" & matchedCount & vbCrLf & _
           "超过限度记录数：" & differenceCount & vbCrLf & vbCrLf & _
           "输出文件：" & vbCrLf & outputPath, vbInformation, "核对净值数据"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = oldCalculation
    Application.EnableEvents = True
    Exit Sub

ErrHandler:
    MsgBox "核对净值数据失败" & vbCrLf & vbCrLf & _
           "错误信息：" & Err.Description, vbCritical, "核对净值数据"
    Resume CleanUp
End Sub


Private Sub WriteSummary(ByVal ws As Worksheet, ByVal sourceRows As Long, ByVal targetRows As Long, _
                         ByVal matchedCount As Long, ByVal withinToleranceCount As Long, _
                         ByVal differenceCount As Long, ByVal unmatchedCount As Long, _
                         ByVal sourceInvalidKeyCount As Long, ByVal targetInvalidKeyCount As Long, _
                         ByVal targetDuplicateCount As Long, ByVal sourceNonNumericNavCount As Long, _
                         ByVal targetNonNumericNavCount As Long)
    Dim summaryData(1 To 14, 1 To 2) As Variant
    summaryData(1, 1) = "执行时间": summaryData(1, 2) = Now
    summaryData(2, 1) = "核对阈值（绝对差）": summaryData(2, 2) = TOLERANCE
    summaryData(3, 1) = SHEET_SOURCE & " 数据行数": summaryData(3, 2) = sourceRows
    summaryData(4, 1) = SHEET_TARGET & " 数据行数": summaryData(4, 2) = targetRows
    summaryData(5, 1) = "成功匹配主键记录数": summaryData(5, 2) = matchedCount
    summaryData(6, 1) = "差值小于阈值记录数": summaryData(6, 2) = withinToleranceCount
    summaryData(7, 1) = "差值不小于阈值记录数": summaryData(7, 2) = differenceCount
    summaryData(8, 1) = "上层表未匹配记录数": summaryData(8, 2) = unmatchedCount
    summaryData(9, 1) = "上层表无效主键记录数": summaryData(9, 2) = sourceInvalidKeyCount
    summaryData(10, 1) = "绘图表无效主键记录数": summaryData(10, 2) = targetInvalidKeyCount
    summaryData(11, 1) = "绘图表重复主键记录数": summaryData(11, 2) = targetDuplicateCount
    summaryData(12, 1) = "上层表非数值净值记录数": summaryData(12, 2) = sourceNonNumericNavCount
    summaryData(13, 1) = "绘图表非数值净值记录数": summaryData(13, 2) = targetNonNumericNavCount
    summaryData(14, 1) = "说明": summaryData(14, 2) = "日期按实际日期值匹配；差异详情仅输出绝对差不小于阈值的记录。"

    ws.Range("A1:B14").Value = summaryData
    ws.Range("A1:A14").Font.Bold = True
    ws.Range("B1").NumberFormat = "yyyy-mm-dd hh:mm:ss"
    ws.Range("B2").NumberFormat = "0.0000"
    ws.Columns("A:B").AutoFit
End Sub

Private Sub WriteDetails(ByVal ws As Worksheet, ByRef detailRows() As Variant, ByVal detailCount As Long)
    Dim headers As Variant
    headers = Array("上层表行号", "绘图表行号", "日期", "信托计划代码", "上层表D列净值", "绘图表D列净值", "差值（上层-绘图）", "绝对差", "差异说明")
    Dim i As Long
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).Value = headers(i)
    Next i
    ws.Range("A1:I1").Font.Bold = True
    ws.Range("A1:I1").Interior.Color = RGB(217, 225, 242)

    If detailCount > 0 Then
        ws.Range("A2").Resize(detailCount, 9).Value = detailRows
        ws.Range("C2:C" & detailCount + 1).NumberFormat = "yyyy-mm-dd"
        ws.Range("E2:H" & detailCount + 1).NumberFormat = "0.000000"
        ws.Range("A1:I" & detailCount + 1).AutoFilter
    Else
        ws.Range("A2").Value = "无绝对差不小于 " & Format(TOLERANCE, "0.0000") & " 的记录。"
    End If
    ws.Columns("A:I").AutoFit
End Sub

Private Function NormalizeDateKey(ByVal value As Variant) As String
    On Error GoTo InvalidDate
    If IsError(value) Then Exit Function
    If IsEmpty(value) Or Len(Trim(CStr(value))) = 0 Then Exit Function
    
    ' Range.Value2 返回 Excel 日期序列值（Double），IsDate 不会可靠识别该值，必须先显式转换。
    Dim dateText As String
    dateText = Trim(CStr(value))
    If Len(dateText) = 8 And Not (dateText Like "*[!0-9]*") Then
        NormalizeDateKey = Format(DateSerial(CInt(Left(dateText, 4)), _
                                              CInt(Mid(dateText, 5, 2)), _
                                              CInt(Right(dateText, 2))), "yyyy-mm-dd")
    ElseIf IsNumeric(value) Then
        If CDbl(value) >= 1 And CDbl(value) <= 2958465 Then
            NormalizeDateKey = Format(DateValue(CDate(CDbl(value))), "yyyy-mm-dd")
        End If
    ElseIf IsDate(value) Then
        NormalizeDateKey = Format(DateValue(CDate(value)), "yyyy-mm-dd")
    End If
    Exit Function
InvalidDate:
    NormalizeDateKey = vbNullString
End Function

Private Function DateFromKey(ByVal dateKey As String) As Date
    DateFromKey = DateSerial(CInt(Left(dateKey, 4)), CInt(Mid(dateKey, 6, 2)), CInt(Right(dateKey, 2)))
End Function

Private Function NormalizeCode(ByVal value As Variant) As String
    If IsError(value) Then Exit Function
    NormalizeCode = Trim(CStr(value))
End Function

Private Function GetWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWorksheet = wb.Worksheets(sheetName)
    On Error GoTo 0
End Function

Private Function LastUsedRow(ByVal ws As Worksheet, ByVal columnLetter As String) As Long
    LastUsedRow = ws.Cells(ws.Rows.Count, columnLetter).End(xlUp).Row
End Function

Private Function BuildOutputPath(ByVal wb As Workbook) As String
    Dim basePath As String
    basePath = wb.Path
    If Len(basePath) = 0 Then basePath = Application.DefaultFilePath
    BuildOutputPath = basePath & Application.PathSeparator & _
                      "净值数据核对_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx"
End Function
