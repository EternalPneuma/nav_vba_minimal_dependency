Option Explicit

' 模块用途：根据"开放日"工作表数据，在"产品分类"工作表中计算每个产品的下一开放日、上一开放日、本次开放日间隔、本期已持续天数。
' 使用方式：将本文件内容复制到 上层产品净值数据库.xlsm 的标准模块中，运行 Data02_CalculateOpenDate。
'
' 计算逻辑：
'   基准日期 = sheet"上层产品净值数据(181)"B列日期（yyyyMMdd格式）的最大值
'   下一开放日 = 开放日列表中 >= 基准日期的最小日期
'   上一开放日 = 开放日列表中 <  基准日期的最大日期
'   上上一开放日 = 开放日列表中 < 上一开放日的最大日期（无则留空）
'   实际间隔 = 下一开放日 - 上一开放日（天数）
'   上次开放实际间隔 = 上一开放日 - 上上一开放日（天数，无则留空）
'   运作时间 = 基准日期 - 上一开放日（天数）
'   日开/周开产品（理论间隔=1/7）直接跳过，不填写开放日测算列。
'   非日开/周开但开放日工作表中无数据的产品，在完成提示中单独列出。
'   基准日期净值 = sheet"上层产品净值数据(181)"中匹配信托计划代码且日期=基准日期的单位净值。

Private Const TARGET_SHEET_NAME As String = "产品分类"
Private Const SOURCE_SHEET_NAME As String = "开放日"
Private Const COL_SEQ As String = "序号"
Private Const COL_OPEN_DATE As String = "开放日"
Private Const COL_THEORETICAL_INTERVAL As String = "理论间隔"

' 待写入的列标题
Private Const COL_NEXT_OPEN As String = "下一开放日"
Private Const COL_PREV_OPEN As String = "上一开放日"
Private Const COL_INTERVAL As String = "实际间隔"
Private Const COL_ELAPSED As String = "运作时间"
Private Const COL_PREV_PREV_OPEN As String = "上上一开放日"
Private Const COL_PREV_INTERVAL As String = "上次开放实际间隔"

' 净值数据工作表相关常量
Private Const NAV_SHEET_NAME As String = "上层产品净值数据(181)"
Private Const COL_TRUST_CODE As String = "信托计划代码"
Private Const COL_NAV_DATE As String = "日期"
Private Const COL_UNIT_NAV As String = "单位净值"

' 基准日期与净值输出列
Private Const COL_BASELINE_DATE As String = "基准日期"
Private Const COL_BASELINE_NAV As String = "基准日期净值"

' 上一/上上一开放日净值列
Private Const COL_PREV_NAV As String = "上一开放日净值"
Private Const COL_PREV_PREV_NAV As String = "上上一开放日净值"

Public Sub Data02_CalculateOpenDate()
    CalculateOpenDaysCore
End Sub

' --- Backward-compatible aliases ---
Private Sub CalculateOpenDaysCore()
    Dim appCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    appCalc = Application.Calculation

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    ' 1. 获取基准日期：sheet"上层产品净值数据(181)"B列日期（yyyyMMdd格式）的最大值
    Dim baselineDate As Date
    baselineDate = GetBaselineDateFromNAV()

    ' 2. 构建净值数据查找表：信托计划代码 → 基准日期的单位净值
    Dim navLookup As Object
    Set navLookup = BuildNAVDateLookup()

    ' 3. 构建开放日索引：序号 → 有序日期集合
    Dim openDayIndex As Object
    Set openDayIndex = BuildOpenDayIndex()

    ' 4. 定位"产品分类"工作表的表头
    Dim wsTarget As Worksheet
    Set wsTarget = ThisWorkbook.Worksheets(TARGET_SHEET_NAME)
    Dim targetHeader As Object
    Set targetHeader = BuildHeaderMap(wsTarget, 1)

    If Not targetHeader.Exists(COL_SEQ) Then
        Err.Raise vbObjectError + 2001, , "目标工作表缺少字段：" & COL_SEQ
    End If

    Dim seqCol As Long
    seqCol = CLng(targetHeader(COL_SEQ))

    If Not targetHeader.Exists(COL_THEORETICAL_INTERVAL) Then
        Err.Raise vbObjectError + 2003, , "目标工作表缺少字段：" & COL_THEORETICAL_INTERVAL
    End If
    Dim theoreticalIntervalCol As Long
    theoreticalIntervalCol = CLng(targetHeader(COL_THEORETICAL_INTERVAL))

    ' 定位信托计划代码列（产品分类 B列，用于 NAV join）
    If Not targetHeader.Exists(COL_TRUST_CODE) Then
        Err.Raise vbObjectError + 2002, , "目标工作表缺少字段：" & COL_TRUST_CODE
    End If
    Dim trustCodeCol As Long
    trustCodeCol = CLng(targetHeader(COL_TRUST_CODE))

    ' 4.5 清理已有输出列，确保每次运行从干净状态开始
    ClearExistingOutputColumns wsTarget, targetHeader

    ' 5. 确保输出列存在，返回其列号
    Dim outputCols As Object
    Set outputCols = EnsureOutputColumns(wsTarget, targetHeader)

    Dim nextCol As Long:         nextCol = CLng(outputCols(COL_NEXT_OPEN))
    Dim prevCol As Long:         prevCol = CLng(outputCols(COL_PREV_OPEN))
    Dim prevPrevCol As Long:     prevPrevCol = CLng(outputCols(COL_PREV_PREV_OPEN))
    Dim intervalCol As Long:     intervalCol = CLng(outputCols(COL_INTERVAL))
    Dim prevIntervalCol As Long: prevIntervalCol = CLng(outputCols(COL_PREV_INTERVAL))
    Dim elapsedCol As Long:      elapsedCol = CLng(outputCols(COL_ELAPSED))
    Dim baselineDateCol As Long: baselineDateCol = CLng(outputCols(COL_BASELINE_DATE))
    Dim baselineNAVCol As Long:  baselineNAVCol = CLng(outputCols(COL_BASELINE_NAV))
    Dim prevNAVCol As Long:      prevNAVCol = CLng(outputCols(COL_PREV_NAV))
    Dim prevPrevNAVCol As Long:  prevPrevNAVCol = CLng(outputCols(COL_PREV_PREV_NAV))

    ' 6. 逐行计算
    Dim lastRow As Long
    lastRow = LastUsedRow(wsTarget)

    Dim processedCount As Long
    Dim skippedCount As Long
    Dim missingOpenDayCount As Long
    Dim missingOpenDayDetails As String
    Dim navMatchedCount As Long
    Dim r As Long

    For r = 2 To lastRow
        Dim seqValue As Variant
        seqValue = wsTarget.Cells(r, seqCol).Value
        If IsEmpty(seqValue) Then GoTo ContinueRow

        Dim seqKey As String
        seqKey = CStr(seqValue)

        Dim trustCode As Variant
        trustCode = wsTarget.Cells(r, trustCodeCol).Value

        ' ============================================================
        ' A. 基准日期 + 基准日期净值：对所有产品写入（包括日开）
        ' ============================================================
        wsTarget.Cells(r, baselineDateCol).Value = baselineDate
        wsTarget.Cells(r, baselineDateCol).NumberFormat = "yyyy-mm-dd"

        wsTarget.Cells(r, baselineNAVCol).Value = LookupNAV(navLookup, trustCode, baselineDate)
        wsTarget.Cells(r, prevNAVCol).Value = Empty
        wsTarget.Cells(r, prevPrevNAVCol).Value = Empty

        If Not IsEmpty(wsTarget.Cells(r, baselineNAVCol).Value) Then
            navMatchedCount = navMatchedCount + 1
        End If

        ' 先按理论间隔识别日开/周开产品，跳过开放日测算。
        If IsDailyOrWeeklyInterval(wsTarget.Cells(r, theoreticalIntervalCol).Value) Then
            skippedCount = skippedCount + 1
            processedCount = processedCount + 1
            GoTo ContinueRow
        End If

        ' ============================================================
        ' B. 开放日相关：仅对有开放日数据的产品计算
        ' ============================================================
        If Not openDayIndex.Exists(seqKey) Then
            missingOpenDayCount = missingOpenDayCount + 1
            missingOpenDayDetails = missingOpenDayDetails & seqKey & "(" & NormalizeText(trustCode) & "), "
            processedCount = processedCount + 1
            GoTo ContinueRow
        End If

        Dim openDates As Collection
        Set openDates = openDayIndex(seqKey)

        Dim prevDate As Variant
        Dim prevPrevDate As Variant
        Dim nextDate As Variant
        prevDate = Empty
        prevPrevDate = Empty
        nextDate = Empty

        FindOpenDates openDates, baselineDate, prevDate, prevPrevDate, nextDate

        ' 写入开放日列
        If Not IsEmpty(nextDate) Then
            wsTarget.Cells(r, nextCol).Value = nextDate
            wsTarget.Cells(r, nextCol).NumberFormat = "yyyy-mm-dd"
        Else
            wsTarget.Cells(r, nextCol).Value = Empty
        End If

        If Not IsEmpty(prevDate) Then
            wsTarget.Cells(r, prevCol).Value = prevDate
            wsTarget.Cells(r, prevCol).NumberFormat = "yyyy-mm-dd"

            wsTarget.Cells(r, elapsedCol).Value = CLng(baselineDate - prevDate)

            If Not IsEmpty(nextDate) Then
                wsTarget.Cells(r, intervalCol).Value = CLng(CDate(nextDate) - CDate(prevDate))
            Else
                wsTarget.Cells(r, intervalCol).Value = Empty
            End If

            ' 上一开放日净值
            wsTarget.Cells(r, prevNAVCol).Value = LookupNAV(navLookup, trustCode, prevDate)
        Else
            wsTarget.Cells(r, prevCol).Value = Empty
            wsTarget.Cells(r, intervalCol).Value = Empty
            wsTarget.Cells(r, elapsedCol).Value = Empty
        End If

        If Not IsEmpty(prevPrevDate) Then
            wsTarget.Cells(r, prevPrevCol).Value = prevPrevDate
            wsTarget.Cells(r, prevPrevCol).NumberFormat = "yyyy-mm-dd"

            If Not IsEmpty(prevDate) Then
                wsTarget.Cells(r, prevIntervalCol).Value = CLng(CDate(prevDate) - CDate(prevPrevDate))
            Else
                wsTarget.Cells(r, prevIntervalCol).Value = Empty
            End If

            ' 上上一开放日净值
            wsTarget.Cells(r, prevPrevNAVCol).Value = LookupNAV(navLookup, trustCode, prevPrevDate)
        Else
            wsTarget.Cells(r, prevPrevCol).Value = Empty
            wsTarget.Cells(r, prevIntervalCol).Value = Empty
        End If

        processedCount = processedCount + 1
ContinueRow:
    Next r

    ' 7. 恢复设置并报告
    Application.Calculation = appCalc
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating

    Dim msg As String
    msg = "开放日测算完成" & vbCrLf & vbCrLf & _
          "基准日期：" & Format$(baselineDate, "yyyy-mm-dd") & vbCrLf & vbCrLf & _
          "处理结果：" & vbCrLf & _
          "计算产品数：" & processedCount & vbCrLf & _
          "跳过（日开/周开）：" & skippedCount & vbCrLf & _
          "未被跳过但开放日未检测到：" & missingOpenDayCount & vbCrLf & _
          "净值匹配数：" & navMatchedCount

    If missingOpenDayCount > 0 Then
        If Len(missingOpenDayDetails) > 2 Then missingOpenDayDetails = Left$(missingOpenDayDetails, Len(missingOpenDayDetails) - 2)
        msg = msg & vbCrLf & vbCrLf & "注意事项：" & vbCrLf & _
              "未检测到开放日的产品：" & vbCrLf & TruncateMessageText(missingOpenDayDetails, 900)
    End If

    MsgBox msg, vbInformation, "开放日测算"
    Exit Sub

CleanFail:
    Application.Calculation = appCalc
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    MsgBox "开放日测算失败" & vbCrLf & vbCrLf & _
           "错误信息：" & err.Description, vbCritical, "开放日测算"
End Sub

Private Function IsDailyOrWeeklyInterval(ByVal intervalValue As Variant) As Boolean
    Dim intervalText As String
    intervalText = NormalizeText(intervalValue)
    If Len(intervalText) = 0 Then Exit Function
    If Not IsNumeric(intervalText) Then Exit Function

    Dim intervalDays As Long
    intervalDays = CLng(intervalText)
    IsDailyOrWeeklyInterval = (intervalDays = 1 Or intervalDays = 7)
End Function

Private Function TruncateMessageText(ByVal textValue As String, ByVal maxLength As Long) As String
    If Len(textValue) <= maxLength Then
        TruncateMessageText = textValue
    Else
        TruncateMessageText = Left$(textValue, maxLength) & "..."
    End If
End Function

' ---------------------------------------------------------------------------
' 构建开放日索引：Dictionary(序号 → Collection of Date)
' ---------------------------------------------------------------------------

Private Function BuildOpenDayIndex() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SOURCE_SHEET_NAME)

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(ws, 1)

    Dim seqCol As Long:   seqCol = CLng(headerMap(COL_SEQ))
    Dim dateCol As Long:  dateCol = CLng(headerMap(COL_OPEN_DATE))

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        Dim seqVal As Variant
        seqVal = ws.Cells(r, seqCol).value
        If IsEmpty(seqVal) Then GoTo ContinueIdx

        Dim dateVal As Variant
        dateVal = ws.Cells(r, dateCol).value
        If Not IsDate(dateVal) Then GoTo ContinueIdx

        Dim seqKey As String
        seqKey = CStr(seqVal)

        If Not dict.Exists(seqKey) Then
            dict.Add seqKey, New Collection
        End If
        dict(seqKey).Add CDate(dateVal)
ContinueIdx:
    Next r

    Set BuildOpenDayIndex = dict
End Function

' ---------------------------------------------------------------------------
' 在日期集合中找到上一/下一/上上一开放日（单次遍历，无需排序）
' ---------------------------------------------------------------------------

Private Sub FindOpenDates(ByVal openDates As Collection, ByVal baselineDate As Date, _
                          ByRef prevDate As Variant, ByRef prevPrevDate As Variant, ByRef nextDate As Variant)
    Dim d As Variant
    For Each d In openDates
        Dim dt As Date
        dt = CDate(d)
        If dt < baselineDate Then
            If IsEmpty(prevDate) Then
                prevDate = dt
            ElseIf dt > CDate(prevDate) Then
                ' 当前prev退为prevPrev
                prevPrevDate = prevDate
                prevDate = dt
            ElseIf IsEmpty(prevPrevDate) Or dt > CDate(prevPrevDate) Then
                prevPrevDate = dt
            End If
        Else
            If IsEmpty(nextDate) Then
                nextDate = dt
            ElseIf dt < CDate(nextDate) Then
                nextDate = dt
            End If
        End If
    Next d
End Sub

' ---------------------------------------------------------------------------
' 清理"产品分类"中已有的输出列（从右到左删除，避免列索引偏移）
' ---------------------------------------------------------------------------

Private Sub ClearExistingOutputColumns(ByVal ws As Worksheet, ByVal headerMap As Object)
    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim c As Long
    For c = lastCol To 1 Step -1
        Dim hdr As String
        hdr = NormalizeText(ws.Cells(1, c).Value)
        If IsOutputColumnName(hdr) Then
            ws.Columns(c).Delete
            ' 同步清理 headerMap，避免持有失效列号
            If headerMap.Exists(hdr) Then
                headerMap.Remove hdr
            End If
        End If
    Next c
End Sub

Private Function IsOutputColumnName(ByVal name As String) As Boolean
    Dim outputNames As Variant
    outputNames = Array(COL_NEXT_OPEN, COL_PREV_OPEN, COL_PREV_PREV_OPEN, _
                        COL_INTERVAL, COL_PREV_INTERVAL, COL_ELAPSED, _
                        COL_BASELINE_DATE, COL_BASELINE_NAV, _
                        COL_PREV_NAV, COL_PREV_PREV_NAV)

    Dim i As Long
    For i = LBound(outputNames) To UBound(outputNames)
        If StrComp(name, CStr(outputNames(i)), vbTextCompare) = 0 Then
            IsOutputColumnName = True
            Exit Function
        End If
    Next i

    IsOutputColumnName = False
End Function

' ---------------------------------------------------------------------------
' 确保输出列存在，返回列名→列号的映射
' ---------------------------------------------------------------------------

Private Function EnsureOutputColumns(ByVal ws As Worksheet, ByVal headerMap As Object) As Object
    Dim outputCols As Object
    Set outputCols = CreateObject("Scripting.Dictionary")

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim newCols As Variant
    ' 列顺序：前4日期 → 中3间隔 → 后4净值
    newCols = Array(COL_PREV_PREV_OPEN, COL_PREV_OPEN, COL_BASELINE_DATE, COL_NEXT_OPEN, _
                    COL_PREV_INTERVAL, COL_INTERVAL, COL_ELAPSED, _
                    COL_PREV_PREV_NAV, COL_PREV_NAV, COL_BASELINE_NAV)

    Dim i As Long
    For i = LBound(newCols) To UBound(newCols)
        Dim colName As String
        colName = CStr(newCols(i))

        If headerMap.Exists(colName) Then
            ' 列已存在，直接用
            outputCols.Add colName, CLng(headerMap(colName))
        Else
            ' 追加到右侧
            lastCol = lastCol + 1
            ws.Cells(1, lastCol).value = colName
            headerMap.Add colName, lastCol
            outputCols.Add colName, lastCol
        End If
    Next i

    Set EnsureOutputColumns = outputCols
End Function

' ---------------------------------------------------------------------------
' 从 sheet"上层产品净值数据(181)" B列获取最大日期（yyyyMMdd 格式）
' ---------------------------------------------------------------------------

Private Function GetBaselineDateFromNAV() As Date
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(NAV_SHEET_NAME)

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(ws, 1)

    If Not headerMap.Exists(COL_NAV_DATE) Then
        Err.Raise vbObjectError + 3001, , NAV_SHEET_NAME & "工作表缺少字段：" & COL_NAV_DATE
    End If

    Dim dateCol As Long
    dateCol = CLng(headerMap(COL_NAV_DATE))

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim maxDate As Date
    Dim found As Boolean
    found = False

    Dim r As Long
    For r = 2 To lastRow
        Dim cellVal As Variant
        cellVal = ws.Cells(r, dateCol).Value
        If Not IsEmpty(cellVal) And Not IsError(cellVal) Then
            Dim parsedDate As Date
            If ParseYYYYMMDD(cellVal, parsedDate) Then
                If Not found Or parsedDate > maxDate Then
                    maxDate = parsedDate
                    found = True
                End If
            End If
        End If
    Next r

    If Not found Then
        Err.Raise vbObjectError + 3002, , NAV_SHEET_NAME & "工作表中未找到有效日期"
    End If

    GetBaselineDateFromNAV = maxDate
End Function

' ---------------------------------------------------------------------------
' 构建净值数据日期查找表：信托计划代码 → (yyyymmdd → 单位净值)
' 用于按任意日期查找产品的单位净值
' ---------------------------------------------------------------------------

Private Function BuildNAVDateLookup() As Object
    Dim outerDict As Object
    Set outerDict = CreateObject("Scripting.Dictionary")
    outerDict.CompareMode = vbTextCompare

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(NAV_SHEET_NAME)

    Dim headerMap As Object
    Set headerMap = BuildHeaderMap(ws, 1)

    If Not headerMap.Exists(COL_NAV_DATE) Then Exit Function
    If Not headerMap.Exists(COL_TRUST_CODE) Then Exit Function
    If Not headerMap.Exists(COL_UNIT_NAV) Then Exit Function

    Dim dateCol As Long:      dateCol = CLng(headerMap(COL_NAV_DATE))
    Dim trustCodeCol As Long: trustCodeCol = CLng(headerMap(COL_TRUST_CODE))
    Dim navCol As Long:       navCol = CLng(headerMap(COL_UNIT_NAV))

    Dim lastRow As Long
    lastRow = LastUsedRow(ws)

    Dim r As Long
    For r = 2 To lastRow
        Dim cellDate As Variant
        cellDate = ws.Cells(r, dateCol).Value
        If IsEmpty(cellDate) Or IsError(cellDate) Then GoTo ContinueNAV

        Dim parsedDate As Date
        If Not ParseYYYYMMDD(cellDate, parsedDate) Then GoTo ContinueNAV
        Dim dateKey As String
        dateKey = Format$(parsedDate, "yyyymmdd")

        Dim trustCode As Variant
        trustCode = ws.Cells(r, trustCodeCol).Value
        If IsEmpty(trustCode) Or IsError(trustCode) Then GoTo ContinueNAV

        Dim navValue As Variant
        navValue = ws.Cells(r, navCol).Value
        If IsError(navValue) Then GoTo ContinueNAV

        Dim codeKey As String
        codeKey = NormalizeText(trustCode)

        If Not outerDict.Exists(codeKey) Then
            Dim innerDict As Object
            Set innerDict = CreateObject("Scripting.Dictionary")
            innerDict.CompareMode = vbTextCompare
            outerDict.Add codeKey, innerDict
        End If

        ' 同一 (code, date) 保留第一条
        If Not outerDict(codeKey).Exists(dateKey) Then
            outerDict(codeKey).Add dateKey, navValue
        End If
ContinueNAV:
    Next r

    Set BuildNAVDateLookup = outerDict
End Function

' ---------------------------------------------------------------------------
' 从 NAV 查找表中按信托计划代码和日期获取单位净值，无匹配返回 Empty
' ---------------------------------------------------------------------------

Private Function LookupNAV(ByVal navLookup As Object, ByVal trustCodeValue As Variant, ByVal targetDate As Variant) As Variant
    If IsEmpty(trustCodeValue) Or IsError(trustCodeValue) Then
        LookupNAV = Empty
        Exit Function
    End If
    If IsEmpty(targetDate) Then
        LookupNAV = Empty
        Exit Function
    End If

    Dim codeKey As String
    codeKey = NormalizeText(trustCodeValue)
    If Not navLookup.Exists(codeKey) Then
        LookupNAV = Empty
        Exit Function
    End If

    Dim dateKey As String
    dateKey = Format$(CDate(targetDate), "yyyymmdd")
    If Not navLookup(codeKey).Exists(dateKey) Then
        LookupNAV = Empty
        Exit Function
    End If

    LookupNAV = navLookup(codeKey)(dateKey)
End Function

' ---------------------------------------------------------------------------
' 解析 yyyyMMdd 格式的日期值（字符串如 "20260525" 或数值 20260525）
' ---------------------------------------------------------------------------

Private Function ParseYYYYMMDD(ByVal value As Variant, ByRef outDate As Date) As Boolean
    If IsDate(value) Then
        ' Excel 日期序列号
        outDate = DateValue(CDate(value))
        ParseYYYYMMDD = True
        Exit Function
    End If

    Dim textValue As String
    textValue = Trim$(CStr(value))

    ' 去除可能的 .0 后缀（Excel 将整数存储为浮点）
    If InStr(textValue, ".") > 0 Then
        textValue = Left$(textValue, InStr(textValue, ".") - 1)
    End If

    If Len(textValue) = 8 And IsNumeric(textValue) Then
        Dim y As Long: y = CLng(Left$(textValue, 4))
        Dim m As Long: m = CLng(Mid$(textValue, 5, 2))
        Dim d As Long: d = CLng(Right$(textValue, 2))
        If m >= 1 And m <= 12 And d >= 1 And d <= 31 Then
            outDate = DateSerial(y, m, d)
            ParseYYYYMMDD = True
            Exit Function
        End If
    End If

    ParseYYYYMMDD = False
End Function

Private Function BuildHeaderMap(ByVal ws As Worksheet, ByVal headerRow As Long) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim lastCol As Long
    lastCol = LastUsedColumn(ws)

    Dim c As Long
    For c = 1 To lastCol
        Dim hdr As String
        hdr = NormalizeText(ws.Cells(headerRow, c).value)
        If Len(hdr) > 0 Then
            If Not dict.Exists(hdr) Then dict.Add hdr, c
        End If
    Next c

    Set BuildHeaderMap = dict
End Function

Private Function LastUsedRow(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = foundCell.Row
    End If
End Function

Private Function LastUsedColumn(ByVal ws As Worksheet) As Long
    Dim foundCell As Range
    Set foundCell = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If foundCell Is Nothing Then
        LastUsedColumn = 1
    Else
        LastUsedColumn = foundCell.Column
    End If
End Function

Private Function NormalizeText(ByVal value As Variant) As String
    If IsError(value) Or IsEmpty(value) Then
        NormalizeText = vbNullString
        Exit Function
    End If

    Dim textValue As String
    textValue = CStr(value)
    textValue = Replace(textValue, ChrW$(12288), " ")
    textValue = Replace(textValue, vbCr, " ")
    textValue = Replace(textValue, vbLf, " ")
    NormalizeText = WorksheetFunction.Trim(textValue)
End Function
