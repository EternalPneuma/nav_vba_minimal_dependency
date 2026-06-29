Option Explicit

'==============================================================
' 模块: 清洗重复数据
' 功能: 对绘图净值数据按"净值日期+产品编号"主键去重,
'       重复时保留物理位置靠后的记录(即最后导入的)
'==============================================================

Public Sub Tool01_CleanDuplicateData()
    
    Dim t0 As Double: t0 = Timer
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim wsData As Worksheet: Set wsData = wbDB.Sheets("绘图净值数据")
    
    '--- 1. 读取数据 ---
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row
    
    If lastRow < 3 Then
        MsgBox "绘图净值数据去重无需处理" & vbCrLf & vbCrLf & _
               "处理结果：" & vbCrLf & _
               "绘图净值数据中数据少于2行，无需去重。", vbInformation, "绘图净值数据去重"
        GoTo CleanUp
    End If
    
    ' A-J列共10列
    Dim dataArr As Variant
    dataArr = wsData.Range("A2:J" & lastRow).Value
    
    Dim totalRows As Long
    totalRows = UBound(dataArr, 1)
    
    '--- 2. 倒序遍历,用Dictionary记录每个主键首次出现的位置 ---
    ' 倒序遍历时,先遇到的就是"物理靠后"的,因此自动保留最后一条
    Dim pkDict As Object
    Set pkDict = CreateObject("Scripting.Dictionary")
    
    Dim keepFlags() As Boolean
    ReDim keepFlags(1 To totalRows)
    
    Dim i As Long, pk As String
    Dim curDate As Variant, curCode As String
    Dim duplicateCount As Long: duplicateCount = 0
    Dim invalidCount As Long: invalidCount = 0
    
    For i = totalRows To 1 Step -1
        curDate = dataArr(i, 1)
        curCode = Trim(CStr(dataArr(i, 2)))
        
        ' 主键校验
        If Len(curCode) = 0 Or IsEmpty(curDate) Then
            ' 无效行(无主键),不保留
            keepFlags(i) = False
            invalidCount = invalidCount + 1
            GoTo NextRow
        End If
        
        ' 构造主键
        If IsDate(curDate) Then
            pk = Format(CDate(curDate), "yyyy-mm-dd") & "|" & curCode
        Else
            pk = CStr(curDate) & "|" & curCode
        End If
        
        ' 倒序遍历下,Dictionary中已存在 → 当前行是更早的重复 → 不保留
        If pkDict.Exists(pk) Then
            keepFlags(i) = False
            duplicateCount = duplicateCount + 1
        Else
            pkDict(pk) = i
            keepFlags(i) = True
        End If
        
NextRow:
    Next i
    
    '--- 3. 如果没有重复也没有无效行,提前退出 ---
    If duplicateCount = 0 And invalidCount = 0 Then
        MsgBox "绘图净值数据去重无需处理" & vbCrLf & vbCrLf & _
               "处理结果：" & vbCrLf & _
               "总记录数：" & totalRows & vbCrLf & _
               "重复数据：0 条" & vbCrLf & _
               "无效数据：0 条", vbInformation, "绘图净值数据去重"
        GoTo CleanUp
    End If
    
    '--- 4. 构造保留后的新数组 ---
    Dim keepCount As Long: keepCount = 0
    For i = 1 To totalRows
        If keepFlags(i) Then keepCount = keepCount + 1
    Next i
    
    If keepCount = 0 Then
        MsgBox "绘图净值数据去重无法继续" & vbCrLf & vbCrLf & _
               "错误信息：清洗后无任何有效数据，已中止操作。", vbCritical, "绘图净值数据去重"
        GoTo CleanUp
    End If
    
    Dim newArr() As Variant
    ReDim newArr(1 To keepCount, 1 To 10)
    
    Dim idx As Long: idx = 0
    Dim c As Long
    For i = 1 To totalRows
        If keepFlags(i) Then
            idx = idx + 1
            For c = 1 To 10
                newArr(idx, c) = dataArr(i, c)
            Next c
        End If
    Next i
    
    '--- 5. 清空原数据区,写回新数据 ---
    wsData.Range("A2:J" & lastRow).ClearContents
    wsData.Range("A2:J" & (1 + keepCount)).Value = newArr
    
    '--- 6. 按日期+编号升序排序(可选,保持与导入模块一致) ---
    Dim finalLastRow As Long
    finalLastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).row
    If finalLastRow >= 3 Then
        wsData.Range("A2:J" & finalLastRow).Sort _
            Key1:=wsData.Range("A2"), Order1:=xlAscending, _
            Key2:=wsData.Range("B2"), Order2:=xlAscending, _
            header:=xlNo
    End If
    
    '--- 7. 提示 ---
    MsgBox "绘图净值数据去重完成" & vbCrLf & vbCrLf & _
           "耗时：" & Format(Timer - t0, "0.00") & " 秒" & vbCrLf & vbCrLf & _
           "处理结果：" & vbCrLf & _
           "原记录数：" & totalRows & vbCrLf & _
           "重复删除行数：" & duplicateCount & vbCrLf & _
           "无效删除行数：" & invalidCount & vbCrLf & _
           "保留记录数：" & keepCount, _
           vbInformation, "绘图净值数据去重"

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub


