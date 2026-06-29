Option Explicit

'==============================================================
' Module: Weekly Recommendation Dependency Update
' Entry : Weekly01_UpdateDependencies
' Output: update 产品要素.产品规模 and 主要底层资产.节点说明
'==============================================================

Private Const SHEET_RAW_NAV As String = "上层产品净值数据(181)"
Private Const SHEET_PRODUCT_ELEMENT As String = "产品要素"
Private Const SHEET_ASSET As String = "主要底层资产"
Private Const SHEET_ASSET_MAPPING As String = "底层资产对应关系"

Public Sub Weekly01_UpdateDependencies()
    Dim t0 As Double: t0 = Timer
    
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldDisplayAlerts As Boolean
    Dim oldCalculation As XlCalculation
    
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    
    Dim wbDB As Workbook: Set wbDB = ThisWorkbook
    Dim wsRaw As Worksheet: Set wsRaw = RequireSheet(wbDB, SHEET_RAW_NAV)
    Dim wsProduct As Worksheet: Set wsProduct = RequireSheet(wbDB, SHEET_PRODUCT_ELEMENT)
    Dim wsAsset As Worksheet: Set wsAsset = RequireSheet(wbDB, SHEET_ASSET)
    Dim wsMapping As Worksheet: Set wsMapping = RequireSheet(wbDB, SHEET_ASSET_MAPPING)
    
    Dim latestNavByCode As Object
    Set latestNavByCode = BuildLatestAssetNavIndex(wsRaw)
    
    Dim productUpdated As Long
    productUpdated = UpdateProductScale(wsProduct, latestNavByCode)
    
    Dim assetUpdated As Long
    assetUpdated = UpdateUnderlyingAssetDescriptions(wsAsset, wsMapping, latestNavByCode)
    
    MsgBox "推荐材料依赖数据更新完成" & vbCrLf & vbCrLf & _
           "耗时：" & Format(Timer - t0, "0.0") & " 秒" & vbCrLf & vbCrLf & _
           "处理结果：" & vbCrLf & _
           "产品规模更新行数：" & productUpdated & vbCrLf & _
           "底层资产节点说明更新行数：" & assetUpdated, vbInformation, "推荐材料依赖数据更新"

CleanExit:
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Exit Sub

EH:
    MsgBox "推荐材料依赖数据更新失败" & vbCrLf & vbCrLf & _
           "错误信息：" & Err.Description, vbExclamation, "推荐材料依赖数据更新"
    Resume CleanExit
End Sub


Private Function BuildLatestAssetNavIndex(ByVal wsRaw As Worksheet) As Object
    Dim headers As Object
    Set headers = GetHeaderMap(wsRaw)
    RequireHeaders headers, Array("日期", "资产净值", "信托计划代码")
    
    Dim dateCol As Long: dateCol = HeaderColumn(headers, "日期")
    Dim navCol As Long: navCol = HeaderColumn(headers, "资产净值")
    Dim codeCol As Long: codeCol = HeaderColumn(headers, "信托计划代码")
    
    Dim latestDateByCode As Object
    Set latestDateByCode = CreateObject("Scripting.Dictionary")
    
    Dim latestNavByCode As Object
    Set latestNavByCode = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, codeCol).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastRow
        Dim productCode As String
        productCode = Trim(CStr(wsRaw.Cells(r, codeCol).Value))
        
        If Len(productCode) > 0 And IsNumeric(wsRaw.Cells(r, navCol).Value) Then
            Dim navDateKey As Long
            navDateKey = NormalizeDateKey(wsRaw.Cells(r, dateCol).Value)
            
            If navDateKey > 0 Then
                If Not latestDateByCode.Exists(productCode) Then
                    latestDateByCode(productCode) = navDateKey
                    latestNavByCode(productCode) = CDbl(wsRaw.Cells(r, navCol).Value)
                ElseIf navDateKey > CLng(latestDateByCode(productCode)) Then
                    latestDateByCode(productCode) = navDateKey
                    latestNavByCode(productCode) = CDbl(wsRaw.Cells(r, navCol).Value)
                End If
            End If
        End If
    Next r
    
    Set BuildLatestAssetNavIndex = latestNavByCode
End Function

Private Function UpdateProductScale(ByVal wsProduct As Worksheet, ByVal latestNavByCode As Object) As Long
    Dim headers As Object
    Set headers = GetHeaderMap(wsProduct)
    RequireHeaders headers, Array("信托计划代码", "产品规模")
    
    Dim codeCol As Long: codeCol = HeaderColumn(headers, "信托计划代码")
    Dim scaleCol As Long: scaleCol = HeaderColumn(headers, "产品规模")
    
    Dim lastRow As Long
    lastRow = wsProduct.Cells(wsProduct.Rows.Count, codeCol).End(xlUp).Row
    
    Dim updatedCount As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim productCode As String
        productCode = Trim(CStr(wsProduct.Cells(r, codeCol).Value))
        
        If latestNavByCode.Exists(productCode) Then
            wsProduct.Cells(r, scaleCol).Value = FormatYiRmb(CDbl(latestNavByCode(productCode)))
            updatedCount = updatedCount + 1
        End If
    Next r
    
    UpdateProductScale = updatedCount
End Function

Private Function UpdateUnderlyingAssetDescriptions(ByVal wsAsset As Worksheet, _
                                                   ByVal wsMapping As Worksheet, _
                                                   ByVal latestNavByCode As Object) As Long
    Dim assetHeaders As Object
    Set assetHeaders = GetHeaderMap(wsAsset)
    RequireHeaders assetHeaders, Array("节点名称", "节点说明")
    
    Dim mappingHeaders As Object
    Set mappingHeaders = GetHeaderMap(wsMapping)
    RequireHeaders mappingHeaders, Array("节点名称", "信托计划代码")
    
    Dim scaleByNodeName As Object
    Set scaleByNodeName = BuildScaleByNodeName(wsMapping, mappingHeaders, latestNavByCode)
    
    Dim nodeCol As Long: nodeCol = HeaderColumn(assetHeaders, "节点名称")
    Dim descCol As Long: descCol = HeaderColumn(assetHeaders, "节点说明")
    Dim otherDescCol As Long
    If assetHeaders.Exists("其他说明") Then otherDescCol = HeaderColumn(assetHeaders, "其他说明")
    
    Dim lastRow As Long
    lastRow = wsAsset.Cells(wsAsset.Rows.Count, nodeCol).End(xlUp).Row
    
    Dim updatedCount As Long
    Dim r As Long
    For r = 2 To lastRow
        Dim nodeName As String
        nodeName = Trim(CStr(wsAsset.Cells(r, nodeCol).Value))
        
        If scaleByNodeName.Exists(nodeName) Then
            wsAsset.Cells(r, descCol).Value = BuildAssetDescription(wsAsset, r, otherDescCol, CDbl(scaleByNodeName(nodeName)))
            updatedCount = updatedCount + 1
        End If
    Next r
    
    UpdateUnderlyingAssetDescriptions = updatedCount
End Function

Private Function BuildAssetDescription(ByVal wsAsset As Worksheet, _
                                       ByVal rowNo As Long, _
                                       ByVal otherDescCol As Long, _
                                       ByVal rawAmount As Double) As String
    Dim scaleText As String
    scaleText = FormatYiRmb(rawAmount)
    
    If otherDescCol <= 0 Then
        BuildAssetDescription = scaleText
        Exit Function
    End If
    
    Dim otherText As String
    otherText = Trim(CStr(wsAsset.Cells(rowNo, otherDescCol).Value))
    
    If Len(otherText) = 0 Then
        BuildAssetDescription = scaleText
    Else
        BuildAssetDescription = otherText & scaleText
    End If
End Function

Private Function BuildScaleByNodeName(ByVal wsMapping As Worksheet, _
                                      ByVal headers As Object, _
                                      ByVal latestNavByCode As Object) As Object
    Dim nodeCol As Long: nodeCol = HeaderColumn(headers, "节点名称")
    Dim codeCol As Long: codeCol = HeaderColumn(headers, "信托计划代码")
    
    Dim scaleByNodeName As Object
    Set scaleByNodeName = CreateObject("Scripting.Dictionary")
    
    Dim lastRow As Long
    lastRow = wsMapping.Cells(wsMapping.Rows.Count, nodeCol).End(xlUp).Row
    
    Dim r As Long
    For r = 2 To lastRow
        Dim nodeName As String
        nodeName = Trim(CStr(wsMapping.Cells(r, nodeCol).Value))
        
        Dim productCode As String
        productCode = Trim(CStr(wsMapping.Cells(r, codeCol).Value))
        
        If Len(nodeName) > 0 And Len(productCode) > 0 Then
            If latestNavByCode.Exists(productCode) Then
                If Not scaleByNodeName.Exists(nodeName) Then
                    scaleByNodeName(nodeName) = 0#
                End If
                scaleByNodeName(nodeName) = CDbl(scaleByNodeName(nodeName)) + CDbl(latestNavByCode(productCode))
            End If
        End If
    Next r
    
    Set BuildScaleByNodeName = scaleByNodeName
End Function

Private Function FormatYiRmb(ByVal rawAmount As Double) As String
    FormatYiRmb = Format(rawAmount / 100000000#, "0.00") & "亿元人民币"
End Function

Private Function NormalizeDateKey(ByVal rawValue As Variant) As Long
    On Error GoTo BadDate
    
    If IsDate(rawValue) Then
        NormalizeDateKey = CLng(Format(CDate(rawValue), "yyyymmdd"))
    ElseIf IsNumeric(rawValue) Then
        NormalizeDateKey = CLng(rawValue)
    ElseIf Len(Trim(CStr(rawValue))) = 8 And IsNumeric(Trim(CStr(rawValue))) Then
        NormalizeDateKey = CLng(Trim(CStr(rawValue)))
    Else
        NormalizeDateKey = 0
    End If
    Exit Function

BadDate:
    NormalizeDateKey = 0
End Function

Private Function GetHeaderMap(ByVal ws As Worksheet) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim lastCol As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Dim c As Long
    For c = 1 To lastCol
        Dim key As String
        key = Trim(CStr(ws.Cells(1, c).Value))
        If Len(key) > 0 Then dict(key) = c
    Next c
    
    Set GetHeaderMap = dict
End Function

Private Sub RequireHeaders(ByVal headers As Object, ByVal names As Variant)
    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Not headers.Exists(CStr(names(i))) Then
            Err.Raise vbObjectError + 301, , "缺少必需字段: " & CStr(names(i))
        End If
    Next i
End Sub

Private Function HeaderColumn(ByVal headers As Object, ByVal headerName As String) As Long
    If Not headers.Exists(headerName) Then
        Err.Raise vbObjectError + 302, , "缺少字段: " & headerName
    End If
    HeaderColumn = CLng(headers(headerName))
End Function

Private Function RequireSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set RequireSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    
    If RequireSheet Is Nothing Then
        Err.Raise vbObjectError + 401, , "未找到 sheet: " & sheetName
    End If
End Function
