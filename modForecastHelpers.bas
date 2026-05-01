Attribute VB_Name = "modForecastHelpers"
Option Explicit

' =========================================================
' Basic Object Helpers
' =========================================================

Public Function GetWorksheet(ByVal wsName As String) As Worksheet
    On Error GoTo ErrHandler
    Set GetWorksheet = ThisWorkbook.Worksheets(wsName)
    Exit Function

ErrHandler:
    MsgBox "Worksheet not found: " & wsName, vbExclamation
    Set GetWorksheet = Nothing
End Function

Public Function GetTable(ByVal wsName As String, ByVal tableName As String) As ListObject
    Dim ws As Worksheet

    On Error GoTo ErrHandler

    Set ws = GetWorksheet(wsName)
    If ws Is Nothing Then Exit Function

    Set GetTable = ws.ListObjects(tableName)
    Exit Function

ErrHandler:
    MsgBox "Table not found: " & tableName & " on sheet " & wsName, vbExclamation
    Set GetTable = Nothing
End Function

Public Function GetHeaderColumn(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long

    GetHeaderColumn = 0
    If lo Is Nothing Then Exit Function

    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(i).name), Trim$(headerName), vbTextCompare) = 0 Then
            GetHeaderColumn = i
            Exit Function
        End If
    Next i
End Function

' =========================================================
' Safe Null / Variant Helpers
' =========================================================

Public Function NzDbl(ByVal v As Variant, Optional ByVal defaultValue As Double = 0) As Double
    If IsError(v) Then
        NzDbl = defaultValue
    ElseIf IsEmpty(v) Then
        NzDbl = defaultValue
    ElseIf Trim$(CStr(v)) = "" Then
        NzDbl = defaultValue
    ElseIf IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = defaultValue
    End If
End Function

Public Function NzStr(ByVal v As Variant, Optional ByVal defaultValue As String = "") As String
    If IsError(v) Then
        NzStr = defaultValue
    ElseIf IsEmpty(v) Then
        NzStr = defaultValue
    Else
        NzStr = Trim$(CStr(v))
        If NzStr = "" Then NzStr = defaultValue
    End If
End Function

Public Function NzDate(ByVal v As Variant, Optional ByVal defaultValue As Date = 0) As Date
    If IsError(v) Then
        NzDate = defaultValue
    ElseIf IsDate(v) Then
        NzDate = CDate(v)
    Else
        NzDate = defaultValue
    End If
End Function

' =========================================================
' UI Value Helpers
' ??:
' ???????,????????,?? UI ???????
' =========================================================

Public Function GetUIValue(ByVal ws As Worksheet, ByVal mainAddress As String, Optional ByVal fallbackRightCells As Long = 1) As Variant
    Dim v As Variant
    Dim rng As Range

    Set rng = ws.Range(mainAddress)
    v = rng.value

    If Trim$(CStr(v)) <> "" Then
        GetUIValue = v
    Else
        GetUIValue = rng.Offset(0, fallbackRightCells).value
    End If
End Function

Public Function GetUIString(ByVal ws As Worksheet, ByVal mainAddress As String, Optional ByVal fallbackRightCells As Long = 1) As String
    GetUIString = UCase$(Trim$(CStr(GetUIValue(ws, mainAddress, fallbackRightCells))))
End Function

Public Function GetUIDouble(ByVal ws As Worksheet, ByVal mainAddress As String, Optional ByVal fallbackRightCells As Long = 1, Optional ByVal defaultValue As Double = 0) As Double
    GetUIDouble = NzDbl(GetUIValue(ws, mainAddress, fallbackRightCells), defaultValue)
End Function

Public Function GetUIDate(ByVal ws As Worksheet, ByVal mainAddress As String, Optional ByVal fallbackRightCells As Long = 1) As Date
    GetUIDate = NzDate(GetUIValue(ws, mainAddress, fallbackRightCells), 0)
End Function

' =========================================================
' General Worksheet Helpers
' =========================================================

Public Function LastUsedRow(ByVal ws As Worksheet, ByVal colNum As Long) As Long
    If ws Is Nothing Then
        LastUsedRow = 1
    Else
        LastUsedRow = ws.Cells(ws.Rows.Count, colNum).End(xlUp).Row
    End If
End Function

Public Sub ClearForecastDailyView()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastCol As Long
    Dim clearRange As Range

    Set ws = GetWorksheet(WS_FORECAST_UI)
    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, COL_VIEW_SKU).End(xlUp).Row
    If lastRow < ROW_VIEW_DATA_START Then lastRow = ROW_VIEW_DATA_START

    lastCol = ws.Cells(ROW_VIEW_HEADER, ws.Columns.Count).End(xlToLeft).Column
    If lastCol < COL_VIEW_DATE_START Then lastCol = COL_VIEW_DATE_START

    Set clearRange = ws.Range(ws.Cells(ROW_VIEW_HEADER, COL_VIEW_SKU), ws.Cells(lastRow, lastCol))
    clearRange.ClearContents
    clearRange.Interior.Pattern = xlNone
End Sub

' =========================================================
' Products_DB Helpers
' Required headers:
'   SKU
'   Product_Name
'   Current_Stock
'   Reorder_Level
'   Lead_Time_Days
' =========================================================

Public Function GetProductNameBySKU(ByVal sku As String) As String
    GetProductNameBySKU = GetProductFieldBySKU(sku, "Product_Name")
End Function

Public Function GetCurrentStockBySKU(ByVal sku As String) As Double
    GetCurrentStockBySKU = NzDbl(GetProductFieldBySKU(sku, "Current_Stock"), 0)
End Function

Public Function GetReorderLevelBySKU(ByVal sku As String) As Double
    GetReorderLevelBySKU = NzDbl(GetProductFieldBySKU(sku, "Reorder_Level"), 0)
End Function

Public Function GetLeadTimeDaysBySKU(ByVal sku As String) As Double
    GetLeadTimeDaysBySKU = NzDbl(GetProductFieldBySKU(sku, "Lead_Time_Days"), 0)
End Function

Private Function GetProductFieldBySKU(ByVal sku As String, ByVal fieldName As String) As Variant
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim skuCol As Long
    Dim targetCol As Long
    Dim r As Long
    Dim c As Long

    Set ws = GetWorksheet(WS_PRODUCTS_DB)
    If ws Is Nothing Then Exit Function

    headerRow = 1
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    skuCol = 0
    targetCol = 0

    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(headerRow, c).value)), "SKU", vbTextCompare) = 0 Then
            skuCol = c
        End If
        If StrComp(Trim$(CStr(ws.Cells(headerRow, c).value)), fieldName, vbTextCompare) = 0 Then
            targetCol = c
        End If
    Next c

    If skuCol = 0 Or targetCol = 0 Then Exit Function

    For r = headerRow + 1 To lastRow
        If StrComp(Trim$(CStr(ws.Cells(r, skuCol).value)), Trim$(sku), vbTextCompare) = 0 Then
            GetProductFieldBySKU = ws.Cells(r, targetCol).value
            Exit Function
        End If
    Next r
End Function

' =========================================================
' Settings Helpers
' Assumption:
'   Settings!B12 = Default Safety Days
' =========================================================

Public Function GetSettingDefaultSafetyDays() As Double
    Dim ws As Worksheet

    Set ws = GetWorksheet(WS_SETTINGS)
    If ws Is Nothing Then
        GetSettingDefaultSafetyDays = 0
        Exit Function
    End If

    GetSettingDefaultSafetyDays = NzDbl(ws.Range("B12").value, 0)
End Function

' =========================================================
' Forecast_UI Scope / SKU Helpers
' =========================================================

Public Function GetSelectedSKUListFromUI() As Collection
    Dim ws As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim result As New Collection
    Dim sku As String

    Set ws = GetWorksheet(WS_FORECAST_UI)
    If ws Is Nothing Then
        Set GetSelectedSKUListFromUI = result
        Exit Function
    End If

    Set rng = ws.Range(RNG_SELECTED_SKUS)

    On Error Resume Next
    For Each cell In rng.Cells
        sku = Trim$(CStr(cell.value))
        If sku <> "" Then
            result.Add sku, UCase$(sku)
        End If
    Next cell
    On Error GoTo 0

    Set GetSelectedSKUListFromUI = result
End Function

Public Function GetScopeSKUList() As Collection
    Dim ws As Worksheet
    Dim applyTo As String
    Dim result As New Collection
    Dim selectedList As Collection
    Dim allList As Collection
    Dim i As Long

    Set ws = GetWorksheet(WS_FORECAST_UI)
    If ws Is Nothing Then
        Set GetScopeSKUList = result
        Exit Function
    End If

    applyTo = UCase$(Trim$(CStr(ws.Range(CELL_APPLY_TO).value)))

    Select Case applyTo
        Case APPLY_TO_SELECTED
            Set selectedList = GetSelectedSKUListFromUI()
            For i = 1 To selectedList.Count
                On Error Resume Next
                result.Add selectedList(i), UCase$(selectedList(i))
                On Error GoTo 0
            Next i

        Case APPLY_TO_ALL
            Set allList = GetAllProductSKUList()
            For i = 1 To allList.Count
                On Error Resume Next
                result.Add allList(i), UCase$(allList(i))
                On Error GoTo 0
            Next i

        Case Else
            ' invalid / blank
    End Select

    Set GetScopeSKUList = result
End Function

Public Function GetAllProductSKUList() As Collection
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim skuCol As Long
    Dim activeCol As Long
    Dim r As Long
    Dim c As Long
    Dim sku As String
    Dim activeText As String
    Dim result As New Collection

    Set ws = GetWorksheet(WS_PRODUCTS_DB)
    If ws Is Nothing Then
        Set GetAllProductSKUList = result
        Exit Function
    End If

    headerRow = 1
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    skuCol = 0
    activeCol = 0

    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(headerRow, c).value)), "SKU", vbTextCompare) = 0 Then skuCol = c
        If StrComp(Trim$(CStr(ws.Cells(headerRow, c).value)), "Active_Status", vbTextCompare) = 0 Then activeCol = c
    Next c

    If skuCol = 0 Then
        Set GetAllProductSKUList = result
        Exit Function
    End If

    On Error Resume Next
    For r = headerRow + 1 To lastRow
        sku = Trim$(CStr(ws.Cells(r, skuCol).value))
        If sku <> "" Then
            If activeCol > 0 Then
                activeText = UCase$(Trim$(CStr(ws.Cells(r, activeCol).value)))
                If activeText = "" Or activeText = "ACTIVE" Then
                    result.Add sku, UCase$(sku)
                End If
            Else
                result.Add sku, UCase$(sku)
            End If
        End If
    Next r
    On Error GoTo 0

    Set GetAllProductSKUList = result
End Function

' =========================================================
' Inventory / Sales History Helpers
' We count negative qty rows as sales / consumption.
' =========================================================

Public Function GetLastNDaysAverage(ByVal sku As String, ByVal daysBack As Long) As Double
    Dim ws As Worksheet
    Dim headerRow As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim c As Long, r As Long

    Dim skuCol As Long
    Dim qtyCol As Long
    Dim dateCol As Long
    Dim typeCol As Long

    Dim rowSKU As String
    Dim rowQty As Double
    Dim rowDate As Variant
    Dim rowType As String

    Dim totalSalesQty As Double
    Dim cutoffDate As Date

    If daysBack <= 0 Then
        GetLastNDaysAverage = 0
        Exit Function
    End If

    Set ws = GetWorksheet(WS_INVENTORY_LOG)
    If ws Is Nothing Then
        GetLastNDaysAverage = 0
        Exit Function
    End If

    headerRow = 1
    lastCol = ws.Cells(headerRow, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    skuCol = 0: qtyCol = 0: dateCol = 0: typeCol = 0

    For c = 1 To lastCol
        Select Case UCase$(Trim$(CStr(ws.Cells(headerRow, c).value)))
            Case "SKU"
                skuCol = c
            Case "QTY_CHANGE", "QUANTITY", "QTY"
                If qtyCol = 0 Then qtyCol = c
            Case "LOG_DATE", "TRANSACTION_DATE", "DATE"
                If dateCol = 0 Then dateCol = c
            Case "LOG_TYPE", "TRANS_TYPE", "TYPE"
                If typeCol = 0 Then typeCol = c
        End Select
    Next c

    If skuCol = 0 Or qtyCol = 0 Or dateCol = 0 Then
        GetLastNDaysAverage = 0
        Exit Function
    End If

    totalSalesQty = 0
    cutoffDate = Date - daysBack + 1

    For r = headerRow + 1 To lastRow
        rowSKU = Trim$(CStr(ws.Cells(r, skuCol).value))
        If StrComp(rowSKU, Trim$(sku), vbTextCompare) = 0 Then
            rowDate = ws.Cells(r, dateCol).value
            If IsDate(rowDate) Then
                If CDate(rowDate) >= cutoffDate And CDate(rowDate) <= Date Then
                    rowQty = NzDbl(ws.Cells(r, qtyCol).value, 0)

                    If typeCol > 0 Then
                        rowType = UCase$(Trim$(CStr(ws.Cells(r, typeCol).value)))
                    Else
                        rowType = ""
                    End If

                    If rowQty < 0 Then
                        totalSalesQty = totalSalesQty + Abs(rowQty)
                    ElseIf rowType = "SALE" Or rowType = "SALES" Then
                        totalSalesQty = totalSalesQty + Abs(rowQty)
                    End If
                End If
            End If
        End If
    Next r

    GetLastNDaysAverage = totalSalesQty / daysBack
End Function

Public Function GetLast7DaysAverage(ByVal sku As String) As Double
    GetLast7DaysAverage = GetLastNDaysAverage(sku, 7)
End Function

Public Function GetLast14DaysAverage(ByVal sku As String) As Double
    GetLast14DaysAverage = GetLastNDaysAverage(sku, 14)
End Function

Public Function GetLast30DaysAverage(ByVal sku As String) As Double
    GetLast30DaysAverage = GetLastNDaysAverage(sku, 30)
End Function

' =========================================================
' Rule Summary Helper
' =========================================================

Public Function BuildRuleSummaryText( _
    ByVal baseMethod As String, _
    ByVal histDays As Long, _
    ByVal trendOn As String, _
    ByVal seasonalityOn As String) As String

    Dim parts As String

    parts = ""

    Select Case UCase$(Trim$(baseMethod))
        Case BASE_METHOD_RECENT_AVG
            parts = "RA(" & histDays & ")"
        Case BASE_METHOD_WEIGHTED_AVG
            parts = "WA"
        Case BASE_METHOD_MANUAL_BASE
            parts = "MANUAL"
        Case Else
            parts = "BASE?"
    End Select

    If UCase$(Trim$(trendOn)) = YES_TEXT Then
        parts = parts & " + Trend"
    End If

    If UCase$(Trim$(seasonalityOn)) = YES_TEXT Then
        parts = parts & " + Seasonality"
    End If

    BuildRuleSummaryText = parts
End Function

' =========================================================
' UI Info Refresh Helper
' =========================================================

Public Sub RefreshForecastRuleInfoBySKU(ByVal sku As String)
    Dim ws As Worksheet

    Set ws = GetWorksheet(WS_FORECAST_UI)
    If ws Is Nothing Then Exit Sub

    If Trim$(sku) = "" Then
        ws.Range(CELL_INFO_CURRENT_STOCK).value = ""
        ws.Range(CELL_INFO_REORDER_LEVEL).value = ""
        ws.Range(CELL_INFO_LEADTIME).value = ""
        ws.Range(CELL_INFO_LAST7AVG).value = ""
        ws.Range(CELL_INFO_LAST30AVG).value = ""
        Exit Sub
    End If

    ws.Range(CELL_INFO_CURRENT_STOCK).value = GetCurrentStockBySKU(sku)
    ws.Range(CELL_INFO_REORDER_LEVEL).value = GetReorderLevelBySKU(sku)
    ws.Range(CELL_INFO_LEADTIME).value = GetLeadTimeDaysBySKU(sku)
    ws.Range(CELL_INFO_LAST7AVG).value = GetLast7DaysAverage(sku)
    ws.Range(CELL_INFO_LAST30AVG).value = GetLast30DaysAverage(sku)
End Sub

Public Function FindHeaderIndexInRange(ByVal rng As Range, ByVal headerName As String) As Long
    Dim c As Long

    FindHeaderIndexInRange = 0
    If rng Is Nothing Then Exit Function

    For c = 1 To rng.Columns.Count
        If StrComp(Trim$(CStr(rng.Cells(1, c).value)), Trim$(headerName), vbTextCompare) = 0 Then
            FindHeaderIndexInRange = c
            Exit Function
        End If
    Next c
End Function
