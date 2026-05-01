Attribute VB_Name = "modForecastView"
Option Explicit

' =========================================================
' Forecast View
' =========================================================

Public Sub LoadForecastView()
    Dim wsUI As Worksheet
    Dim skuList As Collection
    Dim startDate As Date
    Dim viewDays As Long
    Dim viewMode As String

    Set wsUI = GetWorksheet(WS_FORECAST_UI)
    If wsUI Is Nothing Then Exit Sub

    startDate = GetUIDate(wsUI, CELL_VIEW_START_DATE, 1)
    If startDate = 0 Then
        MsgBox "Please enter a valid View Start Date.", vbExclamation
        Exit Sub
    End If

    viewDays = CLng(GetUIDouble(wsUI, CELL_VIEW_DAYS, 1, DEFAULT_VIEW_DAYS))
    If viewDays <= 0 Then
        MsgBox "View Days must be greater than 0.", vbExclamation
        Exit Sub
    End If

    viewMode = GetUIString(wsUI, CELL_VIEW_MODE, 1)
    If viewMode = "" Then viewMode = VIEW_MODE_FINAL

    Set skuList = GetScopeSKUList()
    If skuList Is Nothing Or skuList.Count = 0 Then
        MsgBox "No SKU found for current scope.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo SafeExit

    ClearForecastDailyView
    BuildForecastViewHeaders startDate, viewDays
    LoadForecastViewRows skuList, startDate, viewDays, viewMode

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub BuildForecastViewHeaders(ByVal startDate As Date, ByVal viewDays As Long)
    Dim ws As Worksheet
    Dim i As Long
    Dim colNum As Long

    Set ws = GetWorksheet(WS_FORECAST_UI)
    If ws Is Nothing Then Exit Sub

    ' Fixed headers
    ws.Cells(ROW_VIEW_HEADER, COL_VIEW_SKU).value = "SKU"
    ws.Cells(ROW_VIEW_HEADER, COL_VIEW_PRODUCT).value = "Product Name"
    ws.Cells(ROW_VIEW_HEADER, COL_VIEW_STOCK).value = "Current Stock"
    ws.Cells(ROW_VIEW_HEADER, COL_VIEW_SOURCE).value = "Source"
    ws.Cells(ROW_VIEW_HEADER, COL_VIEW_RULESUMMARY).value = "Rule Summary"

    ' Dynamic date headers
    For i = 0 To viewDays - 1
        colNum = COL_VIEW_DATE_START + i
        ws.Cells(ROW_VIEW_HEADER, colNum).value = startDate + i
        ws.Cells(ROW_VIEW_HEADER, colNum).NumberFormat = "yyyy-mm-dd"
    Next i
End Sub

Public Sub LoadForecastViewRows( _
    ByVal skuList As Collection, _
    ByVal startDate As Date, _
    ByVal viewDays As Long, _
    ByVal viewMode As String)

    Dim ws As Worksheet
    Dim r As Long
    Dim i As Long
    Dim sku As String
    Dim colNum As Long
    Dim targetDate As Date
    Dim displayValue As Variant
    Dim sourceText As String
    Dim ruleSummary As String
    Dim productName As String
    Dim currentStock As Double

    Set ws = GetWorksheet(WS_FORECAST_UI)
    If ws Is Nothing Then Exit Sub

    r = ROW_VIEW_DATA_START

    For i = 1 To skuList.Count
        sku = CStr(skuList(i))

        productName = GetProductNameBySKU(sku)
        currentStock = GetCurrentStockBySKU(sku)
        sourceText = GetForecastSourceBySKU(sku, startDate)
        ruleSummary = GetRuleSummaryBySKU(sku)

        ws.Cells(r, COL_VIEW_SKU).value = sku
        ws.Cells(r, COL_VIEW_PRODUCT).value = productName
        ws.Cells(r, COL_VIEW_STOCK).value = currentStock
        ws.Cells(r, COL_VIEW_SOURCE).value = sourceText
        ws.Cells(r, COL_VIEW_RULESUMMARY).value = ruleSummary

        Dim d As Long
        For d = 0 To viewDays - 1
            colNum = COL_VIEW_DATE_START + d
            targetDate = startDate + d

            displayValue = GetForecastDisplayQty(sku, targetDate, viewMode)
            ws.Cells(r, colNum).value = displayValue
        Next d

        r = r + 1
    Next i
End Sub

Public Function GetForecastDisplayQty(ByVal sku As String, ByVal targetDate As Date, ByVal viewMode As String) As Variant
    Dim lo As ListObject
    Dim lr As ListRow

    Dim skuCol As Long
    Dim dateCol As Long
    Dim sysCol As Long
    Dim overCol As Long
    Dim finalCol As Long

    Dim rowDate As Variant
    Dim rowSKU As String

    Set lo = GetTable(WS_FORECAST_DB, TBL_FORECAST)
    If lo Is Nothing Then Exit Function

    skuCol = GetHeaderColumn(lo, "SKU")
    dateCol = GetHeaderColumn(lo, "Forecast_Date")
    sysCol = GetHeaderColumn(lo, "System_Forecast_Qty")
    overCol = GetHeaderColumn(lo, "Override_Forecast_Qty")
    finalCol = GetHeaderColumn(lo, "Final_Forecast_Qty")

    If skuCol = 0 Or dateCol = 0 Then Exit Function

    For Each lr In lo.ListRows
        rowSKU = Trim$(CStr(lr.Range.Cells(1, skuCol).value))
        If StrComp(rowSKU, Trim$(sku), vbTextCompare) = 0 Then

            rowDate = lr.Range.Cells(1, dateCol).value
            If IsDate(rowDate) Then
                If CLng(CDate(rowDate)) = CLng(targetDate) Then

                    Select Case UCase$(Trim$(viewMode))
                        Case VIEW_MODE_SYSTEM
                            If sysCol > 0 Then
                                GetForecastDisplayQty = lr.Range.Cells(1, sysCol).value
                            End If

                        Case VIEW_MODE_OVERRIDE
                            If overCol > 0 Then
                                GetForecastDisplayQty = lr.Range.Cells(1, overCol).value
                            End If

                        Case Else   ' FINAL
                            If finalCol > 0 Then
                                GetForecastDisplayQty = lr.Range.Cells(1, finalCol).value
                            End If
                    End Select

                    Exit Function
                End If
            End If
        End If
    Next lr

    GetForecastDisplayQty = ""
End Function

Public Function GetForecastSourceBySKU(ByVal sku As String, ByVal startDate As Date) As String
    Dim lo As ListObject
    Dim lr As ListRow

    Dim skuCol As Long
    Dim dateCol As Long
    Dim sourceCol As Long

    Dim rowDate As Variant
    Dim rowSKU As String

    Set lo = GetTable(WS_FORECAST_DB, TBL_FORECAST)
    If lo Is Nothing Then Exit Function

    skuCol = GetHeaderColumn(lo, "SKU")
    dateCol = GetHeaderColumn(lo, "Forecast_Date")
    sourceCol = GetHeaderColumn(lo, "Forecast_Source")

    If skuCol = 0 Or dateCol = 0 Or sourceCol = 0 Then Exit Function

    For Each lr In lo.ListRows
        rowSKU = Trim$(CStr(lr.Range.Cells(1, skuCol).value))
        If StrComp(rowSKU, Trim$(sku), vbTextCompare) = 0 Then

            rowDate = lr.Range.Cells(1, dateCol).value
            If IsDate(rowDate) Then
                If CLng(CDate(rowDate)) = CLng(startDate) Then
                    GetForecastSourceBySKU = NzStr(lr.Range.Cells(1, sourceCol).value, "")
                    Exit Function
                End If
            End If
        End If
    Next lr

    GetForecastSourceBySKU = ""
End Function

Public Function GetRuleSummaryBySKU(ByVal sku As String) As String
    Dim lo As ListObject
    Dim ruleRow As Range

    Dim baseMethod As String
    Dim histDays As Long
    Dim trendOn As String
    Dim seasonalityOn As String

    Set lo = GetTable(WS_FORECAST_RULES_DB, TBL_FORECAST_RULES)
    If lo Is Nothing Then Exit Function

    Set ruleRow = GetForecastRuleRowBySKU(sku)
    If ruleRow Is Nothing Then Exit Function

    baseMethod = NzStr(GetCellByHeader_View(ruleRow, lo, "Base_Method"), "")
    histDays = CLng(NzDbl(GetCellByHeader_View(ruleRow, lo, "Historical_Window_Days"), DEFAULT_HIST_DAYS))
    trendOn = NzStr(GetCellByHeader_View(ruleRow, lo, "Trend_On"), NO_TEXT)

    If GetHeaderColumn(lo, "Seasonality_On") > 0 Then
        seasonalityOn = NzStr(GetCellByHeader_View(ruleRow, lo, "Seasonality_On"), YES_TEXT)
    Else
        seasonalityOn = YES_TEXT
    End If

    GetRuleSummaryBySKU = BuildRuleSummaryText(baseMethod, histDays, trendOn, seasonalityOn)
End Function

' =========================================================
' Local helpers
' =========================================================

Private Function GetCellByHeader_View(ByVal rowRange As Range, ByVal lo As ListObject, ByVal headerName As String) As Variant
    Dim colIndex As Long

    colIndex = GetHeaderColumn(lo, headerName)
    If colIndex = 0 Then Exit Function

    GetCellByHeader_View = rowRange.Cells(1, colIndex).value
End Function

