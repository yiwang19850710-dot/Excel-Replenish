Attribute VB_Name = "modForecastEngine"
Option Explicit

Public Sub GenerateForecast_Run()
    Dim wsUI As Worksheet
    Dim skuList As Collection
    Dim sku As Variant
    Dim runID As String
    Dim startDate As Date
    Dim horizonDays As Long
    Dim replaceExisting As Boolean
    Dim missingOnly As Boolean

    Dim generatedSKUCount As Long
    Dim generatedRowCount As Long
    Dim skuRows As Long

    Set wsUI = GetWorksheet(WS_FORECAST_UI)
    If wsUI Is Nothing Then Exit Sub

    startDate = GetUIDate(wsUI, CELL_GEN_START_DATE, 1)
    If startDate = 0 Then
        MsgBox "Please enter a valid Start Date.", vbExclamation
        Exit Sub
    End If

    horizonDays = CLng(GetUIDouble(wsUI, CELL_GEN_HORIZON_DAYS, 1, 0))
    If horizonDays <= 0 Then
        MsgBox "Horizon Days must be greater than 0.", vbExclamation
        Exit Sub
    End If

    replaceExisting = (GetUIString(wsUI, CELL_GEN_REPLACE_EXISTING, 1) = YES_TEXT)
    missingOnly = (GetUIString(wsUI, CELL_GEN_MISSING_ONLY, 1) = YES_TEXT)

    If replaceExisting And missingOnly Then
        MsgBox "Replace Existing and Generate Missing Only cannot both be YES.", vbExclamation
        Exit Sub
    End If

    If (Not replaceExisting) And (Not missingOnly) Then
        MsgBox "Please set either Replace Existing = YES or Generate Missing Only = YES.", vbExclamation
        Exit Sub
    End If

    Set skuList = GetScopeSKUList()
    If skuList Is Nothing Or skuList.Count = 0 Then
        MsgBox "No SKU found for current scope.", vbExclamation
        Exit Sub
    End If

    runID = NextForecastRunID()
    wsUI.Range(CELL_GEN_RUN_PREVIEW).value = "Generating " & horizonDays & " day(s)..."

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo SafeExit

    generatedSKUCount = 0
    generatedRowCount = 0

    For Each sku In skuList
        skuRows = BuildForecastForSKU(CStr(sku), runID, startDate, horizonDays, replaceExisting, missingOnly)

        If skuRows > 0 Then
            generatedSKUCount = generatedSKUCount + 1
            generatedRowCount = generatedRowCount + skuRows
        End If
    Next sku

    wsUI.Range(CELL_GEN_RUN_PREVIEW).value = runID & " | " & generatedSKUCount & " SKU(s) | " & generatedRowCount & " row(s)"

    MsgBox "Forecast generated successfully." & vbCrLf & vbCrLf & _
           "Run ID: " & runID & vbCrLf & _
           "SKUs: " & generatedSKUCount & vbCrLf & _
           "Rows: " & generatedRowCount, vbInformation

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Function BuildForecastForSKU( _
    ByVal sku As String, _
    ByVal runID As String, _
    ByVal startDate As Date, _
    ByVal horizonDays As Long, _
    ByVal replaceExisting As Boolean, _
    ByVal missingOnly As Boolean) As Long

    Dim loForecast As ListObject
    Dim loRules As ListObject
    Dim ruleRow As Range
    Dim ruleID As String
    Dim productName As String

    Dim baseMethod As String
    Dim histDays As Long
    Dim trendOn As String
    Dim trendSensitivity As String
    Dim noHistoryMethod As String
    Dim noHistoryDefaultQty As Double
    Dim manualBaseQty As Double
    Dim seasonalityOn As String
    Dim promoMultiplier As Double
    Dim promoAddOn As Double
    Dim roundRule As String

    Dim i As Long
    Dim targetDate As Date
    Dim systemQty As Double
    Dim sourceText As String

    Dim baseDemand As Double
    Dim trendFactor As Double
    Dim seasonalityFactor As Double
    Dim finalSystemQty As Double

    Dim newRow As ListRow
    Dim wroteCount As Long

    Set loForecast = GetTable(WS_FORECAST_DB, TBL_FORECAST)
    If loForecast Is Nothing Then Exit Function

    Set loRules = GetTable(WS_FORECAST_RULES_DB, TBL_FORECAST_RULES)
    If loRules Is Nothing Then Exit Function

    If Trim$(sku) = "" Then Exit Function

    ruleID = EnsureForecastRuleForSKU(sku)
    If ruleID = "" Then Exit Function

    Set ruleRow = GetForecastRuleRowBySKU(sku)
    If ruleRow Is Nothing Then Exit Function

    productName = GetProductNameBySKU(sku)

    baseMethod = UCase$(Trim$(CStr(GetCellByHeader_Engine(ruleRow, loRules, "Base_Method"))))
    histDays = CLng(NzDbl(GetCellByHeader_Engine(ruleRow, loRules, "Historical_Window_Days"), DEFAULT_HIST_DAYS))
    trendOn = UCase$(Trim$(CStr(GetCellByHeader_Engine(ruleRow, loRules, "Trend_On"))))
    trendSensitivity = UCase$(Trim$(CStr(GetCellByHeader_Engine(ruleRow, loRules, "Trend_Sensitivity"))))
    noHistoryMethod = UCase$(Trim$(CStr(GetCellByHeader_Engine(ruleRow, loRules, "NoHistory_Method"))))
    noHistoryDefaultQty = NzDbl(GetCellByHeader_Engine(ruleRow, loRules, "NoHistory_Default_Qty"), DEFAULT_NOHISTORY_QTY)
    manualBaseQty = NzDbl(GetCellByHeader_Engine(ruleRow, loRules, "Manual_Base_Qty"), 0)
    promoMultiplier = NzDbl(GetCellByHeader_Engine(ruleRow, loRules, "Promo_Multiplier"), DEFAULT_PROMO_MULTIPLIER)
    promoAddOn = NzDbl(GetCellByHeader_Engine(ruleRow, loRules, "Promo_AddOn"), DEFAULT_PROMO_ADDON)
    roundRule = UCase$(Trim$(CStr(GetCellByHeader_Engine(ruleRow, loRules, "Round_Rule"))))

    If GetHeaderColumn(loRules, "Seasonality_On") > 0 Then
        seasonalityOn = UCase$(Trim$(CStr(GetCellByHeader_Engine(ruleRow, loRules, "Seasonality_On"))))
    Else
        seasonalityOn = DEFAULT_SEASONALITY_ON
    End If

    If replaceExisting Then
        DeleteExistingForecastRows sku, startDate, horizonDays
    End If

    wroteCount = 0

    For i = 0 To horizonDays - 1
        targetDate = startDate + i

        If missingOnly Then
            If ForecastRowExists(sku, targetDate) Then GoTo ContinueLoop
        End If

        baseDemand = CalcBaseDemand(sku, baseMethod, histDays, manualBaseQty, sourceText)

        If baseDemand <= 0 Then
            baseDemand = CalcNoHistoryFallback(sku, noHistoryMethod, noHistoryDefaultQty, sourceText)
        End If

        trendFactor = CalcTrendFactor(sku, trendOn, trendSensitivity)
        seasonalityFactor = CalcSeasonalityFactor(sku, targetDate, seasonalityOn, histDays)

        finalSystemQty = baseDemand * trendFactor * seasonalityFactor
        finalSystemQty = finalSystemQty * promoMultiplier + promoAddOn
        finalSystemQty = ApplyRoundRule(finalSystemQty, roundRule)

        If finalSystemQty < 0 Then finalSystemQty = 0
        systemQty = finalSystemQty

        Set newRow = loForecast.ListRows.Add

        SetCellByHeader_Engine newRow.Range, loForecast, "Forecast_ID", NextForecastID()
        SetCellByHeader_Engine newRow.Range, loForecast, "Run_ID", runID
        SetCellByHeader_Engine newRow.Range, loForecast, "SKU", sku
        SetCellByHeader_Engine newRow.Range, loForecast, "Product_Name", productName
        SetCellByHeader_Engine newRow.Range, loForecast, "Forecast_Date", targetDate
        SetCellByHeader_Engine newRow.Range, loForecast, "System_Forecast_Qty", systemQty
        SetCellByHeader_Engine newRow.Range, loForecast, "Override_Forecast_Qty", ""
        SetCellByHeader_Engine newRow.Range, loForecast, "Final_Forecast_Qty", systemQty
        SetCellByHeader_Engine newRow.Range, loForecast, "Forecast_Source", sourceText
        SetCellByHeader_Engine newRow.Range, loForecast, "Rule_ID", ruleID
        SetCellByHeader_Engine newRow.Range, loForecast, "Run_Date", Now
        SetCellByHeader_Engine newRow.Range, loForecast, "Updated_At", Now

        wroteCount = wroteCount + 1

ContinueLoop:
    Next i

    BuildForecastForSKU = wroteCount
End Function

Public Function CalcBaseDemand( _
    ByVal sku As String, _
    ByVal baseMethod As String, _
    ByVal histDays As Long, _
    ByVal manualBase As Double, _
    ByRef sourceText As String) As Double

    Select Case UCase$(Trim$(baseMethod))
        Case BASE_METHOD_RECENT_AVG
            CalcBaseDemand = CalcRecentAverage(sku, histDays)
            sourceText = "RECENT_AVG"

        Case BASE_METHOD_WEIGHTED_AVG
            CalcBaseDemand = CalcWeightedAverage(sku)
            sourceText = "WEIGHTED_AVG"

        Case BASE_METHOD_MANUAL_BASE
            CalcBaseDemand = manualBase
            sourceText = "MANUAL_BASE"

        Case Else
            CalcBaseDemand = CalcWeightedAverage(sku)
            sourceText = "WEIGHTED_AVG"
    End Select
End Function

Public Function CalcRecentAverage(ByVal sku As String, ByVal histDays As Long) As Double
    If histDays <= 0 Then histDays = DEFAULT_HIST_DAYS
    CalcRecentAverage = GetLastNDaysAverage(sku, histDays)
End Function

Public Function CalcWeightedAverage(ByVal sku As String) As Double
    Dim avg7 As Double, avg14 As Double
    avg7 = GetLast7DaysAverage(sku)
    avg14 = GetLast14DaysAverage(sku)
    CalcWeightedAverage = (avg7 * 0.7) + (avg14 * 0.3)
End Function

Public Function CalcTrendFactor(ByVal sku As String, ByVal trendOn As String, ByVal trendSensitivity As String) As Double
    Dim avg7 As Double, avg14 As Double
    Dim trendRaw As Double
    Dim weightFactor As Double

    If UCase$(Trim$(trendOn)) <> YES_TEXT Then
        CalcTrendFactor = 1
        Exit Function
    End If

    avg7 = GetLast7DaysAverage(sku)
    avg14 = GetLast14DaysAverage(sku)

    If avg14 <= 0 Then
        CalcTrendFactor = 1
        Exit Function
    End If

    trendRaw = avg7 / avg14

    Select Case UCase$(Trim$(trendSensitivity))
        Case TREND_CONSERVATIVE
            weightFactor = 0.25
        Case TREND_AGGRESSIVE
            weightFactor = 0.75
        Case Else
            weightFactor = 0.5
    End Select

    CalcTrendFactor = 1 + (trendRaw - 1) * weightFactor

    If CalcTrendFactor < 0.5 Then CalcTrendFactor = 0.5
    If CalcTrendFactor > 1.5 Then CalcTrendFactor = 1.5
End Function

Public Function CalcSeasonalityFactor(ByVal sku As String, _
                                      ByVal targetDate As Date, _
                                      ByVal seasonalityOn As String, _
                                      ByVal historicalDays As Long) As Double

    If UCase$(Trim$(seasonalityOn)) <> YES_TEXT Then
        CalcSeasonalityFactor = 1
        Exit Function
    End If

    CalcSeasonalityFactor = CalcAutoSeasonalityFactor(sku, targetDate, historicalDays)
End Function

Public Function CalcAutoSeasonalityFactor(ByVal sku As String, _
                                          ByVal targetDate As Date, _
                                          ByVal historicalDays As Long) As Double

    Dim wsS As Worksheet
    Dim rng As Range
    Dim lastRow As Long

    Dim colSKU As Long
    Dim colQty As Long
    Dim colDate As Long

    Dim i As Long
    Dim salesDate As Variant
    Dim qty As Double

    Dim totalQty As Double
    Dim totalDays As Long
    Dim weekdayQty As Double
    Dim weekdayCount As Long

    Dim overallAvg As Double
    Dim weekdayAvg As Double
    Dim factor As Double
    Dim targetWeekday As Long

    On Error GoTo FailSafe

    If historicalDays <= 0 Then historicalDays = DEFAULT_AUTO_SEASONALITY_WINDOW_DAYS

    Set wsS = ThisWorkbook.Worksheets(WS_SALES_DB)
    lastRow = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then GoTo NotEnoughData

    Set rng = wsS.Range("A1").CurrentRegion

    colSKU = FindHeaderIndexInRange(rng, "SKU")
    colQty = FindHeaderIndexInRange(rng, "Qty")
    colDate = FindHeaderIndexInRange(rng, "Sales_Date")

    If colSKU = 0 Or colQty = 0 Or colDate = 0 Then GoTo NotEnoughData

    targetWeekday = Weekday(targetDate, vbMonday)

    totalQty = 0
    totalDays = 0
    weekdayQty = 0
    weekdayCount = 0

    For i = 2 To rng.Rows.Count
        If Trim$(CStr(rng.Cells(i, colSKU).value)) = sku Then
            salesDate = rng.Cells(i, colDate).value

            If IsDate(salesDate) Then
                If CDate(salesDate) >= targetDate - historicalDays And CDate(salesDate) < targetDate Then
                    qty = NzDbl(rng.Cells(i, colQty).value, 0)

                    totalQty = totalQty + qty
                    totalDays = totalDays + 1

                    If Weekday(CDate(salesDate), vbMonday) = targetWeekday Then
                        weekdayQty = weekdayQty + qty
                        weekdayCount = weekdayCount + 1
                    End If
                End If
            End If
        End If
    Next i

    If totalDays < MIN_AUTO_SEASONALITY_TOTAL_DAYS Then GoTo NotEnoughData
    If weekdayCount < MIN_AUTO_SEASONALITY_WEEKDAY_COUNT Then GoTo NotEnoughData

    overallAvg = totalQty / totalDays
    If overallAvg <= 0 Then GoTo NotEnoughData

    weekdayAvg = weekdayQty / weekdayCount
    factor = weekdayAvg / overallAvg

    If factor < MIN_AUTO_SEASONALITY_FACTOR Then factor = MIN_AUTO_SEASONALITY_FACTOR
    If factor > MAX_AUTO_SEASONALITY_FACTOR Then factor = MAX_AUTO_SEASONALITY_FACTOR

    CalcAutoSeasonalityFactor = Round(factor, 2)
    Exit Function

NotEnoughData:
    CalcAutoSeasonalityFactor = 1
    Exit Function

FailSafe:
    CalcAutoSeasonalityFactor = 1
End Function

Public Function CalcNoHistoryFallback( _
    ByVal sku As String, _
    ByVal noHistoryMethod As String, _
    ByVal noHistoryDefaultQty As Double, _
    ByRef sourceText As String) As Double

    Dim reorderLevel As Double
    Dim safetyDays As Double

    Select Case UCase$(Trim$(noHistoryMethod))
        Case NOHISTORY_ZERO
            CalcNoHistoryFallback = 0
            sourceText = "NOHISTORY_ZERO"

        Case NOHISTORY_MANUAL_DEFAULT
            CalcNoHistoryFallback = noHistoryDefaultQty
            sourceText = "NOHISTORY_MANUAL"

        Case NOHISTORY_REORDER_BASED
            reorderLevel = GetReorderLevelBySKU(sku)
            safetyDays = GetSettingDefaultSafetyDays()
            If safetyDays <= 0 Then safetyDays = 1

            If reorderLevel > 0 Then
                CalcNoHistoryFallback = reorderLevel / safetyDays
                sourceText = "NOHISTORY_REORDER"
            Else
                CalcNoHistoryFallback = noHistoryDefaultQty
                sourceText = "NOHISTORY_MANUAL"
            End If

        Case Else
            CalcNoHistoryFallback = noHistoryDefaultQty
            sourceText = "NOHISTORY_MANUAL"
    End Select
End Function

Public Function ApplyRoundRule(ByVal valueIn As Double, ByVal roundRule As String) As Double
    Select Case UCase$(Trim$(roundRule))
        Case ROUND_RULE_ROUNDUP
            ApplyRoundRule = Application.WorksheetFunction.RoundUp(valueIn, 0)
        Case ROUND_RULE_ROUNDDOWN
            ApplyRoundRule = Application.WorksheetFunction.RoundDown(valueIn, 0)
        Case Else
            ApplyRoundRule = WorksheetFunction.Round(valueIn, 0)
    End Select
End Function

Public Function NextForecastRunID() As String
    Dim lo As ListObject
    Dim lr As ListRow
    Dim runCol As Long
    Dim txt As String
    Dim maxNum As Long
    Dim n As Long
    Dim todayPart As String

    Set lo = GetTable(WS_FORECAST_DB, TBL_FORECAST)
    If lo Is Nothing Then
        NextForecastRunID = "RUN-" & Format$(Date, "yyyymmdd") & "-001"
        Exit Function
    End If

    runCol = GetHeaderColumn(lo, "Run_ID")
    If runCol = 0 Then
        NextForecastRunID = "RUN-" & Format$(Date, "yyyymmdd") & "-001"
        Exit Function
    End If

    todayPart = "RUN-" & Format$(Date, "yyyymmdd") & "-"
    maxNum = 0

    For Each lr In lo.ListRows
        txt = NzStr(lr.Range.Cells(1, runCol).value, "")
        If UCase$(Left$(txt, Len(todayPart))) = UCase$(todayPart) Then
            n = val(Mid$(txt, Len(todayPart) + 1))
            If n > maxNum Then maxNum = n
        End If
    Next lr

    NextForecastRunID = todayPart & Format$(maxNum + 1, "000")
End Function

Public Function NextForecastID() As String
    Dim lo As ListObject
    Dim lr As ListRow
    Dim idCol As Long
    Dim txt As String
    Dim maxNum As Long
    Dim n As Long

    Set lo = GetTable(WS_FORECAST_DB, TBL_FORECAST)
    If lo Is Nothing Then
        NextForecastID = "FC-000001"
        Exit Function
    End If

    idCol = GetHeaderColumn(lo, "Forecast_ID")
    If idCol = 0 Then
        NextForecastID = "FC-000001"
        Exit Function
    End If

    maxNum = 0

    For Each lr In lo.ListRows
        txt = NzStr(lr.Range.Cells(1, idCol).value, "")
        If UCase$(Left$(txt, 3)) = "FC-" Then
            n = val(Mid$(txt, 4))
            If n > maxNum Then maxNum = n
        End If
    Next lr

    NextForecastID = "FC-" & Format$(maxNum + 1, "000000")
End Function

Public Sub DeleteExistingForecastRows(ByVal sku As String, ByVal startDate As Date, ByVal horizonDays As Long)
    Dim lo As ListObject
    Dim skuCol As Long
    Dim dateCol As Long
    Dim i As Long
    Dim rowSKU As String
    Dim rowDate As Variant
    Dim endDate As Date

    Set lo = GetTable(WS_FORECAST_DB, TBL_FORECAST)
    If lo Is Nothing Then Exit Sub

    skuCol = GetHeaderColumn(lo, "SKU")
    dateCol = GetHeaderColumn(lo, "Forecast_Date")
    If skuCol = 0 Or dateCol = 0 Then Exit Sub

    endDate = startDate + horizonDays - 1

    For i = lo.ListRows.Count To 1 Step -1
        rowSKU = Trim$(CStr(lo.DataBodyRange.Cells(i, skuCol).value))
        rowDate = lo.DataBodyRange.Cells(i, dateCol).value

        If StrComp(rowSKU, sku, vbTextCompare) = 0 Then
            If IsDate(rowDate) Then
                If CDate(rowDate) >= startDate And CDate(rowDate) <= endDate Then
                    lo.ListRows(i).Delete
                End If
            End If
        End If
    Next i
End Sub

Public Function ForecastRowExists(ByVal sku As String, ByVal targetDate As Date) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow
    Dim skuCol As Long
    Dim dateCol As Long
    Dim rowSKU As String
    Dim rowDate As Variant

    Set lo = GetTable(WS_FORECAST_DB, TBL_FORECAST)
    If lo Is Nothing Then Exit Function

    skuCol = GetHeaderColumn(lo, "SKU")
    dateCol = GetHeaderColumn(lo, "Forecast_Date")
    If skuCol = 0 Or dateCol = 0 Then Exit Function

    For Each lr In lo.ListRows
        rowSKU = Trim$(CStr(lr.Range.Cells(1, skuCol).value))
        rowDate = lr.Range.Cells(1, dateCol).value

        If StrComp(rowSKU, sku, vbTextCompare) = 0 Then
            If IsDate(rowDate) Then
                If CLng(CDate(rowDate)) = CLng(targetDate) Then
                    ForecastRowExists = True
                    Exit Function
                End If
            End If
        End If
    Next lr
End Function

Private Function GetCellByHeader_Engine(ByVal rowRange As Range, ByVal lo As ListObject, ByVal headerName As String) As Variant
    Dim colIndex As Long
    colIndex = GetHeaderColumn(lo, headerName)
    If colIndex = 0 Then Exit Function
    GetCellByHeader_Engine = rowRange.Cells(1, colIndex).value
End Function

Private Sub SetCellByHeader_Engine(ByVal rowRange As Range, ByVal lo As ListObject, ByVal headerName As String, ByVal newValue As Variant)
    Dim colIndex As Long
    colIndex = GetHeaderColumn(lo, headerName)
    If colIndex = 0 Then Exit Sub
    rowRange.Cells(1, colIndex).value = newValue
End Sub

