Attribute VB_Name = "modForecastRules"
Option Explicit

Public Sub ForecastRule_SetupUIValidation()
    Dim wsUI As Worksheet

    Set wsUI = GetWorksheet(WS_FORECAST_UI)
    If wsUI Is Nothing Then Exit Sub

    On Error Resume Next
    wsUI.Range(CELL_SEASONALITY_ON).Validation.Delete
    On Error GoTo 0

    With wsUI.Range(CELL_SEASONALITY_ON).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="YES,NO"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Seasonality On"
        .ErrorTitle = "Invalid Value"
        .InputMessage = "Choose YES or NO."
        .ErrorMessage = "Please choose YES or NO."
        .ShowInput = True
        .ShowError = True
    End With
End Sub

Public Sub ForecastRule_Save()
    Dim wsUI As Worksheet
    Dim lo As ListObject
    Dim skuList As Collection
    Dim sku As Variant
    Dim savedCount As Long

    Set wsUI = GetWorksheet(WS_FORECAST_UI)
    If wsUI Is Nothing Then Exit Sub

    ForecastRule_SetupUIValidation

    Set lo = GetTable(WS_FORECAST_RULES_DB, TBL_FORECAST_RULES)
    If lo Is Nothing Then Exit Sub

    Set skuList = GetScopeSKUList()

    If skuList.Count = 0 Then
        MsgBox "No SKU found for current scope. Please check Apply To / SKU selection.", vbExclamation
        Exit Sub
    End If

    savedCount = 0

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo SafeExit

    For Each sku In skuList
        SaveRuleForSingleSKU CStr(sku)
        savedCount = savedCount + 1
    Next sku

    MsgBox savedCount & " forecast rule(s) saved.", vbInformation

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub SaveRuleForSingleSKU(ByVal sku As String)
    Dim lo As ListObject
    Dim targetRow As Range
    Dim newRow As ListRow
    Dim ruleID As String
    Dim productName As String
    Dim seasonalityOn As String

    Dim wsUI As Worksheet
    Dim baseMethod As String
    Dim histDays As Long
    Dim trendOn As String
    Dim trendSensitivity As String
    Dim noHistoryMethod As String
    Dim noHistoryDefaultQty As Double
    Dim manualBaseQty As Double
    Dim promoMultiplier As Double
    Dim promoAddOn As Double
    Dim roundRule As String

    Set lo = GetTable(WS_FORECAST_RULES_DB, TBL_FORECAST_RULES)
    If lo Is Nothing Then Exit Sub
    If Trim$(sku) = "" Then Exit Sub

    Set wsUI = GetWorksheet(WS_FORECAST_UI)
    If wsUI Is Nothing Then Exit Sub

    baseMethod = UCase$(Trim$(CStr(wsUI.Range(CELL_BASE_METHOD).value)))
    histDays = CLng(NzDbl(wsUI.Range(CELL_HIST_DAYS).value, DEFAULT_HIST_DAYS))
    trendOn = UCase$(Trim$(CStr(wsUI.Range(CELL_TREND_ON).value)))
    trendSensitivity = UCase$(Trim$(CStr(wsUI.Range(CELL_TREND_SENSITIVITY).value)))
    noHistoryMethod = UCase$(Trim$(CStr(wsUI.Range(CELL_NOHISTORY_METHOD).value)))
    noHistoryDefaultQty = NzDbl(wsUI.Range(CELL_NOHISTORY_DEFAULT_QTY).value, DEFAULT_NOHISTORY_QTY)
    manualBaseQty = NzDbl(wsUI.Range(CELL_MANUAL_BASE_QTY).value, 0)
    promoMultiplier = NzDbl(wsUI.Range(CELL_PROMO_MULTIPLIER).value, DEFAULT_PROMO_MULTIPLIER)
    promoAddOn = NzDbl(wsUI.Range(CELL_PROMO_ADDON).value, DEFAULT_PROMO_ADDON)
    seasonalityOn = UCase$(Trim$(CStr(wsUI.Range(CELL_SEASONALITY_ON).value)))
    roundRule = UCase$(Trim$(CStr(wsUI.Range(CELL_ROUND_RULE).value)))

    If baseMethod = "" Then baseMethod = BASE_METHOD_WEIGHTED_AVG
    If histDays <= 0 Then histDays = DEFAULT_HIST_DAYS
    If trendOn = "" Then trendOn = YES_TEXT
    If trendSensitivity = "" Then trendSensitivity = TREND_NORMAL
    If noHistoryMethod = "" Then noHistoryMethod = NOHISTORY_REORDER_BASED
    If seasonalityOn <> YES_TEXT Then seasonalityOn = NO_TEXT
    If roundRule = "" Then roundRule = ROUND_RULE_ROUND

    productName = GetProductNameBySKU(sku)

    Set targetRow = GetForecastRuleRowBySKU(sku)

    If targetRow Is Nothing Then
        ruleID = NextForecastRuleID()
        Set newRow = lo.ListRows.Add

        WriteRuleRow newRow.Range, lo, ruleID, sku, productName, baseMethod, histDays, trendOn, trendSensitivity, _
                     noHistoryMethod, noHistoryDefaultQty, manualBaseQty, seasonalityOn, promoMultiplier, promoAddOn, roundRule
    Else
        ruleID = NzStr(targetRow.Cells(1, GetHeaderColumn(lo, "Rule_ID")).value, "")
        If ruleID = "" Then ruleID = NextForecastRuleID()

        WriteRuleRow targetRow, lo, ruleID, sku, productName, baseMethod, histDays, trendOn, trendSensitivity, _
                     noHistoryMethod, noHistoryDefaultQty, manualBaseQty, seasonalityOn, promoMultiplier, promoAddOn, roundRule
    End If
End Sub

Private Sub WriteRuleRow( _
    ByVal targetRow As Range, _
    ByVal lo As ListObject, _
    ByVal ruleID As String, _
    ByVal sku As String, _
    ByVal productName As String, _
    ByVal baseMethod As String, _
    ByVal histDays As Long, _
    ByVal trendOn As String, _
    ByVal trendSensitivity As String, _
    ByVal noHistoryMethod As String, _
    ByVal noHistoryDefaultQty As Double, _
    ByVal manualBaseQty As Double, _
    ByVal seasonalityOn As String, _
    ByVal promoMultiplier As Double, _
    ByVal promoAddOn As Double, _
    ByVal roundRule As String)

    SetCellByHeader targetRow, lo, "Rule_ID", ruleID
    SetCellByHeader targetRow, lo, "SKU", sku
    SetCellByHeader targetRow, lo, "Product_Name", productName
    SetCellByHeader targetRow, lo, "Base_Method", baseMethod
    SetCellByHeader targetRow, lo, "Historical_Window_Days", histDays
    SetCellByHeader targetRow, lo, "Trend_On", trendOn
    SetCellByHeader targetRow, lo, "Trend_Sensitivity", trendSensitivity
    SetCellByHeader targetRow, lo, "NoHistory_Method", noHistoryMethod
    SetCellByHeader targetRow, lo, "NoHistory_Default_Qty", noHistoryDefaultQty
    SetCellByHeader targetRow, lo, "Manual_Base_Qty", manualBaseQty

    If GetHeaderColumn(lo, "Seasonality_On") > 0 Then
        SetCellByHeader targetRow, lo, "Seasonality_On", seasonalityOn
    End If

    SetCellByHeader targetRow, lo, "Promo_Multiplier", promoMultiplier
    SetCellByHeader targetRow, lo, "Promo_AddOn", promoAddOn
    SetCellByHeader targetRow, lo, "Round_Rule", roundRule
    SetCellByHeader targetRow, lo, "Updated_At", Now
End Sub

Public Sub ForecastRule_Load()
    Dim wsUI As Worksheet
    Dim applyTo As String
    Dim sku As String
    Dim targetRow As Range
    Dim lo As ListObject

    Set wsUI = GetWorksheet(WS_FORECAST_UI)
    If wsUI Is Nothing Then Exit Sub

    ForecastRule_SetupUIValidation

    applyTo = UCase$(Trim$(CStr(wsUI.Range(CELL_APPLY_TO).value)))

    Select Case applyTo
        Case APPLY_TO_SELECTED
            sku = GetFirstSelectedSKU()

        Case APPLY_TO_ALL
            MsgBox "Load Rule is only available in SELECTED_SKUS mode." & vbCrLf & _
                   "Please switch Apply To to SELECTED_SKUS.", vbInformation
            Exit Sub

        Case Else
            MsgBox "Please select a valid Apply To option.", vbExclamation
            Exit Sub
    End Select

    If Trim$(sku) = "" Then
        MsgBox "No selected SKU available to load rule.", vbExclamation
        Exit Sub
    End If

    Set lo = GetTable(WS_FORECAST_RULES_DB, TBL_FORECAST_RULES)
    If lo Is Nothing Then Exit Sub

    Set targetRow = GetForecastRuleRowBySKU(sku)
    If targetRow Is Nothing Then
        MsgBox "No saved rule found for SKU: " & sku, vbInformation
        Exit Sub
    End If

    LoadRuleRowToUI targetRow, lo
    RefreshForecastRuleInfoBySKU sku

    MsgBox "Forecast rule loaded for SKU: " & sku, vbInformation
End Sub

Private Sub LoadRuleRowToUI(ByVal srcRow As Range, ByVal lo As ListObject)
    Dim wsUI As Worksheet

    Set wsUI = GetWorksheet(WS_FORECAST_UI)
    If wsUI Is Nothing Then Exit Sub

    wsUI.Range(CELL_BASE_METHOD).value = GetCellByHeader(srcRow, lo, "Base_Method")
    wsUI.Range(CELL_HIST_DAYS).value = GetCellByHeader(srcRow, lo, "Historical_Window_Days")
    wsUI.Range(CELL_TREND_ON).value = GetCellByHeader(srcRow, lo, "Trend_On")
    wsUI.Range(CELL_TREND_SENSITIVITY).value = GetCellByHeader(srcRow, lo, "Trend_Sensitivity")
    wsUI.Range(CELL_NOHISTORY_METHOD).value = GetCellByHeader(srcRow, lo, "NoHistory_Method")
    wsUI.Range(CELL_NOHISTORY_DEFAULT_QTY).value = GetCellByHeader(srcRow, lo, "NoHistory_Default_Qty")
    wsUI.Range(CELL_MANUAL_BASE_QTY).value = GetCellByHeader(srcRow, lo, "Manual_Base_Qty")
    wsUI.Range(CELL_PROMO_MULTIPLIER).value = GetCellByHeader(srcRow, lo, "Promo_Multiplier")
    wsUI.Range(CELL_PROMO_ADDON).value = GetCellByHeader(srcRow, lo, "Promo_AddOn")

    If GetHeaderColumn(lo, "Seasonality_On") > 0 Then
        wsUI.Range(CELL_SEASONALITY_ON).value = NzStr(GetCellByHeader(srcRow, lo, "Seasonality_On"), DEFAULT_SEASONALITY_ON)
    Else
        wsUI.Range(CELL_SEASONALITY_ON).value = DEFAULT_SEASONALITY_ON
    End If

    wsUI.Range(CELL_ROUND_RULE).value = GetCellByHeader(srcRow, lo, "Round_Rule")

    wsUI.Range(CELL_INFO_RULE_ID).value = GetCellByHeader(srcRow, lo, "Rule_ID")
    wsUI.Range(CELL_INFO_LAST_UPDATED).value = GetCellByHeader(srcRow, lo, "Updated_At")
End Sub

Public Sub ForecastRule_ApplyToScope()
    ForecastRule_Save
End Sub

Public Sub ForecastRule_ResetInputs()
    Dim wsUI As Worksheet

    Set wsUI = GetWorksheet(WS_FORECAST_UI)
    If wsUI Is Nothing Then Exit Sub

    ForecastRule_SetupUIValidation

    Application.EnableEvents = False
    On Error GoTo SafeExit

    wsUI.Range(CELL_BASE_METHOD).value = BASE_METHOD_WEIGHTED_AVG
    wsUI.Range(CELL_HIST_DAYS).value = DEFAULT_HIST_DAYS
    wsUI.Range(CELL_TREND_ON).value = YES_TEXT
    wsUI.Range(CELL_TREND_SENSITIVITY).value = TREND_NORMAL
    wsUI.Range(CELL_NOHISTORY_METHOD).value = NOHISTORY_REORDER_BASED
    wsUI.Range(CELL_NOHISTORY_DEFAULT_QTY).value = DEFAULT_NOHISTORY_QTY
    wsUI.Range(CELL_MANUAL_BASE_QTY).value = ""
    wsUI.Range(CELL_PROMO_MULTIPLIER).value = DEFAULT_PROMO_MULTIPLIER
    wsUI.Range(CELL_PROMO_ADDON).value = DEFAULT_PROMO_ADDON
    wsUI.Range(CELL_SEASONALITY_ON).value = DEFAULT_SEASONALITY_ON
    wsUI.Range(CELL_ROUND_RULE).value = ROUND_RULE_ROUND

    wsUI.Range(CELL_INFO_RULE_ID).value = ""
    wsUI.Range(CELL_INFO_LAST_UPDATED).value = ""

SafeExit:
    Application.EnableEvents = True
End Sub

Public Function NextForecastRuleID() As String
    Dim lo As ListObject
    Dim lr As ListRow
    Dim maxNum As Long
    Dim txt As String
    Dim n As Long
    Dim ruleIDCol As Long

    Set lo = GetTable(WS_FORECAST_RULES_DB, TBL_FORECAST_RULES)
    If lo Is Nothing Then
        NextForecastRuleID = "FR-00001"
        Exit Function
    End If

    ruleIDCol = GetHeaderColumn(lo, "Rule_ID")
    If ruleIDCol = 0 Then
        NextForecastRuleID = "FR-00001"
        Exit Function
    End If

    maxNum = 0

    For Each lr In lo.ListRows
        txt = NzStr(lr.Range.Cells(1, ruleIDCol).value, "")
        If UCase$(Left$(txt, 3)) = "FR-" Then
            n = val(Mid$(txt, 4))
            If n > maxNum Then maxNum = n
        End If
    Next lr

    NextForecastRuleID = "FR-" & Format$(maxNum + 1, "00000")
End Function

Public Function GetForecastRuleRowBySKU(ByVal sku As String) As Range
    Dim lo As ListObject
    Dim lr As ListRow
    Dim skuCol As Long

    Set lo = GetTable(WS_FORECAST_RULES_DB, TBL_FORECAST_RULES)
    If lo Is Nothing Then Exit Function

    skuCol = GetHeaderColumn(lo, "SKU")
    If skuCol = 0 Then Exit Function

    For Each lr In lo.ListRows
        If StrComp(Trim$(CStr(lr.Range.Cells(1, skuCol).value)), Trim$(sku), vbTextCompare) = 0 Then
            Set GetForecastRuleRowBySKU = lr.Range
            Exit Function
        End If
    Next lr
End Function

Public Function EnsureForecastRuleForSKU(ByVal sku As String) As String
    Dim targetRow As Range
    Dim lo As ListObject
    Dim ruleID As String
    Dim productName As String
    Dim newRow As ListRow

    Set lo = GetTable(WS_FORECAST_RULES_DB, TBL_FORECAST_RULES)
    If lo Is Nothing Then Exit Function

    Set targetRow = GetForecastRuleRowBySKU(sku)

    If targetRow Is Nothing Then
        ruleID = NextForecastRuleID()
        productName = GetProductNameBySKU(sku)

        Set newRow = lo.ListRows.Add

        WriteRuleRow newRow.Range, lo, ruleID, sku, productName, _
                     BASE_METHOD_WEIGHTED_AVG, DEFAULT_HIST_DAYS, YES_TEXT, TREND_NORMAL, _
                     NOHISTORY_REORDER_BASED, DEFAULT_NOHISTORY_QTY, 0, DEFAULT_SEASONALITY_ON, _
                     DEFAULT_PROMO_MULTIPLIER, DEFAULT_PROMO_ADDON, ROUND_RULE_ROUND

        EnsureForecastRuleForSKU = ruleID
    Else
        EnsureForecastRuleForSKU = NzStr(GetCellByHeader(targetRow, lo, "Rule_ID"), "")
        If EnsureForecastRuleForSKU = "" Then
            EnsureForecastRuleForSKU = NextForecastRuleID()
            SetCellByHeader targetRow, lo, "Rule_ID", EnsureForecastRuleForSKU
        End If
    End If
End Function

Private Function GetFirstSelectedSKU() As String
    Dim col As Collection
    Set col = GetSelectedSKUListFromUI()

    If col.Count > 0 Then
        GetFirstSelectedSKU = CStr(col(1))
    Else
        GetFirstSelectedSKU = ""
    End If
End Function

Private Function GetCellByHeader(ByVal rowRange As Range, ByVal lo As ListObject, ByVal headerName As String) As Variant
    Dim colIndex As Long
    colIndex = GetHeaderColumn(lo, headerName)
    If colIndex = 0 Then Exit Function
    GetCellByHeader = rowRange.Cells(1, colIndex).value
End Function

Private Sub SetCellByHeader(ByVal rowRange As Range, ByVal lo As ListObject, ByVal headerName As String, ByVal newValue As Variant)
    Dim colIndex As Long
    colIndex = GetHeaderColumn(lo, headerName)
    If colIndex = 0 Then Exit Sub
    rowRange.Cells(1, colIndex).value = newValue
End Sub

