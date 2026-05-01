VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Worksheet_Activate()
    On Error Resume Next
    Application.EnableEvents = True
    InitForecastUIDefaults
    ForecastRule_SetupUIValidation
    ApplyScopeGrayOut
    On Error GoTo 0
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Application.EnableEvents = False Then Exit Sub

    Application.EnableEvents = False

    If Not Intersect(Target, Me.Range(CELL_APPLY_TO)) Is Nothing Then
        Dim applyTo As String
        applyTo = UCase$(Trim$(CStr(Me.Range(CELL_APPLY_TO).value)))

        Select Case applyTo
            Case APPLY_TO_SELECTED, APPLY_TO_ALL
                If applyTo = APPLY_TO_ALL Then
                    Me.Range(RNG_SELECTED_SKUS).ClearContents
                End If
                ApplyScopeGrayOut

            Case Else
                MsgBox "Apply To must be SELECTED_SKUS or ALL_PRODUCTS.", vbExclamation
                Me.Range(CELL_APPLY_TO).value = APPLY_TO_SELECTED
                ApplyScopeGrayOut
        End Select
    End If

    If Not Intersect(Target, Me.Range(CELL_SKU_SELECTOR)) Is Nothing Then
        Dim sku As String
        sku = Trim$(CStr(Me.Range(CELL_SKU_SELECTOR).value))

        If sku <> "" Then
            RefreshForecastRuleInfoBySKU sku
        Else
            RefreshForecastRuleInfoBySKU ""
        End If
    End If

    If Not Intersect(Target, Me.Range(CELL_VIEW_START_DATE)) Is Nothing Then
        If Trim$(CStr(Me.Range(CELL_VIEW_START_DATE).value)) <> "" Then
            If Not IsDate(Me.Range(CELL_VIEW_START_DATE).value) Then
                MsgBox "Please enter a valid View Start Date.", vbExclamation
                Me.Range(CELL_VIEW_START_DATE).ClearContents
            End If
        End If
    End If

    If Not Intersect(Target, Me.Range(CELL_VIEW_DAYS)) Is Nothing Then
        Dim viewDays As Long
        viewDays = CLng(NzDbl(Me.Range(CELL_VIEW_DAYS).value, 0))

        Select Case viewDays
            Case 30, 60, 90
            Case Else
                MsgBox "View Days must be 30, 60, or 90.", vbExclamation
                Me.Range(CELL_VIEW_DAYS).value = DEFAULT_VIEW_DAYS
        End Select
    End If

    If Not Intersect(Target, Me.Range(CELL_VIEW_MODE)) Is Nothing Then
        Dim vm As String
        vm = UCase$(Trim$(CStr(Me.Range(CELL_VIEW_MODE).value)))

        Select Case vm
            Case VIEW_MODE_FINAL, VIEW_MODE_SYSTEM, VIEW_MODE_OVERRIDE
                Me.Range(CELL_VIEW_MODE).value = vm
            Case Else
                MsgBox "View Mode must be FINAL, SYSTEM, or OVERRIDE.", vbExclamation
                Me.Range(CELL_VIEW_MODE).value = VIEW_MODE_FINAL
        End Select
    End If

    If Not Intersect(Target, Me.Range(CELL_GEN_START_DATE)) Is Nothing Then
        If Trim$(CStr(Me.Range(CELL_GEN_START_DATE).value)) <> "" Then
            If Not IsDate(Me.Range(CELL_GEN_START_DATE).value) Then
                MsgBox "Please enter a valid Generate Start Date.", vbExclamation
                Me.Range(CELL_GEN_START_DATE).ClearContents
            End If
        End If
    End If

    If Not Intersect(Target, Me.Range(CELL_GEN_HORIZON_DAYS)) Is Nothing Then
        Dim horizonVal As Long
        horizonVal = CLng(NzDbl(Me.Range(CELL_GEN_HORIZON_DAYS).value, 0))

        If horizonVal <= 0 Then
            MsgBox "Horizon Days must be greater than 0.", vbExclamation
            Me.Range(CELL_GEN_HORIZON_DAYS).value = DEFAULT_HORIZON_DAYS
        Else
            Me.Range(CELL_GEN_HORIZON_DAYS).value = horizonVal
        End If
    End If

    If Not Intersect(Target, Me.Range(CELL_GEN_REPLACE_EXISTING)) Is Nothing Then
        Dim reVal As String
        reVal = UCase$(Trim$(CStr(Me.Range(CELL_GEN_REPLACE_EXISTING).value)))

        Select Case reVal
            Case YES_TEXT, NO_TEXT
                Me.Range(CELL_GEN_REPLACE_EXISTING).value = reVal
            Case Else
                MsgBox "Replace Existing must be YES or NO.", vbExclamation
                Me.Range(CELL_GEN_REPLACE_EXISTING).value = YES_TEXT
        End Select
    End If

    If Not Intersect(Target, Me.Range(CELL_GEN_MISSING_ONLY)) Is Nothing Then
        Dim gmVal As String
        gmVal = UCase$(Trim$(CStr(Me.Range(CELL_GEN_MISSING_ONLY).value)))

        Select Case gmVal
            Case YES_TEXT, NO_TEXT
                Me.Range(CELL_GEN_MISSING_ONLY).value = gmVal
            Case Else
                MsgBox "Generate Missing Only must be YES or NO.", vbExclamation
                Me.Range(CELL_GEN_MISSING_ONLY).value = NO_TEXT
        End Select
    End If

    If Not Intersect(Target, Me.Range(CELL_SEASONALITY_ON)) Is Nothing Then
        Dim sVal As String
        sVal = UCase$(Trim$(CStr(Me.Range(CELL_SEASONALITY_ON).value)))

        Select Case sVal
            Case YES_TEXT, NO_TEXT
                Me.Range(CELL_SEASONALITY_ON).value = sVal
            Case Else
                MsgBox "Seasonality On must be YES or NO.", vbExclamation
                Me.Range(CELL_SEASONALITY_ON).value = DEFAULT_SEASONALITY_ON
        End Select
    End If

SafeExit:
    Application.EnableEvents = True
End Sub

Private Sub InitForecastUIDefaults()
    If Trim$(CStr(Me.Range(CELL_APPLY_TO).value)) = "" Then
        Me.Range(CELL_APPLY_TO).value = APPLY_TO_SELECTED
    End If

    If Trim$(CStr(Me.Range(CELL_BASE_METHOD).value)) = "" Then
        Me.Range(CELL_BASE_METHOD).value = BASE_METHOD_WEIGHTED_AVG
    End If

    If Trim$(CStr(Me.Range(CELL_HIST_DAYS).value)) = "" Then
        Me.Range(CELL_HIST_DAYS).value = DEFAULT_HIST_DAYS
    End If

    If Trim$(CStr(Me.Range(CELL_TREND_ON).value)) = "" Then
        Me.Range(CELL_TREND_ON).value = YES_TEXT
    End If

    If Trim$(CStr(Me.Range(CELL_TREND_SENSITIVITY).value)) = "" Then
        Me.Range(CELL_TREND_SENSITIVITY).value = TREND_NORMAL
    End If

    If Trim$(CStr(Me.Range(CELL_NOHISTORY_METHOD).value)) = "" Then
        Me.Range(CELL_NOHISTORY_METHOD).value = NOHISTORY_REORDER_BASED
    End If

    If Trim$(CStr(Me.Range(CELL_NOHISTORY_DEFAULT_QTY).value)) = "" Then
        Me.Range(CELL_NOHISTORY_DEFAULT_QTY).value = DEFAULT_NOHISTORY_QTY
    End If

    If Trim$(CStr(Me.Range(CELL_PROMO_MULTIPLIER).value)) = "" Then
        Me.Range(CELL_PROMO_MULTIPLIER).value = DEFAULT_PROMO_MULTIPLIER
    End If

    If Trim$(CStr(Me.Range(CELL_PROMO_ADDON).value)) = "" Then
        Me.Range(CELL_PROMO_ADDON).value = DEFAULT_PROMO_ADDON
    End If

    If Trim$(CStr(Me.Range(CELL_SEASONALITY_ON).value)) = "" Then
        Me.Range(CELL_SEASONALITY_ON).value = DEFAULT_SEASONALITY_ON
    End If

    If Trim$(CStr(Me.Range(CELL_ROUND_RULE).value)) = "" Then
        Me.Range(CELL_ROUND_RULE).value = ROUND_RULE_ROUND
    End If

    If Trim$(CStr(Me.Range(CELL_GEN_START_DATE).value)) = "" Then
        Me.Range(CELL_GEN_START_DATE).value = Date
    End If

    If Trim$(CStr(Me.Range(CELL_GEN_HORIZON_DAYS).value)) = "" Then
        Me.Range(CELL_GEN_HORIZON_DAYS).value = DEFAULT_HORIZON_DAYS
    End If

    If Trim$(CStr(Me.Range(CELL_GEN_REPLACE_EXISTING).value)) = "" Then
        Me.Range(CELL_GEN_REPLACE_EXISTING).value = YES_TEXT
    End If

    If Trim$(CStr(Me.Range(CELL_GEN_MISSING_ONLY).value)) = "" Then
        Me.Range(CELL_GEN_MISSING_ONLY).value = NO_TEXT
    End If

    If Trim$(CStr(Me.Range(CELL_VIEW_START_DATE).value)) = "" Then
        Me.Range(CELL_VIEW_START_DATE).value = Date
    End If

    If Trim$(CStr(Me.Range(CELL_VIEW_DAYS).value)) = "" Then
        Me.Range(CELL_VIEW_DAYS).value = DEFAULT_VIEW_DAYS
    End If

    If Trim$(CStr(Me.Range(CELL_VIEW_MODE).value)) = "" Then
        Me.Range(CELL_VIEW_MODE).value = VIEW_MODE_FINAL
    End If
End Sub

Private Sub ApplyScopeGrayOut()
    Dim applyTo As String

    applyTo = UCase$(Trim$(CStr(Me.Range(CELL_APPLY_TO).value)))

    If applyTo = APPLY_TO_ALL Then
        With Me.Range(CELL_SKU_SELECTOR)
            .Interior.Color = RGB(217, 217, 217)
            .Font.Color = RGB(127, 127, 127)
        End With

        With Me.Range(RNG_SELECTED_SKUS)
            .Interior.Color = RGB(217, 217, 217)
            .Font.Color = RGB(127, 127, 127)
        End With
    Else
        With Me.Range(CELL_SKU_SELECTOR)
            .Interior.Color = RGB(255, 255, 204)
            .Font.Color = RGB(0, 0, 0)
        End With

        With Me.Range(RNG_SELECTED_SKUS)
            .Interior.Color = RGB(255, 255, 204)
            .Font.Color = RGB(0, 0, 0)
        End With
    End If
End Sub
