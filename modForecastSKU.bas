Attribute VB_Name = "modForecastSKU"
Option Explicit

' =========================================================
' Forecast SKU List Actions
' =========================================================

Public Sub Forecast_AddSKU()
    Dim ws As Worksheet
    Dim sku As String
    Dim rng As Range
    Dim cell As Range
    Dim added As Boolean

    Set ws = GetWorksheet(WS_FORECAST_UI)
    If ws Is Nothing Then Exit Sub

    ' 只允许在 SELECTED_SKUS 模式下添加
    If UCase$(Trim$(CStr(ws.Range(CELL_APPLY_TO).value))) <> APPLY_TO_SELECTED Then
        MsgBox "Add SKU is only available when Apply To = SELECTED_SKUS.", vbExclamation
        Exit Sub
    End If

    sku = Trim$(CStr(ws.Range(CELL_SKU_SELECTOR).value))
    If sku = "" Then
        MsgBox "Please select a SKU first.", vbExclamation
        Exit Sub
    End If

    ' 检查 SKU 是否存在于 Products_DB
    If GetProductNameBySKU(sku) = "" Then
        MsgBox "SKU not found in Products_DB: " & sku, vbExclamation
        Exit Sub
    End If

    Set rng = ws.Range(RNG_SELECTED_SKUS)

    ' 去重检查
    For Each cell In rng.Cells
        If StrComp(Trim$(CStr(cell.value)), sku, vbTextCompare) = 0 Then
            MsgBox "SKU already exists in selected list.", vbInformation
            Exit Sub
        End If
    Next cell

    ' 找第一个空位
    added = False
    For Each cell In rng.Cells
        If Trim$(CStr(cell.value)) = "" Then
            cell.value = sku
            added = True
            Exit For
        End If
    Next cell

    If Not added Then
        MsgBox "Selected SKU list is full. Maximum is " & MAX_SELECTED_SKUS & " SKU(s).", vbExclamation
        Exit Sub
    End If

    ' 刷新右侧产品信息（用当前 selector）
    RefreshForecastRuleInfoBySKU sku
End Sub

Public Sub Forecast_ClearSKUList()
    Dim ws As Worksheet

    Set ws = GetWorksheet(WS_FORECAST_UI)
    If ws Is Nothing Then Exit Sub

    ws.Range(RNG_SELECTED_SKUS).ClearContents
End Sub
