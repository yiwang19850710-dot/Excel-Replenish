VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet13"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()

    On Error Resume Next
    SetupStockHealthSKUValidation
    RefreshTopInfoFromFirstSKUInList
    On Error GoTo 0

End Sub

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo SafeExit
    If Application.EnableEvents = False Then Exit Sub

    Application.EnableEvents = False

    ' B3 selector changed -> preview info only
    If Not Intersect(Target, Me.Range("B3")) Is Nothing Then
        StockHealth_FillProductInfo
    End If

    ' editable forecast/supply area for up to 5 SKU blocks
    If Not Intersect(Target, Me.Range("B17:CM60")) Is Nothing Then
        Application.ScreenUpdating = False
        Me.Calculate
        ReapplyAllStockHealthDSIColors
        Application.ScreenUpdating = True
    End If

SafeExit:
    Application.EnableEvents = True

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    On Error Resume Next
    If Target.CountLarge > 1 Then Exit Sub

    If Not Intersect(Target, Me.Range("B4:B8")) Is Nothing Then
        If Trim$(CStr(Target.value)) <> "" Then
            RefreshTopInfoBySKU Trim$(CStr(Target.value))
        End If
    End If

    On Error GoTo 0

End Sub
