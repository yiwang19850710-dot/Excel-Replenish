VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo SafeExit
    If Application.EnableEvents = False Then Exit Sub

    Application.EnableEvents = False

    ' SKU changed -> restore product name / unit cost formulas
    If Not Intersect(Target, Me.Range("B7")) Is Nothing Then
        Me.Range("B8").Formula = "=IFERROR(XLOOKUP(B7,tblProducts[SKU],tblProducts[Product_Name]),"""")"
        Me.Range("B10").Formula = "=IFERROR(XLOOKUP(B7,tblProducts[SKU],tblProducts[Unit_Cost]),"""")"
        Me.Range("B11").Formula = "=IFERROR(B9*B10,0)"
    End If

    ' Qty or Unit Cost changed -> total cost updates immediately
    If Not Intersect(Target, Me.Range("B9:B10")) Is Nothing Then
        Me.Range("B11").Formula = "=IFERROR(B9*B10,0)"
    End If

SafeExit:
    Application.EnableEvents = True

End Sub

Private Sub Worksheet_Activate()

    On Error Resume Next
    Application.EnableEvents = False
    Me.Range("B11").Formula = "=IFERROR(B9*B10,0)"
    Application.EnableEvents = True

End Sub
