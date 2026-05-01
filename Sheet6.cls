VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim changedCell As Range
    Dim r As Long

    On Error GoTo SafeExit
    If Application.EnableEvents = False Then Exit Sub

    Application.EnableEvents = False

    ' 1) SKU changed -> auto fill Product Name + Unit Price + Line Total formula
    If Not Intersect(Target, Me.Range("A9:A18")) Is Nothing Then
        For Each changedCell In Intersect(Target, Me.Range("A9:A18"))
            r = changedCell.Row

            Me.Range("B" & r).Formula = "=IFERROR(XLOOKUP(A" & r & ",tblProducts[SKU],tblProducts[Product_Name]),"""")"
            Me.Range("D" & r).Formula = "=IFERROR(XLOOKUP(A" & r & ",tblProducts[SKU],tblProducts[Unit_Cost]),"""")"
            Me.Range("E" & r).Formula = "=IFERROR(C" & r & "*D" & r & ",0)"
        Next changedCell
    End If

    ' 2) Qty or Unit Price changed -> refresh Line Total
    If Not Intersect(Target, Me.Range("C9:D18")) Is Nothing Then
        For Each changedCell In Intersect(Target, Me.Range("C9:D18"))
            r = changedCell.Row
            Me.Range("E" & r).Formula = "=IFERROR(C" & r & "*D" & r & ",0)"
        Next changedCell
    End If

SafeExit:
    Application.EnableEvents = True

End Sub

Private Sub Worksheet_Activate()

    Dim r As Long

    On Error Resume Next
    Application.EnableEvents = False

    For r = 9 To 18
        If Trim$(CStr(Me.Range("A" & r).value)) <> "" Then
            If Trim$(CStr(Me.Range("B" & r).value)) = "" Then
                Me.Range("B" & r).Formula = "=IFERROR(XLOOKUP(A" & r & ",tblProducts[SKU],tblProducts[Product_Name]),"""")"
            End If

            If Trim$(CStr(Me.Range("D" & r).value)) = "" Or Me.Range("D" & r).value = 0 Then
                Me.Range("D" & r).Formula = "=IFERROR(XLOOKUP(A" & r & ",tblProducts[SKU],tblProducts[Unit_Cost]),"""")"
            End If
        End If

        Me.Range("E" & r).Formula = "=IFERROR(C" & r & "*D" & r & ",0)"
    Next r

    Application.EnableEvents = True

End Sub
