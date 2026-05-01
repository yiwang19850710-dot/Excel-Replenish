VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet29"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    On Error Resume Next
    InventoryHistory_SetupUI
    On Error GoTo 0
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Application.EnableEvents = False Then Exit Sub

    Application.EnableEvents = False

    If Not Intersect(Target, Me.Range("B5:B6")) Is Nothing Then
        Dim c As Range
        For Each c In Intersect(Target, Me.Range("B5:B6"))
            If Trim$(CStr(c.value)) <> "" Then
                If Not IsDate(c.value) Then
                    MsgBox "Please enter a valid date.", vbExclamation
                    c.ClearContents
                End If
            End If
        Next c
    End If

    If Not Intersect(Target, Me.Range("B7")) Is Nothing Then
        Dim mv As String
        mv = UCase$(Trim$(CStr(Me.Range("B7").value)))

        Select Case mv
            Case "ALL", "RECEIVING", "SALES", "ADJUSTMENT"
            Case Else
                MsgBox "Movement Type must be ALL, RECEIVING, SALES, or ADJUSTMENT.", vbExclamation
                Me.Range("B7").value = "ALL"
        End Select
    End If

SafeExit:
    Application.EnableEvents = True
End Sub
