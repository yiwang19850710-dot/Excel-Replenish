VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSelectInventoryLocation 
   Caption         =   "Select Inventory Location"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7365
   OleObjectBlob   =   "frmSelectInventoryLocation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSelectInventoryLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedColumn As Long
Private mLocationCols As Object

Public Sub LoadLocations(ByVal ws As Worksheet, ByVal locationCols As Object)

    Dim k As Variant

    Set mLocationCols = locationCols
    Me.cboLocation.Clear

    For Each k In locationCols.Keys
        Me.cboLocation.AddItem CStr(k)
    Next k

    If Me.cboLocation.ListCount > 0 Then
        Me.cboLocation.ListIndex = 0
    End If

End Sub

Private Sub btnOK_Click()

    If Me.cboLocation.ListIndex < 0 Then
        MsgBox "Please select one inventory location.", vbExclamation
        Exit Sub
    End If

    SelectedColumn = CLng(mLocationCols(Me.cboLocation.value))
    Me.Hide

End Sub

Private Sub btnCancel_Click()

    SelectedColumn = 0
    Me.Hide

End Sub

Private Sub lblMessage_Click()

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    If SelectedColumn = 0 Then SelectedColumn = 0

End Sub
