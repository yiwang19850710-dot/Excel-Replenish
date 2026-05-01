VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDuplicateOrderAction 
   Caption         =   "Duplicate Order Detected"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355.001
   OleObjectBlob   =   "frmDuplicateOrderAction.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmDuplicateOrderAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public SelectedAction As String

Public Sub SetOrderInfo(ByVal externalOrderNo As String)

    Me.lblMessage.caption = _
        "Duplicate Lines Found" & vbCrLf & _
        "External No: " & externalOrderNo & vbCrLf & vbCrLf & _
        "Please choose how to handle this duplicate line."

End Sub

Private Sub btnSkipOne_Click()
    SelectedAction = "SKIP_ONE"
    Me.Hide
End Sub

Private Sub btnSkipAll_Click()
    SelectedAction = "SKIP_ALL"
    Me.Hide
End Sub

Private Sub btnReplaceOne_Click()
    SelectedAction = "REPLACE_ONE"
    Me.Hide
End Sub

Private Sub btnReplaceAll_Click()
    SelectedAction = "REPLACE_ALL"
    Me.Hide
End Sub

Private Sub btnStop_Click()
    SelectedAction = "STOP"
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If SelectedAction = "" Then SelectedAction = "STOP"
End Sub
