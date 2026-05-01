VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    Const COL_LATEST_PO As Long = 8   ' H
    Const FIRST_DATA_ROW As Long = 11

    Dim clickedCell As Range
    Dim poNo As String
    Dim wsPurchaseUI As Worksheet

    On Error GoTo SafeExit

    Set clickedCell = Target.Range

    If clickedCell.Row < FIRST_DATA_ROW Then Exit Sub
    If clickedCell.Column <> COL_LATEST_PO Then Exit Sub

    poNo = Trim$(CStr(clickedCell.value))
    If poNo = "" Then Exit Sub

    Set wsPurchaseUI = ThisWorkbook.Worksheets("Purchase_UI")

    Application.EnableEvents = False
    wsPurchaseUI.Range("B3").value = poNo
    Purchase_Load
    wsPurchaseUI.Activate

SafeExit:
    Application.EnableEvents = True

End Sub
