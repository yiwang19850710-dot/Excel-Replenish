VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet34"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    If Intersect(Target, Me.Range("B3")) Is Nothing Then Exit Sub

    On Error GoTo CleanExit
    Application.EnableEvents = False

    Me.Range("B4").value = GetProductNameBySKU_POInquiry(Trim$(CStr(Me.Range("B3").value)))

CleanExit:
    Application.EnableEvents = True

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    If Target.Row < 12 Then Exit Sub
    If Target.Column < 1 Or Target.Column > 18 Then Exit Sub

    On Error GoTo CleanExit
    Application.EnableEvents = False

    Me.Range(Me.Cells(12, 1), Me.Cells(Me.Rows.Count, 18)).Interior.Pattern = xlNone
    Me.Range(Me.Cells(Target.Row, 1), Me.Cells(Target.Row, 18)).Interior.Color = RGB(221, 235, 247)

CleanExit:
    Application.EnableEvents = True

End Sub

Private Function GetProductNameBySKU_POInquiry(ByVal sku As String) As String

    Dim tbl As ListObject
    Dim i As Long
    Dim colSKU As Long
    Dim colName As Long

    If sku = "" Then
        GetProductNameBySKU_POInquiry = ""
        Exit Function
    End If

    Set tbl = ThisWorkbook.Worksheets("Products_DB").ListObjects("tblProducts")

    colSKU = tbl.ListColumns("SKU").Index
    colName = tbl.ListColumns("Product_Name").Index

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If StrComp(Trim$(CStr(tbl.DataBodyRange.Cells(i, colSKU).value)), sku, vbTextCompare) = 0 Then
            GetProductNameBySKU_POInquiry = tbl.DataBodyRange.Cells(i, colName).value
            Exit Function
        End If
    Next i

    GetProductNameBySKU_POInquiry = ""

End Function
