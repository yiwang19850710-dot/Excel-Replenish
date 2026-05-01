Attribute VB_Name = "modPurchasePrint"
Option Explicit

' ==========================================
' Purchase Order Print Module - Brand Style+
' ==========================================

Private Const WS_PURCHASE_UI As String = "Purchase_UI"
Private Const WS_PURCHASE_DB As String = "Purchase_DB"
Private Const WS_PO_PRINT As String = "PO_Print"
Private Const WS_SETUP_UI As String = "Setup_UI"
Private Const WS_SUPPLIERS_DB As String = "Suppliers_DB"

Private Const TBL_PURCHASE As String = "tblPurchase"
Private Const TBL_SUPPLIERS As String = "tblSuppliers"

' Purchase_UI current PO cell
Private Const CELL_UI_PO_NO As String = "B3"

Public Sub PurchasePreview_PO()
    If BuildPOPrintPage() Then
        ThisWorkbook.Worksheets(WS_PO_PRINT).PrintPreview
    End If
End Sub

Public Sub PurchasePrint_PO()
    If BuildPOPrintPage() Then
        ThisWorkbook.Worksheets(WS_PO_PRINT).PrintOut
    End If
End Sub

Public Function BuildPOPrintPage() As Boolean

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim wsP As Worksheet
    Dim lo As ListObject

    Dim poNo As String
    Dim poDate As Variant
    Dim supplierName As String
    Dim notesText As String
    Dim receiveStatus As String
    Dim paymentStatus As String
    Dim amountPaid As Double
    Dim balanceDue As Double
    Dim subtotal As Double

    Dim companyName As String
    Dim companyAddress As String
    Dim companyEmail As String
    Dim companyPhone As String

    Dim suppContact As String
    Dim suppEmail As String
    Dim suppPhone As String
    Dim suppAddress As String
    Dim suppCountry As String

    Dim colPONo As Long
    Dim colPODate As Long
    Dim colSupplierName As Long
    Dim colSKU As Long
    Dim colProduct As Long
    Dim colQty As Long
    Dim colUnitCost As Long
    Dim colLineTotal As Long
    Dim colAmountPaid As Long
    Dim colBalanceDue As Long
    Dim colPaymentStatus As Long
    Dim colReceiveStatus As Long
    Dim colNotes As Long

    Dim i As Long
    Dim outRow As Long
    Dim matchedCount As Long

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(WS_PURCHASE_UI)
    Set wsDB = ThisWorkbook.Worksheets(WS_PURCHASE_DB)
    Set wsP = ThisWorkbook.Worksheets(WS_PO_PRINT)

    On Error Resume Next
    Set lo = wsDB.ListObjects(TBL_PURCHASE)
    On Error GoTo ErrHandler

    If lo Is Nothing Then
        MsgBox "Purchase table '" & TBL_PURCHASE & "' not found in Purchase_DB.", vbCritical
        Exit Function
    End If

    poNo = Trim$(CStr(wsUI.Range(CELL_UI_PO_NO).value))
    If poNo = "" Then
        MsgBox "Please load or enter a Purchase ID first in Purchase_UI.", vbExclamation
        Exit Function
    End If

    If lo.DataBodyRange Is Nothing Then
        MsgBox "Purchase_DB has no data.", vbExclamation
        Exit Function
    End If

    colPONo = GetHeaderColumnByName(lo, "Purchase_Order_No")
    colPODate = GetHeaderColumnByName(lo, "Purchase_Date")
    colSupplierName = GetHeaderColumnByName(lo, "Supplier_Name")
    colSKU = GetHeaderColumnByName(lo, "SKU")
    colProduct = GetHeaderColumnByName(lo, "Product_Name")
    colQty = GetHeaderColumnByName(lo, "Qty")
    colUnitCost = GetHeaderColumnByName(lo, "Unit_Cost")
    colLineTotal = GetHeaderColumnByName(lo, "Line_Total")
    colAmountPaid = GetHeaderColumnByName(lo, "Amount_Paid")
    colBalanceDue = GetHeaderColumnByName(lo, "Balance_Due")
    colPaymentStatus = GetHeaderColumnByName(lo, "Payment_Status")
    colReceiveStatus = GetHeaderColumnByName(lo, "Receive_Status")
    colNotes = GetHeaderColumnByName(lo, "Notes")

    If colPONo = 0 Or colPODate = 0 Or colSupplierName = 0 Or _
       colSKU = 0 Or colProduct = 0 Or colQty = 0 Or _
       colUnitCost = 0 Or colLineTotal = 0 Then
        MsgBox "Purchase_DB is missing one or more required columns.", vbCritical
        Exit Function
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    PreparePOPrintSheet wsP
    InsertCompanyLogoPO wsP

    matchedCount = 0
    subtotal = 0
    amountPaid = 0
    balanceDue = 0
    notesText = ""
    supplierName = ""
    receiveStatus = ""
    paymentStatus = ""

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(i, colPONo).value)), poNo, vbTextCompare) = 0 Then

            If matchedCount = 0 Then
                poDate = lo.DataBodyRange.Cells(i, colPODate).value
                supplierName = Trim$(CStr(lo.DataBodyRange.Cells(i, colSupplierName).value))

                If colAmountPaid > 0 Then amountPaid = NzD(lo.DataBodyRange.Cells(i, colAmountPaid).value)
                If colBalanceDue > 0 Then balanceDue = NzD(lo.DataBodyRange.Cells(i, colBalanceDue).value)
                If colPaymentStatus > 0 Then paymentStatus = Trim$(CStr(lo.DataBodyRange.Cells(i, colPaymentStatus).value))
                If colReceiveStatus > 0 Then receiveStatus = Trim$(CStr(lo.DataBodyRange.Cells(i, colReceiveStatus).value))
            End If

            If colNotes > 0 Then
                If Trim$(notesText) = "" Then
                    notesText = Trim$(CStr(lo.DataBodyRange.Cells(i, colNotes).value))
                End If
            End If

            subtotal = subtotal + NzD(lo.DataBodyRange.Cells(i, colLineTotal).value)
            matchedCount = matchedCount + 1
        End If
    Next i

    If matchedCount = 0 Then
        MsgBox "Purchase Order not found in Purchase_DB: " & poNo, vbExclamation
        GoTo SafeExit
    End If

    ' company info
    companyName = GetNamedText("CompanyName")
    companyAddress = GetNamedText("CompanyAddress")
    companyEmail = GetNamedText("CompanyEmail")
    companyPhone = GetNamedText("CompanyPhone")

    ' supplier info
    GetSupplierDetails supplierName, suppContact, suppEmail, suppPhone, suppAddress, suppCountry

    ' -----------------------------
    ' Header section
    ' -----------------------------
    With wsP.Range("B1:H1")
        .Merge
        .value = "PURCHASE ORDER"
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Company block
    With wsP.Range("A3:D3")
        .Merge
        .value = IIf(companyName <> "", companyName, "Company")
        .Font.Bold = True
    End With
    With wsP.Range("A4:D4")
        .Merge
        .value = companyAddress
    End With
    With wsP.Range("A5:D5")
        .Merge
        .value = BuildInlineContact(companyEmail, companyPhone)
    End With

    ' Supplier block
    With wsP.Range("E3:H3")
        .Merge
        .value = "Supplier: " & supplierName
        .Font.Bold = True
    End With
    With wsP.Range("E4:H4")
        .Merge
        .value = BuildSupplierLine(suppContact, suppEmail, suppPhone)
    End With
    With wsP.Range("E5:H5")
        .Merge
        .value = BuildSupplierAddress(suppAddress, suppCountry)
    End With

    ' PO details row
    wsP.Range("A7").value = "PO No"
    wsP.Range("B7").value = poNo

    wsP.Range("D7").value = "PO Date"
    wsP.Range("E7").value = poDate

    wsP.Range("G7").value = "Print Time"
    wsP.Range("H7").value = Now

    wsP.Range("A8").value = "Receive Status"
    wsP.Range("B8").value = receiveStatus

    wsP.Range("D8").value = "Payment Status"
    wsP.Range("E8").value = paymentStatus

    wsP.Range("A7:H8").Font.Bold = True
    wsP.Range("E7").NumberFormat = "yyyy-mm-dd"
    wsP.Range("H7").NumberFormat = "yyyy-mm-dd hh:mm"

    ' Notes
    wsP.Range("A10").value = "Notes"
    wsP.Range("B10:H10").Merge
    wsP.Range("B10").value = notesText
    wsP.Range("A10:H10").Borders.LineStyle = xlContinuous
    wsP.Range("B10").WrapText = True

    ' -----------------------------
    ' Line item header
    ' -----------------------------
    wsP.Range("A12").value = "SKU"

    wsP.Range("B12:D12").Merge
    wsP.Range("B12").value = "Product Name"

    wsP.Range("E12").value = "Qty"
    wsP.Range("F12").value = "Unit Cost"

    wsP.Range("G12:H12").Merge
    wsP.Range("G12").value = "Line Total"

    With wsP.Range("A12:H12")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' -----------------------------
    ' Line items
    ' -----------------------------
    outRow = 13

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(i, colPONo).value)), poNo, vbTextCompare) = 0 Then

            wsP.Range("A" & outRow).value = lo.DataBodyRange.Cells(i, colSKU).value

            wsP.Range("B" & outRow & ":D" & outRow).Merge
            wsP.Range("B" & outRow).value = lo.DataBodyRange.Cells(i, colProduct).value

            wsP.Range("E" & outRow).value = NzD(lo.DataBodyRange.Cells(i, colQty).value)
            wsP.Range("F" & outRow).value = NzD(lo.DataBodyRange.Cells(i, colUnitCost).value)

            wsP.Range("G" & outRow & ":H" & outRow).Merge
            wsP.Range("G" & outRow).value = NzD(lo.DataBodyRange.Cells(i, colLineTotal).value)

            With wsP.Range("A" & outRow & ":H" & outRow)
                .Borders.LineStyle = xlContinuous
                .VerticalAlignment = xlCenter
            End With

            outRow = outRow + 1
        End If
    Next i

    ' -----------------------------
    ' Totals section
    ' -----------------------------
    wsP.Range("F" & outRow + 1).value = "Subtotal"
    wsP.Range("G" & outRow + 1 & ":H" & outRow + 1).Merge
    wsP.Range("G" & outRow + 1).value = subtotal

    wsP.Range("F" & outRow + 2).value = "Amount Paid"
    wsP.Range("G" & outRow + 2 & ":H" & outRow + 2).Merge
    wsP.Range("G" & outRow + 2).value = amountPaid

    wsP.Range("F" & outRow + 3).value = "Balance Due"
    wsP.Range("G" & outRow + 3 & ":H" & outRow + 3).Merge
    wsP.Range("G" & outRow + 3).value = balanceDue

    With wsP.Range("F" & outRow + 1 & ":H" & outRow + 3)
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
    End With

    ' -----------------------------
    ' Formatting
    ' -----------------------------
    wsP.Columns("A").ColumnWidth = 16
    wsP.Columns("B").ColumnWidth = 18
    wsP.Columns("C").ColumnWidth = 10
    wsP.Columns("D").ColumnWidth = 10
    wsP.Columns("E").ColumnWidth = 10
    wsP.Columns("F").ColumnWidth = 12
    wsP.Columns("G").ColumnWidth = 10
    wsP.Columns("H").ColumnWidth = 14

    wsP.Range("E13:H" & outRow + 3).NumberFormat = "#,##0.00"
    wsP.Range("H7").EntireColumn.ColumnWidth = 16

    wsP.Range("A3:H10").WrapText = True
    wsP.Rows("1:10").RowHeight = 22
    wsP.Rows("12:" & outRow + 3).RowHeight = 20

    ' -----------------------------
    ' Page setup
    ' -----------------------------
    With wsP.PageSetup
        .Orientation = xlPortrait
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$12:$12"
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True
    End With

    BuildPOPrintPage = True

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Function

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "PO Print"
    Resume SafeExit

End Function

Private Sub PreparePOPrintSheet(ByVal ws As Worksheet)

    ws.Cells.Clear
    ws.Cells.Font.name = "Calibri"
    ws.Cells.Font.Size = 10

    On Error Resume Next
    ws.DrawingObjects.Delete
    On Error GoTo 0

    ws.Rows.RowHeight = 20
End Sub

Private Sub InsertCompanyLogoPO(ByVal wsTarget As Worksheet)

    Dim wsSetup As Worksheet
    Dim shp As Shape
    Dim newShp As Shape
    Dim s As Shape

    On Error Resume Next
    Set wsSetup = ThisWorkbook.Worksheets(WS_SETUP_UI)
    Set shp = wsSetup.Shapes("CompanyLogo")
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    For Each s In wsTarget.Shapes
        If s.name Like "PrintLogo*" Then
            s.Delete
        End If
    Next s

    shp.Copy
    wsTarget.Paste

    Set newShp = wsTarget.Shapes(wsTarget.Shapes.Count)
    newShp.name = "PrintLogo"

    With newShp
        .Top = wsTarget.Range("A1").Top
        .Left = wsTarget.Range("A1").Left
        .LockAspectRatio = msoTrue
        .Height = 40
    End With
End Sub

Private Sub GetSupplierDetails(ByVal supplierName As String, _
                               ByRef contactPerson As String, _
                               ByRef emailAddr As String, _
                               ByRef phoneNo As String, _
                               ByRef addressText As String, _
                               ByRef countryText As String)

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim i As Long

    Dim colName As Long
    Dim colContact As Long
    Dim colEmail As Long
    Dim colPhone As Long
    Dim colAddress As Long
    Dim colCountry As Long

    contactPerson = ""
    emailAddr = ""
    phoneNo = ""
    addressText = ""
    countryText = ""

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_SUPPLIERS_DB)
    Set lo = ws.ListObjects(TBL_SUPPLIERS)
    On Error GoTo 0

    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    colName = GetHeaderColumnByName(lo, "Supplier_Name")
    colContact = GetHeaderColumnByName(lo, "Contact_Person")
    colEmail = GetHeaderColumnByName(lo, "Email")
    colPhone = GetHeaderColumnByName(lo, "Phone")
    colAddress = GetHeaderColumnByName(lo, "Address")
    colCountry = GetHeaderColumnByName(lo, "Country")

    If colName = 0 Then Exit Sub

    For i = 1 To lo.ListRows.Count
        If StrComp(Trim$(CStr(lo.DataBodyRange.Cells(i, colName).value)), supplierName, vbTextCompare) = 0 Then
            If colContact > 0 Then contactPerson = Trim$(CStr(lo.DataBodyRange.Cells(i, colContact).value))
            If colEmail > 0 Then emailAddr = Trim$(CStr(lo.DataBodyRange.Cells(i, colEmail).value))
            If colPhone > 0 Then phoneNo = Trim$(CStr(lo.DataBodyRange.Cells(i, colPhone).value))
            If colAddress > 0 Then addressText = Trim$(CStr(lo.DataBodyRange.Cells(i, colAddress).value))
            If colCountry > 0 Then countryText = Trim$(CStr(lo.DataBodyRange.Cells(i, colCountry).value))
            Exit For
        End If
    Next i
End Sub

Private Function GetNamedText(ByVal nm As String) As String
    On Error Resume Next
    GetNamedText = Trim$(CStr(ThisWorkbook.Names(nm).RefersToRange.value))
    On Error GoTo 0
End Function

Private Function BuildInlineContact(ByVal emailAddr As String, ByVal phoneNo As String) As String
    If emailAddr <> "" And phoneNo <> "" Then
        BuildInlineContact = emailAddr & " | " & phoneNo
    ElseIf emailAddr <> "" Then
        BuildInlineContact = emailAddr
    Else
        BuildInlineContact = phoneNo
    End If
End Function

Private Function BuildSupplierLine(ByVal contactPerson As String, ByVal emailAddr As String, ByVal phoneNo As String) As String
    Dim s As String

    s = ""
    If contactPerson <> "" Then s = contactPerson
    If emailAddr <> "" Then
        If s <> "" Then s = s & " | "
        s = s & emailAddr
    End If
    If phoneNo <> "" Then
        If s <> "" Then s = s & " | "
        s = s & phoneNo
    End If

    BuildSupplierLine = s
End Function

Private Function BuildSupplierAddress(ByVal addressText As String, ByVal countryText As String) As String
    If addressText <> "" And countryText <> "" Then
        BuildSupplierAddress = addressText & ", " & countryText
    ElseIf addressText <> "" Then
        BuildSupplierAddress = addressText
    Else
        BuildSupplierAddress = countryText
    End If
End Function

Private Function GetHeaderColumnByName(ByVal lo As ListObject, ByVal headerName As String) As Long
    Dim i As Long

    GetHeaderColumnByName = 0
    If lo Is Nothing Then Exit Function

    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(i).name), Trim$(headerName), vbTextCompare) = 0 Then
            GetHeaderColumnByName = i
            Exit Function
        End If
    Next i
End Function

Private Function NzD(ByVal v As Variant) As Double
    If IsError(v) Then
        NzD = 0
    ElseIf IsNumeric(v) Then
        NzD = CDbl(v)
    Else
        NzD = 0
    End If
End Function

