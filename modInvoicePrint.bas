Attribute VB_Name = "modInvoicePrint"
Option Explicit

'========================================
' Sheet / Table Names
'========================================
Private Const SHEET_UI As String = "Invoice_UI"
Private Const SHEET_PRINT As String = "Invoice_Print"
Private Const SHEET_SETTINGS As String = "Settings"
Private Const SHEET_SETUP_UI As String = "Setup_UI"
Private Const SHEET_INVOICE_DB As String = "Invoice_DB"
Private Const SHEET_SALES_DB As String = "Sales_DB"
Private Const SHEET_CUSTOMERS_DB As String = "Customers_DB"

Private Const TABLE_SETTINGS As String = "tblSettings"
Private Const TABLE_INVOICES As String = "tblInvoices"
Private Const TABLE_SALES As String = "tblSales"
Private Const TABLE_CUSTOMERS As String = "tblCustomers"

'========================================
' Invoice_UI Mapping
'========================================
Private Const UI_INVOICE_LOOKUP As String = "B3"
Private Const UI_INVOICE_NO As String = "B4"
Private Const UI_SALES_ORDER_NO As String = "B6"

'========================================
' Print Layout Settings
'========================================
Private Const LOGO_ANCHOR As String = "A2"
Private Const DETAIL_HEADER_ROW As Long = 14
Private Const DETAIL_START_ROW As Long = 15

'========================================
' Public Entry Points
'========================================

Public Sub InvoicePrint_Preview()

    Dim invoiceNo As String
    Dim wsPrint As Worksheet
    
    On Error GoTo ErrHandler
    
    invoiceNo = ResolveInvoiceNoFromUI()
    If invoiceNo = "" Then Exit Sub
    
    Set wsPrint = ThisWorkbook.Worksheets(SHEET_PRINT)
    wsPrint.Visible = xlSheetVisible
    
    BuildInvoicePrint invoiceNo
    wsPrint.PrintPreview
    
    wsPrint.Visible = xlSheetHidden
    Exit Sub

ErrHandler:
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_PRINT).Visible = xlSheetHidden
    MsgBox "Error: " & Err.Description, vbCritical

End Sub

Public Sub InvoicePrint_SavePDF()

    Dim wsPrint As Worksheet
    Dim invoiceNo As String
    Dim outputPath As String
    
    On Error GoTo ErrHandler
    
    invoiceNo = ResolveInvoiceNoFromUI()
    If invoiceNo = "" Then Exit Sub
    
    Set wsPrint = ThisWorkbook.Worksheets(SHEET_PRINT)
    wsPrint.Visible = xlSheetVisible
    
    BuildInvoicePrint invoiceNo
    
    If ThisWorkbook.Path = "" Then
        outputPath = Application.DefaultFilePath & "\" & CleanFileName(invoiceNo) & ".pdf"
    Else
        outputPath = ThisWorkbook.Path & "\" & CleanFileName(invoiceNo) & ".pdf"
    End If
    
    wsPrint.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    wsPrint.Visible = xlSheetHidden
    MsgBox "PDF saved to:" & vbCrLf & outputPath, vbInformation
    Exit Sub

ErrHandler:
    On Error Resume Next
    ThisWorkbook.Worksheets(SHEET_PRINT).Visible = xlSheetHidden
    MsgBox "Error: " & Err.Description, vbCritical

End Sub

'========================================
' Main Render
'========================================

Private Sub BuildInvoicePrint(ByVal invoiceNo As String)

    Dim wsPrint As Worksheet
    Dim wsSettings As Worksheet
    Dim wsInvoiceDB As Worksheet
    Dim wsSalesDB As Worksheet
    Dim wsCustomersDB As Worksheet
    
    Dim tblSettings As ListObject
    Dim tblInvoices As ListObject
    Dim tblSales As ListObject
    Dim tblCustomers As ListObject
    
    Dim invoiceRow As Long
    Dim customerRow As Long
    
    Dim invoiceDate As Variant
    Dim salesOrderNo As String
    Dim customerID As String
    Dim customerName As String
    Dim customerAddress As String
    Dim customerPhone As String
    Dim customerEmail As String
    
    Dim companyName As String
    Dim companyAddress As String
    Dim companyPhone As String
    Dim companyEmail As String
    
    Dim subtotal As Double
    Dim tax1Name As String
    Dim tax1Rate As Double
    Dim tax1Amount As Double
    Dim tax2Name As String
    Dim tax2Rate As Double
    Dim tax2Amount As Double
    Dim grandTotal As Double
    Dim amountPaid As Double
    Dim balanceDue As Double
    Dim invoiceStatus As String
    Dim notesText As String
    
    Dim lastDetailRow As Long
    Dim summaryStartRow As Long
    Dim footerStartRow As Long
    
    On Error GoTo ErrHandler
    
    Set wsPrint = ThisWorkbook.Worksheets(SHEET_PRINT)
    Set wsSettings = ThisWorkbook.Worksheets(SHEET_SETTINGS)
    Set wsInvoiceDB = ThisWorkbook.Worksheets(SHEET_INVOICE_DB)
    Set wsSalesDB = ThisWorkbook.Worksheets(SHEET_SALES_DB)
    Set wsCustomersDB = ThisWorkbook.Worksheets(SHEET_CUSTOMERS_DB)
    
    Set tblSettings = wsSettings.ListObjects(TABLE_SETTINGS)
    Set tblInvoices = wsInvoiceDB.ListObjects(TABLE_INVOICES)
    Set tblSales = wsSalesDB.ListObjects(TABLE_SALES)
    Set tblCustomers = wsCustomersDB.ListObjects(TABLE_CUSTOMERS)
    
    invoiceRow = FindRowByExactValue(tblInvoices, "Invoice_No", invoiceNo)
    If invoiceRow = 0 Then
        MsgBox "Invoice not found in Invoice_DB.", vbExclamation
        Exit Sub
    End If
    
    ' Invoice header
    invoiceDate = GetTableValue(tblInvoices, invoiceRow, "Invoice_Date")
    salesOrderNo = Trim(CStr(GetTableValue(tblInvoices, invoiceRow, "Sales_Order_No")))
    customerID = Trim(CStr(GetTableValue(tblInvoices, invoiceRow, "Customer_ID")))
    customerName = Trim(CStr(GetTableValue(tblInvoices, invoiceRow, "Customer_Name")))
    
    subtotal = CDbl(NzNumber(GetTableValue(tblInvoices, invoiceRow, "Subtotal")))
    tax1Name = Trim(CStr(GetTableValue(tblInvoices, invoiceRow, "Tax1_Name")))
    tax1Rate = CDbl(NzNumber(GetTableValue(tblInvoices, invoiceRow, "Tax1_Rate")))
    tax1Amount = CDbl(NzNumber(GetTableValue(tblInvoices, invoiceRow, "Tax1_Amount")))
    tax2Name = Trim(CStr(GetTableValue(tblInvoices, invoiceRow, "Tax2_Name")))
    tax2Rate = CDbl(NzNumber(GetTableValue(tblInvoices, invoiceRow, "Tax2_Rate")))
    tax2Amount = CDbl(NzNumber(GetTableValue(tblInvoices, invoiceRow, "Tax2_Amount")))
    grandTotal = CDbl(NzNumber(GetTableValue(tblInvoices, invoiceRow, "Grand_Total")))
    amountPaid = CDbl(NzNumber(GetTableValue(tblInvoices, invoiceRow, "Amount_Paid")))
    balanceDue = CDbl(NzNumber(GetTableValue(tblInvoices, invoiceRow, "Balance_Due")))
    invoiceStatus = Trim(CStr(GetTableValue(tblInvoices, invoiceRow, "Invoice_Status")))
    notesText = Trim(CStr(GetTableValue(tblInvoices, invoiceRow, "Notes")))
    
    ' Company data
    companyName = Trim(CStr(GetTableValue(tblSettings, 1, "Company_Name")))
    companyAddress = Trim(CStr(GetTableValue(tblSettings, 1, "Address")))
    companyPhone = Trim(CStr(GetTableValue(tblSettings, 1, "Phone")))
    companyEmail = Trim(CStr(GetTableValue(tblSettings, 1, "Email")))
    
    ' Customer data
    customerRow = 0
    If customerID <> "" Then
        customerRow = FindRowByExactValue(tblCustomers, "Customer_ID", customerID)
    End If
    
    If customerRow > 0 Then
        customerAddress = Trim(CStr(GetTableValue(tblCustomers, customerRow, "Address")))
        customerPhone = Trim(CStr(GetTableValue(tblCustomers, customerRow, "Phone")))
        customerEmail = Trim(CStr(GetTableValue(tblCustomers, customerRow, "Email")))
    Else
        customerAddress = ""
        customerPhone = ""
        customerEmail = ""
    End If
    
    PreparePrintSheet wsPrint
    LoadLogoToInvoicePrint
    
    '========================================
    ' Header
    '========================================
    ' Remove duplicated seller/company block beside logo
    wsPrint.Range("I2").value = "INVOICE"
    wsPrint.Range("I4").value = "Invoice Number:"
    wsPrint.Range("K4").value = invoiceNo
    wsPrint.Range("I5").value = "Invoice Date:"
    wsPrint.Range("K5").value = invoiceDate
    
    '========================================
    ' Customer / Order
    '========================================
    wsPrint.Range("A8").value = "BILL TO"
    wsPrint.Range("A9").value = customerName
    wsPrint.Range("A10").value = customerAddress
    wsPrint.Range("A11").value = customerPhone
    wsPrint.Range("A12").value = customerEmail
    
    wsPrint.Range("I8").value = "Sales Order No:"
    wsPrint.Range("K8").value = salesOrderNo
    
    '========================================
    ' Detail table
    '========================================
    WriteDetailHeader wsPrint, DETAIL_HEADER_ROW
    lastDetailRow = WriteDetailLines(wsPrint, tblSales, salesOrderNo, DETAIL_START_ROW, tax1Rate, tax2Rate)
    
    '========================================
    ' Summary
    '========================================
    summaryStartRow = lastDetailRow + 2
    WriteSummarySection wsPrint, summaryStartRow, subtotal, tax1Name, tax1Rate, tax1Amount, tax2Name, tax2Rate, tax2Amount, grandTotal, amountPaid, balanceDue, invoiceStatus
    
    '========================================
    ' Footer
    '========================================
    footerStartRow = summaryStartRow + 9
    
    wsPrint.Cells(footerStartRow, 1).value = "Notes:"
    wsPrint.Cells(footerStartRow + 1, 1).value = notesText
    
    wsPrint.Cells(footerStartRow + 4, 1).value = "Thank you for your business."
    wsPrint.Cells(footerStartRow + 6, 1).value = companyName
    wsPrint.Cells(footerStartRow + 7, 1).value = companyAddress
    wsPrint.Cells(footerStartRow + 8, 1).value = companyPhone
    wsPrint.Cells(footerStartRow + 9, 1).value = companyEmail
    
    FormatPrintSheet wsPrint, footerStartRow + 10
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical

End Sub

'========================================
' Resolve Invoice Number
'========================================

Private Function ResolveInvoiceNoFromUI() As String

    Dim wsUI As Worksheet
    Dim wsInvoiceDB As Worksheet
    Dim tblInvoices As ListObject
    
    Dim invoiceLookup As String
    Dim invoiceNo As String
    Dim salesOrderNo As String
    Dim invoiceRow As Long
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsInvoiceDB = ThisWorkbook.Worksheets(SHEET_INVOICE_DB)
    Set tblInvoices = wsInvoiceDB.ListObjects(TABLE_INVOICES)
    
    invoiceLookup = Trim(CStr(wsUI.Range(UI_INVOICE_LOOKUP).value))
    invoiceNo = Trim(CStr(wsUI.Range(UI_INVOICE_NO).value))
    salesOrderNo = Trim(CStr(wsUI.Range(UI_SALES_ORDER_NO).value))
    
    ResolveInvoiceNoFromUI = ""
    
    If invoiceLookup <> "" Then
        ResolveInvoiceNoFromUI = invoiceLookup
        Exit Function
    End If
    
    If invoiceNo <> "" Then
        ResolveInvoiceNoFromUI = invoiceNo
        Exit Function
    End If
    
    If salesOrderNo <> "" Then
        invoiceRow = FindRowByExactValue(tblInvoices, "Sales_Order_No", salesOrderNo)
        If invoiceRow > 0 Then
            ResolveInvoiceNoFromUI = CStr(GetTableValue(tblInvoices, invoiceRow, "Invoice_No"))
            Exit Function
        Else
            MsgBox "No invoice found for this Sales Order No.", vbExclamation
            Exit Function
        End If
    End If
    
    ResolveInvoiceNoFromUI = InputBox("Enter Invoice Number:")
    ResolveInvoiceNoFromUI = Trim(ResolveInvoiceNoFromUI)

End Function

'========================================
' Detail Rendering
'========================================

Private Sub WriteDetailHeader(ws As Worksheet, ByVal headerRow As Long)

    ws.Cells(headerRow, 1).value = "ITEM"
    ws.Cells(headerRow, 6).value = "SKU"
    ws.Cells(headerRow, 8).value = "TAX1"
    ws.Cells(headerRow, 9).value = "TAX2"
    ws.Cells(headerRow, 10).value = "QTY"
    ws.Cells(headerRow, 11).value = "UNIT PRICE"
    ws.Cells(headerRow, 12).value = "LINE TOTAL"
    
    With ws.Range(ws.Cells(headerRow, 1), ws.Cells(headerRow, 12))
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

End Sub

Private Function WriteDetailLines(ws As Worksheet, tblSales As ListObject, ByVal salesOrderNo As String, ByVal startRow As Long, ByVal tax1Rate As Double, ByVal tax2Rate As Double) As Long

    Dim i As Long
    Dim r As Long
    Dim qtyVal As Double
    Dim unitPriceVal As Double
    Dim lineTotalVal As Double
    
    r = startRow
    
    For i = 1 To tblSales.ListRows.Count
        If Trim(CStr(GetTableValue(tblSales, i, "Sales_Order_No"))) = salesOrderNo Then
            
            qtyVal = CDbl(NzNumber(GetTableValue(tblSales, i, "Qty")))
            unitPriceVal = CDbl(NzNumber(GetTableValue(tblSales, i, "Unit_Price")))
            lineTotalVal = CDbl(NzNumber(GetTableValue(tblSales, i, "Line_Total")))
            
            ws.Cells(r, 1).value = CStr(GetTableValue(tblSales, i, "Product_Name"))
            ws.Cells(r, 6).value = CStr(GetTableValue(tblSales, i, "SKU"))
            ws.Cells(r, 8).value = Format(tax1Rate, "0.00%")
            ws.Cells(r, 9).value = Format(tax2Rate, "0.000%")
            ws.Cells(r, 10).value = qtyVal
            ws.Cells(r, 11).value = unitPriceVal
            ws.Cells(r, 12).value = lineTotalVal
            
            ws.Cells(r, 1).HorizontalAlignment = xlLeft
            ws.Cells(r, 6).HorizontalAlignment = xlLeft
            ws.Cells(r, 8).HorizontalAlignment = xlCenter
            ws.Cells(r, 9).HorizontalAlignment = xlCenter
            ws.Cells(r, 10).HorizontalAlignment = xlCenter
            ws.Cells(r, 11).HorizontalAlignment = xlRight
            ws.Cells(r, 12).HorizontalAlignment = xlRight
            
            ws.Cells(r, 11).NumberFormat = "#,##0.00"
            ws.Cells(r, 12).NumberFormat = "#,##0.00"
            
            r = r + 1
        End If
    Next i
    
    If r = startRow Then
        ws.Cells(r, 1).value = "(No line items found)"
        r = r + 1
    End If
    
    ws.Range(ws.Cells(startRow, 1), ws.Cells(r - 1, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    WriteDetailLines = r - 1

End Function

Private Sub WriteSummarySection(ws As Worksheet, ByVal startRow As Long, ByVal subtotal As Double, ByVal tax1Name As String, ByVal tax1Rate As Double, ByVal tax1Amount As Double, ByVal tax2Name As String, ByVal tax2Rate As Double, ByVal tax2Amount As Double, ByVal grandTotal As Double, ByVal amountPaid As Double, ByVal balanceDue As Double, ByVal invoiceStatus As String)

    ws.Cells(startRow, 10).value = "Subtotal"
    ws.Cells(startRow, 12).value = subtotal
    
    ws.Cells(startRow + 1, 10).value = tax1Name & " (" & Format(tax1Rate, "0.00%") & ")"
    ws.Cells(startRow + 1, 12).value = tax1Amount
    
    ws.Cells(startRow + 2, 10).value = tax2Name & " (" & Format(tax2Rate, "0.000%") & ")"
    ws.Cells(startRow + 2, 12).value = tax2Amount
    
    ws.Cells(startRow + 3, 10).value = "Grand Total"
    ws.Cells(startRow + 3, 12).value = grandTotal
    
    ws.Cells(startRow + 4, 10).value = "Amount Paid"
    ws.Cells(startRow + 4, 12).value = amountPaid
    
    ws.Cells(startRow + 5, 10).value = "Balance Due"
    ws.Cells(startRow + 5, 12).value = balanceDue
    
    ws.Cells(startRow + 6, 10).value = "Status"
    ws.Cells(startRow + 6, 12).value = invoiceStatus
    
    With ws.Range(ws.Cells(startRow, 12), ws.Cells(startRow + 5, 12))
        .NumberFormat = "#,##0.00"
        .HorizontalAlignment = xlRight
    End With
    
    ws.Cells(startRow + 3, 10).Font.Bold = True
    ws.Cells(startRow + 3, 12).Font.Bold = True
    
    ws.Range(ws.Cells(startRow + 3, 10), ws.Cells(startRow + 3, 12)).Borders(xlEdgeTop).LineStyle = xlContinuous
    ws.Range(ws.Cells(startRow + 3, 10), ws.Cells(startRow + 3, 12)).Borders(xlEdgeBottom).LineStyle = xlContinuous

End Sub

'========================================
' Sheet Prep / Formatting
'========================================

Private Sub PreparePrintSheet(ws As Worksheet)

    ws.Cells.ClearContents
    ws.Cells.UnMerge
    ws.Cells.Font.Bold = False
    ws.Cells.Font.Size = 10
    DeleteShapeIfExists ws, "CompanyLogo_Print"

End Sub

Private Sub FormatPrintSheet(ws As Worksheet, ByVal LastUsedRow As Long)

    With ws.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperLetter
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .PrintGridlines = False
        .PrintArea = "$A$1:$L$" & Application.Max(LastUsedRow, 36)
    End With
    
    ws.Range("A1:L80").Font.name = "Aptos Narrow"
    ws.Range("A1:L80").Font.Size = 10
    
    ' Title
    ws.Range("I2").Font.Bold = True
    ws.Range("I2").Font.Size = 28
    
    ' Bill To
    ws.Range("A8").Font.Bold = True
    ws.Range("A8").Font.Size = 11
    ws.Range("A9:A12").HorizontalAlignment = xlLeft
    ws.Range("A9:A12").VerticalAlignment = xlTop
    
    ' Right-side info
    ws.Range("I8").Font.Bold = True
    
    ' Footer company name larger and bold
    If Trim(CStr(ws.Cells(LastUsedRow - 4, 1).value)) <> "" Then
        ws.Cells(LastUsedRow - 4, 1).Font.Bold = True
        ws.Cells(LastUsedRow - 4, 1).Font.Size = 12
    End If

End Sub

'========================================
' Logo
'========================================

Private Sub LoadLogoToInvoicePrint()

    Dim wsSetup As Worksheet
    Dim wsPrint As Worksheet
    Dim shp As Shape
    Dim newShp As Shape
    
    Set wsSetup = ThisWorkbook.Worksheets(SHEET_SETUP_UI)
    Set wsPrint = ThisWorkbook.Worksheets(SHEET_PRINT)
    
    DeleteShapeIfExists wsPrint, "CompanyLogo_Print"
    
    For Each shp In wsSetup.Shapes
        If shp.name = "CompanyLogo" Then
            shp.Copy
            wsPrint.Paste
            
            Set newShp = wsPrint.Shapes(wsPrint.Shapes.Count)
            newShp.name = "CompanyLogo_Print"
            
            newShp.Left = wsPrint.Range(LOGO_ANCHOR).Left
            newShp.Top = wsPrint.Range(LOGO_ANCHOR).Top
            newShp.LockAspectRatio = msoTrue
            newShp.Placement = xlMoveAndSize
            
            If newShp.Width > 180 Then newShp.Width = 180
            If newShp.Height > 90 Then newShp.Height = 90
            
            Exit For
        End If
    Next shp

End Sub

Private Sub DeleteShapeIfExists(ws As Worksheet, shapeName As String)

    Dim shp As Shape
    
    For Each shp In ws.Shapes
        If shp.name = shapeName Then
            shp.Delete
            Exit For
        End If
    Next shp

End Sub

'========================================
' Generic Helpers
'========================================

Private Function FindRowByExactValue(tbl As ListObject, colName As String, lookupValue As String) As Long

    Dim i As Long
    
    FindRowByExactValue = 0
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns(colName).DataBodyRange.Cells(i, 1).value)) = Trim(lookupValue) Then
            FindRowByExactValue = i
            Exit Function
        End If
    Next i

End Function

Private Function GetTableValue(tbl As ListObject, rowIndex As Long, colName As String) As Variant
    GetTableValue = tbl.ListColumns(colName).DataBodyRange.Cells(rowIndex, 1).value
End Function

Private Function NzNumber(val As Variant) As Double
    If Trim(CStr(val)) = "" Then
        NzNumber = 0
    ElseIf IsNumeric(val) Then
        NzNumber = CDbl(val)
    Else
        NzNumber = 0
    End If
End Function

Private Function CleanFileName(ByVal s As String) As String

    Dim badChars As Variant
    Dim i As Long
    
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    CleanFileName = s
    
    For i = LBound(badChars) To UBound(badChars)
        CleanFileName = Replace(CleanFileName, badChars(i), "_")
    Next i

End Function


