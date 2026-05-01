Attribute VB_Name = "ModInvoice"
Option Explicit

Private Const SHEET_UI As String = "Invoice_UI"
Private Const TABLE_SALES As String = "tblSales"
Private Const TABLE_INVOICES As String = "tblInvoices"
Private Const TABLE_SETTINGS As String = "tblSettings"

Private Const SHEET_PWD As String = ""

' UI Mapping - updated to current layout
Private Const CELL_INVOICE_LOOKUP As String = "B3"
Private Const CELL_INVOICE_NO As String = "B4"
Private Const CELL_INVOICE_DATE As String = "B5"
Private Const CELL_SALES_ORDER_NO As String = "B6"
Private Const CELL_CUSTOMER_NAME As String = "B8"
Private Const CELL_SUBTOTAL As String = "B9"
Private Const CELL_TAX1_NAME As String = "B10"
Private Const CELL_TAX1_RATE As String = "B11"
Private Const CELL_TAX1_AMOUNT As String = "B12"
Private Const CELL_TAX2_NAME As String = "B13"
Private Const CELL_TAX2_RATE As String = "B14"
Private Const CELL_TAX2_AMOUNT As String = "B15"
Private Const CELL_GRAND_TOTAL As String = "B16"
Private Const CELL_AMOUNT_PAID As String = "B17"
Private Const CELL_BALANCE_DUE As String = "B18"
Private Const CELL_INVOICE_STATUS As String = "B19"
Private Const CELL_NOTES As String = "B20"

Public Sub Invoice_New()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    
    wsUI.Unprotect Password:=SHEET_PWD
    ClearInvoiceForm wsUI
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True

End Sub

Public Sub Invoice_LoadSales()

    Dim wsUI As Worksheet
    Dim wsSales As Worksheet
    Dim wsSettings As Worksheet
    
    Dim tblSales As ListObject
    Dim tblSettings As ListObject
    
    Dim salesOrderNo As String
    Dim salesRow As Long
    
    Dim customerName As String
    Dim subtotal As Double
    
    Dim tax1Name As String
    Dim tax2Name As String
    Dim tax2OnTax1 As Boolean
    
    Dim tax1Rate As Double
    Dim tax2Rate As Double
    Dim tax1Amount As Double
    Dim tax2Amount As Double
    Dim grandTotal As Double
    Dim amountPaid As Double
    Dim balanceDue As Double
    Dim invoiceStatus As String
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsSales = ThisWorkbook.Worksheets("Sales_DB")
    Set wsSettings = ThisWorkbook.Worksheets("Settings")
    
    Set tblSales = wsSales.ListObjects(TABLE_SALES)
    Set tblSettings = wsSettings.ListObjects(TABLE_SETTINGS)
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    salesOrderNo = Trim(CStr(wsUI.Range(CELL_SALES_ORDER_NO).value))
    
    If salesOrderNo = "" Then
        MsgBox "Please select Sales Order No.", vbExclamation
        GoTo SafeExit
    End If
    
    salesRow = FindRowByExactValue(tblSales, "Sales_Order_No", salesOrderNo)
    
    If salesRow = 0 Then
        MsgBox "Sales Order No not found in Sales_DB.", vbCritical
        GoTo SafeExit
    End If
    
    customerName = Trim(CStr(GetTableValue(tblSales, salesRow, "Customer_Name")))
    subtotal = CDbl(NzNumber(GetTableValue(tblSales, salesRow, "Line_Total")))
    
    tax1Name = Trim(CStr(GetTableValue(tblSettings, 1, "Tax1_Name")))
    tax1Rate = CDbl(NzNumber(GetTableValue(tblSettings, 1, "Tax1_Rate")))
    tax2Name = Trim(CStr(GetTableValue(tblSettings, 1, "Tax2_Name")))
    tax2Rate = CDbl(NzNumber(GetTableValue(tblSettings, 1, "Tax2_Rate")))
    tax2OnTax1 = ToBoolean(GetTableValue(tblSettings, 1, "Tax2_On_Tax1"))
    
    tax1Amount = Round(subtotal * tax1Rate, 2)
    
    If tax2OnTax1 Then
        tax2Amount = Round((subtotal + tax1Amount) * tax2Rate, 2)
    Else
        tax2Amount = Round(subtotal * tax2Rate, 2)
    End If
    
    grandTotal = Round(subtotal + tax1Amount + tax2Amount, 2)
    amountPaid = 0
    balanceDue = grandTotal
    invoiceStatus = "Unpaid"
    
    wsUI.Range(CELL_CUSTOMER_NAME).value = customerName
    wsUI.Range(CELL_SUBTOTAL).value = subtotal
    wsUI.Range(CELL_TAX1_NAME).value = tax1Name
    wsUI.Range(CELL_TAX1_RATE).value = tax1Rate
    wsUI.Range(CELL_TAX1_AMOUNT).value = tax1Amount
    wsUI.Range(CELL_TAX2_NAME).value = tax2Name
    wsUI.Range(CELL_TAX2_RATE).value = tax2Rate
    wsUI.Range(CELL_TAX2_AMOUNT).value = tax2Amount
    wsUI.Range(CELL_GRAND_TOTAL).value = grandTotal
    wsUI.Range(CELL_AMOUNT_PAID).value = amountPaid
    wsUI.Range(CELL_BALANCE_DUE).value = balanceDue
    wsUI.Range(CELL_INVOICE_STATUS).value = invoiceStatus
    
    MsgBox "Sales data loaded successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

Public Sub Invoice_Save()

    Dim wsUI As Worksheet
    Dim wsInvoices As Worksheet
    Dim wsSales As Worksheet
    
    Dim tblInvoices As ListObject
    Dim tblSales As ListObject
    
    Dim invoiceNo As String
    Dim invoiceDate As Variant
    Dim salesOrderNo As String
    Dim salesRow As Long
    Dim existingInvoiceRow As Long
    
    Dim customerID As String
    Dim customerName As String
    
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
    Dim notes As String
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsInvoices = ThisWorkbook.Worksheets("Invoice_DB")
    Set wsSales = ThisWorkbook.Worksheets("Sales_DB")
    
    Set tblInvoices = wsInvoices.ListObjects(TABLE_INVOICES)
    Set tblSales = wsSales.ListObjects(TABLE_SALES)
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    If Trim(CStr(wsUI.Range(CELL_SALES_ORDER_NO).value)) = "" Then
        MsgBox "Please select Sales Order No.", vbExclamation
        GoTo SafeExit
    End If
    
    If Trim(CStr(wsUI.Range(CELL_INVOICE_DATE).value)) = "" Then
        MsgBox "Please enter Invoice Date.", vbExclamation
        GoTo SafeExit
    End If
    
    If Not IsDate(wsUI.Range(CELL_INVOICE_DATE).value) Then
        MsgBox "Invoice Date is invalid.", vbExclamation
        GoTo SafeExit
    End If
    
    salesOrderNo = Trim(CStr(wsUI.Range(CELL_SALES_ORDER_NO).value))
    invoiceDate = wsUI.Range(CELL_INVOICE_DATE).value
    
    salesRow = FindRowByExactValue(tblSales, "Sales_Order_No", salesOrderNo)
    If salesRow = 0 Then
        MsgBox "Sales Order No not found in Sales_DB.", vbCritical
        GoTo SafeExit
    End If
    
    existingInvoiceRow = FindRowByExactValue(tblInvoices, "Sales_Order_No", salesOrderNo)
    If existingInvoiceRow > 0 Then
        MsgBox "This Sales Order No has already been invoiced.", vbExclamation
        GoTo SafeExit
    End If
    
    invoiceNo = GenerateInvoiceNo(tblInvoices)
    wsUI.Range(CELL_INVOICE_NO).value = invoiceNo
    
    customerID = Trim(CStr(GetTableValue(tblSales, salesRow, "Customer_ID")))
    customerName = Trim(CStr(wsUI.Range(CELL_CUSTOMER_NAME).value))
    
    subtotal = CDbl(NzNumber(wsUI.Range(CELL_SUBTOTAL).value))
    tax1Name = Trim(CStr(wsUI.Range(CELL_TAX1_NAME).value))
    tax1Rate = CDbl(NzNumber(wsUI.Range(CELL_TAX1_RATE).value))
    tax1Amount = CDbl(NzNumber(wsUI.Range(CELL_TAX1_AMOUNT).value))
    tax2Name = Trim(CStr(wsUI.Range(CELL_TAX2_NAME).value))
    tax2Rate = CDbl(NzNumber(wsUI.Range(CELL_TAX2_RATE).value))
    tax2Amount = CDbl(NzNumber(wsUI.Range(CELL_TAX2_AMOUNT).value))
    grandTotal = CDbl(NzNumber(wsUI.Range(CELL_GRAND_TOTAL).value))
    amountPaid = CDbl(NzNumber(wsUI.Range(CELL_AMOUNT_PAID).value))
    balanceDue = CDbl(NzNumber(wsUI.Range(CELL_BALANCE_DUE).value))
    invoiceStatus = Trim(CStr(wsUI.Range(CELL_INVOICE_STATUS).value))
    notes = Trim(CStr(wsUI.Range(CELL_NOTES).value))
    
    With tblInvoices.ListRows.Add
        .Range(1, 1).value = invoiceNo
        .Range(1, 2).value = invoiceDate
        .Range(1, 3).value = salesOrderNo
        .Range(1, 4).value = customerID
        .Range(1, 5).value = customerName
        .Range(1, 6).value = subtotal
        .Range(1, 7).value = tax1Name
        .Range(1, 8).value = tax1Rate
        .Range(1, 9).value = tax1Amount
        .Range(1, 10).value = tax2Name
        .Range(1, 11).value = tax2Rate
        .Range(1, 12).value = tax2Amount
        .Range(1, 13).value = grandTotal
        .Range(1, 14).value = amountPaid
        .Range(1, 15).value = balanceDue
        .Range(1, 16).value = invoiceStatus
        .Range(1, 17).value = notes
        .Range(1, 18).value = Date
        .Range(1, 19).value = Date
    End With
    
    SetTableValue tblSales, salesRow, "Invoice_Status", "Invoiced"
    SetTableValue tblSales, salesRow, "Updated_At", Date
    
    MsgBox "Invoice saved successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

Public Sub Invoice_LoadInvoice()

    Dim wsUI As Worksheet
    Dim wsInvoices As Worksheet
    Dim tblInvoices As ListObject
    
    Dim invoiceLookup As String
    Dim invoiceNo As String
    Dim salesOrderNo As String
    Dim invoiceRow As Long
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsInvoices = ThisWorkbook.Worksheets("Invoice_DB")
    Set tblInvoices = wsInvoices.ListObjects(TABLE_INVOICES)
    
    invoiceLookup = Trim(CStr(wsUI.Range(CELL_INVOICE_LOOKUP).value))
    invoiceNo = Trim(CStr(wsUI.Range(CELL_INVOICE_NO).value))
    salesOrderNo = Trim(CStr(wsUI.Range(CELL_SALES_ORDER_NO).value))
    
    If invoiceLookup = "" And invoiceNo = "" And salesOrderNo = "" Then
        MsgBox "Please enter Invoice Lookup, Invoice No, or Sales Order No first.", vbExclamation
        Exit Sub
    End If
    
    invoiceRow = 0
    
    If invoiceLookup <> "" Then
        invoiceRow = FindRowByExactValue(tblInvoices, "Invoice_No", invoiceLookup)
    End If
    
    If invoiceRow = 0 And invoiceNo <> "" Then
        invoiceRow = FindRowByExactValue(tblInvoices, "Invoice_No", invoiceNo)
    End If
    
    If invoiceRow = 0 And salesOrderNo <> "" Then
        invoiceRow = FindRowByExactValue(tblInvoices, "Sales_Order_No", salesOrderNo)
    End If
    
    If invoiceRow = 0 Then
        MsgBox "Invoice not found in Invoice_DB.", vbExclamation
        Exit Sub
    End If
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    wsUI.Range(CELL_INVOICE_LOOKUP).value = GetTableValue(tblInvoices, invoiceRow, "Invoice_No")
    wsUI.Range(CELL_INVOICE_NO).value = GetTableValue(tblInvoices, invoiceRow, "Invoice_No")
    wsUI.Range(CELL_INVOICE_DATE).value = GetTableValue(tblInvoices, invoiceRow, "Invoice_Date")
    wsUI.Range(CELL_SALES_ORDER_NO).value = GetTableValue(tblInvoices, invoiceRow, "Sales_Order_No")
    wsUI.Range(CELL_CUSTOMER_NAME).value = GetTableValue(tblInvoices, invoiceRow, "Customer_Name")
    wsUI.Range(CELL_SUBTOTAL).value = GetTableValue(tblInvoices, invoiceRow, "Subtotal")
    wsUI.Range(CELL_TAX1_NAME).value = GetTableValue(tblInvoices, invoiceRow, "Tax1_Name")
    wsUI.Range(CELL_TAX1_RATE).value = GetTableValue(tblInvoices, invoiceRow, "Tax1_Rate")
    wsUI.Range(CELL_TAX1_AMOUNT).value = GetTableValue(tblInvoices, invoiceRow, "Tax1_Amount")
    wsUI.Range(CELL_TAX2_NAME).value = GetTableValue(tblInvoices, invoiceRow, "Tax2_Name")
    wsUI.Range(CELL_TAX2_RATE).value = GetTableValue(tblInvoices, invoiceRow, "Tax2_Rate")
    wsUI.Range(CELL_TAX2_AMOUNT).value = GetTableValue(tblInvoices, invoiceRow, "Tax2_Amount")
    wsUI.Range(CELL_GRAND_TOTAL).value = GetTableValue(tblInvoices, invoiceRow, "Grand_Total")
    wsUI.Range(CELL_AMOUNT_PAID).value = GetTableValue(tblInvoices, invoiceRow, "Amount_Paid")
    wsUI.Range(CELL_BALANCE_DUE).value = GetTableValue(tblInvoices, invoiceRow, "Balance_Due")
    wsUI.Range(CELL_INVOICE_STATUS).value = GetTableValue(tblInvoices, invoiceRow, "Invoice_Status")
    wsUI.Range(CELL_NOTES).value = GetTableValue(tblInvoices, invoiceRow, "Notes")
    
    MsgBox "Invoice loaded successfully.", vbInformation
    
SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

Private Sub ClearInvoiceForm(wsUI As Worksheet)

    wsUI.Range(CELL_INVOICE_LOOKUP).value = ""
    wsUI.Range(CELL_INVOICE_NO).value = ""
    wsUI.Range(CELL_INVOICE_DATE).value = ""
    wsUI.Range(CELL_SALES_ORDER_NO).value = ""
    wsUI.Range(CELL_CUSTOMER_NAME).value = ""
    wsUI.Range(CELL_SUBTOTAL).value = ""
    wsUI.Range(CELL_TAX1_NAME).value = ""
    wsUI.Range(CELL_TAX1_RATE).value = ""
    wsUI.Range(CELL_TAX1_AMOUNT).value = ""
    wsUI.Range(CELL_TAX2_NAME).value = ""
    wsUI.Range(CELL_TAX2_RATE).value = ""
    wsUI.Range(CELL_TAX2_AMOUNT).value = ""
    wsUI.Range(CELL_GRAND_TOTAL).value = ""
    wsUI.Range(CELL_AMOUNT_PAID).value = ""
    wsUI.Range(CELL_BALANCE_DUE).value = ""
    wsUI.Range(CELL_INVOICE_STATUS).value = ""
    wsUI.Range(CELL_NOTES).value = ""

End Sub

Private Function GenerateInvoiceNo(tbl As ListObject) As String
    GenerateInvoiceNo = "INV" & Format(tbl.ListRows.Count + 1, "00000")
End Function

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

Private Sub SetTableValue(tbl As ListObject, rowIndex As Long, colName As String, val As Variant)
    tbl.ListColumns(colName).DataBodyRange.Cells(rowIndex, 1).value = val
End Sub

Private Function NzNumber(val As Variant) As Double
    If Trim(CStr(val)) = "" Then
        NzNumber = 0
    ElseIf IsNumeric(val) Then
        NzNumber = CDbl(val)
    Else
        NzNumber = 0
    End If
End Function

Private Function ToBoolean(val As Variant) As Boolean

    Dim s As String
    s = LCase(Trim(CStr(val)))
    
    Select Case s
        Case "yes", "true", "1", "y"
            ToBoolean = True
        Case Else
            ToBoolean = False
    End Select

End Function
