Attribute VB_Name = "modPayment"
Option Explicit

Private Const SHEET_UI As String = "Payment_UI"
Private Const TABLE_PAYMENT As String = "tblPayment"
Private Const TABLE_CUSTOMERS As String = "tblCustomers"
Private Const TABLE_SUPPLIERS As String = "tblSuppliers"
Private Const TABLE_INVOICE As String = "tblInvoices"
Private Const TABLE_PURCHASE As String = "tblPurchase"

' UI Mapping (updated: B6/B7 swapped)
Private Const CELL_ID As String = "B3"
Private Const CELL_DATE As String = "B4"
Private Const CELL_TYPE As String = "B5"
Private Const CELL_REF As String = "B6"
Private Const CELL_PARTY As String = "B7"
Private Const CELL_AMOUNT As String = "B8"
Private Const CELL_METHOD As String = "B9"
Private Const CELL_NOTES As String = "B10"

Private Const TYPE_CUSTOMER_RECEIPT As String = "CUSTOMER_RECEIPT"
Private Const TYPE_SUPPLIER_PAYMENT As String = "SUPPLIER_PAYMENT"

Public Sub Payment_New()
    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    ClearPaymentForm wsUI
End Sub

Public Sub Payment_Save()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tblPay As ListObject
    Dim paymentID As String

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Payment_DB")
    Set tblPay = wsDB.ListObjects(TABLE_PAYMENT)

    If Trim$(CStr(wsUI.Range(CELL_ID).value)) <> "" Then
        MsgBox "Payment ID should be blank when creating a new payment." & vbCrLf & _
               "Click New Payment first, or use Update Payment for an existing record.", vbExclamation, "Payment Save"
        Exit Sub
    End If

    ValidatePaymentReadyToSave False, "", "", 0

    paymentID = GeneratePaymentID(tblPay)
    wsUI.Range(CELL_ID).value = paymentID

    SavePaymentUsingID paymentID

    MsgBox "Payment saved successfully.", vbInformation, "Payment Save"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Payment Save"

End Sub

Public Sub Payment_Load()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tblPay As ListObject
    Dim paymentID As String
    Dim rowNum As Long

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Payment_DB")
    Set tblPay = wsDB.ListObjects(TABLE_PAYMENT)

    paymentID = Trim$(CStr(wsUI.Range(CELL_ID).value))

    If paymentID = "" Then
        MsgBox "Please enter Payment ID first.", vbExclamation, "Payment Load"
        Exit Sub
    End If

    rowNum = FindPaymentRowByID(tblPay, paymentID)
    If rowNum = 0 Then
        MsgBox "Payment ID not found.", vbExclamation, "Payment Load"
        Exit Sub
    End If

    ClearPaymentForm wsUI
    wsUI.Range(CELL_ID).value = paymentID

    wsUI.Range(CELL_DATE).value = tblPay.ListColumns("Payment_Date").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_TYPE).value = tblPay.ListColumns("Payment_Type").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_REF).value = tblPay.ListColumns("Reference_No").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_PARTY).value = tblPay.ListColumns("Party_Name").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_AMOUNT).value = tblPay.ListColumns("Amount").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_METHOD).value = tblPay.ListColumns("Method").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_NOTES).value = tblPay.ListColumns("Notes").DataBodyRange.Cells(rowNum, 1).value

    MsgBox "Payment loaded successfully.", vbInformation, "Payment Load"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Payment Load"

End Sub

Public Sub Payment_Update()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim wsInv As Worksheet
    Dim wsPO As Worksheet

    Dim tblPay As ListObject
    Dim tblInv As ListObject
    Dim tblPO As ListObject

    Dim paymentID As String
    Dim rowNum As Long

    Dim oldRow As Variant
    Dim oldRef As String
    Dim oldType As String
    Dim oldAmt As Double

    Dim invSnap As Object
    Dim poSnap As Object
    Dim updateStarted As Boolean

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Payment_DB")
    Set wsInv = ThisWorkbook.Worksheets("Invoice_DB")
    Set wsPO = ThisWorkbook.Worksheets("Purchase_DB")

    Set tblPay = wsDB.ListObjects(TABLE_PAYMENT)
    Set tblInv = wsInv.ListObjects(TABLE_INVOICE)
    Set tblPO = wsPO.ListObjects(TABLE_PURCHASE)

    paymentID = Trim$(CStr(wsUI.Range(CELL_ID).value))

    If paymentID = "" Then
        MsgBox "Please load an existing Payment ID before updating.", vbExclamation, "Payment Update"
        Exit Sub
    End If

    rowNum = FindPaymentRowByID(tblPay, paymentID)
    If rowNum = 0 Then
        MsgBox "Payment ID not found in Payment_DB.", vbExclamation, "Payment Update"
        Exit Sub
    End If

    oldRow = tblPay.ListRows(rowNum).Range.value
    oldRef = Trim$(CStr(tblPay.ListColumns("Reference_No").DataBodyRange.Cells(rowNum, 1).value))
    oldType = Trim$(CStr(tblPay.ListColumns("Payment_Type").DataBodyRange.Cells(rowNum, 1).value))
    oldAmt = NzNumber(tblPay.ListColumns("Amount").DataBodyRange.Cells(rowNum, 1).value)

    ValidatePaymentReadyToSave True, oldType, oldRef, oldAmt

    If MsgBox("Update this payment?" & vbCrLf & _
              "Payment ID: " & paymentID, vbQuestion + vbYesNo, "Confirm Update") <> vbYes Then
        Exit Sub
    End If

    Set invSnap = CreateObject("Scripting.Dictionary")
    Set poSnap = CreateObject("Scripting.Dictionary")

    SnapshotInvoiceByRef tblInv, oldType, oldRef, invSnap
    SnapshotPurchaseByRef tblPO, oldType, oldRef, poSnap

    updateStarted = True

    RollbackPaymentImpact tblInv, tblPO, oldType, oldRef, oldAmt

    tblPay.ListRows(rowNum).Delete

    SavePaymentUsingID paymentID

    MsgBox "Payment updated successfully.", vbInformation, "Payment Update"
    Exit Sub

ErrHandler:
    If updateStarted Then
        On Error Resume Next
        RestoreInvoiceSnapshots tblInv, invSnap
        RestorePurchaseSnapshots tblPO, poSnap
        DeletePaymentRowsByID tblPay, paymentID
        RestoreTableRow tblPay, oldRow
        On Error GoTo 0
    End If

    MsgBox "Error: " & Err.Description, vbCritical, "Payment Update"

End Sub

Private Sub SavePaymentUsingID(ByVal paymentID As String)

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim wsCustomers As Worksheet
    Dim wsSuppliers As Worksheet
    Dim wsInv As Worksheet
    Dim wsPO As Worksheet

    Dim tblPay As ListObject
    Dim tblCust As ListObject
    Dim tblSupp As ListObject
    Dim tblInv As ListObject
    Dim tblPO As ListObject

    Dim paymentDate As Variant
    Dim paymentType As String
    Dim partyName As String
    Dim refNo As String
    Dim amount As Double
    Dim methodText As String
    Dim notes As String

    Dim partyType As String
    Dim partyID As String
    Dim referenceType As String

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Payment_DB")
    Set wsCustomers = ThisWorkbook.Worksheets("Customers_DB")
    Set wsSuppliers = ThisWorkbook.Worksheets("Suppliers_DB")
    Set wsInv = ThisWorkbook.Worksheets("Invoice_DB")
    Set wsPO = ThisWorkbook.Worksheets("Purchase_DB")

    Set tblPay = wsDB.ListObjects(TABLE_PAYMENT)
    Set tblCust = wsCustomers.ListObjects(TABLE_CUSTOMERS)
    Set tblSupp = wsSuppliers.ListObjects(TABLE_SUPPLIERS)
    Set tblInv = wsInv.ListObjects(TABLE_INVOICE)
    Set tblPO = wsPO.ListObjects(TABLE_PURCHASE)

    paymentDate = wsUI.Range(CELL_DATE).value
    paymentType = Trim$(CStr(wsUI.Range(CELL_TYPE).value))
    partyName = Trim$(CStr(wsUI.Range(CELL_PARTY).value))
    refNo = Trim$(CStr(wsUI.Range(CELL_REF).value))
    amount = Round(NzNumber(wsUI.Range(CELL_AMOUNT).value), 2)
    methodText = Trim$(CStr(wsUI.Range(CELL_METHOD).value))
    notes = Trim$(CStr(wsUI.Range(CELL_NOTES).value))

    Select Case paymentType
        Case TYPE_CUSTOMER_RECEIPT
            partyType = "CUSTOMER"
            referenceType = "INVOICE"
            partyID = FindIDByName(tblCust, "Customer_Name", "Customer_ID", partyName)
            If partyID = "" Then Err.Raise vbObjectError + 3101, , "Customer not found in Customers_DB."
            ApplyPaymentToInvoice tblInv, refNo, amount

        Case TYPE_SUPPLIER_PAYMENT
            partyType = "SUPPLIER"
            referenceType = "PURCHASE"
            partyID = FindIDByName(tblSupp, "Supplier_Name", "Supplier_ID", partyName)
            If partyID = "" Then Err.Raise vbObjectError + 3102, , "Supplier not found in Suppliers_DB."
            ApplyPaymentToPurchase tblPO, refNo, amount

        Case Else
            Err.Raise vbObjectError + 3103, , "Invalid Payment Type."
    End Select

    With tblPay.ListRows.Add
        .Range(1, GetCol(tblPay, "Payment_ID")).value = paymentID
        .Range(1, GetCol(tblPay, "Payment_Date")).value = paymentDate
        .Range(1, GetCol(tblPay, "Payment_Type")).value = paymentType
        .Range(1, GetCol(tblPay, "Party_Type")).value = partyType
        .Range(1, GetCol(tblPay, "Party_ID")).value = partyID
        .Range(1, GetCol(tblPay, "Party_Name")).value = partyName
        .Range(1, GetCol(tblPay, "Reference_No")).value = refNo
        .Range(1, GetCol(tblPay, "Reference_Type")).value = referenceType
        .Range(1, GetCol(tblPay, "Amount")).value = amount
        .Range(1, GetCol(tblPay, "Method")).value = methodText
        .Range(1, GetCol(tblPay, "Notes")).value = notes
        .Range(1, GetCol(tblPay, "Created_At")).value = Date
        .Range(1, GetCol(tblPay, "Updated_At")).value = Date
    End With

End Sub

Private Sub ValidatePaymentReadyToSave(ByVal isUpdate As Boolean, _
                                       ByVal oldType As String, _
                                       ByVal oldRef As String, _
                                       ByVal oldAmt As Double)

    Dim wsUI As Worksheet
    Dim wsInv As Worksheet
    Dim wsPO As Worksheet
    Dim wsCust As Worksheet
    Dim wsSupp As Worksheet

    Dim tblInv As ListObject
    Dim tblPO As ListObject
    Dim tblCust As ListObject
    Dim tblSupp As ListObject

    Dim paymentType As String
    Dim partyName As String
    Dim refNo As String
    Dim amount As Double
    Dim releaseAmt As Double

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsInv = ThisWorkbook.Worksheets("Invoice_DB")
    Set wsPO = ThisWorkbook.Worksheets("Purchase_DB")
    Set wsCust = ThisWorkbook.Worksheets("Customers_DB")
    Set wsSupp = ThisWorkbook.Worksheets("Suppliers_DB")

    Set tblInv = wsInv.ListObjects(TABLE_INVOICE)
    Set tblPO = wsPO.ListObjects(TABLE_PURCHASE)
    Set tblCust = wsCust.ListObjects(TABLE_CUSTOMERS)
    Set tblSupp = wsSupp.ListObjects(TABLE_SUPPLIERS)

    If Trim$(CStr(wsUI.Range(CELL_DATE).value)) = "" Then Err.Raise vbObjectError + 3001, , "Please enter Payment Date."
    If Not IsDate(wsUI.Range(CELL_DATE).value) Then Err.Raise vbObjectError + 3002, , "Payment Date is invalid."

    paymentType = Trim$(CStr(wsUI.Range(CELL_TYPE).value))
    If paymentType <> TYPE_CUSTOMER_RECEIPT And paymentType <> TYPE_SUPPLIER_PAYMENT Then
        Err.Raise vbObjectError + 3003, , "Payment Type must be CUSTOMER_RECEIPT or SUPPLIER_PAYMENT."
    End If

    refNo = Trim$(CStr(wsUI.Range(CELL_REF).value))
    If refNo = "" Then Err.Raise vbObjectError + 3005, , "Please enter Reference No."

    partyName = Trim$(CStr(wsUI.Range(CELL_PARTY).value))
    If partyName = "" Then Err.Raise vbObjectError + 3004, , "Please enter Party Name."

    If Not IsNumeric(wsUI.Range(CELL_AMOUNT).value) Or CDbl(wsUI.Range(CELL_AMOUNT).value) <= 0 Then
        Err.Raise vbObjectError + 3006, , "Amount must be greater than 0."
    End If

    amount = Round(CDbl(wsUI.Range(CELL_AMOUNT).value), 2)

    releaseAmt = 0
    If isUpdate Then
        If oldType = paymentType And oldRef = refNo Then
            releaseAmt = Round(oldAmt, 2)
        End If
    End If

    If paymentType = TYPE_CUSTOMER_RECEIPT Then
        If FindIDByName(tblCust, "Customer_Name", "Customer_ID", partyName) = "" Then
            Err.Raise vbObjectError + 3007, , "Customer not found in Customers_DB."
        End If
        ValidateInvoicePayment tblInv, refNo, partyName, amount, releaseAmt
    Else
        If FindIDByName(tblSupp, "Supplier_Name", "Supplier_ID", partyName) = "" Then
            Err.Raise vbObjectError + 3008, , "Supplier not found in Suppliers_DB."
        End If
        ValidatePurchasePayment tblPO, refNo, partyName, amount, releaseAmt
    End If

End Sub

Private Sub ValidateInvoicePayment(ByVal tblInv As ListObject, _
                                   ByVal invoiceNo As String, _
                                   ByVal customerName As String, _
                                   ByVal amount As Double, _
                                   ByVal releaseAmt As Double)

    Dim rowNum As Long
    Dim invoiceCustomer As String
    Dim totalAmt As Double
    Dim paidAmt As Double
    Dim availableAmt As Double
    Dim inputAmt As Double

    rowNum = FindRowByRef(tblInv, "Invoice_No", invoiceNo)
    If rowNum = 0 Then Err.Raise vbObjectError + 3011, , "Invoice not found: " & invoiceNo

    invoiceCustomer = Trim$(CStr(tblInv.ListColumns("Customer_Name").DataBodyRange.Cells(rowNum, 1).value))
    If invoiceCustomer <> customerName Then
        Err.Raise vbObjectError + 3012, , "Customer does not match the selected Invoice."
    End If

    totalAmt = Round(NzNumber(tblInv.ListColumns("Grand_Total").DataBodyRange.Cells(rowNum, 1).value), 2)
    paidAmt = Round(NzNumber(tblInv.ListColumns("Amount_Paid").DataBodyRange.Cells(rowNum, 1).value), 2)
    availableAmt = Round(totalAmt - paidAmt + releaseAmt, 2)
    inputAmt = Round(amount, 2)

    If inputAmt > availableAmt Then
        Err.Raise vbObjectError + 3013, , "Payment amount exceeds remaining invoice balance." & vbCrLf & _
                                          "Available balance: " & Format(availableAmt, "0.00")
    End If

End Sub

Private Sub ValidatePurchasePayment(ByVal tblPO As ListObject, _
                                    ByVal poNo As String, _
                                    ByVal supplierName As String, _
                                    ByVal amount As Double, _
                                    ByVal releaseAmt As Double)

    Dim totalAmt As Double
    Dim paidAmt As Double
    Dim remainingAmt As Double
    Dim poSupplier As String
    Dim inputAmt As Double

    poSupplier = GetPurchaseSupplierName(tblPO, poNo)
    If poSupplier = "" Then Err.Raise vbObjectError + 3021, , "Purchase Order not found: " & poNo
    If poSupplier <> supplierName Then Err.Raise vbObjectError + 3022, , "Supplier does not match the selected Purchase Order."

    totalAmt = Round(GetPurchaseTotalAmount(tblPO, poNo), 2)
    paidAmt = Round(GetPurchasePaidAmount(tblPO, poNo), 2)
    remainingAmt = Round(totalAmt - paidAmt + releaseAmt, 2)
    inputAmt = Round(amount, 2)

    If inputAmt > remainingAmt Then
        Err.Raise vbObjectError + 3023, , "Payment amount exceeds remaining PO balance." & vbCrLf & _
                                          "Available balance: " & Format(remainingAmt, "0.00")
    End If

End Sub

Private Sub ApplyPaymentToInvoice(ByVal tblInv As ListObject, ByVal invoiceNo As String, ByVal amount As Double)

    Dim rowNum As Long
    Dim totalAmt As Double
    Dim paidAmt As Double
    Dim balanceDue As Double

    rowNum = FindRowByRef(tblInv, "Invoice_No", invoiceNo)
    If rowNum = 0 Then Err.Raise vbObjectError + 3031, , "Invoice not found while applying payment."

    totalAmt = Round(NzNumber(tblInv.ListColumns("Grand_Total").DataBodyRange.Cells(rowNum, 1).value), 2)
    paidAmt = Round(NzNumber(tblInv.ListColumns("Amount_Paid").DataBodyRange.Cells(rowNum, 1).value) + amount, 2)
    balanceDue = Round(totalAmt - paidAmt, 2)
    If balanceDue < 0 Then balanceDue = 0

    tblInv.ListColumns("Amount_Paid").DataBodyRange.Cells(rowNum, 1).value = paidAmt
    tblInv.ListColumns("Balance_Due").DataBodyRange.Cells(rowNum, 1).value = balanceDue
    tblInv.ListColumns("Invoice_Status").DataBodyRange.Cells(rowNum, 1).value = GetPaymentStatus(totalAmt, paidAmt)

    If HasColumn(tblInv, "Updated_At") Then
        tblInv.ListColumns("Updated_At").DataBodyRange.Cells(rowNum, 1).value = Date
    End If

End Sub

Private Sub ApplyPaymentToPurchase(ByVal tblPO As ListObject, ByVal poNo As String, ByVal amount As Double)

    Dim i As Long
    Dim totalAmt As Double
    Dim newPaid As Double
    Dim newBalance As Double
    Dim paymentStatus As String

    totalAmt = Round(GetPurchaseTotalAmount(tblPO, poNo), 2)
    newPaid = Round(GetPurchasePaidAmount(tblPO, poNo) + amount, 2)
    newBalance = Round(totalAmt - newPaid, 2)
    If newBalance < 0 Then newBalance = 0

    paymentStatus = GetPaymentStatus(totalAmt, newPaid)

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            tblPO.ListColumns("Amount_Paid").DataBodyRange.Cells(i, 1).value = newPaid
            tblPO.ListColumns("Balance_Due").DataBodyRange.Cells(i, 1).value = newBalance

            If HasColumn(tblPO, "Payment_Status") Then
                tblPO.ListColumns("Payment_Status").DataBodyRange.Cells(i, 1).value = paymentStatus
            End If

            If HasColumn(tblPO, "Updated_At") Then
                tblPO.ListColumns("Updated_At").DataBodyRange.Cells(i, 1).value = Date
            End If
        End If
    Next i

End Sub

Private Sub RollbackPaymentImpact(ByVal tblInv As ListObject, ByVal tblPO As ListObject, ByVal paymentType As String, ByVal refNo As String, ByVal amount As Double)

    Dim rowNum As Long
    Dim totalAmt As Double
    Dim paidAmt As Double
    Dim balanceDue As Double
    Dim paymentStatus As String
    Dim i As Long

    If paymentType = TYPE_CUSTOMER_RECEIPT Then

        rowNum = FindRowByRef(tblInv, "Invoice_No", refNo)
        If rowNum > 0 Then
            totalAmt = Round(NzNumber(tblInv.ListColumns("Grand_Total").DataBodyRange.Cells(rowNum, 1).value), 2)
            paidAmt = Round(NzNumber(tblInv.ListColumns("Amount_Paid").DataBodyRange.Cells(rowNum, 1).value) - amount, 2)
            If paidAmt < 0 Then paidAmt = 0

            balanceDue = Round(totalAmt - paidAmt, 2)
            If balanceDue < 0 Then balanceDue = 0

            tblInv.ListColumns("Amount_Paid").DataBodyRange.Cells(rowNum, 1).value = paidAmt
            tblInv.ListColumns("Balance_Due").DataBodyRange.Cells(rowNum, 1).value = balanceDue
            tblInv.ListColumns("Invoice_Status").DataBodyRange.Cells(rowNum, 1).value = GetPaymentStatus(totalAmt, paidAmt)
        End If

    ElseIf paymentType = TYPE_SUPPLIER_PAYMENT Then

        totalAmt = Round(GetPurchaseTotalAmount(tblPO, refNo), 2)
        paidAmt = Round(GetPurchasePaidAmount(tblPO, refNo) - amount, 2)
        If paidAmt < 0 Then paidAmt = 0

        balanceDue = Round(totalAmt - paidAmt, 2)
        If balanceDue < 0 Then balanceDue = 0

        paymentStatus = GetPaymentStatus(totalAmt, paidAmt)

        For i = 1 To tblPO.ListRows.Count
            If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = refNo Then
                tblPO.ListColumns("Amount_Paid").DataBodyRange.Cells(i, 1).value = paidAmt
                tblPO.ListColumns("Balance_Due").DataBodyRange.Cells(i, 1).value = balanceDue
                If HasColumn(tblPO, "Payment_Status") Then
                    tblPO.ListColumns("Payment_Status").DataBodyRange.Cells(i, 1).value = paymentStatus
                End If
            End If
        Next i

    End If

End Sub

Private Sub ClearPaymentForm(ByVal wsUI As Worksheet)
    wsUI.Range(CELL_ID).value = ""
    wsUI.Range(CELL_DATE).value = ""
    wsUI.Range(CELL_TYPE).value = ""
    wsUI.Range(CELL_REF).value = ""
    wsUI.Range(CELL_PARTY).value = ""
    wsUI.Range(CELL_AMOUNT).value = ""
    wsUI.Range(CELL_METHOD).value = ""
    wsUI.Range(CELL_NOTES).value = ""
End Sub

Private Function GeneratePaymentID(ByVal tbl As ListObject) As String

    Dim maxNum As Long, i As Long, s As String, n As Long
    maxNum = 0

    If tbl.ListRows.Count = 0 Then
        GeneratePaymentID = "PM00001"
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        s = Trim$(CStr(tbl.ListColumns("Payment_ID").DataBodyRange.Cells(i, 1).value))
        If Left$(s, 2) = "PM" Then
            If IsNumeric(Mid$(s, 3)) Then
                n = CLng(Mid$(s, 3))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i

    GeneratePaymentID = "PM" & Format$(maxNum + 1, "00000")

End Function

Private Function FindPaymentRowByID(ByVal tbl As ListObject, ByVal paymentID As String) As Long

    Dim i As Long
    FindPaymentRowByID = 0

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("Payment_ID").DataBodyRange.Cells(i, 1).value)) = paymentID Then
            FindPaymentRowByID = i
            Exit Function
        End If
    Next i

End Function

Private Function FindRowByRef(ByVal tbl As ListObject, ByVal refCol As String, ByVal refValue As String) As Long

    Dim i As Long
    FindRowByRef = 0

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns(refCol).DataBodyRange.Cells(i, 1).value)) = refValue Then
            FindRowByRef = i
            Exit Function
        End If
    Next i

End Function

Private Function FindIDByName(ByVal tbl As ListObject, ByVal nameCol As String, ByVal idCol As String, ByVal partyName As String) As String

    Dim i As Long
    FindIDByName = ""

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns(nameCol).DataBodyRange.Cells(i, 1).value)) = partyName Then
            FindIDByName = Trim$(CStr(tbl.ListColumns(idCol).DataBodyRange.Cells(i, 1).value))
            Exit Function
        End If
    Next i

End Function

Private Function GetPurchaseSupplierName(ByVal tblPO As ListObject, ByVal poNo As String) As String

    Dim i As Long
    GetPurchaseSupplierName = ""

    If tblPO.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            GetPurchaseSupplierName = Trim$(CStr(tblPO.ListColumns("Supplier_Name").DataBodyRange.Cells(i, 1).value))
            Exit Function
        End If
    Next i

End Function

Private Function GetPurchaseTotalAmount(ByVal tblPO As ListObject, ByVal poNo As String) As Double

    Dim i As Long, totalAmt As Double
    totalAmt = 0

    If tblPO.DataBodyRange Is Nothing Then
        GetPurchaseTotalAmount = 0
        Exit Function
    End If

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            totalAmt = totalAmt + NzNumber(tblPO.ListColumns("Line_Total").DataBodyRange.Cells(i, 1).value)
        End If
    Next i

    GetPurchaseTotalAmount = Round(totalAmt, 2)

End Function

Private Function GetPurchasePaidAmount(ByVal tblPO As ListObject, ByVal poNo As String) As Double

    Dim i As Long
    GetPurchasePaidAmount = 0

    If tblPO.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            GetPurchasePaidAmount = Round(NzNumber(tblPO.ListColumns("Amount_Paid").DataBodyRange.Cells(i, 1).value), 2)
            Exit Function
        End If
    Next i

End Function

Private Function GetPaymentStatus(ByVal totalAmt As Double, ByVal paidAmt As Double) As String
    totalAmt = Round(totalAmt, 2)
    paidAmt = Round(paidAmt, 2)

    If paidAmt <= 0 Then
        GetPaymentStatus = "Unpaid"
    ElseIf paidAmt < totalAmt Then
        GetPaymentStatus = "Partial"
    Else
        GetPaymentStatus = "Paid"
    End If
End Function

Private Sub SnapshotInvoiceByRef(ByVal tblInv As ListObject, ByVal paymentType As String, ByVal refNo As String, ByVal snap As Object)

    Dim rowNum As Long, key As String
    If paymentType <> TYPE_CUSTOMER_RECEIPT Then Exit Sub

    rowNum = FindRowByRef(tblInv, "Invoice_No", refNo)
    If rowNum > 0 Then
        key = CStr(rowNum)
        If Not snap.Exists(key) Then snap.Add key, tblInv.ListRows(rowNum).Range.value
    End If

End Sub

Private Sub SnapshotPurchaseByRef(ByVal tblPO As ListObject, ByVal paymentType As String, ByVal refNo As String, ByVal snap As Object)

    Dim i As Long, key As String
    If paymentType <> TYPE_SUPPLIER_PAYMENT Then Exit Sub

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = refNo Then
            key = CStr(i)
            If Not snap.Exists(key) Then snap.Add key, tblPO.ListRows(i).Range.value
        End If
    Next i

End Sub

Private Sub RestoreInvoiceSnapshots(ByVal tblInv As ListObject, ByVal snap As Object)
    Dim k As Variant
    If snap Is Nothing Then Exit Sub
    For Each k In snap.Keys
        tblInv.ListRows(CLng(k)).Range.value = snap(k)
    Next k
End Sub

Private Sub RestorePurchaseSnapshots(ByVal tblPO As ListObject, ByVal snap As Object)
    Dim k As Variant
    If snap Is Nothing Then Exit Sub
    For Each k In snap.Keys
        tblPO.ListRows(CLng(k)).Range.value = snap(k)
    Next k
End Sub

Private Sub DeletePaymentRowsByID(ByVal tbl As ListObject, ByVal paymentID As String)

    Dim i As Long
    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For i = tbl.ListRows.Count To 1 Step -1
        If Trim$(CStr(tbl.ListColumns("Payment_ID").DataBodyRange.Cells(i, 1).value)) = paymentID Then
            tbl.ListRows(i).Delete
        End If
    Next i

End Sub

Private Sub RestoreTableRow(ByVal tbl As ListObject, ByVal rowData As Variant)

    Dim newRow As ListRow
    If IsEmpty(rowData) Then Exit Sub

    Set newRow = tbl.ListRows.Add
    newRow.Range.value = rowData

End Sub

Private Function GetCol(ByVal tbl As ListObject, ByVal colName As String) As Long
    GetCol = tbl.ListColumns(colName).Index
End Function

Private Function HasColumn(ByVal tbl As ListObject, ByVal colName As String) As Boolean
    Dim lc As ListColumn
    HasColumn = False

    For Each lc In tbl.ListColumns
        If StrComp(Trim$(lc.name), Trim$(colName), vbTextCompare) = 0 Then
            HasColumn = True
            Exit Function
        End If
    Next lc
End Function

Private Function NzNumber(ByVal v As Variant) As Double
    If IsError(v) Then
        NzNumber = 0
    ElseIf IsNumeric(v) Then
        NzNumber = CDbl(v)
    ElseIf Trim$(CStr(v)) = "" Then
        NzNumber = 0
    Else
        NzNumber = 0
    End If
End Function

