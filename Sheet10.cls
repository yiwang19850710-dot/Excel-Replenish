VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    On Error GoTo SafeExit
    If Application.EnableEvents = False Then Exit Sub

    Application.EnableEvents = False

    ' Payment Type changed
    If Not Intersect(Target, Me.Range("B5")) Is Nothing Then
        HandlePaymentTypeChanged
    End If

    ' Reference No changed
    If Not Intersect(Target, Me.Range("B6")) Is Nothing Then
        AutoFillPartyNameByReference
    End If

SafeExit:
    Application.EnableEvents = True

End Sub

Private Sub HandlePaymentTypeChanged()

    Me.Range("B6").value = ""
    Me.Range("B7").value = ""

    ApplyReferenceValidationByType

End Sub

Private Sub AutoFillPartyNameByReference()

    Dim paymentType As String
    Dim refNo As String
    Dim partyName As String

    paymentType = Trim$(CStr(Me.Range("B5").value))
    refNo = Trim$(CStr(Me.Range("B6").value))

    Me.Range("B7").value = ""

    If refNo = "" Then Exit Sub

    Select Case paymentType

        Case "CUSTOMER_RECEIPT"
            partyName = GetInvoiceCustomerName(refNo)
            If partyName <> "" Then
                Me.Range("B7").value = partyName
            Else
                MsgBox "Invoice No not found: " & refNo, vbExclamation, "Reference Not Found"
                Me.Range("B6").value = ""
            End If

        Case "SUPPLIER_PAYMENT"
            partyName = GetPurchaseSupplierName(refNo)
            If partyName <> "" Then
                Me.Range("B7").value = partyName
            Else
                MsgBox "Purchase Order No not found: " & refNo, vbExclamation, "Reference Not Found"
                Me.Range("B6").value = ""
            End If

        Case Else
            MsgBox "Please select Payment Type first." & vbCrLf & _
                   "Use CUSTOMER_RECEIPT or SUPPLIER_PAYMENT.", vbExclamation, "Payment Type Required"
            Me.Range("B6").value = ""
    End Select

End Sub

Private Sub ApplyReferenceValidationByType()

    Dim paymentType As String
    Dim formulaText As String

    paymentType = Trim$(CStr(Me.Range("B5").value))

    Me.Range("B6").Validation.Delete

    Select Case paymentType
        Case "CUSTOMER_RECEIPT"
            formulaText = "=nmPaymentInvoiceRefs"

        Case "SUPPLIER_PAYMENT"
            formulaText = "=nmPaymentPurchaseRefs"

        Case Else
            Exit Sub
    End Select

    With Me.Range("B6").Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formulaText
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Reference No"
        .ErrorTitle = "Invalid Reference"
        .InputMessage = "Please choose a valid reference number from the list."
        .ErrorMessage = "Please choose a valid reference number from the dropdown list."
        .ShowInput = True
        .ShowError = True
    End With

End Sub

Private Function GetInvoiceCustomerName(ByVal invoiceNo As String) As String

    Dim wsInv As Worksheet
    Dim tblInv As ListObject
    Dim i As Long

    GetInvoiceCustomerName = ""

    Set wsInv = ThisWorkbook.Worksheets("Invoice_DB")
    Set tblInv = wsInv.ListObjects("tblInvoices")

    If tblInv.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tblInv.ListRows.Count
        If Trim$(CStr(tblInv.ListColumns("Invoice_No").DataBodyRange.Cells(i, 1).value)) = invoiceNo Then
            GetInvoiceCustomerName = Trim$(CStr(tblInv.ListColumns("Customer_Name").DataBodyRange.Cells(i, 1).value))
            Exit Function
        End If
    Next i

End Function

Private Function GetPurchaseSupplierName(ByVal poNo As String) As String

    Dim wsPO As Worksheet
    Dim tblPO As ListObject
    Dim i As Long

    GetPurchaseSupplierName = ""

    Set wsPO = ThisWorkbook.Worksheets("Purchase_DB")
    Set tblPO = wsPO.ListObjects("tblPurchase")

    If tblPO.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            GetPurchaseSupplierName = Trim$(CStr(tblPO.ListColumns("Supplier_Name").DataBodyRange.Cells(i, 1).value))
            Exit Function
        End If
    Next i

End Function

