Attribute VB_Name = "modPurchase"
Option Explicit

Private Const SHEET_UI As String = "Purchase_UI"
Private Const SHEET_PURCHASE_DB As String = "Purchase_DB"
Private Const SHEET_SUPPLIERS As String = "Suppliers_DB"
Private Const SHEET_PRODUCTS As String = "Products_DB"

Private Const TABLE_PURCHASE As String = "tblPurchase"
Private Const TABLE_SUPPLIERS As String = "tblSuppliers"
Private Const TABLE_PRODUCTS As String = "tblProducts"

' Header UI
Private Const CELL_PO_ID As String = "B3"
Private Const CELL_PO_DATE As String = "B4"
Private Const CELL_SUPPLIER As String = "B5"
Private Const CELL_NOTES As String = "B6"

' Lines
Private Const START_ROW As Long = 9
Private Const END_ROW As Long = 18

Private Const COL_SKU As String = "A"
Private Const COL_NAME As String = "B"
Private Const COL_QTY As String = "C"
Private Const COL_UNIT As String = "D"
Private Const COL_LINE_TOTAL As String = "E"

' Summary / status
Private Const CELL_SUBTOTAL As String = "E20"
Private Const CELL_RECEIVED_STATUS As String = "B22"
Private Const CELL_PAYMENT_STATUS As String = "E22"
Private Const CELL_AMOUNT_PAID As String = "B23"
Private Const CELL_BALANCE_DUE As String = "E23"

'==================================================
' NEW
'==================================================
Public Sub Purchase_New()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    ClearPurchaseForm wsUI
    RestorePurchaseLineFormulas wsUI

End Sub

'==================================================
' SAVE / CREATE PURCHASE
'==================================================
Public Sub Purchase_Save()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tblPO As ListObject
    Dim poNo As String

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets(SHEET_PURCHASE_DB)
    Set tblPO = wsDB.ListObjects(TABLE_PURCHASE)

    If Trim$(CStr(wsUI.Range(CELL_PO_ID).value)) <> "" Then
        MsgBox "Purchase ID should be blank when creating a new purchase." & vbCrLf & _
               "Click New Purchase first, or use Update Purchase for an existing purchase.", vbExclamation, "Create Purchase"
        Exit Sub
    End If

    ValidatePurchaseReadyToSave

    poNo = GeneratePurchaseID(tblPO)
    wsUI.Range(CELL_PO_ID).value = poNo

    SavePurchaseUsingID poNo

    MsgBox "Purchase created successfully.", vbInformation, "Create Purchase"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Create Purchase"

End Sub

'==================================================
' LOAD PURCHASE
'==================================================
Public Sub Purchase_Load()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tblPO As ListObject
    Dim poNo As String
    Dim i As Long
    Dim uiRow As Long
    Dim firstRow As Long

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets(SHEET_PURCHASE_DB)
    Set tblPO = wsDB.ListObjects(TABLE_PURCHASE)

    poNo = Trim$(CStr(wsUI.Range(CELL_PO_ID).value))
    If poNo = "" Then
        MsgBox "Please enter Purchase ID first.", vbExclamation, "Load Purchase"
        Exit Sub
    End If

    If Not PurchaseExists(tblPO, poNo) Then
        MsgBox "Purchase ID not found in Purchase_DB.", vbExclamation, "Load Purchase"
        Exit Sub
    End If

    ClearPurchaseForm wsUI
    wsUI.Range(CELL_PO_ID).value = poNo

    firstRow = FindFirstPurchaseRow(tblPO, poNo)
    If firstRow = 0 Then
        MsgBox "Purchase ID not found.", vbExclamation, "Load Purchase"
        Exit Sub
    End If

    wsUI.Range(CELL_PO_DATE).value = tblPO.ListColumns("Purchase_Date").DataBodyRange.Cells(firstRow, 1).value
    wsUI.Range(CELL_SUPPLIER).value = tblPO.ListColumns("Supplier_Name").DataBodyRange.Cells(firstRow, 1).value
    wsUI.Range(CELL_NOTES).value = tblPO.ListColumns("Notes").DataBodyRange.Cells(firstRow, 1).value

    uiRow = START_ROW
    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            If uiRow > END_ROW Then Exit For

            wsUI.Range(COL_SKU & uiRow).value = tblPO.ListColumns("SKU").DataBodyRange.Cells(i, 1).value
            wsUI.Range(COL_NAME & uiRow).value = tblPO.ListColumns("Product_Name").DataBodyRange.Cells(i, 1).value
            wsUI.Range(COL_QTY & uiRow).value = tblPO.ListColumns("Qty").DataBodyRange.Cells(i, 1).value
            wsUI.Range(COL_UNIT & uiRow).value = tblPO.ListColumns("Unit_Cost").DataBodyRange.Cells(i, 1).value
            wsUI.Range(COL_LINE_TOTAL & uiRow).value = tblPO.ListColumns("Line_Total").DataBodyRange.Cells(i, 1).value

            uiRow = uiRow + 1
        End If
    Next i

    UpdatePurchaseSummaryUI poNo

    MsgBox "Purchase loaded successfully.", vbInformation, "Load Purchase"
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Load Purchase"

End Sub

'==================================================
' UPDATE PURCHASE
'==================================================
Public Sub Purchase_Update()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tblPO As ListObject
    Dim poNo As String
    Dim oldRows As Collection
    Dim updateStarted As Boolean

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets(SHEET_PURCHASE_DB)
    Set tblPO = wsDB.ListObjects(TABLE_PURCHASE)

    poNo = Trim$(CStr(wsUI.Range(CELL_PO_ID).value))

    If poNo = "" Then
        MsgBox "Please load an existing Purchase ID before updating.", vbExclamation, "Update Purchase"
        Exit Sub
    End If

    If Not PurchaseExists(tblPO, poNo) Then
        MsgBox "Purchase ID not found in Purchase_DB.", vbExclamation, "Update Purchase"
        Exit Sub
    End If

    ' Safety rule:
    ' if any receiving already happened, do not allow editing PO lines
    If PurchaseHasReceivedQty(tblPO, poNo) Then
        MsgBox "This purchase already has receiving activity." & vbCrLf & vbCrLf & _
               "To protect inventory history, this PO cannot be updated anymore." & vbCrLf & _
               "Please keep this PO as history and create a new Purchase Order if changes are needed.", _
               vbExclamation, "Update Not Allowed"
        Exit Sub
    End If

    ValidatePurchaseReadyToSave

    If MsgBox("Update this purchase?" & vbCrLf & _
              "Purchase ID: " & poNo, vbQuestion + vbYesNo, "Confirm Update") <> vbYes Then
        Exit Sub
    End If

    Set oldRows = BackupPurchaseRows(tblPO, poNo)
    updateStarted = True

    DeletePurchaseRowsByID tblPO, poNo
    SavePurchaseUsingID poNo

    MsgBox "Purchase updated successfully.", vbInformation, "Update Purchase"
    Exit Sub

ErrHandler:
    If updateStarted Then
        On Error Resume Next
        DeletePurchaseRowsByID tblPO, poNo
        RestoreTableRowsFromCollection tblPO, oldRows
        On Error GoTo 0
    End If

    MsgBox "Error: " & Err.Description, vbCritical, "Update Purchase"

End Sub

'==================================================
' SHARED SAVE
'==================================================
Private Sub SavePurchaseUsingID(ByVal poNo As String)

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim wsSuppliers As Worksheet
    Dim wsProducts As Worksheet

    Dim tblPO As ListObject
    Dim tblSup As ListObject
    Dim tblProd As ListObject

    Dim purchaseDate As Variant
    Dim supplierName As String
    Dim supplierID As String
    Dim notes As String

    Dim i As Long
    Dim sku As String
    Dim productName As String
    Dim qty As Double
    Dim unitCost As Double
    Dim lineTotal As Double
    Dim productID As String
    Dim existingReceived As Double
    Dim remainingQty As Double

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets(SHEET_PURCHASE_DB)
    Set wsSuppliers = ThisWorkbook.Worksheets(SHEET_SUPPLIERS)
    Set wsProducts = ThisWorkbook.Worksheets(SHEET_PRODUCTS)

    Set tblPO = wsDB.ListObjects(TABLE_PURCHASE)
    Set tblSup = wsSuppliers.ListObjects(TABLE_SUPPLIERS)
    Set tblProd = wsProducts.ListObjects(TABLE_PRODUCTS)

    purchaseDate = wsUI.Range(CELL_PO_DATE).value
    supplierName = Trim$(CStr(wsUI.Range(CELL_SUPPLIER).value))
    notes = Trim$(CStr(wsUI.Range(CELL_NOTES).value))

    supplierID = FindPartyIDByName(tblSup, "Supplier_Name", "Supplier_ID", supplierName)
    If supplierID = "" Then
        Err.Raise vbObjectError + 4101, , "Supplier not found in Suppliers_DB."
    End If

    For i = START_ROW To END_ROW

        sku = Trim$(CStr(wsUI.Range(COL_SKU & i).value))
        If sku <> "" Then

            productName = Trim$(CStr(wsUI.Range(COL_NAME & i).value))
            qty = NzNumber(wsUI.Range(COL_QTY & i).value)
            unitCost = Round(NzNumber(wsUI.Range(COL_UNIT & i).value), 2)
            lineTotal = Round(qty * unitCost, 2)

            productID = FindPartyIDByName(tblProd, "SKU", "Product_ID", sku)
            If productID = "" Then
                Err.Raise vbObjectError + 4102, , "SKU not found in Products_DB: " & sku
            End If

            If productName = "" Then
                productName = FindTextByKey(tblProd, "SKU", sku, "Product_Name")
            End If

            existingReceived = 0
            remainingQty = qty

            With tblPO.ListRows.Add
                .Range(1, GetCol(tblPO, "Purchase_Order_No")).value = poNo
                .Range(1, GetCol(tblPO, "Purchase_Date")).value = purchaseDate
                .Range(1, GetCol(tblPO, "Supplier_ID")).value = supplierID
                .Range(1, GetCol(tblPO, "Supplier_Name")).value = supplierName
                .Range(1, GetCol(tblPO, "Product_ID")).value = productID
                .Range(1, GetCol(tblPO, "SKU")).value = sku
                .Range(1, GetCol(tblPO, "Product_Name")).value = productName
                .Range(1, GetCol(tblPO, "Qty")).value = qty
                .Range(1, GetCol(tblPO, "Received_Qty")).value = existingReceived
                .Range(1, GetCol(tblPO, "Remaining_Qty")).value = remainingQty
                .Range(1, GetCol(tblPO, "Unit_Cost")).value = unitCost
                .Range(1, GetCol(tblPO, "Line_Total")).value = lineTotal

                If HasColumn(tblPO, "Amount_Paid") Then
                    .Range(1, GetCol(tblPO, "Amount_Paid")).value = 0
                End If
                If HasColumn(tblPO, "Balance_Due") Then
                    .Range(1, GetCol(tblPO, "Balance_Due")).value = lineTotal
                End If
                If HasColumn(tblPO, "Payment_Status") Then
                    .Range(1, GetCol(tblPO, "Payment_Status")).value = "Unpaid"
                End If

                If HasColumn(tblPO, "Status") Then
                    .Range(1, GetCol(tblPO, "Status")).value = "Open"
                End If
                If HasColumn(tblPO, "Line_Status") Then
                    .Range(1, GetCol(tblPO, "Line_Status")).value = "Open"
                End If

                .Range(1, GetCol(tblPO, "Notes")).value = notes
                .Range(1, GetCol(tblPO, "Created_At")).value = Date
                If HasColumn(tblPO, "Updated_At") Then
                    .Range(1, GetCol(tblPO, "Updated_At")).value = Date
                End If
            End With

        End If
    Next i

    UpdatePurchaseSummaryUI poNo

End Sub

'==================================================
' VALIDATION
'==================================================
Private Sub ValidatePurchaseReadyToSave()

    Dim wsUI As Worksheet
    Dim wsSuppliers As Worksheet
    Dim wsProducts As Worksheet

    Dim tblSup As ListObject
    Dim tblProd As ListObject

    Dim supplierName As String
    Dim purchaseDate As Variant
    Dim hasLine As Boolean
    Dim i As Long
    Dim sku As String
    Dim qty As Double
    Dim unitCost As Double

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsSuppliers = ThisWorkbook.Worksheets(SHEET_SUPPLIERS)
    Set wsProducts = ThisWorkbook.Worksheets(SHEET_PRODUCTS)

    Set tblSup = wsSuppliers.ListObjects(TABLE_SUPPLIERS)
    Set tblProd = wsProducts.ListObjects(TABLE_PRODUCTS)

    purchaseDate = wsUI.Range(CELL_PO_DATE).value
    supplierName = Trim$(CStr(wsUI.Range(CELL_SUPPLIER).value))

    If Trim$(CStr(purchaseDate)) = "" Then
        Err.Raise vbObjectError + 4201, , "Please enter Purchase Date."
    End If

    If Not IsDate(purchaseDate) Then
        Err.Raise vbObjectError + 4202, , "Purchase Date is invalid."
    End If

    If supplierName = "" Then
        Err.Raise vbObjectError + 4203, , "Please enter Supplier Name."
    End If

    If FindPartyIDByName(tblSup, "Supplier_Name", "Supplier_ID", supplierName) = "" Then
        Err.Raise vbObjectError + 4204, , "Supplier not found in Suppliers_DB."
    End If

    hasLine = False

    For i = START_ROW To END_ROW
        sku = Trim$(CStr(wsUI.Range(COL_SKU & i).value))

        If sku <> "" Then
            hasLine = True

            If FindPartyIDByName(tblProd, "SKU", "Product_ID", sku) = "" Then
                Err.Raise vbObjectError + 4205, , "SKU not found in Products_DB: " & sku
            End If

            qty = NzNumber(wsUI.Range(COL_QTY & i).value)
            If qty <= 0 Then
                Err.Raise vbObjectError + 4206, , "Qty must be greater than 0 for line " & i & "."
            End If

            unitCost = NzNumber(wsUI.Range(COL_UNIT & i).value)
            If unitCost < 0 Then
                Err.Raise vbObjectError + 4207, , "Unit Price cannot be negative for line " & i & "."
            End If
        End If
    Next i

    If Not hasLine Then
        Err.Raise vbObjectError + 4208, , "Please enter at least one purchase line."
    End If

End Sub

'==================================================
' SUMMARY UI
'==================================================
Private Sub UpdatePurchaseSummaryUI(ByVal poNo As String)

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tblPO As ListObject

    Dim subtotal As Double
    Dim amountPaid As Double
    Dim balanceDue As Double
    Dim receivedStatus As String
    Dim paymentStatus As String

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets(SHEET_PURCHASE_DB)
    Set tblPO = wsDB.ListObjects(TABLE_PURCHASE)

    subtotal = GetPurchaseSubtotal(tblPO, poNo)
    amountPaid = GetPurchaseAmountPaid(tblPO, poNo)
    balanceDue = GetPurchaseBalanceDue(tblPO, poNo)
    receivedStatus = GetPurchaseReceivedStatus(tblPO, poNo)
    paymentStatus = GetPurchasePaymentStatus(tblPO, poNo)

    wsUI.Range(CELL_SUBTOTAL).value = Round(subtotal, 2)
    wsUI.Range(CELL_RECEIVED_STATUS).value = receivedStatus
    wsUI.Range(CELL_PAYMENT_STATUS).value = paymentStatus
    wsUI.Range(CELL_AMOUNT_PAID).value = Round(amountPaid, 2)
    wsUI.Range(CELL_BALANCE_DUE).value = Round(balanceDue, 2)

End Sub

'==================================================
' UI HELPERS
'==================================================
Private Sub ClearPurchaseForm(ByVal wsUI As Worksheet)

    Dim i As Long

    wsUI.Range(CELL_PO_ID).value = ""
    wsUI.Range(CELL_PO_DATE).value = ""
    wsUI.Range(CELL_SUPPLIER).value = ""
    wsUI.Range(CELL_NOTES).value = ""

    For i = START_ROW To END_ROW
        wsUI.Range(COL_SKU & i).value = ""
        wsUI.Range(COL_NAME & i).value = ""
        wsUI.Range(COL_QTY & i).value = ""
        wsUI.Range(COL_UNIT & i).value = ""
        wsUI.Range(COL_LINE_TOTAL & i).value = ""
    Next i

    wsUI.Range(CELL_SUBTOTAL).value = ""
    wsUI.Range(CELL_RECEIVED_STATUS).value = ""
    wsUI.Range(CELL_PAYMENT_STATUS).value = ""
    wsUI.Range(CELL_AMOUNT_PAID).value = ""
    wsUI.Range(CELL_BALANCE_DUE).value = ""

End Sub

Private Sub RestorePurchaseLineFormulas(ByVal wsUI As Worksheet)

    Dim i As Long

    For i = START_ROW To END_ROW
        wsUI.Range(COL_NAME & i).Formula = "=IFERROR(XLOOKUP(" & COL_SKU & i & ",tblProducts[SKU],tblProducts[Product_Name]),"""")"
        wsUI.Range(COL_UNIT & i).Formula = "=IFERROR(XLOOKUP(" & COL_SKU & i & ",tblProducts[SKU],tblProducts[Unit_Cost]),"""")"
        wsUI.Range(COL_LINE_TOTAL & i).Formula = "=IFERROR(" & COL_QTY & i & "*" & COL_UNIT & i & ",0)"
    Next i

End Sub

'==================================================
' PURCHASE STATUS HELPERS
'==================================================
Private Function GetPurchaseSubtotal(ByVal tblPO As ListObject, ByVal poNo As String) As Double

    Dim i As Long
    Dim totalAmt As Double

    totalAmt = 0

    If tblPO.DataBodyRange Is Nothing Then
        GetPurchaseSubtotal = 0
        Exit Function
    End If

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            totalAmt = totalAmt + NzNumber(tblPO.ListColumns("Line_Total").DataBodyRange.Cells(i, 1).value)
        End If
    Next i

    GetPurchaseSubtotal = Round(totalAmt, 2)

End Function

Private Function GetPurchaseAmountPaid(ByVal tblPO As ListObject, ByVal poNo As String) As Double

    Dim i As Long
    GetPurchaseAmountPaid = 0

    If tblPO.DataBodyRange Is Nothing Then Exit Function
    If Not HasColumn(tblPO, "Amount_Paid") Then Exit Function

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            GetPurchaseAmountPaid = Round(NzNumber(tblPO.ListColumns("Amount_Paid").DataBodyRange.Cells(i, 1).value), 2)
            Exit Function
        End If
    Next i

End Function

Private Function GetPurchaseBalanceDue(ByVal tblPO As ListObject, ByVal poNo As String) As Double

    Dim i As Long
    GetPurchaseBalanceDue = 0

    If tblPO.DataBodyRange Is Nothing Then Exit Function
    If Not HasColumn(tblPO, "Balance_Due") Then Exit Function

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            GetPurchaseBalanceDue = Round(NzNumber(tblPO.ListColumns("Balance_Due").DataBodyRange.Cells(i, 1).value), 2)
            Exit Function
        End If
    Next i

End Function

Private Function GetPurchasePaymentStatus(ByVal tblPO As ListObject, ByVal poNo As String) As String

    Dim i As Long
    GetPurchasePaymentStatus = ""

    If tblPO.DataBodyRange Is Nothing Then Exit Function
    If Not HasColumn(tblPO, "Payment_Status") Then Exit Function

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            GetPurchasePaymentStatus = Trim$(CStr(tblPO.ListColumns("Payment_Status").DataBodyRange.Cells(i, 1).value))
            Exit Function
        End If
    Next i

End Function

Private Function GetPurchaseReceivedStatus(ByVal tblPO As ListObject, ByVal poNo As String) As String

    Dim i As Long
    Dim allZero As Boolean
    Dim allClosed As Boolean
    Dim qty As Double
    Dim receivedQty As Double

    allZero = True
    allClosed = True

    If tblPO.DataBodyRange Is Nothing Then
        GetPurchaseReceivedStatus = ""
        Exit Function
    End If

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then

            qty = NzNumber(tblPO.ListColumns("Qty").DataBodyRange.Cells(i, 1).value)
            receivedQty = NzNumber(tblPO.ListColumns("Received_Qty").DataBodyRange.Cells(i, 1).value)

            If receivedQty > 0 Then allZero = False
            If receivedQty < qty Then allClosed = False
        End If
    Next i

    If allZero Then
        GetPurchaseReceivedStatus = "Open"
    ElseIf allClosed Then
        GetPurchaseReceivedStatus = "Closed"
    Else
        GetPurchaseReceivedStatus = "Partial"
    End If

End Function

Private Function PurchaseHasReceivedQty(ByVal tblPO As ListObject, ByVal poNo As String) As Boolean

    Dim i As Long
    PurchaseHasReceivedQty = False

    If tblPO.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            If NzNumber(tblPO.ListColumns("Received_Qty").DataBodyRange.Cells(i, 1).value) > 0 Then
                PurchaseHasReceivedQty = True
                Exit Function
            End If
        End If
    Next i

End Function

'==================================================
' DB HELPERS
'==================================================
Private Function GeneratePurchaseID(ByVal tbl As ListObject) As String

    Dim i As Long
    Dim s As String
    Dim n As Long
    Dim maxNum As Long

    maxNum = 0

    If tbl.ListRows.Count = 0 Then
        GeneratePurchaseID = "PO00001"
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        s = Trim$(CStr(tbl.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value))
        If Left$(s, 2) = "PO" Then
            If IsNumeric(Mid$(s, 3)) Then
                n = CLng(Mid$(s, 3))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i

    GeneratePurchaseID = "PO" & Format$(maxNum + 1, "00000")

End Function

Private Function PurchaseExists(ByVal tbl As ListObject, ByVal poNo As String) As Boolean

    Dim i As Long
    PurchaseExists = False

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            PurchaseExists = True
            Exit Function
        End If
    Next i

End Function

Private Function FindFirstPurchaseRow(ByVal tbl As ListObject, ByVal poNo As String) As Long

    Dim i As Long
    FindFirstPurchaseRow = 0

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            FindFirstPurchaseRow = i
            Exit Function
        End If
    Next i

End Function

Private Function BackupPurchaseRows(ByVal tbl As ListObject, ByVal poNo As String) As Collection

    Dim rowsBackup As New Collection
    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then
        Set BackupPurchaseRows = rowsBackup
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            rowsBackup.Add tbl.ListRows(i).Range.value
        End If
    Next i

    Set BackupPurchaseRows = rowsBackup

End Function

Private Sub DeletePurchaseRowsByID(ByVal tbl As ListObject, ByVal poNo As String)

    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For i = tbl.ListRows.Count To 1 Step -1
        If Trim$(CStr(tbl.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            tbl.ListRows(i).Delete
        End If
    Next i

End Sub

Private Sub RestoreTableRowsFromCollection(ByVal tbl As ListObject, ByVal rowsBackup As Collection)

    Dim i As Long
    Dim newRow As ListRow

    If rowsBackup Is Nothing Then Exit Sub
    If rowsBackup.Count = 0 Then Exit Sub

    For i = 1 To rowsBackup.Count
        Set newRow = tbl.ListRows.Add
        newRow.Range.value = rowsBackup(i)
    Next i

End Sub

'==================================================
' GENERIC HELPERS
'==================================================
Private Function FindPartyIDByName(ByVal tbl As ListObject, ByVal keyCol As String, ByVal idCol As String, ByVal keyValue As String) As String

    Dim i As Long
    FindPartyIDByName = ""

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns(keyCol).DataBodyRange.Cells(i, 1).value)) = keyValue Then
            FindPartyIDByName = Trim$(CStr(tbl.ListColumns(idCol).DataBodyRange.Cells(i, 1).value))
            Exit Function
        End If
    Next i

End Function

Private Function FindTextByKey(ByVal tbl As ListObject, ByVal keyCol As String, ByVal keyValue As String, ByVal returnCol As String) As String

    Dim i As Long
    FindTextByKey = ""

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns(keyCol).DataBodyRange.Cells(i, 1).value)) = keyValue Then
            FindTextByKey = Trim$(CStr(tbl.ListColumns(returnCol).DataBodyRange.Cells(i, 1).value))
            Exit Function
        End If
    Next i

End Function

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

