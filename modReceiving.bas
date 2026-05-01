Attribute VB_Name = "modReceiving"
Option Explicit

Private Const SHEET_UI As String = "Receiving_UI"
Private Const TABLE_RECEIVING As String = "tblReceiving"
Private Const TABLE_PRODUCTS As String = "tblProducts"
Private Const TABLE_SUPPLIERS As String = "tblSuppliers"
Private Const TABLE_LOG As String = "tblInventoryLog"
Private Const TABLE_PURCHASE As String = "tblPurchase"
Private Const SHEET_PWD As String = ""

Private Const NON_PO_LABEL As String = "NON-PO RECEIPT"

' UI Mapping
Private Const CELL_ID As String = "B3"
Private Const CELL_DATE As String = "B4"
Private Const CELL_SUPPLIER As String = "B5"
Private Const CELL_PO As String = "B6"
Private Const CELL_SKU As String = "B7"
Private Const CELL_NAME As String = "B8"
Private Const CELL_QTY As String = "B9"
Private Const CELL_UNIT_COST As String = "B10"
Private Const CELL_TOTAL_COST As String = "B11"
Private Const CELL_TRACKING As String = "B12"
Private Const CELL_NOTES As String = "B13"

'==================================================
' NEW RECEIVING
'==================================================
Public Sub Receiving_New()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    wsUI.Unprotect Password:=SHEET_PWD
    ClearReceivingForm wsUI
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True

End Sub

'==================================================
' SAVE NEW RECEIVING
'==================================================
Public Sub Receiving_Save()

    Dim wsUI As Worksheet
    Dim wsReceiving As Worksheet
    Dim tblR As ListObject
    Dim receivingID As String

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsReceiving = ThisWorkbook.Worksheets("Receiving_DB")
    Set tblR = wsReceiving.ListObjects(TABLE_RECEIVING)

    wsUI.Unprotect Password:=SHEET_PWD

    If Trim$(CStr(wsUI.Range(CELL_ID).value)) <> "" Then
        MsgBox "Receiving ID should be blank when creating a new receiving record." & vbCrLf & _
               "Click New Receiving first, or use Update Receiving for an existing record.", vbExclamation, "Create Receiving"
        GoTo SafeExit
    End If

    ValidateReceivingReadyToSave False, "", "", 0

    receivingID = GenerateReceivingID(tblR)
    wsUI.Range(CELL_ID).value = receivingID

    SaveReceivingUsingID receivingID

    MsgBox "Stock received successfully.", vbInformation, "Receiving Saved"

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Receiving Save"
    Resume SafeExit

End Sub

'==================================================
' LOAD RECEIVING
'==================================================
Public Sub Receiving_Load()

    Dim wsUI As Worksheet
    Dim wsReceiving As Worksheet
    Dim tblR As ListObject
    Dim receivingID As String
    Dim rowNum As Long

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsReceiving = ThisWorkbook.Worksheets("Receiving_DB")
    Set tblR = wsReceiving.ListObjects(TABLE_RECEIVING)

    receivingID = Trim$(CStr(wsUI.Range(CELL_ID).value))

    If receivingID = "" Then
        MsgBox "Please enter Receiving ID first.", vbExclamation, "Load Receiving"
        Exit Sub
    End If

    rowNum = FindReceivingRowByID(tblR, receivingID)
    If rowNum = 0 Then
        MsgBox "Receiving ID not found.", vbExclamation, "Load Receiving"
        Exit Sub
    End If

    wsUI.Unprotect Password:=SHEET_PWD
    ClearReceivingForm wsUI
    wsUI.Range(CELL_ID).value = receivingID

    wsUI.Range(CELL_DATE).value = tblR.ListColumns("Receiving_Date").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_SUPPLIER).value = tblR.ListColumns("Supplier_Name").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_PO).value = tblR.ListColumns("Purchase_Order_No").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_SKU).value = tblR.ListColumns("SKU").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_NAME).value = tblR.ListColumns("Product_Name").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_QTY).value = tblR.ListColumns("Quantity").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_UNIT_COST).value = tblR.ListColumns("Unit_Cost").DataBodyRange.Cells(rowNum, 1).value

    If HasColumn(tblR, "Tracking_No") Then
        wsUI.Range(CELL_TRACKING).value = tblR.ListColumns("Tracking_No").DataBodyRange.Cells(rowNum, 1).value
    Else
        wsUI.Range(CELL_TRACKING).value = ""
    End If

    wsUI.Range(CELL_NOTES).value = tblR.ListColumns("Notes").DataBodyRange.Cells(rowNum, 1).value

    RestoreReceivingFormulas wsUI

    MsgBox "Receiving record loaded successfully.", vbInformation, "Load Receiving"

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Receiving Load"
    Resume SafeExit

End Sub

'==================================================
' UPDATE RECEIVING
'==================================================
Public Sub Receiving_Update()

    Dim wsUI As Worksheet
    Dim wsReceiving As Worksheet
    Dim wsProducts As Worksheet
    Dim wsLog As Worksheet
    Dim wsPurchase As Worksheet

    Dim tblR As ListObject
    Dim tblP As ListObject
    Dim tblL As ListObject
    Dim tblPO As ListObject

    Dim receivingID As String
    Dim rowNum As Long

    Dim oldSKU As String
    Dim oldQty As Double
    Dim oldSupplier As String
    Dim oldUnitCost As Double
    Dim oldNotes As String
    Dim oldTracking As String
    Dim oldPO As String

    Dim newSKU As String
    Dim newQty As Double
    Dim newDate As Variant
    Dim newSupplier As String
    Dim newUnitCost As Double
    Dim newNotes As String
    Dim newTracking As String
    Dim newPO As String

    Dim stockFieldsChanged As Boolean
    Dim poFieldsChanged As Boolean

    Dim productRow As Long
    Dim currentStock As Double
    Dim rolledBackStock As Double
    Dim i As Long

    Dim oldReceivingRow As Variant
    Dim oldLogRows As Collection
    Dim stockSnapshot As Object
    Dim purchaseSnapshot As Object
    Dim updateStarted As Boolean

    Dim oldPORow As Long
    Dim newPORow As Long

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsReceiving = ThisWorkbook.Worksheets("Receiving_DB")
    Set wsProducts = ThisWorkbook.Worksheets("Products_DB")
    Set wsLog = ThisWorkbook.Worksheets("Inventory_Log")
    Set wsPurchase = ThisWorkbook.Worksheets("Purchase_DB")

    Set tblR = wsReceiving.ListObjects(TABLE_RECEIVING)
    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)
    Set tblL = wsLog.ListObjects(TABLE_LOG)
    Set tblPO = wsPurchase.ListObjects(TABLE_PURCHASE)

    wsUI.Unprotect Password:=SHEET_PWD

    receivingID = Trim$(CStr(wsUI.Range(CELL_ID).value))

    If receivingID = "" Then
        MsgBox "Please load an existing Receiving ID before updating.", vbExclamation, "Update Receiving"
        GoTo SafeExit
    End If

    rowNum = FindReceivingRowByID(tblR, receivingID)
    If rowNum = 0 Then
        MsgBox "Receiving ID not found in Receiving_DB.", vbExclamation, "Update Receiving"
        GoTo SafeExit
    End If

    oldSupplier = Trim$(CStr(tblR.ListColumns("Supplier_Name").DataBodyRange.Cells(rowNum, 1).value))
    oldPO = Trim$(CStr(tblR.ListColumns("Purchase_Order_No").DataBodyRange.Cells(rowNum, 1).value))
    oldSKU = Trim$(CStr(tblR.ListColumns("SKU").DataBodyRange.Cells(rowNum, 1).value))
    oldQty = NzNumber(tblR.ListColumns("Quantity").DataBodyRange.Cells(rowNum, 1).value)
    oldUnitCost = NzNumber(tblR.ListColumns("Unit_Cost").DataBodyRange.Cells(rowNum, 1).value)
    oldNotes = Trim$(CStr(tblR.ListColumns("Notes").DataBodyRange.Cells(rowNum, 1).value))

    If HasColumn(tblR, "Tracking_No") Then
        oldTracking = Trim$(CStr(tblR.ListColumns("Tracking_No").DataBodyRange.Cells(rowNum, 1).value))
    Else
        oldTracking = ""
    End If

    newDate = wsUI.Range(CELL_DATE).value
    newSupplier = Trim$(CStr(wsUI.Range(CELL_SUPPLIER).value))
    newPO = NormalizePOValueForSave(Trim$(CStr(wsUI.Range(CELL_PO).value)))
    newSKU = Trim$(CStr(wsUI.Range(CELL_SKU).value))
    newQty = NzNumber(wsUI.Range(CELL_QTY).value)
    newUnitCost = NzNumber(wsUI.Range(CELL_UNIT_COST).value)
    newTracking = Trim$(CStr(wsUI.Range(CELL_TRACKING).value))
    newNotes = Trim$(CStr(wsUI.Range(CELL_NOTES).value))

    ValidateReceivingReadyToSave True, oldPO, oldSKU, oldQty

    stockFieldsChanged = False
    If oldSKU <> newSKU Then stockFieldsChanged = True
    If oldQty <> newQty Then stockFieldsChanged = True

    poFieldsChanged = False
    If oldPO <> newPO Then poFieldsChanged = True

    If stockFieldsChanged Then
        productRow = FindProductRowBySKU(tblP, oldSKU)
        If productRow > 0 Then
            currentStock = NzNumber(tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value)
            rolledBackStock = currentStock - oldQty

            If rolledBackStock < 0 Then
                MsgBox "This receiving record cannot be updated for Quantity or SKU." & vbCrLf & vbCrLf & _
                       "Reason:" & vbCrLf & _
                       "This receipt has already affected later inventory activity. Changing Quantity or SKU now would make stock negative or break inventory history." & vbCrLf & vbCrLf & _
                       "Recommended action:" & vbCrLf & _
                       "Please keep this receiving record as history and use an Inventory Adjustment entry to correct the stock difference." & vbCrLf & vbCrLf & _
                       "Example:" & vbCrLf & _
                       "If received quantity was entered as 859 but should be 800, create an adjustment of -59.", _
                       vbExclamation, "Receiving Update Not Allowed"
                GoTo SafeExit
            End If
        End If
    End If

    If MsgBox("Update this receiving record?" & vbCrLf & _
              "Receiving ID: " & receivingID, vbQuestion + vbYesNo, "Confirm Update") <> vbYes Then
        GoTo SafeExit
    End If

    oldReceivingRow = tblR.ListRows(rowNum).Range.value
    Set oldLogRows = BackupInventoryLogRows(tblL, "RECEIVING", receivingID)
    Set stockSnapshot = CreateObject("Scripting.Dictionary")
    Set purchaseSnapshot = CreateObject("Scripting.Dictionary")

    SnapshotProductStockBySKU tblP, stockSnapshot, oldSKU
    SnapshotProductStockBySKU tblP, stockSnapshot, newSKU

    oldPORow = FindPurchaseRowByPOAndSKU(tblPO, oldPO, oldSKU)
    newPORow = FindPurchaseRowByPOAndSKU(tblPO, newPO, newSKU)

    SnapshotPurchaseRow tblPO, purchaseSnapshot, oldPORow
    SnapshotPurchaseRow tblPO, purchaseSnapshot, newPORow

    updateStarted = True

    If Not stockFieldsChanged And Not poFieldsChanged Then

        With tblR.ListRows(rowNum).Range
            .Cells(1, GetCol(tblR, "Receiving_Date")).value = newDate
            .Cells(1, GetCol(tblR, "Supplier_Name")).value = newSupplier
            .Cells(1, GetCol(tblR, "Purchase_Order_No")).value = newPO
            .Cells(1, GetCol(tblR, "Product_Name")).value = wsUI.Range(CELL_NAME).value
            .Cells(1, GetCol(tblR, "Unit_Cost")).value = newUnitCost
            .Cells(1, GetCol(tblR, "Total_Cost")).value = Round(newQty * newUnitCost, 2)
            .Cells(1, GetCol(tblR, "Notes")).value = newNotes

            If HasColumn(tblR, "Tracking_No") Then
                .Cells(1, GetCol(tblR, "Tracking_No")).value = newTracking
            End If

            If HasColumn(tblR, "Updated_At") Then
                .Cells(1, GetCol(tblR, "Updated_At")).value = Date
            End If
        End With

        UpdateReceivingSupplierID tblR, rowNum, newSupplier

        For i = tblL.ListRows.Count To 1 Step -1
            If Trim$(CStr(tblL.ListColumns("Tran_Type").DataBodyRange.Cells(i, 1).value)) = "RECEIVING" _
               And Trim$(CStr(tblL.ListColumns("Ref_No").DataBodyRange.Cells(i, 1).value)) = receivingID Then

                tblL.ListColumns("Log_Date").DataBodyRange.Cells(i, 1).value = newDate
                tblL.ListColumns("Unit_Cost").DataBodyRange.Cells(i, 1).value = newUnitCost
                tblL.ListColumns("Total_Cost").DataBodyRange.Cells(i, 1).value = Round(newQty * newUnitCost, 2)
                tblL.ListColumns("Party_Name").DataBodyRange.Cells(i, 1).value = newSupplier
                tblL.ListColumns("Notes").DataBodyRange.Cells(i, 1).value = newNotes

                UpdateInventoryLogSupplierID tblL, i, newSupplier
            End If
        Next i

        wsUI.Range(CELL_TOTAL_COST).Formula = "=IFERROR(" & CELL_QTY & "*" & CELL_UNIT_COST & ",0)"
        MsgBox "Receiving record updated successfully.", vbInformation, "Update Receiving"
        GoTo SafeExit

    End If

    productRow = FindProductRowBySKU(tblP, oldSKU)
    If productRow > 0 Then
        currentStock = NzNumber(tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value)
        rolledBackStock = currentStock - oldQty
        tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value = rolledBackStock
    End If

    RollbackReceiptFromPurchase tblPO, oldPO, oldSKU, oldQty

    DeleteInventoryLogRowsByRef tblL, "RECEIVING", receivingID
    tblR.ListRows(rowNum).Delete

    SaveReceivingUsingID receivingID

    MsgBox "Receiving record updated successfully.", vbInformation, "Update Receiving"

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    If updateStarted Then
        On Error Resume Next
        RestoreProductStocks tblP, stockSnapshot
        DeleteInventoryLogRowsByRef tblL, "RECEIVING", receivingID
        DeleteReceivingRowsByID tblR, receivingID
        RestoreTableRows tblR, oldReceivingRow
        RestoreTableRowsFromCollection tblL, oldLogRows
        RestorePurchaseSnapshots tblPO, purchaseSnapshot
        On Error GoTo 0
    End If

    MsgBox "Error: " & Err.Description, vbCritical, "Receiving Update"
    Resume SafeExit

End Sub

'==================================================
' SHARED SAVE LOGIC
'==================================================
Private Sub SaveReceivingUsingID(ByVal receivingID As String)

    Dim wsUI As Worksheet
    Dim wsReceiving As Worksheet
    Dim wsProducts As Worksheet
    Dim wsSuppliers As Worksheet
    Dim wsLog As Worksheet
    Dim wsPurchase As Worksheet

    Dim tblR As ListObject
    Dim tblP As ListObject
    Dim tblS As ListObject
    Dim tblL As ListObject
    Dim tblPO As ListObject

    Dim receivingDate As Variant
    Dim supplierName As String
    Dim supplierID As String
    Dim poValue As String
    Dim sku As String
    Dim productName As String
    Dim qty As Double
    Dim unitCost As Double
    Dim totalCost As Double
    Dim trackingNo As String
    Dim notes As String

    Dim productRow As Long
    Dim newRow As ListRow
    Dim productID As String
    Dim currentStock As Double
    Dim newStock As Double
    Dim i As Long

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsReceiving = ThisWorkbook.Worksheets("Receiving_DB")
    Set wsProducts = ThisWorkbook.Worksheets("Products_DB")
    Set wsSuppliers = ThisWorkbook.Worksheets("Suppliers_DB")
    Set wsLog = ThisWorkbook.Worksheets("Inventory_Log")
    Set wsPurchase = ThisWorkbook.Worksheets("Purchase_DB")

    Set tblR = wsReceiving.ListObjects(TABLE_RECEIVING)
    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)
    Set tblS = wsSuppliers.ListObjects(TABLE_SUPPLIERS)
    Set tblL = wsLog.ListObjects(TABLE_LOG)
    Set tblPO = wsPurchase.ListObjects(TABLE_PURCHASE)

    receivingDate = wsUI.Range(CELL_DATE).value
    supplierName = Trim$(CStr(wsUI.Range(CELL_SUPPLIER).value))
    poValue = NormalizePOValueForSave(Trim$(CStr(wsUI.Range(CELL_PO).value)))
    sku = Trim$(CStr(wsUI.Range(CELL_SKU).value))
    productName = Trim$(CStr(wsUI.Range(CELL_NAME).value))
    qty = NzNumber(wsUI.Range(CELL_QTY).value)
    unitCost = NzNumber(wsUI.Range(CELL_UNIT_COST).value)
    totalCost = Round(qty * unitCost, 2)
    trackingNo = Trim$(CStr(wsUI.Range(CELL_TRACKING).value))
    notes = Trim$(CStr(wsUI.Range(CELL_NOTES).value))

    wsUI.Range(CELL_TOTAL_COST).Formula = "=IFERROR(" & CELL_QTY & "*" & CELL_UNIT_COST & ",0)"

    supplierID = ""
    For i = 1 To tblS.ListRows.Count
        If Trim$(CStr(tblS.ListColumns("Supplier_Name").DataBodyRange.Cells(i, 1).value)) = supplierName Then
            supplierID = Trim$(CStr(tblS.ListColumns("Supplier_ID").DataBodyRange.Cells(i, 1).value))
            Exit For
        End If
    Next i

    If supplierID = "" Then
        Err.Raise vbObjectError + 1301, , "Supplier not found in Suppliers_DB."
    End If

    productRow = FindProductRowBySKU(tblP, sku)
    If productRow = 0 Then
        Err.Raise vbObjectError + 1302, , "Product SKU not found in Products_DB."
    End If

    If productName = "" Then
        productName = Trim$(CStr(tblP.ListColumns("Product_Name").DataBodyRange.Cells(productRow, 1).value))
        wsUI.Range(CELL_NAME).value = productName
    End If

    productID = Trim$(CStr(tblP.ListColumns("Product_ID").DataBodyRange.Cells(productRow, 1).value))
    currentStock = NzNumber(tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value)
    newStock = currentStock + qty

    Set newRow = tblR.ListRows.Add

    With newRow.Range
        .Cells(1, GetCol(tblR, "Receiving_ID")).value = receivingID
        .Cells(1, GetCol(tblR, "Receiving_Date")).value = receivingDate
        .Cells(1, GetCol(tblR, "Supplier_ID")).value = supplierID
        .Cells(1, GetCol(tblR, "Supplier_Name")).value = supplierName
        .Cells(1, GetCol(tblR, "Purchase_Order_No")).value = poValue
        .Cells(1, GetCol(tblR, "Product_ID")).value = productID
        .Cells(1, GetCol(tblR, "SKU")).value = sku
        .Cells(1, GetCol(tblR, "Product_Name")).value = productName
        .Cells(1, GetCol(tblR, "Quantity")).value = qty
        .Cells(1, GetCol(tblR, "Unit_Cost")).value = unitCost
        .Cells(1, GetCol(tblR, "Total_Cost")).value = totalCost
        .Cells(1, GetCol(tblR, "Notes")).value = notes

        If HasColumn(tblR, "Tracking_No") Then
            .Cells(1, GetCol(tblR, "Tracking_No")).value = trackingNo
        End If

        .Cells(1, GetCol(tblR, "Created_At")).value = Date
        If HasColumn(tblR, "Updated_At") Then
            .Cells(1, GetCol(tblR, "Updated_At")).value = Date
        End If
    End With

    tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value = newStock

    With tblL.ListRows.Add
        .Range(1, 1).value = GenerateLogID(tblL)
        .Range(1, 2).value = receivingDate
        .Range(1, 3).value = "RECEIVING"
        .Range(1, 4).value = receivingID
        .Range(1, 5).value = productID
        .Range(1, 6).value = sku
        .Range(1, 7).value = productName
        .Range(1, 8).value = qty
        .Range(1, 9).value = newStock
        .Range(1, 10).value = unitCost
        .Range(1, 11).value = totalCost
        .Range(1, 12).value = "SUPPLIER"
        .Range(1, 13).value = supplierID
        .Range(1, 14).value = supplierName
        .Range(1, 15).value = notes
        .Range(1, 16).value = Date
    End With

    ApplyReceiptToPurchase tblPO, poValue, sku, qty

End Sub

'==================================================
' VALIDATION
'==================================================
Private Sub ValidateReceivingForm()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    If Trim$(CStr(wsUI.Range(CELL_DATE).value)) = "" Then
        Err.Raise vbObjectError + 1311, , "Please enter Receiving Date."
    End If

    If Not IsDate(wsUI.Range(CELL_DATE).value) Then
        Err.Raise vbObjectError + 1312, , "Receiving Date is invalid."
    End If

    If Trim$(CStr(wsUI.Range(CELL_SUPPLIER).value)) = "" Then
        Err.Raise vbObjectError + 1313, , "Please enter Supplier Name."
    End If

    If Trim$(CStr(wsUI.Range(CELL_SKU).value)) = "" Then
        Err.Raise vbObjectError + 1314, , "Please enter Product SKU."
    End If

    If Not IsNumeric(wsUI.Range(CELL_QTY).value) Or NzNumber(wsUI.Range(CELL_QTY).value) <= 0 Then
        Err.Raise vbObjectError + 1315, , "Quantity must be greater than 0."
    End If

    If Not IsBlankOrNumeric(wsUI.Range(CELL_UNIT_COST).value) Then
        Err.Raise vbObjectError + 1316, , "Unit Cost must be numeric."
    End If

End Sub

Private Sub ValidateReceivingReadyToSave(ByVal isUpdate As Boolean, _
                                         ByVal oldPO As String, _
                                         ByVal oldSKU As String, _
                                         ByVal oldQty As Double)

    Dim wsUI As Worksheet
    Dim wsProducts As Worksheet
    Dim wsSuppliers As Worksheet
    Dim wsPurchase As Worksheet

    Dim tblP As ListObject
    Dim tblS As ListObject
    Dim tblPO As ListObject

    Dim supplierName As String
    Dim sku As String
    Dim poValue As String
    Dim productRow As Long
    Dim i As Long
    Dim supplierFound As Boolean
    Dim releaseQty As Double

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsProducts = ThisWorkbook.Worksheets("Products_DB")
    Set wsSuppliers = ThisWorkbook.Worksheets("Suppliers_DB")
    Set wsPurchase = ThisWorkbook.Worksheets("Purchase_DB")

    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)
    Set tblS = wsSuppliers.ListObjects(TABLE_SUPPLIERS)
    Set tblPO = wsPurchase.ListObjects(TABLE_PURCHASE)

    ValidateReceivingForm

    supplierName = Trim$(CStr(wsUI.Range(CELL_SUPPLIER).value))
    sku = Trim$(CStr(wsUI.Range(CELL_SKU).value))
    poValue = NormalizePOValueForSave(Trim$(CStr(wsUI.Range(CELL_PO).value)))

    supplierFound = False
    For i = 1 To tblS.ListRows.Count
        If Trim$(CStr(tblS.ListColumns("Supplier_Name").DataBodyRange.Cells(i, 1).value)) = supplierName Then
            supplierFound = True
            Exit For
        End If
    Next i

    If Not supplierFound Then
        Err.Raise vbObjectError + 1321, , "Supplier not found in Suppliers_DB."
    End If

    productRow = FindProductRowBySKU(tblP, sku)
    If productRow = 0 Then
        Err.Raise vbObjectError + 1322, , "Product SKU not found in Products_DB."
    End If

    If poValue = NON_PO_LABEL Then
        If MsgBox("No Purchase Order No was entered." & vbCrLf & vbCrLf & _
                  "Do you want to continue as a NON-PO RECEIPT?" & vbCrLf & _
                  "The system will save '" & NON_PO_LABEL & "' into Receiving_DB.", _
                  vbQuestion + vbYesNo, "Confirm Non-PO Receipt") <> vbYes Then
            Err.Raise vbObjectError + 1323, , "Receiving was cancelled because no PO was confirmed."
        End If
    Else
        releaseQty = 0
        If isUpdate Then
            If oldPO = poValue And oldSKU = sku Then releaseQty = oldQty
        End If

        ValidatePOReceipt tblPO, poValue, supplierName, sku, NzNumber(wsUI.Range(CELL_QTY).value), releaseQty
    End If

End Sub

'==================================================
' PO LINKAGE
'==================================================
Private Sub ValidatePOReceipt(ByVal tblPO As ListObject, _
                              ByVal poNo As String, _
                              ByVal supplierName As String, _
                              ByVal sku As String, _
                              ByVal receiptQty As Double, _
                              ByVal releaseQty As Double)

    Dim poExists As Boolean
    Dim rowNum As Long
    Dim i As Long
    Dim orderedQty As Double
    Dim receivedQty As Double
    Dim availableQty As Double
    Dim poSupplier As String

    poExists = False
    rowNum = 0

    If tblPO.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 1331, , "Purchase_DB is empty. PO cannot be validated."
    End If

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo Then
            poExists = True
            Exit For
        End If
    Next i

    If Not poExists Then
        Err.Raise vbObjectError + 1332, , "PO not found in Purchase_DB: " & poNo
    End If

    rowNum = FindPurchaseRowByPOAndSKU(tblPO, poNo, sku)
    If rowNum = 0 Then
        Err.Raise vbObjectError + 1333, , "This SKU does not exist under the selected PO." & vbCrLf & _
                                          "PO: " & poNo & vbCrLf & _
                                          "SKU: " & sku
    End If

    poSupplier = Trim$(CStr(tblPO.ListColumns("Supplier_Name").DataBodyRange.Cells(rowNum, 1).value))
    If poSupplier <> supplierName Then
        Err.Raise vbObjectError + 1334, , "Supplier does not match the selected PO." & vbCrLf & _
                                          "PO Supplier: " & poSupplier & vbCrLf & _
                                          "Receiving Supplier: " & supplierName
    End If

    orderedQty = NzNumber(tblPO.ListColumns("Qty").DataBodyRange.Cells(rowNum, 1).value)
    receivedQty = NzNumber(tblPO.ListColumns("Received_Qty").DataBodyRange.Cells(rowNum, 1).value)

    availableQty = orderedQty - receivedQty + releaseQty

    If receiptQty > availableQty Then
        Err.Raise vbObjectError + 1335, , "Receiving quantity exceeds remaining PO quantity." & vbCrLf & _
                                          "PO: " & poNo & vbCrLf & _
                                          "SKU: " & sku & vbCrLf & _
                                          "Available to receive: " & availableQty & vbCrLf & _
                                          "This receipt qty: " & receiptQty
    End If

End Sub

Private Sub ApplyReceiptToPurchase(ByVal tblPO As ListObject, ByVal poNo As String, ByVal sku As String, ByVal qty As Double)

    Dim rowNum As Long
    Dim orderedQty As Double
    Dim receivedQty As Double
    Dim remainingQty As Double
    Dim statusText As String

    If poNo = NON_PO_LABEL Then Exit Sub
    If Trim$(poNo) = "" Then Exit Sub

    rowNum = FindPurchaseRowByPOAndSKU(tblPO, poNo, sku)
    If rowNum = 0 Then
        Err.Raise vbObjectError + 1341, , "Cannot update Purchase_DB. Matching PO/SKU not found."
    End If

    orderedQty = NzNumber(tblPO.ListColumns("Qty").DataBodyRange.Cells(rowNum, 1).value)
    receivedQty = NzNumber(tblPO.ListColumns("Received_Qty").DataBodyRange.Cells(rowNum, 1).value) + qty
    remainingQty = orderedQty - receivedQty
    If remainingQty < 0 Then remainingQty = 0

    statusText = GetPOStatusText(orderedQty, receivedQty)

    tblPO.ListColumns("Received_Qty").DataBodyRange.Cells(rowNum, 1).value = receivedQty
    tblPO.ListColumns("Remaining_Qty").DataBodyRange.Cells(rowNum, 1).value = remainingQty

    If HasColumn(tblPO, "Status") Then
        tblPO.ListColumns("Status").DataBodyRange.Cells(rowNum, 1).value = statusText
    End If
    If HasColumn(tblPO, "Line_Status") Then
        tblPO.ListColumns("Line_Status").DataBodyRange.Cells(rowNum, 1).value = statusText
    End If
    If HasColumn(tblPO, "Updated_At") Then
        tblPO.ListColumns("Updated_At").DataBodyRange.Cells(rowNum, 1).value = Date
    End If

End Sub

Private Sub RollbackReceiptFromPurchase(ByVal tblPO As ListObject, ByVal poNo As String, ByVal sku As String, ByVal qty As Double)

    Dim rowNum As Long
    Dim orderedQty As Double
    Dim receivedQty As Double
    Dim remainingQty As Double
    Dim statusText As String

    If poNo = NON_PO_LABEL Then Exit Sub
    If Trim$(poNo) = "" Then Exit Sub

    rowNum = FindPurchaseRowByPOAndSKU(tblPO, poNo, sku)
    If rowNum = 0 Then Exit Sub

    orderedQty = NzNumber(tblPO.ListColumns("Qty").DataBodyRange.Cells(rowNum, 1).value)
    receivedQty = NzNumber(tblPO.ListColumns("Received_Qty").DataBodyRange.Cells(rowNum, 1).value) - qty
    If receivedQty < 0 Then receivedQty = 0

    remainingQty = orderedQty - receivedQty
    If remainingQty < 0 Then remainingQty = 0

    statusText = GetPOStatusText(orderedQty, receivedQty)

    tblPO.ListColumns("Received_Qty").DataBodyRange.Cells(rowNum, 1).value = receivedQty
    tblPO.ListColumns("Remaining_Qty").DataBodyRange.Cells(rowNum, 1).value = remainingQty

    If HasColumn(tblPO, "Status") Then
        tblPO.ListColumns("Status").DataBodyRange.Cells(rowNum, 1).value = statusText
    End If
    If HasColumn(tblPO, "Line_Status") Then
        tblPO.ListColumns("Line_Status").DataBodyRange.Cells(rowNum, 1).value = statusText
    End If
    If HasColumn(tblPO, "Updated_At") Then
        tblPO.ListColumns("Updated_At").DataBodyRange.Cells(rowNum, 1).value = Date
    End If

End Sub

Private Function GetPOStatusText(ByVal orderedQty As Double, ByVal receivedQty As Double) As String
    If receivedQty <= 0 Then
        GetPOStatusText = "Open"
    ElseIf receivedQty < orderedQty Then
        GetPOStatusText = "Partial"
    Else
        GetPOStatusText = "Closed"
    End If
End Function

Private Function FindPurchaseRowByPOAndSKU(ByVal tblPO As ListObject, ByVal poNo As String, ByVal sku As String) As Long

    Dim i As Long
    FindPurchaseRowByPOAndSKU = 0

    If tblPO.DataBodyRange Is Nothing Then Exit Function
    If poNo = NON_PO_LABEL Then Exit Function
    If Trim$(poNo) = "" Then Exit Function

    For i = 1 To tblPO.ListRows.Count
        If Trim$(CStr(tblPO.ListColumns("Purchase_Order_No").DataBodyRange.Cells(i, 1).value)) = poNo _
           And Trim$(CStr(tblPO.ListColumns("SKU").DataBodyRange.Cells(i, 1).value)) = sku Then
            FindPurchaseRowByPOAndSKU = i
            Exit Function
        End If
    Next i

End Function

Private Function NormalizePOValueForSave(ByVal poValue As String) As String
    If Trim$(poValue) = "" Then
        NormalizePOValueForSave = NON_PO_LABEL
    Else
        NormalizePOValueForSave = Trim$(poValue)
    End If
End Function

'==================================================
' SUPPORTING HELPERS
'==================================================
Private Sub UpdateReceivingSupplierID(ByVal tblR As ListObject, ByVal rowNum As Long, ByVal supplierName As String)

    Dim tblS As ListObject
    Dim i As Long
    Dim supplierID As String

    Set tblS = ThisWorkbook.Worksheets("Suppliers_DB").ListObjects(TABLE_SUPPLIERS)

    supplierID = ""
    For i = 1 To tblS.ListRows.Count
        If Trim$(CStr(tblS.ListColumns("Supplier_Name").DataBodyRange.Cells(i, 1).value)) = supplierName Then
            supplierID = Trim$(CStr(tblS.ListColumns("Supplier_ID").DataBodyRange.Cells(i, 1).value))
            Exit For
        End If
    Next i

    If supplierID <> "" Then
        tblR.DataBodyRange.Rows(rowNum).Cells(1, GetCol(tblR, "Supplier_ID")).value = supplierID
    End If

End Sub

Private Sub UpdateInventoryLogSupplierID(ByVal tblL As ListObject, ByVal rowNum As Long, ByVal supplierName As String)

    Dim tblS As ListObject
    Dim i As Long
    Dim supplierID As String

    Set tblS = ThisWorkbook.Worksheets("Suppliers_DB").ListObjects(TABLE_SUPPLIERS)

    supplierID = ""
    For i = 1 To tblS.ListRows.Count
        If Trim$(CStr(tblS.ListColumns("Supplier_Name").DataBodyRange.Cells(i, 1).value)) = supplierName Then
            supplierID = Trim$(CStr(tblS.ListColumns("Supplier_ID").DataBodyRange.Cells(i, 1).value))
            Exit For
        End If
    Next i

    If supplierID <> "" Then
        tblL.ListColumns("Party_ID").DataBodyRange.Cells(rowNum, 1).value = supplierID
    End If

End Sub

Private Sub RestoreReceivingFormulas(ByVal wsUI As Worksheet)

    wsUI.Range(CELL_NAME).Formula = "=IFERROR(XLOOKUP(" & CELL_SKU & ",tblProducts[SKU],tblProducts[Product_Name]),"""")"
    wsUI.Range(CELL_UNIT_COST).Formula = "=IFERROR(XLOOKUP(" & CELL_SKU & ",tblProducts[SKU],tblProducts[Unit_Cost]),"""")"
    wsUI.Range(CELL_TOTAL_COST).Formula = "=IFERROR(" & CELL_QTY & "*" & CELL_UNIT_COST & ",0)"

End Sub

Private Sub ClearReceivingForm(ByVal wsUI As Worksheet)

    wsUI.Range(CELL_ID).value = ""
    wsUI.Range(CELL_DATE).value = ""
    wsUI.Range(CELL_SUPPLIER).value = ""
    wsUI.Range(CELL_PO).value = ""
    wsUI.Range(CELL_SKU).value = ""

    wsUI.Range(CELL_NAME).Formula = ""
    wsUI.Range(CELL_QTY).value = ""
    wsUI.Range(CELL_UNIT_COST).Formula = ""
    wsUI.Range(CELL_TOTAL_COST).Formula = ""
    wsUI.Range(CELL_TRACKING).value = ""
    wsUI.Range(CELL_NOTES).value = ""

    RestoreReceivingFormulas wsUI

End Sub

Private Function GenerateReceivingID(ByVal tbl As ListObject) As String

    Dim i As Long, s As String, n As Long, maxNum As Long

    maxNum = 0

    If tbl.ListRows.Count = 0 Then
        GenerateReceivingID = "R00001"
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        s = Trim$(CStr(tbl.ListColumns("Receiving_ID").DataBodyRange.Cells(i, 1).value))
        If Left$(s, 1) = "R" Then
            If IsNumeric(Mid$(s, 2)) Then
                n = CLng(Mid$(s, 2))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i

    GenerateReceivingID = "R" & Format$(maxNum + 1, "00000")

End Function

Private Function FindReceivingRowByID(ByVal tbl As ListObject, ByVal receivingID As String) As Long

    Dim i As Long
    FindReceivingRowByID = 0

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("Receiving_ID").DataBodyRange.Cells(i, 1).value)) = receivingID Then
            FindReceivingRowByID = i
            Exit Function
        End If
    Next i

End Function

Private Function FindProductRowBySKU(ByVal tbl As ListObject, ByVal sku As String) As Long

    Dim i As Long
    FindProductRowBySKU = 0

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("SKU").DataBodyRange.Cells(i, 1).value)) = sku Then
            FindProductRowBySKU = i
            Exit Function
        End If
    Next i

End Function

Private Function GenerateLogID(ByVal tbl As ListObject) As String

    Dim maxNum As Long
    Dim i As Long
    Dim s As String
    Dim n As Long

    maxNum = 0

    If tbl.ListRows.Count = 0 Then
        GenerateLogID = "L00001"
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        s = Trim$(CStr(tbl.ListColumns("Log_ID").DataBodyRange.Cells(i, 1).value))
        If Left$(s, 1) = "L" Then
            If IsNumeric(Mid$(s, 2)) Then
                n = CLng(Mid$(s, 2))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i

    GenerateLogID = "L" & Format$(maxNum + 1, "00000")

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

Private Function IsBlankOrNumeric(ByVal v As Variant) As Boolean
    If Trim$(CStr(v)) = "" Then
        IsBlankOrNumeric = True
    ElseIf IsNumeric(v) Then
        IsBlankOrNumeric = True
    Else
        IsBlankOrNumeric = False
    End If
End Function

Private Function BackupInventoryLogRows(ByVal tbl As ListObject, ByVal tranType As String, ByVal refNo As String) As Collection
    Dim rowsBackup As New Collection
    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then
        Set BackupInventoryLogRows = rowsBackup
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("Tran_Type").DataBodyRange.Cells(i, 1).value)) = tranType _
           And Trim$(CStr(tbl.ListColumns("Ref_No").DataBodyRange.Cells(i, 1).value)) = refNo Then
            rowsBackup.Add tbl.ListRows(i).Range.value
        End If
    Next i

    Set BackupInventoryLogRows = rowsBackup
End Function

Private Sub SnapshotProductStockBySKU(ByVal tblP As ListObject, ByVal stockSnapshot As Object, ByVal sku As String)
    Dim rowNum As Long

    sku = Trim$(sku)
    If sku = "" Then Exit Sub
    If stockSnapshot.Exists(sku) Then Exit Sub

    rowNum = FindProductRowBySKU(tblP, sku)
    If rowNum > 0 Then
        stockSnapshot(sku) = NzNumber(tblP.ListColumns("Current_Stock").DataBodyRange.Cells(rowNum, 1).value)
    End If
End Sub

Private Sub RestoreProductStocks(ByVal tblP As ListObject, ByVal stockSnapshot As Object)
    Dim sku As Variant
    Dim rowNum As Long

    If stockSnapshot Is Nothing Then Exit Sub

    For Each sku In stockSnapshot.Keys
        rowNum = FindProductRowBySKU(tblP, CStr(sku))
        If rowNum > 0 Then
            tblP.ListColumns("Current_Stock").DataBodyRange.Cells(rowNum, 1).value = stockSnapshot(sku)
        End If
    Next sku
End Sub

Private Sub DeleteInventoryLogRowsByRef(ByVal tbl As ListObject, ByVal tranType As String, ByVal refNo As String)
    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For i = tbl.ListRows.Count To 1 Step -1
        If Trim$(CStr(tbl.ListColumns("Tran_Type").DataBodyRange.Cells(i, 1).value)) = tranType _
           And Trim$(CStr(tbl.ListColumns("Ref_No").DataBodyRange.Cells(i, 1).value)) = refNo Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

Private Sub DeleteReceivingRowsByID(ByVal tbl As ListObject, ByVal receivingID As String)
    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For i = tbl.ListRows.Count To 1 Step -1
        If Trim$(CStr(tbl.ListColumns("Receiving_ID").DataBodyRange.Cells(i, 1).value)) = receivingID Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

Private Sub RestoreTableRows(ByVal tbl As ListObject, ByVal rowData As Variant)
    Dim newRow As ListRow

    If IsEmpty(rowData) Then Exit Sub

    Set newRow = tbl.ListRows.Add
    newRow.Range.value = rowData
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

Private Sub SnapshotPurchaseRow(ByVal tblPO As ListObject, ByVal snapshot As Object, ByVal rowNum As Long)
    If rowNum <= 0 Then Exit Sub
    If snapshot.Exists(CStr(rowNum)) Then Exit Sub
    snapshot.Add CStr(rowNum), tblPO.ListRows(rowNum).Range.value
End Sub

Private Sub RestorePurchaseSnapshots(ByVal tblPO As ListObject, ByVal snapshot As Object)
    Dim k As Variant

    If snapshot Is Nothing Then Exit Sub

    For Each k In snapshot.Keys
        If CLng(k) > 0 And CLng(k) <= tblPO.ListRows.Count Then
            tblPO.ListRows(CLng(k)).Range.value = snapshot(k)
        End If
    Next k
End Sub

