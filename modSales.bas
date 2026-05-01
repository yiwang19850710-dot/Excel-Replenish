Attribute VB_Name = "modSales"
Option Explicit

Private Const SHEET_UI As String = "Sales_UI"

Private Const TABLE_PRODUCTS As String = "tblProducts"
Private Const TABLE_CUSTOMERS As String = "tblCustomers"
Private Const TABLE_SALES As String = "tblSales"
Private Const TABLE_LOG As String = "tblInventoryLog"

Private Const SHEET_PWD As String = ""

' Header UI Mapping
Private Const CELL_ID As String = "B3"
Private Const CELL_DATE As String = "B4"
Private Const CELL_CUSTOMER As String = "B5"
Private Const CELL_NOTES As String = "B6"

' Sales line area
Private Const FIRST_LINE_ROW As Long = 9
Private Const LAST_LINE_ROW As Long = 18

Private Const COL_SKU As String = "A"
Private Const COL_NAME As String = "B"
Private Const COL_QTY As String = "C"
Private Const COL_PRICE As String = "D"
Private Const COL_TOTAL As String = "E"

Private Const CELL_SUBTOTAL As String = "E20"

'==================================================
' CREATE NEW SALES
'==================================================
Public Sub Sales_Save()

    Dim wsUI As Worksheet
    Dim wsSales As Worksheet
    Dim tblS As ListObject
    Dim salesID As String

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsSales = ThisWorkbook.Worksheets("Sales_DB")
    Set tblS = wsSales.ListObjects(TABLE_SALES)

    wsUI.Unprotect Password:=SHEET_PWD

    If Trim$(CStr(wsUI.Range(CELL_ID).value)) <> "" Then
        MsgBox "Sales ID should be blank when creating a new sales order." & vbCrLf & _
               "Click New Sales first, or use Update Sales for an existing order.", vbExclamation
        GoTo SafeExit
    End If

    ValidateSalesReadyToSave False

    salesID = GenerateSalesID(tblS)
    wsUI.Range(CELL_ID).value = salesID

    SaveSalesUsingID salesID

    MsgBox "Sales order saved successfully.", vbInformation
    ClearSalesForm wsUI

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' NEW SALES
'==================================================
Public Sub Sales_New()

    Dim wsUI As Worksheet

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    wsUI.Unprotect Password:=SHEET_PWD
    ClearSalesForm wsUI
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True

End Sub

'==================================================
' LOAD SALES
'==================================================
Public Sub Sales_Load()

    Dim wsUI As Worksheet
    Dim wsSales As Worksheet
    Dim tblS As ListObject

    Dim salesID As String
    Dim i As Long
    Dim lineRow As Long
    Dim found As Boolean

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsSales = ThisWorkbook.Worksheets("Sales_DB")
    Set tblS = wsSales.ListObjects(TABLE_SALES)

    salesID = Trim$(CStr(wsUI.Range(CELL_ID).value))

    If salesID = "" Then
        MsgBox "Please enter Sales ID first.", vbExclamation
        Exit Sub
    End If

    wsUI.Unprotect Password:=SHEET_PWD
    ClearSalesForm wsUI
    wsUI.Range(CELL_ID).value = salesID

    lineRow = FIRST_LINE_ROW
    found = False

    If Not tblS.DataBodyRange Is Nothing Then
        For i = 1 To tblS.ListRows.Count
            If Trim$(CStr(tblS.ListColumns("Sales_Order_No").DataBodyRange.Cells(i, 1).value)) = salesID Then

                found = True

                If lineRow > LAST_LINE_ROW Then
                    MsgBox "This sales order has more lines than the current UI supports.", vbExclamation
                    Exit For
                End If

                If Trim$(CStr(wsUI.Range(CELL_DATE).value)) = "" Then
                    wsUI.Range(CELL_DATE).value = tblS.ListColumns("Sales_Date").DataBodyRange.Cells(i, 1).value
                    wsUI.Range(CELL_CUSTOMER).value = tblS.ListColumns("Customer_Name").DataBodyRange.Cells(i, 1).value
                    wsUI.Range(CELL_NOTES).value = tblS.ListColumns("Notes").DataBodyRange.Cells(i, 1).value
                End If

                wsUI.Cells(lineRow, COL_SKU).value = tblS.ListColumns("SKU").DataBodyRange.Cells(i, 1).value
                wsUI.Cells(lineRow, COL_NAME).value = tblS.ListColumns("Product_Name").DataBodyRange.Cells(i, 1).value
                wsUI.Cells(lineRow, COL_QTY).value = tblS.ListColumns("Qty").DataBodyRange.Cells(i, 1).value
                wsUI.Cells(lineRow, COL_PRICE).value = tblS.ListColumns("Unit_Price").DataBodyRange.Cells(i, 1).value
                wsUI.Cells(lineRow, COL_TOTAL).value = tblS.ListColumns("Line_Total").DataBodyRange.Cells(i, 1).value

                lineRow = lineRow + 1
            End If
        Next i
    End If

    If Not found Then
        MsgBox "Sales ID not found in Sales_DB.", vbExclamation
    Else
        MsgBox "Sales order loaded successfully.", vbInformation
    End If

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' UPDATE SALES (SAFE VERSION)
'==================================================
Public Sub Sales_Update()

    Dim wsUI As Worksheet
    Dim wsSales As Worksheet
    Dim wsProducts As Worksheet
    Dim wsLog As Worksheet

    Dim tblS As ListObject
    Dim tblP As ListObject
    Dim tblL As ListObject

    Dim salesID As String
    Dim oldSalesRows As Collection
    Dim oldLogRows As Collection
    Dim impactedSKUs As Object
    Dim stockSnapshot As Object
    Dim updateStarted As Boolean

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsSales = ThisWorkbook.Worksheets("Sales_DB")
    Set wsProducts = ThisWorkbook.Worksheets("Products_DB")
    Set wsLog = ThisWorkbook.Worksheets("Inventory_Log")

    Set tblS = wsSales.ListObjects(TABLE_SALES)
    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)
    Set tblL = wsLog.ListObjects(TABLE_LOG)

    wsUI.Unprotect Password:=SHEET_PWD

    salesID = Trim$(CStr(wsUI.Range(CELL_ID).value))

    If salesID = "" Then
        MsgBox "Please load an existing Sales ID before updating.", vbExclamation
        GoTo SafeExit
    End If

    If Not SalesIDExists(tblS, salesID) Then
        MsgBox "Sales ID not found in Sales_DB.", vbExclamation
        GoTo SafeExit
    End If

    ' 1) Validate current UI first
    ValidateSalesReadyToSave True

    If MsgBox("Update this sales order?" & vbCrLf & _
              "Sales ID: " & salesID, vbQuestion + vbYesNo, "Confirm Update") <> vbYes Then
        GoTo SafeExit
    End If

    ' 2) Snapshot current state before any mutation
    Set oldSalesRows = BackupSalesRows(tblS, salesID)
    Set oldLogRows = BackupInventoryLogRows_Sales(tblL, "SALE", salesID)
    Set impactedSKUs = CreateObject("Scripting.Dictionary")
    Set stockSnapshot = CreateObject("Scripting.Dictionary")

    CollectOldSalesSKUs tblS, salesID, impactedSKUs
    CollectUISalesSKUs wsUI, impactedSKUs
    SnapshotSalesStocks tblP, impactedSKUs, stockSnapshot

    updateStarted = True

    ' 3) Safe order: rollback stock -> delete old logs -> delete old rows -> write new rows
    RestoreStockFromOldSales tblS, tblP, salesID
    DeleteInventoryLogRowsByRef_Sales tblL, "SALE", salesID
    DeleteSalesRowsByID tblS, salesID

    SaveSalesUsingID salesID

    MsgBox "Sales order updated successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    If updateStarted Then
        On Error Resume Next
        DeleteInventoryLogRowsByRef_Sales tblL, "SALE", salesID
        DeleteSalesRowsByID tblS, salesID
        RestoreTableRowsFromCollection_Sales tblS, oldSalesRows
        RestoreTableRowsFromCollection_Sales tblL, oldLogRows
        RestoreProductStocks_Sales tblP, stockSnapshot
        On Error GoTo 0
    End If

    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' PRE-VALIDATION LAYER
'==================================================
Private Sub ValidateSalesReadyToSave(ByVal isUpdate As Boolean)

    Dim wsUI As Worksheet
    Dim wsProducts As Worksheet
    Dim wsCustomers As Worksheet
    Dim wsSales As Worksheet

    Dim tblP As ListObject
    Dim tblC As ListObject
    Dim tblS As ListObject

    Dim r As Long
    Dim i As Long
    Dim hasLine As Boolean

    Dim customerName As String
    Dim customerFound As Boolean
    Dim sku As String
    Dim productName As String
    Dim qty As Double
    Dim availableStock As Double
    Dim salesID As String
    Dim productRow As Long

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsProducts = ThisWorkbook.Worksheets("Products_DB")
    Set wsCustomers = ThisWorkbook.Worksheets("Customers_DB")
    Set wsSales = ThisWorkbook.Worksheets("Sales_DB")

    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)
    Set tblC = wsCustomers.ListObjects(TABLE_CUSTOMERS)
    Set tblS = wsSales.ListObjects(TABLE_SALES)

    If Trim$(CStr(wsUI.Range(CELL_CUSTOMER).value)) = "" Then
        Err.Raise vbObjectError + 1401, , "Please select Customer Name."
    End If

    If Trim$(CStr(wsUI.Range(CELL_DATE).value)) = "" Then
        Err.Raise vbObjectError + 1402, , "Please enter Sales Date."
    End If

    If Not IsDate(wsUI.Range(CELL_DATE).value) Then
        Err.Raise vbObjectError + 1403, , "Sales Date is invalid."
    End If

    customerName = Trim$(CStr(wsUI.Range(CELL_CUSTOMER).value))
    customerFound = False

    If Not tblC.DataBodyRange Is Nothing Then
        For i = 1 To tblC.ListRows.Count
            If Trim$(CStr(tblC.ListColumns("Customer_Name").DataBodyRange.Cells(i, 1).value)) = customerName Then
                customerFound = True
                Exit For
            End If
        Next i
    End If

    If Not customerFound Then
        Err.Raise vbObjectError + 1404, , "Customer not found in Customers_DB."
    End If

    hasLine = False
    For r = FIRST_LINE_ROW To LAST_LINE_ROW
        sku = Trim$(CStr(wsUI.Cells(r, COL_SKU).value))
        If sku <> "" Then
            hasLine = True
            Exit For
        End If
    Next r

    If Not hasLine Then
        Err.Raise vbObjectError + 1405, , "Please enter at least one sales line."
    End If

    salesID = Trim$(CStr(wsUI.Range(CELL_ID).value))

    For r = FIRST_LINE_ROW To LAST_LINE_ROW

        sku = Trim$(CStr(wsUI.Cells(r, COL_SKU).value))

        If sku <> "" Then

            If Not IsNumeric(wsUI.Cells(r, COL_QTY).value) Or CDbl(wsUI.Cells(r, COL_QTY).value) <= 0 Then
                Err.Raise vbObjectError + 1406, , "Row " & r & ": Quantity must be greater than 0."
            End If

            If Not IsNumeric(wsUI.Cells(r, COL_PRICE).value) Or CDbl(wsUI.Cells(r, COL_PRICE).value) < 0 Then
                Err.Raise vbObjectError + 1407, , "Row " & r & ": Unit Price must be valid."
            End If

            qty = CDbl(wsUI.Cells(r, COL_QTY).value)
            productName = Trim$(CStr(wsUI.Cells(r, COL_NAME).value))

            productRow = FindProductRowBySKU(tblP, sku)
            If productRow = 0 Then
                Err.Raise vbObjectError + 1408, , "Row " & r & ": Product SKU not found in Products_DB."
            End If

            availableStock = GetAvailableStockForSales(tblP, tblS, sku, salesID, isUpdate)

            If qty > availableStock Then
                Err.Raise vbObjectError + 1409, , _
                    "Row " & r & ": Insufficient stock." & vbCrLf & _
                    "SKU: " & sku & vbCrLf & _
                    "Available stock: " & availableStock & vbCrLf & _
                    "Requested quantity: " & qty
            End If

            If productName = "" Then
                wsUI.Cells(r, COL_NAME).value = Trim$(CStr(tblP.ListColumns("Product_Name").DataBodyRange.Cells(productRow, 1).value))
            End If

            wsUI.Cells(r, COL_TOTAL).value = Round(qty * CDbl(wsUI.Cells(r, COL_PRICE).value), 2)

        End If
    Next r

End Sub

'==================================================
' SHARED SAVE LOGIC
'==================================================
Private Sub SaveSalesUsingID(ByVal salesID As String)

    Dim wsUI As Worksheet
    Dim wsProducts As Worksheet
    Dim wsCustomers As Worksheet
    Dim wsSales As Worksheet
    Dim wsLog As Worksheet

    Dim tblP As ListObject
    Dim tblC As ListObject
    Dim tblS As ListObject
    Dim tblL As ListObject

    Dim r As Long
    Dim i As Long

    Dim salesDate As Variant
    Dim customerName As String
    Dim customerID As String
    Dim notes As String

    Dim sku As String
    Dim productName As String
    Dim qty As Double
    Dim unitPrice As Double
    Dim lineTotal As Double

    Dim productRow As Long
    Dim currentStock As Double
    Dim newStock As Double
    Dim productID As String

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsProducts = ThisWorkbook.Worksheets("Products_DB")
    Set wsCustomers = ThisWorkbook.Worksheets("Customers_DB")
    Set wsSales = ThisWorkbook.Worksheets("Sales_DB")
    Set wsLog = ThisWorkbook.Worksheets("Inventory_Log")

    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)
    Set tblC = wsCustomers.ListObjects(TABLE_CUSTOMERS)
    Set tblS = wsSales.ListObjects(TABLE_SALES)
    Set tblL = wsLog.ListObjects(TABLE_LOG)

    salesDate = wsUI.Range(CELL_DATE).value
    customerName = Trim$(CStr(wsUI.Range(CELL_CUSTOMER).value))
    notes = Trim$(CStr(wsUI.Range(CELL_NOTES).value))

    customerID = ""
    If Not tblC.DataBodyRange Is Nothing Then
        For i = 1 To tblC.ListRows.Count
            If Trim$(CStr(tblC.ListColumns("Customer_Name").DataBodyRange.Cells(i, 1).value)) = customerName Then
                customerID = Trim$(CStr(tblC.ListColumns("Customer_ID").DataBodyRange.Cells(i, 1).value))
                Exit For
            End If
        Next i
    End If

    If customerID = "" Then
        Err.Raise vbObjectError + 1411, , "Customer not found in Customers_DB."
    End If

    For r = FIRST_LINE_ROW To LAST_LINE_ROW

        sku = Trim$(CStr(wsUI.Cells(r, COL_SKU).value))

        If sku <> "" Then

            qty = CDbl(wsUI.Cells(r, COL_QTY).value)
            unitPrice = CDbl(wsUI.Cells(r, COL_PRICE).value)
            productName = Trim$(CStr(wsUI.Cells(r, COL_NAME).value))
            lineTotal = Round(qty * unitPrice, 2)

            productRow = FindProductRowBySKU(tblP, sku)
            If productRow = 0 Then
                Err.Raise vbObjectError + 1412, , "Product SKU not found while saving."
            End If

            productID = Trim$(CStr(tblP.ListColumns("Product_ID").DataBodyRange.Cells(productRow, 1).value))
            currentStock = NzNumber(tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value)
            newStock = currentStock - qty

            If newStock < 0 Then
                Err.Raise vbObjectError + 1413, , "Stock would become negative while saving."
            End If

            With tblS.ListRows.Add
                .Range(1, 1).value = salesID
                .Range(1, 2).value = salesDate
                .Range(1, 3).value = customerID
                .Range(1, 4).value = customerName
                .Range(1, 5).value = productID
                .Range(1, 6).value = sku
                .Range(1, 7).value = productName
                .Range(1, 8).value = qty
                .Range(1, 9).value = unitPrice
                .Range(1, 10).value = ""
                .Range(1, 11).value = 0
                .Range(1, 12).value = lineTotal
                .Range(1, 13).value = lineTotal
                .Range(1, 14).value = "Not Invoiced"
                .Range(1, 15).value = "Unpaid"
                .Range(1, 16).value = notes
                .Range(1, 17).value = Date
                .Range(1, 18).value = Date
            End With

            tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value = newStock

            With tblL.ListRows.Add
                .Range(1, 1).value = GenerateLogID(tblL)
                .Range(1, 2).value = salesDate
                .Range(1, 3).value = "SALE"
                .Range(1, 4).value = salesID
                .Range(1, 5).value = productID
                .Range(1, 6).value = sku
                .Range(1, 7).value = productName
                .Range(1, 8).value = -qty
                .Range(1, 9).value = newStock
                .Range(1, 10).value = unitPrice
                .Range(1, 11).value = lineTotal
                .Range(1, 12).value = "CUSTOMER"
                .Range(1, 13).value = customerID
                .Range(1, 14).value = customerName
                .Range(1, 15).value = notes
                .Range(1, 16).value = Date
            End With

        End If
    Next r

End Sub

'==================================================
' UI HELPERS
'==================================================
Private Sub ClearSalesForm(ByVal wsUI As Worksheet)

    Dim r As Long

    wsUI.Range(CELL_ID).value = ""
    wsUI.Range(CELL_DATE).value = ""
    wsUI.Range(CELL_CUSTOMER).value = ""
    wsUI.Range(CELL_NOTES).value = ""

    For r = FIRST_LINE_ROW To LAST_LINE_ROW
        wsUI.Cells(r, COL_SKU).value = ""
        wsUI.Cells(r, COL_NAME).Formula = ""
        wsUI.Cells(r, COL_QTY).value = ""
        wsUI.Cells(r, COL_PRICE).value = ""
        wsUI.Cells(r, COL_TOTAL).Formula = ""
    Next r

    For r = FIRST_LINE_ROW To LAST_LINE_ROW
        wsUI.Cells(r, COL_NAME).Formula = "=IFERROR(XLOOKUP(A" & r & ",tblProducts[SKU],tblProducts[Product_Name]),"""")"
        wsUI.Cells(r, COL_TOTAL).Formula = "=IFERROR(C" & r & "*D" & r & ",0)"
    Next r

    wsUI.Range(CELL_SUBTOTAL).Formula = "=SUM(E" & FIRST_LINE_ROW & ":E" & LAST_LINE_ROW & ")"

End Sub

'==================================================
' STOCK / VALIDATION HELPERS
'==================================================
Private Function GetAvailableStockForSales(ByVal tblP As ListObject, ByVal tblS As ListObject, ByVal sku As String, ByVal salesID As String, ByVal isUpdate As Boolean) As Double

    Dim productRow As Long
    Dim currentStock As Double
    Dim releasedQty As Double
    Dim i As Long

    productRow = FindProductRowBySKU(tblP, sku)
    If productRow = 0 Then
        GetAvailableStockForSales = 0
        Exit Function
    End If

    currentStock = NzNumber(tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value)
    releasedQty = 0

    If isUpdate And salesID <> "" Then
        If Not tblS.DataBodyRange Is Nothing Then
            For i = 1 To tblS.ListRows.Count
                If Trim$(CStr(tblS.ListColumns("Sales_Order_No").DataBodyRange.Cells(i, 1).value)) = salesID _
                   And Trim$(CStr(tblS.ListColumns("SKU").DataBodyRange.Cells(i, 1).value)) = sku Then
                    releasedQty = releasedQty + NzNumber(tblS.ListColumns("Qty").DataBodyRange.Cells(i, 1).value)
                End If
            Next i
        End If
    End If

    GetAvailableStockForSales = currentStock + releasedQty

End Function

Private Function GenerateSalesID(ByVal tbl As ListObject) As String

    Dim maxNum As Long
    Dim i As Long
    Dim s As String
    Dim n As Long

    maxNum = 0

    If tbl.ListRows.Count = 0 Then
        GenerateSalesID = "SO00001"
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        s = Trim$(CStr(tbl.ListColumns("Sales_Order_No").DataBodyRange.Cells(i, 1).value))
        If Left$(s, 2) = "SO" Then
            If IsNumeric(Mid$(s, 3)) Then
                n = CLng(Mid$(s, 3))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i

    GenerateSalesID = "SO" & Format$(maxNum + 1, "00000")

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

Private Function SalesIDExists(ByVal tbl As ListObject, ByVal salesID As String) As Boolean

    Dim i As Long

    SalesIDExists = False

    If tbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("Sales_Order_No").DataBodyRange.Cells(i, 1).value)) = salesID Then
            SalesIDExists = True
            Exit Function
        End If
    Next i

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

'==================================================
' SAFE UPDATE HELPERS
'==================================================
Private Function BackupSalesRows(ByVal tbl As ListObject, ByVal salesID As String) As Collection
    Dim rowsBackup As New Collection
    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then
        Set BackupSalesRows = rowsBackup
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("Sales_Order_No").DataBodyRange.Cells(i, 1).value)) = salesID Then
            rowsBackup.Add tbl.ListRows(i).Range.value
        End If
    Next i

    Set BackupSalesRows = rowsBackup
End Function

Private Function BackupInventoryLogRows_Sales(ByVal tbl As ListObject, ByVal tranType As String, ByVal refNo As String) As Collection
    Dim rowsBackup As New Collection
    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then
        Set BackupInventoryLogRows_Sales = rowsBackup
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("Tran_Type").DataBodyRange.Cells(i, 1).value)) = tranType _
           And Trim$(CStr(tbl.ListColumns("Ref_No").DataBodyRange.Cells(i, 1).value)) = refNo Then
            rowsBackup.Add tbl.ListRows(i).Range.value
        End If
    Next i

    Set BackupInventoryLogRows_Sales = rowsBackup
End Function

Private Sub CollectOldSalesSKUs(ByVal tbl As ListObject, ByVal salesID As String, ByVal dict As Object)
    Dim i As Long
    Dim sku As String

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To tbl.ListRows.Count
        If Trim$(CStr(tbl.ListColumns("Sales_Order_No").DataBodyRange.Cells(i, 1).value)) = salesID Then
            sku = Trim$(CStr(tbl.ListColumns("SKU").DataBodyRange.Cells(i, 1).value))
            If sku <> "" Then
                If Not dict.Exists(sku) Then dict.Add sku, True
            End If
        End If
    Next i
End Sub

Private Sub CollectUISalesSKUs(ByVal wsUI As Worksheet, ByVal dict As Object)
    Dim r As Long
    Dim sku As String

    For r = FIRST_LINE_ROW To LAST_LINE_ROW
        sku = Trim$(CStr(wsUI.Cells(r, COL_SKU).value))
        If sku <> "" Then
            If Not dict.Exists(sku) Then dict.Add sku, True
        End If
    Next r
End Sub

Private Sub SnapshotSalesStocks(ByVal tblP As ListObject, ByVal impactedSKUs As Object, ByVal stockSnapshot As Object)
    Dim sku As Variant
    Dim rowNum As Long

    If impactedSKUs Is Nothing Then Exit Sub

    For Each sku In impactedSKUs.Keys
        rowNum = FindProductRowBySKU(tblP, CStr(sku))
        If rowNum > 0 Then
            stockSnapshot(CStr(sku)) = NzNumber(tblP.ListColumns("Current_Stock").DataBodyRange.Cells(rowNum, 1).value)
        End If
    Next sku
End Sub

Private Sub RestoreStockFromOldSales(ByVal tblS As ListObject, ByVal tblP As ListObject, ByVal salesID As String)
    Dim i As Long
    Dim sku As String
    Dim qty As Double
    Dim productRow As Long
    Dim currentStock As Double

    If tblS.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To tblS.ListRows.Count
        If Trim$(CStr(tblS.ListColumns("Sales_Order_No").DataBodyRange.Cells(i, 1).value)) = salesID Then
            sku = Trim$(CStr(tblS.ListColumns("SKU").DataBodyRange.Cells(i, 1).value))
            qty = NzNumber(tblS.ListColumns("Qty").DataBodyRange.Cells(i, 1).value)

            productRow = FindProductRowBySKU(tblP, sku)
            If productRow > 0 Then
                currentStock = NzNumber(tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value)
                tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value = currentStock + qty
            End If
        End If
    Next i
End Sub

Private Sub DeleteInventoryLogRowsByRef_Sales(ByVal tbl As ListObject, ByVal tranType As String, ByVal refNo As String)
    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For i = tbl.ListRows.Count To 1 Step -1
        If Trim$(CStr(tbl.ListColumns("Tran_Type").DataBodyRange.Cells(i, 1).value)) = tranType _
           And Trim$(CStr(tbl.ListColumns("Ref_No").DataBodyRange.Cells(i, 1).value)) = refNo Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

Private Sub DeleteSalesRowsByID(ByVal tbl As ListObject, ByVal salesID As String)
    Dim i As Long

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    For i = tbl.ListRows.Count To 1 Step -1
        If Trim$(CStr(tbl.ListColumns("Sales_Order_No").DataBodyRange.Cells(i, 1).value)) = salesID Then
            tbl.ListRows(i).Delete
        End If
    Next i
End Sub

Private Sub RestoreTableRowsFromCollection_Sales(ByVal tbl As ListObject, ByVal rowsBackup As Collection)
    Dim i As Long
    Dim newRow As ListRow

    If rowsBackup Is Nothing Then Exit Sub
    If rowsBackup.Count = 0 Then Exit Sub

    For i = 1 To rowsBackup.Count
        Set newRow = tbl.ListRows.Add
        newRow.Range.value = rowsBackup(i)
    Next i
End Sub

Private Sub RestoreProductStocks_Sales(ByVal tblP As ListObject, ByVal stockSnapshot As Object)
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

