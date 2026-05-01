Attribute VB_Name = "modShopifyImport"
Option Explicit

Private gShopifyFilePath As String

Private Const WS_IMPORT_UI As String = "Import_UI"
Private Const WS_VALIDATION As String = "Import_Validation"
Private Const WS_SALES_DB As String = "Sales_DB"
Private Const WS_PRODUCTS_DB As String = "Products_DB"
Private Const WS_CUSTOMERS_DB As String = "Customers_DB"
Private Const WS_INVENTORY_LOG As String = "Inventory_Log"

Private Const TBL_SALES As String = "tblSales"
Private Const TBL_PRODUCTS As String = "tblProducts"
Private Const TBL_CUSTOMERS As String = "tblCustomers"
Private Const TBL_LOG As String = "tblInventoryLog"

' Shared Import UI cells
Private Const UI_CELL_FILE As String = "B6"
Private Const UI_CELL_VALIDATION As String = "B7"
Private Const UI_CELL_RESULT As String = "B8"

Private Const UI_CELL_OPTION As String = "B9"

Private Const RUN_OPTION_STRICT As String = "STOP_ON_ERROR"
Private Const RUN_OPTION_VALID_ONLY As String = "IMPORT_VALID_ONLY"

' =========================================
' Public Buttons
' =========================================
Public Sub ShopifyImport_SelectFile()

    Dim fd As Object
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)
    Set fd = Application.FileDialog(3)

    With fd
        .Title = "Select Shopify Order CSV File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "All Files", "*.*"

        If .Show <> -1 Then Exit Sub

        gShopifyFilePath = .SelectedItems(1)
    End With

    ws.Range(UI_CELL_FILE).value = gShopifyFilePath
    ws.Range(UI_CELL_VALIDATION).ClearContents
    ws.Range(UI_CELL_RESULT).ClearContents

    Shopify_ClearValidationLink

End Sub

Public Sub ShopifyImport_Validate()

    Dim mappedRows As Collection
    Dim ok As Boolean

    Set mappedRows = New Collection
    ok = Shopify_MapAndValidateFile(mappedRows, True)

    If ok Then
        MsgBox "Shopify validation passed.", vbInformation
    Else
        MsgBox "Shopify validation failed. Click Last Validation to open Import_Validation.", vbExclamation
    End If

End Sub

Public Sub ShopifyImport_Run()

    Dim mappedRows As Collection
    Dim ok As Boolean
    Dim runOption As String
    Dim validCount As Long
    Dim errCount As Long

    Set mappedRows = New Collection
    ok = Shopify_MapAndValidateFile(mappedRows, True)

    runOption = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_OPTION).value)))
    If runOption = "" Then runOption = RUN_OPTION_STRICT

    GetShopifyValidationCounts errCount, validCount

    If runOption = RUN_OPTION_STRICT Then
        If Not ok Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT BLOCKED - VALIDATION FAILED"
            MsgBox "Shopify import stopped because validation failed.", vbExclamation
            Exit Sub
        End If

    ElseIf runOption = RUN_OPTION_VALID_ONLY Then
        If validCount = 0 Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "NO VALID ROWS TO IMPORT"
            MsgBox "There are no VALID rows to import.", vbExclamation
            Exit Sub
        End If

    Else
        MsgBox "Invalid Import Option. Use STOP_ON_ERROR or IMPORT_VALID_ONLY.", vbExclamation
        Exit Sub
    End If

    Shopify_RunMappedRows mappedRows

End Sub

Public Sub ShopifyImport_Clear()

    gShopifyFilePath = ""

    With ThisWorkbook.Worksheets(WS_IMPORT_UI)
        .Range(UI_CELL_FILE).ClearContents
        .Range(UI_CELL_VALIDATION).ClearContents
        .Range(UI_CELL_RESULT).ClearContents
        .Range(UI_CELL_VALIDATION).Interior.Pattern = xlNone
        .Range(UI_CELL_VALIDATION).Font.Underline = xlUnderlineStyleNone
        .Range(UI_CELL_VALIDATION).Font.ColorIndex = xlAutomatic
    End With

    Shopify_ClearValidationLink

End Sub

' =========================================
' Core: Map + Validate
' =========================================
Private Function Shopify_MapAndValidateFile(ByRef mappedRows As Collection, ByVal writeReport As Boolean) As Boolean

    Dim filePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim hdr As Object

    Dim orderCol As Long, createdCol As Long, shipNameCol As Long
    Dim billNameCol As Long, custNameCol As Long
    Dim skuCol As Long, prodNameCol As Long, qtyCol As Long
    Dim priceCol As Long, notesCol As Long, lineDiscCol As Long
    Dim lastRow As Long, i As Long
    Dim results As Collection
    Dim res As Object

    Dim salesTbl As ListObject, prodTbl As ListObject, custTbl As ListObject
    Dim dictProducts As Object, dictCustomers As Object, dictExistingOrders As Object
    Dim dictGeneratedSO As Object, dictTempStock As Object
    Dim dictOrderCustomer As Object

    Dim validCount As Long, errCount As Long, processedCount As Long
    Dim nextSONum As Long
    Dim errMsg As String
    Dim orderText As String, skuText As String

    On Error GoTo ErrHandler

    ProgressStart "Validating Shopify Orders", "Loading master data..."

    filePath = Shopify_GetFilePath()
    If filePath = "" Then
        ProgressEnd
        MsgBox "Please select Shopify file first.", vbExclamation
        Shopify_MapAndValidateFile = False
        Exit Function
    End If

    Set salesTbl = Shopify_GetTableSafe(WS_SALES_DB, TBL_SALES)
    Set prodTbl = Shopify_GetTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)
    Set custTbl = Shopify_GetTableSafe(WS_CUSTOMERS_DB, TBL_CUSTOMERS)

    If salesTbl Is Nothing Or prodTbl Is Nothing Then
        ProgressEnd
        MsgBox "Sales_DB or Products_DB table not found.", vbCritical
        Shopify_MapAndValidateFile = False
        Exit Function
    End If

    Set dictProducts = Shopify_BuildProductDict(prodTbl)
    Set dictCustomers = Shopify_BuildCustomerDict(custTbl)
    Set dictExistingOrders = Shopify_BuildExistingSalesOrderDict(salesTbl)
    Set dictGeneratedSO = CreateObject("Scripting.Dictionary")
    Set dictTempStock = Shopify_BuildTempStockDict(prodTbl)
    Set dictOrderCustomer = CreateObject("Scripting.Dictionary")
    Set results = New Collection
    Set mappedRows = New Collection

    nextSONum = Shopify_GetNextSalesOrderSeq(salesTbl)

    ProgressStep "Opening Shopify order file..."
    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Worksheets(1)

    ProgressStep "Reading headers..."
    Set hdr = Shopify_BuildHeaderMap(ws)

    orderCol = Shopify_FindCol(hdr, Array("NAME", "ORDERNAME"))
    createdCol = Shopify_FindCol(hdr, Array("CREATEDAT", "ORDERDATE", "PAIDAT"))
    shipNameCol = Shopify_FindCol(hdr, Array("SHIPPINGNAME"))
    billNameCol = Shopify_FindCol(hdr, Array("BILLINGNAME"))
    custNameCol = Shopify_FindCol(hdr, Array("CUSTOMERNAME"))
    skuCol = Shopify_FindCol(hdr, Array("LINEITEMSKU", "SKU"))
    prodNameCol = Shopify_FindCol(hdr, Array("LINEITEMNAME", "PRODUCTNAME", "ITEMNAME"))
    qtyCol = Shopify_FindCol(hdr, Array("LINEITEMQUANTITY", "QUANTITY"))
    priceCol = Shopify_FindCol(hdr, Array("LINEITEMPRICE", "PRICE"))
    notesCol = Shopify_FindCol(hdr, Array("NOTES", "NOTE"))
    lineDiscCol = Shopify_FindCol(hdr, Array("LINEITEMDISCOUNT", "LINEITEMDISCOUNTAMOUNT", "LINEITEMDISCOUNTVALUE"))
    

    If orderCol = 0 Then Shopify_AddHeaderError results, "Missing Shopify header: Name"
    If createdCol = 0 Then Shopify_AddHeaderError results, "Missing Shopify header: Created at"
    If skuCol = 0 Then Shopify_AddHeaderError results, "Missing Shopify header: Lineitem sku"
    If qtyCol = 0 Then Shopify_AddHeaderError results, "Missing Shopify header: Lineitem quantity"

    If results.Count > 0 Then
        wb.Close False
        Set wb = Nothing

        If writeReport Then Shopify_WriteValidationReport results
        Shopify_UpdateValidationUI False, 0, results.Count

        ProgressEnd
        Shopify_MapAndValidateFile = False
        Exit Function
    End If

    If WorksheetFunction.CountA(ws.Cells) = 0 Then
        wb.Close False
        Set wb = Nothing
        ProgressEnd
        MsgBox "Shopify file is empty.", vbExclamation
        Shopify_MapAndValidateFile = False
        Exit Function
    End If

    lastRow = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row

    For i = 2 To lastRow

        orderText = ""
        skuText = ""

        If orderCol > 0 Then orderText = Trim$(CStr(ws.Cells(i, orderCol).value))
        If skuCol > 0 Then skuText = Trim$(CStr(ws.Cells(i, skuCol).value))

        ProgressUpdate "Validating Shopify rows", i - 1, lastRow - 1, _
            "Order: " & orderText & " | SKU: " & skuText

        If Not Shopify_RowHasAnyData(ws, i, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column) Then GoTo nextRow

        processedCount = processedCount + 1
        
Set res = Shopify_ValidateAndMapOneRow( _
    ws, i, _
    orderCol, createdCol, shipNameCol, billNameCol, custNameCol, _
    skuCol, prodNameCol, qtyCol, priceCol, notesCol, lineDiscCol, _
    hdr, _
    dictProducts, dictCustomers, dictExistingOrders, dictGeneratedSO, dictTempStock, dictOrderCustomer, nextSONum)

        results.Add res

        If UCase$(res("Status")) = "VALID" Then
            validCount = validCount + 1
            mappedRows.Add res
        Else
            errCount = errCount + 1
        End If

nextRow:
    Next i

    If processedCount = 0 Then
        Set res = CreateObject("Scripting.Dictionary")
        Shopify_FillEmptyValidationRow res
        res("Row_No") = 0
        res("Status") = "ERROR"
        res("Message") = "No Shopify data rows found."
        results.Add res
        errCount = errCount + 1
    End If

    wb.Close False
    Set wb = Nothing

    ProgressStep "Writing validation report..."

    If writeReport Then Shopify_WriteValidationReport results
    Shopify_UpdateValidationUI (errCount = 0), validCount, errCount

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    Shopify_MapAndValidateFile = (errCount = 0)
    Exit Function

ErrHandler:
    errMsg = Err.Description

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0

    ProgressEnd
    MsgBox "Shopify validation error: " & errMsg, vbCritical
    Shopify_MapAndValidateFile = False

End Function

Private Function Shopify_ValidateAndMapOneRow( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal orderCol As Long, _
    ByVal createdCol As Long, _
    ByVal shipNameCol As Long, _
    ByVal billNameCol As Long, _
    ByVal custNameCol As Long, _
    ByVal skuCol As Long, _
    ByVal prodNameCol As Long, _
    ByVal qtyCol As Long, _
    ByVal priceCol As Long, _
    ByVal notesCol As Long, _
    ByVal lineDiscCol As Long, _
    ByVal hdr As Object, _
    ByVal dictProducts As Object, _
    ByVal dictCustomers As Object, _
    ByVal dictExistingOrders As Object, _
    ByVal dictGeneratedSO As Object, _
    ByVal dictTempStock As Object, _
    ByVal dictOrderCustomer As Object, _
    ByRef nextSONum As Long) As Object

    Dim res As Object
    Dim extOrderNo As String, salesOrderNo As String
    Dim salesDateRaw As Variant, salesDateVal As Variant
    Dim customerName As String, customerID As String
    Dim customerEmail As String, customerPhone As String
    Dim customerCountry As String, customerAddress As String
    Dim sku As String, inputProductName As String, productName As String, productID As String
    Dim qty As Double, unitPrice As Double
    Dim discountType As String, discountValue As Double
    Dim notesText As String
    Dim subtotal As Double, lineTotal As Double
    Dim unitCost As Double, sellingPrice As Double
    Dim msg As String
    Dim prod As Object

    Set res = CreateObject("Scripting.Dictionary")

    extOrderNo = Trim$(CStr(ws.Cells(rowNum, orderCol).value))
    salesDateRaw = ws.Cells(rowNum, createdCol).value

    customerName = Trim$(CStr(Shopify_FirstNonBlank( _
                    Shopify_GetCell(ws, rowNum, shipNameCol), _
                    Shopify_GetCell(ws, rowNum, billNameCol), _
                    Shopify_GetCell(ws, rowNum, custNameCol))))

    If extOrderNo <> "" Then
        If customerName <> "" Then
            dictOrderCustomer(UCase$(extOrderNo)) = customerName
        ElseIf dictOrderCustomer.Exists(UCase$(extOrderNo)) Then
            customerName = dictOrderCustomer(UCase$(extOrderNo))
        End If
    End If

    sku = UCase$(Trim$(CStr(Shopify_GetCell(ws, rowNum, skuCol))))
    inputProductName = Trim$(CStr(Shopify_GetCell(ws, rowNum, prodNameCol)))
    qty = Shopify_NzDbl(Shopify_GetCell(ws, rowNum, qtyCol))
    unitPrice = Shopify_NzDbl(Shopify_GetCell(ws, rowNum, priceCol))
    notesText = Trim$(CStr(Shopify_GetCell(ws, rowNum, notesCol)))

    customerEmail = Trim$(CStr(Shopify_GetCell(ws, rowNum, Shopify_FindCol(hdr, Array("EMAIL", "CUSTOMEREMAIL")))))
    customerPhone = Trim$(CStr(Shopify_GetCell(ws, rowNum, Shopify_FindCol(hdr, Array("PHONE", "CUSTOMERPHONE", "SHIPPINGPHONE", "BILLINGPHONE")))))
    customerCountry = Trim$(CStr(Shopify_GetCell(ws, rowNum, Shopify_FindCol(hdr, Array("SHIPPINGCOUNTRY", "SHIPCOUNTRY", "COUNTRY")))))

    customerAddress = Trim$(CStr(Shopify_GetCell(ws, rowNum, Shopify_FindCol(hdr, Array("SHIPPINGADDRESS1", "SHIPADDRESS1", "BILLINGADDRESS1")))))
    If Trim$(CStr(Shopify_GetCell(ws, rowNum, Shopify_FindCol(hdr, Array("SHIPPINGADDRESS2", "SHIPADDRESS2", "BILLINGADDRESS2"))))) <> "" Then
        customerAddress = Trim$(customerAddress & " " & Trim$(CStr(Shopify_GetCell(ws, rowNum, Shopify_FindCol(hdr, Array("SHIPPINGADDRESS2", "SHIPADDRESS2", "BILLINGADDRESS2"))))))
    End If

    discountValue = Shopify_NzDbl(Shopify_GetCell(ws, rowNum, lineDiscCol))
    If discountValue > 0 Then
        discountType = "AMOUNT"
    Else
        discountType = ""
        discountValue = 0
    End If

    salesDateVal = Shopify_ParseDateValue(salesDateRaw)

    If extOrderNo = "" Then msg = Shopify_AppendMsg(msg, "Shopify order Name is blank")
    If IsEmpty(salesDateVal) Then msg = Shopify_AppendMsg(msg, "Created at invalid")
    If customerName = "" Then msg = Shopify_AppendMsg(msg, "Customer name missing")
    If sku = "" Then msg = Shopify_AppendMsg(msg, "SKU required")
    If qty <= 0 Then msg = Shopify_AppendMsg(msg, "Qty must be > 0")

    salesOrderNo = Shopify_ResolveSalesOrderNo("", extOrderNo, dictGeneratedSO, nextSONum)

    If dictExistingOrders.Exists(UCase$(salesOrderNo)) Then
        msg = Shopify_AppendMsg(msg, "Sales_Order_No already exists in Sales_DB")
    End If

    If sku <> "" Then
        If dictProducts.Exists(UCase$(sku)) Then
            Set prod = dictProducts(UCase$(sku))

            productID = prod("Product_ID")
            productName = prod("Product_Name")
            unitCost = prod("Unit_Cost")
            sellingPrice = prod("Selling_Price")

            If productName = "" Then msg = Shopify_AppendMsg(msg, "Product master missing Product_Name")
            If unitPrice <= 0 Then unitPrice = sellingPrice

            If dictTempStock.Exists(UCase$(sku)) Then
                If dictTempStock(UCase$(sku)) < qty Then
                    msg = Shopify_AppendMsg(msg, "Insufficient stock for SKU")
                Else
                    dictTempStock(UCase$(sku)) = dictTempStock(UCase$(sku)) - qty
                End If
            Else
                msg = Shopify_AppendMsg(msg, "SKU stock record missing")
            End If
        Else
            msg = Shopify_AppendMsg(msg, "SKU not found in Products_DB")
        End If
    End If

    If unitPrice < 0 Then msg = Shopify_AppendMsg(msg, "Unit_Price cannot be negative")

    subtotal = qty * unitPrice
    lineTotal = subtotal - discountValue

    If discountValue < 0 Or discountValue > subtotal Then
        msg = Shopify_AppendMsg(msg, "Discount_Value invalid")
    End If

    If customerID = "" Then
        If dictCustomers.Exists(UCase$(customerName)) Then
            customerID = dictCustomers(UCase$(customerName))
        End If
    End If

    res("Row_No") = rowNum
    res("Status") = IIf(msg = "", "VALID", "ERROR")
    res("Message") = IIf(msg = "", "OK", msg)

    res("Sales_Order_No") = salesOrderNo
    If IsEmpty(salesDateVal) Then
        res("Sales_Date") = Empty
    Else
        res("Sales_Date") = salesDateVal
    End If

    res("Customer_ID") = customerID
    res("Customer_Name") = customerName
    res("Customer_Email") = customerEmail
    res("Customer_Phone") = customerPhone
    res("Customer_Address") = customerAddress
    res("Customer_Country") = customerCountry

    res("Product_ID") = productID
    res("SKU") = sku
    res("Product_Name") = productName
    res("Qty") = qty
    res("Unit_Price") = unitPrice
    res("Discount_Type") = discountType
    res("Discount_Value") = discountValue
    res("Line_Subtotal") = subtotal
    res("Line_Total") = lineTotal
    res("Notes") = notesText
    res("Source") = "SHOPIFY"
    res("External_Order_No") = extOrderNo
    res("Unit_Cost") = unitCost

    Set Shopify_ValidateAndMapOneRow = res

End Function
Private Function Shopify_ResolveSalesOrderNo( _
    ByVal salesOrderNo As String, _
    ByVal externalOrderNo As String, _
    ByVal dictGenerated As Object, _
    ByRef nextSONum As Long) As String

    Dim key As String

    salesOrderNo = Trim$(salesOrderNo)
    externalOrderNo = Trim$(externalOrderNo)

    If salesOrderNo <> "" Then
        Shopify_ResolveSalesOrderNo = salesOrderNo
        Exit Function
    End If

    If externalOrderNo <> "" Then
        key = UCase$(externalOrderNo)
        If dictGenerated.Exists(key) Then
            Shopify_ResolveSalesOrderNo = dictGenerated(key)
        Else
            Shopify_ResolveSalesOrderNo = Shopify_GenerateSalesOrderNo(nextSONum)
            dictGenerated.Add key, Shopify_ResolveSalesOrderNo
            nextSONum = nextSONum + 1
        End If
    Else
        Shopify_ResolveSalesOrderNo = Shopify_GenerateSalesOrderNo(nextSONum)
        nextSONum = nextSONum + 1
    End If

End Function

Private Function Shopify_GenerateSalesOrderNo(ByVal seqNum As Long) As String
    Shopify_GenerateSalesOrderNo = "S" & Format$(seqNum, "000000")
End Function

' =========================================
' Run Import
' =========================================
Private Sub Shopify_RunMappedRows(ByVal mappedRows As Collection)

    Dim salesTbl As ListObject, prodTbl As ListObject, logTbl As ListObject, custTbl As ListObject
    Dim prodRows As Object, custDict As Object, existingExtDict As Object
    Dim skippedOrders As Object, replacedOrders As Object
    Dim i As Long
    Dim rowObj As Object
    Dim newStock As Double
    Dim orderText As String, skuText As String, extKey As String
    Dim importCount As Long
    Dim skipAll As Boolean, replaceAll As Boolean
    Dim action As String

    On Error GoTo ErrHandler

    Set salesTbl = Shopify_GetTableSafe(WS_SALES_DB, TBL_SALES)
    Set prodTbl = Shopify_GetTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)
    Set logTbl = Shopify_GetTableSafe(WS_INVENTORY_LOG, TBL_LOG)
    Set custTbl = Shopify_GetTableSafe(WS_CUSTOMERS_DB, TBL_CUSTOMERS)

    If salesTbl Is Nothing Or prodTbl Is Nothing Or logTbl Is Nothing Or custTbl Is Nothing Then
        MsgBox "Required tables not found.", vbCritical
        Exit Sub
    End If

    Set prodRows = Shopify_BuildProductRowDict(prodTbl)
    Set custDict = Shopify_BuildCustomerDict(custTbl)
    Set existingExtDict = Shopify_BuildExistingExternalOrderDict(salesTbl)
    Set skippedOrders = CreateObject("Scripting.Dictionary")
    Set replacedOrders = CreateObject("Scripting.Dictionary")

    ProgressStart "Running Shopify Order Import", "Writing sales and inventory records..."

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For i = 1 To mappedRows.Count

        Set rowObj = mappedRows(i)

        orderText = Trim$(CStr(rowObj("External_Order_No")))
        If orderText = "" Then orderText = Trim$(CStr(rowObj("Sales_Order_No")))
        skuText = Trim$(CStr(rowObj("SKU")))
        extKey = UCase$(Trim$(CStr(rowObj("External_Order_No"))))

        ProgressUpdate "Importing Shopify rows", i, mappedRows.Count, _
            "Order: " & orderText & " | SKU: " & skuText

        If extKey <> "" And existingExtDict.Exists(extKey) Then

            If skippedOrders.Exists(extKey) Then GoTo nextRow

            If Not replacedOrders.Exists(extKey) Then

                If skipAll Then
                    skippedOrders(extKey) = True
                    GoTo nextRow

                ElseIf replaceAll Then
                    Shopify_ReplaceExistingExternalOrder salesTbl, prodTbl, logTbl, prodRows, rowObj("External_Order_No")
                    replacedOrders(extKey) = True

                Else
                    action = Shopify_AskDuplicateAction(rowObj("External_Order_No"))

                    Select Case action
                        Case "SKIP_ONE"
                            skippedOrders(extKey) = True
                            GoTo nextRow

                        Case "SKIP_ALL"
                            skipAll = True
                            skippedOrders(extKey) = True
                            GoTo nextRow

                        Case "REPLACE_ONE"
                            Shopify_ReplaceExistingExternalOrder salesTbl, prodTbl, logTbl, prodRows, rowObj("External_Order_No")
                            replacedOrders(extKey) = True

                        Case "REPLACE_ALL"
                            replaceAll = True
                            Shopify_ReplaceExistingExternalOrder salesTbl, prodTbl, logTbl, prodRows, rowObj("External_Order_No")
                            replacedOrders(extKey) = True

                        Case Else
                            GoTo StopImport
                    End Select
                End If

            End If
        End If

        Shopify_EnsureCustomerExists custTbl, custDict, rowObj

        Shopify_WriteSalesRow salesTbl, rowObj

        newStock = Shopify_DeductProductStock(prodTbl, prodRows, rowObj("SKU"), rowObj("Qty"))
        Shopify_WriteInventoryLogRow logTbl, rowObj, newStock

        importCount = importCount + 1

nextRow:
    Next i

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    With ThisWorkbook.Worksheets(WS_IMPORT_UI)
        .Range(UI_CELL_RESULT).value = "IMPORT SUCCESS - " & importCount & " row(s)"
    End With

    MsgBox "Shopify import completed: " & importCount & " row(s).", vbInformation
    Exit Sub

StopImport:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ProgressEnd
    ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT STOPPED BY USER"
    MsgBox "Shopify import stopped by user.", vbExclamation
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ProgressEnd
    MsgBox "Shopify import error: " & Err.Description, vbCritical

End Sub

Private Sub Shopify_WriteSalesRow(ByVal salesTbl As ListObject, ByVal rowObj As Object)

    Dim lr As ListRow
    Set lr = salesTbl.ListRows.Add

    Shopify_SetCellByHeader lr.Range, salesTbl, "Sales_Order_No", rowObj("Sales_Order_No")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Sales_Date", rowObj("Sales_Date")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Customer_ID", rowObj("Customer_ID")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Customer_Name", rowObj("Customer_Name")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Product_ID", rowObj("Product_ID")
    Shopify_SetCellByHeader lr.Range, salesTbl, "SKU", rowObj("SKU")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Product_Name", rowObj("Product_Name")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Qty", rowObj("Qty")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Unit_Price", rowObj("Unit_Price")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Discount_Type", rowObj("Discount_Type")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Discount_Value", rowObj("Discount_Value")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Line_Subtotal", rowObj("Line_Subtotal")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Line_Total", rowObj("Line_Total")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Invoice_Status", "Not Invoiced"
    Shopify_SetCellByHeader lr.Range, salesTbl, "Payment_Status", "Unpaid"
    Shopify_SetCellByHeader lr.Range, salesTbl, "Notes", rowObj("Notes")

    ' Use Shopify order created time as Created_At
    Shopify_SetCellByHeader lr.Range, salesTbl, "Created_At", rowObj("Sales_Date")
    Shopify_SetCellByHeader lr.Range, salesTbl, "Updated_At", Now

    Shopify_SetCellByHeader lr.Range, salesTbl, "Source", "SHOPIFY"
    Shopify_SetCellByHeader lr.Range, salesTbl, "External_Order_No", rowObj("External_Order_No")

End Sub
Private Function Shopify_DeductProductStock( _
    ByVal prodTbl As ListObject, _
    ByVal prodRows As Object, _
    ByVal sku As String, _
    ByVal qty As Double) As Double

    Dim rr As Range
    Dim colStock As Long

    colStock = Shopify_GetHeaderColumn(prodTbl, "Current_Stock")

    If prodRows.Exists(UCase$(sku)) Then
        Set rr = prodRows(UCase$(sku))
        rr.Cells(1, colStock).value = Shopify_NzDbl(rr.Cells(1, colStock).value) - qty
        Shopify_DeductProductStock = Shopify_NzDbl(rr.Cells(1, colStock).value)
    Else
        Shopify_DeductProductStock = 0
    End If

End Function

Private Sub Shopify_WriteInventoryLogRow(ByVal logTbl As ListObject, ByVal rowObj As Object, ByVal balanceAfter As Double)

    Dim lr As ListRow
    Set lr = logTbl.ListRows.Add

    Shopify_SetCellByHeader lr.Range, logTbl, "Log_ID", Shopify_NextLogID(logTbl)
    Shopify_SetCellByHeader lr.Range, logTbl, "Log_Date", rowObj("Sales_Date")
    Shopify_SetCellByHeader lr.Range, logTbl, "Tran_Type", "SALE"
    Shopify_SetCellByHeader lr.Range, logTbl, "Ref_No", rowObj("Sales_Order_No")
    Shopify_SetCellByHeader lr.Range, logTbl, "Product_ID", rowObj("Product_ID")
    Shopify_SetCellByHeader lr.Range, logTbl, "SKU", rowObj("SKU")
    Shopify_SetCellByHeader lr.Range, logTbl, "Product_Name", rowObj("Product_Name")
    Shopify_SetCellByHeader lr.Range, logTbl, "Qty_Change", -rowObj("Qty")
    Shopify_SetCellByHeader lr.Range, logTbl, "Balance_After", balanceAfter
    Shopify_SetCellByHeader lr.Range, logTbl, "Unit_Cost", rowObj("Unit_Cost")
    Shopify_SetCellByHeader lr.Range, logTbl, "Total_Cost", rowObj("Unit_Cost") * rowObj("Qty")
    Shopify_SetCellByHeader lr.Range, logTbl, "Party_Type", "CUSTOMER"
    Shopify_SetCellByHeader lr.Range, logTbl, "Party_ID", rowObj("Customer_ID")
    Shopify_SetCellByHeader lr.Range, logTbl, "Party_Name", rowObj("Customer_Name")
    Shopify_SetCellByHeader lr.Range, logTbl, "Notes", rowObj("Notes")
    Shopify_SetCellByHeader lr.Range, logTbl, "Created_At", Now

End Sub

' =========================================
' Validation Report UI
' =========================================
Private Sub Shopify_UpdateValidationUI(ByVal isValid As Boolean, ByVal validCount As Long, ByVal errCount As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)

    Shopify_ClearValidationLink

    With ws.Range(UI_CELL_VALIDATION)
        .Hyperlinks.Delete
        .Font.Underline = xlUnderlineStyleNone
        .Font.ColorIndex = xlAutomatic
        .Interior.Pattern = xlSolid

        If isValid Then
            .value = "VALID - " & validCount & " row(s)"
            .Interior.Color = RGB(226, 239, 218)
        Else
            .value = "ERROR - " & errCount & " row(s) (Click to view)"
            .Interior.Color = RGB(255, 199, 206)

            ws.Hyperlinks.Add Anchor:=ws.Range(UI_CELL_VALIDATION), _
                              Address:="", _
                              SubAddress:="'" & WS_VALIDATION & "'!A1", _
                              TextToDisplay:=ws.Range(UI_CELL_VALIDATION).value

            ws.Range(UI_CELL_VALIDATION).Font.Color = RGB(0, 0, 255)
            ws.Range(UI_CELL_VALIDATION).Font.Underline = xlUnderlineStyleSingle
        End If
    End With

End Sub

Private Sub Shopify_ClearValidationLink()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)

    On Error Resume Next
    ws.Range(UI_CELL_VALIDATION).Hyperlinks.Delete
    On Error GoTo 0

End Sub

Private Sub Shopify_WriteValidationReport(ByVal results As Collection)

    Dim ws As Worksheet
    Dim i As Long, r As Long
    Dim res As Object

    Set ws = Shopify_GetOrCreateSheet(WS_VALIDATION)
    ws.Cells.Clear

    Shopify_AddValidationBackLink ws

    ws.Range("A3:T3").value = Array( _
        "Row_No", "Status", "Sales_Order_No", "Sales_Date", "Customer_ID", "Customer_Name", _
        "Product_ID", "SKU", "Product_Name", "Qty", "Unit_Price", "Discount_Type", _
        "Discount_Value", "Line_Subtotal", "Line_Total", "Notes", "Source", _
        "External_Order_No", "Message", "Unit_Cost")

    ws.Rows(3).Font.Bold = True
    ws.Rows(3).Interior.Color = RGB(217, 225, 242)

    r = 4
    For i = 1 To results.Count
        Set res = results(i)

        ws.Cells(r, 1).value = res("Row_No")
        ws.Cells(r, 2).value = res("Status")
        ws.Cells(r, 3).value = res("Sales_Order_No")
        ws.Cells(r, 4).value = res("Sales_Date")
        ws.Cells(r, 5).value = res("Customer_ID")
        ws.Cells(r, 6).value = res("Customer_Name")
        ws.Cells(r, 7).value = res("Product_ID")
        ws.Cells(r, 8).value = res("SKU")
        ws.Cells(r, 9).value = res("Product_Name")
        ws.Cells(r, 10).value = res("Qty")
        ws.Cells(r, 11).value = res("Unit_Price")
        ws.Cells(r, 12).value = res("Discount_Type")
        ws.Cells(r, 13).value = res("Discount_Value")
        ws.Cells(r, 14).value = res("Line_Subtotal")
        ws.Cells(r, 15).value = res("Line_Total")
        ws.Cells(r, 16).value = res("Notes")
        ws.Cells(r, 17).value = res("Source")
        ws.Cells(r, 18).value = res("External_Order_No")
        ws.Cells(r, 19).value = res("Message")
        ws.Cells(r, 20).value = res("Unit_Cost")

        If UCase$(Trim$(res("Status"))) = "ERROR" Then
            ws.Rows(r).Interior.Color = RGB(255, 230, 230)
        End If

        r = r + 1
    Next i

    ws.Columns("A:T").AutoFit
If r > 4 Then
    ws.Range("D4:D" & r - 1).NumberFormat = "yyyy-mm-dd"
    ws.Range("J4:O" & r - 1).NumberFormat = "#,##0.00"
End If

If r > 3 Then
    ws.Range("A3:T" & r - 1).Borders.LineStyle = xlContinuous
End If

End Sub

Private Sub Shopify_AddValidationBackLink(ByVal ws As Worksheet)

    On Error Resume Next
    ws.Hyperlinks.Delete
    On Error GoTo 0

    ws.Range("A1").value = "<< Back to Import_UI"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Color = RGB(0, 0, 255)
    ws.Range("A1").Font.Underline = xlUnderlineStyleSingle

    ws.Hyperlinks.Add Anchor:=ws.Range("A1"), _
                      Address:="", _
                      SubAddress:="'" & WS_IMPORT_UI & "'!A1", _
                      TextToDisplay:="<< Back to Import_UI"

    ws.Range("A2").value = "Shopify Validation Details"
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Size = 14

End Sub

Private Sub Shopify_AddHeaderError(ByVal results As Collection, ByVal msg As String)

    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    Shopify_FillEmptyValidationRow res
    res("Row_No") = 0
    res("Status") = "ERROR"
    res("Message") = msg
    results.Add res

End Sub

Private Sub Shopify_FillEmptyValidationRow(ByVal res As Object)

    res("Row_No") = 0
    res("Status") = ""
    res("Sales_Order_No") = ""
    res("Sales_Date") = ""
    res("Customer_ID") = ""
    res("Customer_Name") = ""
    res("Product_ID") = ""
    res("SKU") = ""
    res("Product_Name") = ""
    res("Qty") = 0
    res("Unit_Price") = 0
    res("Discount_Type") = ""
    res("Discount_Value") = 0
    res("Line_Subtotal") = 0
    res("Line_Total") = 0
    res("Notes") = ""
    res("Source") = ""
    res("External_Order_No") = ""
    res("Message") = ""
    res("Unit_Cost") = 0

End Sub

' =========================================
' Helpers
' =========================================
Private Function Shopify_GetFilePath() As String
    If Trim$(gShopifyFilePath) <> "" Then
        Shopify_GetFilePath = gShopifyFilePath
    Else
        Shopify_GetFilePath = Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_FILE).value))
    End If
End Function

Private Function Shopify_BuildHeaderMap(ByVal ws As Worksheet) As Object

    Dim dict As Object
    Dim lastCol As Long, c As Long
    Dim h As String

    Set dict = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        h = Shopify_NormalizeHeader(CStr(ws.Cells(1, c).value))
        If h <> "" Then dict(h) = c
    Next c

    Set Shopify_BuildHeaderMap = dict

End Function

Private Function Shopify_NormalizeHeader(ByVal txt As String) As String

    txt = UCase$(Trim$(txt))
    txt = Replace(txt, " ", "")
    txt = Replace(txt, "_", "")
    txt = Replace(txt, "-", "")
    txt = Replace(txt, ".", "")
    txt = Replace(txt, "/", "")
    txt = Replace(txt, "#", "")
    txt = Replace(txt, "*", "")
    txt = Replace(txt, "(", "")
    txt = Replace(txt, ")", "")

    Shopify_NormalizeHeader = txt

End Function

Private Function Shopify_FindCol(ByVal hdr As Object, ByVal aliases As Variant) As Long

    Dim i As Long, key As String

    For i = LBound(aliases) To UBound(aliases)
        key = Shopify_NormalizeHeader(CStr(aliases(i)))
        If hdr.Exists(key) Then
            Shopify_FindCol = hdr(key)
            Exit Function
        End If
    Next i

End Function

Private Function Shopify_GetCell(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colNum As Long) As Variant
    If colNum > 0 Then
        Shopify_GetCell = ws.Cells(rowNum, colNum).value
    Else
        Shopify_GetCell = Empty
    End If
End Function

Private Function Shopify_FirstNonBlank(ParamArray vals() As Variant) As String

    Dim i As Long
    For i = LBound(vals) To UBound(vals)
        If Trim$(CStr(vals(i))) <> "" Then
            Shopify_FirstNonBlank = Trim$(CStr(vals(i)))
            Exit Function
        End If
    Next i

End Function

Private Function Shopify_ParseDateValue(ByVal v As Variant) As Variant
    If IsDate(v) Then
        Shopify_ParseDateValue = CDate(v)
    Else
        Shopify_ParseDateValue = Empty
    End If
End Function

Private Function Shopify_RowHasAnyData(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal maxCol As Long) As Boolean

    Dim c As Long
    For c = 1 To maxCol
        If Trim$(CStr(ws.Cells(rowNum, c).value)) <> "" Then
            Shopify_RowHasAnyData = True
            Exit Function
        End If
    Next c

End Function

Private Function Shopify_AppendMsg(ByVal baseMsg As String, ByVal addMsg As String) As String
    If baseMsg = "" Then
        Shopify_AppendMsg = addMsg
    Else
        Shopify_AppendMsg = baseMsg & " | " & addMsg
    End If
End Function

Private Function Shopify_NzDbl(ByVal v As Variant) As Double
    If IsError(v) Then
        Shopify_NzDbl = 0
    ElseIf IsNumeric(v) Then
        Shopify_NzDbl = CDbl(v)
    Else
        Shopify_NzDbl = 0
    End If
End Function

Private Function Shopify_GetTableSafe(ByVal wsName As String, ByVal preferredTableName As String) As ListObject

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set Shopify_GetTableSafe = ws.ListObjects(preferredTableName)
    On Error GoTo 0

    If Shopify_GetTableSafe Is Nothing Then
        If ws.ListObjects.Count > 0 Then
            Set Shopify_GetTableSafe = ws.ListObjects(1)
        End If
    End If

End Function

Private Function Shopify_GetHeaderColumn(ByVal lo As ListObject, ByVal headerName As String) As Long

    Dim i As Long

    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(i).name), Trim$(headerName), vbTextCompare) = 0 Then
            Shopify_GetHeaderColumn = i
            Exit Function
        End If
    Next i

End Function

Private Sub Shopify_SetCellByHeader(ByVal rowRange As Range, ByVal lo As ListObject, ByVal headerName As String, ByVal newValue As Variant)

    Dim c As Long
    c = Shopify_GetHeaderColumn(lo, headerName)
    If c > 0 Then rowRange.Cells(1, c).value = newValue

End Sub

Private Function Shopify_BuildProductDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object, prod As Object
    Dim i As Long
    Dim cPID As Long, cSKU As Long, cPN As Long, cUC As Long, cSP As Long, cCS As Long
    Dim keySKU As String

    Set dict = CreateObject("Scripting.Dictionary")

    cPID = Shopify_GetHeaderColumn(prodTbl, "Product_ID")
    cSKU = Shopify_GetHeaderColumn(prodTbl, "SKU")
    cPN = Shopify_GetHeaderColumn(prodTbl, "Product_Name")
    cUC = Shopify_GetHeaderColumn(prodTbl, "Unit_Cost")
    cSP = Shopify_GetHeaderColumn(prodTbl, "Selling_Price")
    cCS = Shopify_GetHeaderColumn(prodTbl, "Current_Stock")

    For i = 1 To prodTbl.ListRows.Count

        keySKU = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))

        If keySKU <> "" Then
            Set prod = CreateObject("Scripting.Dictionary")
            prod("Product_ID") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cPID).value))
            prod("Product_Name") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cPN).value))
            prod("Unit_Cost") = Shopify_NzDbl(prodTbl.DataBodyRange.Cells(i, cUC).value)
            prod("Selling_Price") = Shopify_NzDbl(prodTbl.DataBodyRange.Cells(i, cSP).value)
            prod("Current_Stock") = Shopify_NzDbl(prodTbl.DataBodyRange.Cells(i, cCS).value)

            If Not dict.Exists(keySKU) Then
                dict.Add keySKU, prod
            End If
        End If

    Next i

    Set Shopify_BuildProductDict = dict

End Function

Private Function Shopify_BuildProductRowDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cSKU As Long
    Dim keySKU As String

    Set dict = CreateObject("Scripting.Dictionary")
    cSKU = Shopify_GetHeaderColumn(prodTbl, "SKU")

    For i = 1 To prodTbl.ListRows.Count
        keySKU = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))
        If keySKU <> "" Then
            If Not dict.Exists(keySKU) Then
                dict.Add keySKU, prodTbl.DataBodyRange.Rows(i)
            End If
        End If
    Next i

    Set Shopify_BuildProductRowDict = dict

End Function

Private Function Shopify_BuildTempStockDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cSKU As Long, cCS As Long
    Dim keySKU As String

    Set dict = CreateObject("Scripting.Dictionary")
    cSKU = Shopify_GetHeaderColumn(prodTbl, "SKU")
    cCS = Shopify_GetHeaderColumn(prodTbl, "Current_Stock")

    For i = 1 To prodTbl.ListRows.Count
        keySKU = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))
        If keySKU <> "" Then
            dict(keySKU) = Shopify_NzDbl(prodTbl.DataBodyRange.Cells(i, cCS).value)
        End If
    Next i

    Set Shopify_BuildTempStockDict = dict

End Function

Private Function Shopify_BuildCustomerDict(ByVal custTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cID As Long, cName As Long
    Dim keyName As String

    Set dict = CreateObject("Scripting.Dictionary")

    If custTbl Is Nothing Then
        Set Shopify_BuildCustomerDict = dict
        Exit Function
    End If

    If custTbl.DataBodyRange Is Nothing Then
        Set Shopify_BuildCustomerDict = dict
        Exit Function
    End If

    cID = Shopify_GetHeaderColumn(custTbl, "Customer_ID")
    cName = Shopify_GetHeaderColumn(custTbl, "Customer_Name")

    If cID = 0 Or cName = 0 Then
        Set Shopify_BuildCustomerDict = dict
        Exit Function
    End If

    For i = 1 To custTbl.ListRows.Count
        keyName = UCase$(Trim$(CStr(custTbl.DataBodyRange.Cells(i, cName).value)))
        If keyName <> "" Then
            dict(keyName) = Trim$(CStr(custTbl.DataBodyRange.Cells(i, cID).value))
        End If
    Next i

    Set Shopify_BuildCustomerDict = dict

End Function

Private Function Shopify_BuildExistingSalesOrderDict(ByVal salesTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cSO As Long
    Dim keySO As String

    Set dict = CreateObject("Scripting.Dictionary")

    cSO = Shopify_GetHeaderColumn(salesTbl, "Sales_Order_No")
    If cSO = 0 Then
        Set Shopify_BuildExistingSalesOrderDict = dict
        Exit Function
    End If

    If salesTbl.DataBodyRange Is Nothing Then
        Set Shopify_BuildExistingSalesOrderDict = dict
        Exit Function
    End If

    For i = 1 To salesTbl.ListRows.Count
        keySO = UCase$(Trim$(CStr(salesTbl.DataBodyRange.Cells(i, cSO).value)))
        If keySO <> "" Then dict(keySO) = True
    Next i

    Set Shopify_BuildExistingSalesOrderDict = dict

End Function

Private Function Shopify_GetNextSalesOrderSeq(ByVal salesTbl As ListObject) As Long

    Dim i As Long, cSO As Long
    Dim soNo As String, n As Long, maxN As Long
    Dim digitsOnly As String

    cSO = Shopify_GetHeaderColumn(salesTbl, "Sales_Order_No")
    If cSO = 0 Then
        Shopify_GetNextSalesOrderSeq = 1
        Exit Function
    End If

    If salesTbl.DataBodyRange Is Nothing Then
        Shopify_GetNextSalesOrderSeq = 1
        Exit Function
    End If

    For i = 1 To salesTbl.ListRows.Count
        soNo = Trim$(CStr(salesTbl.DataBodyRange.Cells(i, cSO).value))
        digitsOnly = Shopify_ExtractDigits(soNo)
        If digitsOnly <> "" Then
            n = CLng(digitsOnly)
            If n > maxN Then maxN = n
        End If
    Next i

    Shopify_GetNextSalesOrderSeq = maxN + 1

End Function

Private Function Shopify_ExtractDigits(ByVal txt As String) As String

    Dim i As Long, ch As String

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch Like "#" Then Shopify_ExtractDigits = Shopify_ExtractDigits & ch
    Next i

End Function

Private Function Shopify_NextLogID(ByVal logTbl As ListObject) As String

    Dim c As Long, i As Long
    Dim txt As String, n As Long, maxN As Long

    c = Shopify_GetHeaderColumn(logTbl, "Log_ID")
    If c = 0 Or logTbl.DataBodyRange Is Nothing Then
        Shopify_NextLogID = "L000001"
        Exit Function
    End If

    For i = 1 To logTbl.ListRows.Count
        txt = Trim$(CStr(logTbl.DataBodyRange.Cells(i, c).value))
        If UCase$(Left$(txt, 1)) = "L" Then
            If IsNumeric(Mid$(txt, 2)) Then
                n = CLng(Mid$(txt, 2))
                If n > maxN Then maxN = n
            End If
        End If
    Next i

    Shopify_NextLogID = "L" & Format$(maxN + 1, "000000")

End Function

Private Function Shopify_GetOrCreateSheet(ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set Shopify_GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If Shopify_GetOrCreateSheet Is Nothing Then
        Set Shopify_GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        Shopify_GetOrCreateSheet.name = sheetName
    End If

End Function

Private Sub GetShopifyValidationCounts(ByRef errCount As Long, ByRef validCount As Long)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim s As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_VALIDATION)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 4 To lastRow
        s = UCase$(Trim$(CStr(ws.Cells(i, 2).value)))

        If s = "VALID" Then
            validCount = validCount + 1
        ElseIf s = "ERROR" Then
            errCount = errCount + 1
        End If
    Next i

End Sub


Private Function Shopify_AskDuplicateAction(ByVal externalOrderNo As String) As String

    With frmDuplicateOrderAction
        .SelectedAction = ""
        .SetOrderInfo externalOrderNo
        .Show vbModal
        Shopify_AskDuplicateAction = .SelectedAction
    End With

End Function

Private Function Shopify_BuildExistingExternalOrderDict(ByVal salesTbl As ListObject) As Object

    Dim dict As Object
    Dim cExt As Long
    Dim i As Long
    Dim extNo As String

    Set dict = CreateObject("Scripting.Dictionary")

    cExt = Shopify_GetHeaderColumn(salesTbl, "External_Order_No")

    If cExt = 0 Or salesTbl.DataBodyRange Is Nothing Then
        Set Shopify_BuildExistingExternalOrderDict = dict
        Exit Function
    End If

    For i = 1 To salesTbl.ListRows.Count
        extNo = UCase$(Trim$(CStr(salesTbl.DataBodyRange.Cells(i, cExt).value)))
        If extNo <> "" Then dict(extNo) = True
    Next i

    Set Shopify_BuildExistingExternalOrderDict = dict

End Function

Private Sub Shopify_EnsureCustomerExists(ByVal custTbl As ListObject, ByVal custDict As Object, ByVal rowObj As Object)

    Dim custName As String
    Dim keyName As String
    Dim newID As String
    Dim lr As ListRow
    Dim existingRow As Range

    Dim custEmail As String
    Dim custPhone As String
    Dim custAddress As String
    Dim custCountry As String

    custName = Trim$(CStr(rowObj("Customer_Name")))
    If custName = "" Then Exit Sub

    keyName = UCase$(custName)

    custEmail = Shopify_GetDictText(rowObj, "Customer_Email")
    custPhone = Shopify_GetDictText(rowObj, "Customer_Phone")
    custAddress = Shopify_GetDictText(rowObj, "Customer_Address")
    custCountry = Shopify_GetDictText(rowObj, "Customer_Country")

    If custDict.Exists(keyName) Then
        rowObj("Customer_ID") = custDict(keyName)

        Set existingRow = Shopify_FindCustomerRowByID(custTbl, CStr(rowObj("Customer_ID")))
        If Not existingRow Is Nothing Then
            Shopify_UpdateCustomerBlankFields custTbl, existingRow, custEmail, custPhone, custAddress, custCountry
        End If

        Exit Sub
    End If

    newID = Shopify_NextCustomerID(custTbl)

    Set lr = custTbl.ListRows.Add

    Shopify_SetCellByHeader lr.Range, custTbl, "Customer_ID", newID
    Shopify_SetCellByHeader lr.Range, custTbl, "Customer_Name", custName
    Shopify_SetCellByHeader lr.Range, custTbl, "Contact_Person", custName
    Shopify_SetCellByHeader lr.Range, custTbl, "Email", custEmail
    Shopify_SetCellByHeader lr.Range, custTbl, "Phone", custPhone
    Shopify_SetCellByHeader lr.Range, custTbl, "Address", custAddress
    Shopify_SetCellByHeader lr.Range, custTbl, "Country", custCountry
    Shopify_SetCellByHeader lr.Range, custTbl, "Notes", "Auto-created from Shopify order import"
    Shopify_SetCellByHeader lr.Range, custTbl, "Active_Status", "Active"
    Shopify_SetCellByHeader lr.Range, custTbl, "Created_At", Date
    Shopify_SetCellByHeader lr.Range, custTbl, "Updated_At", Date

    custDict(keyName) = newID
    rowObj("Customer_ID") = newID

End Sub
Private Function Shopify_NextCustomerID(ByVal custTbl As ListObject) As String

    Dim cID As Long
    Dim i As Long
    Dim txt As String
    Dim n As Long
    Dim maxN As Long

    cID = Shopify_GetHeaderColumn(custTbl, "Customer_ID")

    If cID = 0 Or custTbl.DataBodyRange Is Nothing Then
        Shopify_NextCustomerID = "C00001"
        Exit Function
    End If

    For i = 1 To custTbl.ListRows.Count
        txt = Trim$(CStr(custTbl.DataBodyRange.Cells(i, cID).value))

        If UCase$(Left$(txt, 1)) = "C" Then
            If IsNumeric(Mid$(txt, 2)) Then
                n = CLng(Mid$(txt, 2))
                If n > maxN Then maxN = n
            End If
        End If
    Next i

    Shopify_NextCustomerID = "C" & Format$(maxN + 1, "00000")

End Function

Private Sub Shopify_ReplaceExistingExternalOrder( _
    ByVal salesTbl As ListObject, _
    ByVal prodTbl As ListObject, _
    ByVal logTbl As ListObject, _
    ByVal prodRows As Object, _
    ByVal externalOrderNo As String)

    Dim cExt As Long, cSO As Long, cSKU As Long, cQty As Long
    Dim i As Long
    Dim extKey As String
    Dim soNo As String, sku As String
    Dim qty As Double
    Dim soDict As Object

    Set soDict = CreateObject("Scripting.Dictionary")

    extKey = UCase$(Trim$(externalOrderNo))

    cExt = Shopify_GetHeaderColumn(salesTbl, "External_Order_No")
    cSO = Shopify_GetHeaderColumn(salesTbl, "Sales_Order_No")
    cSKU = Shopify_GetHeaderColumn(salesTbl, "SKU")
    cQty = Shopify_GetHeaderColumn(salesTbl, "Qty")

    If cExt = 0 Or cSO = 0 Or cSKU = 0 Or cQty = 0 Then Exit Sub
    If salesTbl.DataBodyRange Is Nothing Then Exit Sub

    For i = salesTbl.ListRows.Count To 1 Step -1

        If UCase$(Trim$(CStr(salesTbl.DataBodyRange.Cells(i, cExt).value))) = extKey Then

            soNo = Trim$(CStr(salesTbl.DataBodyRange.Cells(i, cSO).value))
            sku = Trim$(CStr(salesTbl.DataBodyRange.Cells(i, cSKU).value))
            qty = Shopify_NzDbl(salesTbl.DataBodyRange.Cells(i, cQty).value)

            If soNo <> "" Then soDict(soNo) = True

            Shopify_RestoreProductStock prodTbl, prodRows, sku, qty

            salesTbl.ListRows(i).Delete

        End If

    Next i

    Shopify_DeleteInventoryLogsBySalesOrders logTbl, soDict

End Sub

Private Sub Shopify_RestoreProductStock( _
    ByVal prodTbl As ListObject, _
    ByVal prodRows As Object, _
    ByVal sku As String, _
    ByVal qty As Double)

    Dim rr As Range
    Dim cStock As Long

    cStock = Shopify_GetHeaderColumn(prodTbl, "Current_Stock")
    If cStock = 0 Then Exit Sub

    If prodRows.Exists(UCase$(sku)) Then
        Set rr = prodRows(UCase$(sku))
        rr.Cells(1, cStock).value = Shopify_NzDbl(rr.Cells(1, cStock).value) + qty
    End If

End Sub

Private Sub Shopify_DeleteInventoryLogsBySalesOrders(ByVal logTbl As ListObject, ByVal soDict As Object)

    Dim cRef As Long
    Dim i As Long
    Dim refNo As String

    If soDict Is Nothing Then Exit Sub
    If soDict.Count = 0 Then Exit Sub

    cRef = Shopify_GetHeaderColumn(logTbl, "Ref_No")
    If cRef = 0 Or logTbl.DataBodyRange Is Nothing Then Exit Sub

    For i = logTbl.ListRows.Count To 1 Step -1
        refNo = Trim$(CStr(logTbl.DataBodyRange.Cells(i, cRef).value))

        If soDict.Exists(refNo) Then
            logTbl.ListRows(i).Delete
        End If
    Next i

End Sub


Private Function Shopify_GetDictText(ByVal dict As Object, ByVal key As String) As String

    If dict Is Nothing Then
        Shopify_GetDictText = ""
    ElseIf dict.Exists(key) Then
        Shopify_GetDictText = Trim$(CStr(dict(key)))
    Else
        Shopify_GetDictText = ""
    End If

End Function


Private Function Shopify_FindCustomerRowByID(ByVal custTbl As ListObject, ByVal customerID As String) As Range

    Dim cID As Long
    Dim i As Long

    cID = Shopify_GetHeaderColumn(custTbl, "Customer_ID")
    If cID = 0 Then Exit Function
    If custTbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To custTbl.ListRows.Count
        If Trim$(CStr(custTbl.DataBodyRange.Cells(i, cID).value)) = Trim$(customerID) Then
            Set Shopify_FindCustomerRowByID = custTbl.DataBodyRange.Rows(i)
            Exit Function
        End If
    Next i

End Function

Private Sub Shopify_UpdateCustomerBlankFields( _
    ByVal custTbl As ListObject, _
    ByVal rowRange As Range, _
    ByVal custEmail As String, _
    ByVal custPhone As String, _
    ByVal custAddress As String, _
    ByVal custCountry As String)

    Shopify_SetIfBlankByHeader custTbl, rowRange, "Email", custEmail
    Shopify_SetIfBlankByHeader custTbl, rowRange, "Phone", custPhone
    Shopify_SetIfBlankByHeader custTbl, rowRange, "Address", custAddress
    Shopify_SetIfBlankByHeader custTbl, rowRange, "Country", custCountry

    Shopify_SetCellByHeader rowRange, custTbl, "Updated_At", Date

End Sub

Private Sub Shopify_SetIfBlankByHeader(ByVal lo As ListObject, ByVal rowRange As Range, ByVal headerName As String, ByVal newValue As String)

    Dim c As Long
    c = Shopify_GetHeaderColumn(lo, headerName)

    If c = 0 Then Exit Sub
    If Trim$(newValue) = "" Then Exit Sub

    If Trim$(CStr(rowRange.Cells(1, c).value)) = "" Then
        rowRange.Cells(1, c).value = newValue
    End If

End Sub

