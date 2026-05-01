Attribute VB_Name = "modImportSales"
Option Explicit

Private gImportFilePath As String

Private Const WS_IMPORT_UI As String = "Import_UI"
Private Const WS_SALES_DB As String = "Sales_DB"
Private Const WS_PRODUCTS_DB As String = "Products_DB"
Private Const WS_INVENTORY_LOG As String = "Inventory_Log"
Private Const WS_CUSTOMERS_DB As String = "Customers_DB"
Private Const WS_VALIDATION As String = "Import_Validation"

Private Const UI_CELL_FILE As String = "B6"
Private Const UI_CELL_VALIDATION As String = "B7"
Private Const UI_CELL_RESULT As String = "B8"

Private Const UI_CELL_OPTION As String = "B9"

Private Const RUN_OPTION_STRICT As String = "STOP_ON_ERROR"
Private Const RUN_OPTION_VALID_ONLY As String = "IMPORT_VALID_ONLY"

Private Const TBL_SALES As String = "tblSales"
Private Const TBL_PRODUCTS As String = "tblProducts"
Private Const TBL_LOG As String = "tblInventoryLog"
Private Const TBL_CUSTOMERS As String = "tblCustomers"



' ===============================
' Public Buttons
' ===============================
Public Sub ImportSales_DownloadTemplate()

    Dim wb As Workbook
    Dim wsInst As Worksheet
    Dim wsData As Worksheet
    Dim savePath As Variant

    On Error GoTo ErrHandler

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="Sales_Import_Template.xlsx", _
        FileFilter:="Excel Files (*.xlsx), *.xlsx")

    If VarType(savePath) = vbBoolean Then Exit Sub

    ProgressStart "Building Sales Import Template", "Preparing workbook..."

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = Workbooks.Add

    ProgressStep "Building instruction sheet..."
    Set wsInst = wb.Sheets(1)
    wsInst.name = "Instructions"
    BuildImportInstructionsSheet wsInst

    ProgressStep "Building data sheet..."
    Set wsData = wb.Sheets.Add(After:=wsInst)
    wsData.name = "Data"
    BuildImportDataSheet wsData

    ProgressStep "Saving template..."
    wb.SaveAs CStr(savePath), FileFormat:=xlOpenXMLWorkbook
    wb.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    MsgBox "Sales import template downloaded.", vbInformation
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ProgressEnd
    MsgBox "Download Template error: " & Err.Description, vbCritical

End Sub

Public Sub ImportSales_SelectFile()

    Dim fd As Object
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)
    Set fd = Application.FileDialog(3)

    With fd
        .Title = "Select Sales Import File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel and CSV Files", "*.xlsx;*.xls;*.csv"
        .Filters.Add "All Files", "*.*"

        If .Show <> -1 Then Exit Sub

        gImportFilePath = .SelectedItems(1)
    End With

    ws.Range(UI_CELL_FILE).value = gImportFilePath
    ws.Range(UI_CELL_VALIDATION).ClearContents
    ws.Range(UI_CELL_RESULT).ClearContents

    ClearValidationLink
    ClearValidationSheetNav

End Sub

Public Sub ImportSales_Validate()

    Dim ok As Boolean

    ok = ValidateImportFile(True)

    If ok Then
        MsgBox "Validation passed.", vbInformation
    Else
        MsgBox "Validation failed. Click Last Validation to open Import_Validation.", vbExclamation
    End If

End Sub

Public Sub ImportSales_Run()

    Dim ok As Boolean
    Dim runOption As String
    Dim validCount As Long
    Dim errCount As Long

    ok = ValidateImportFile(True)

    runOption = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_OPTION).value)))
    If runOption = "" Then runOption = RUN_OPTION_STRICT

    GetSalesValidationCounts errCount, validCount

    If runOption = RUN_OPTION_STRICT Then
        If Not ok Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT BLOCKED - VALIDATION FAILED"
            MsgBox "Import stopped because validation failed. Click Last Validation to view details.", vbExclamation
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

    RunValidatedImport runOption

End Sub

Public Sub ImportSales_Clear()

    gImportFilePath = ""

    With ThisWorkbook.Worksheets(WS_IMPORT_UI)
        .Range(UI_CELL_FILE).ClearContents
        .Range(UI_CELL_VALIDATION).ClearContents
        .Range(UI_CELL_RESULT).ClearContents
        .Range(UI_CELL_VALIDATION).Interior.Pattern = xlNone
        .Range(UI_CELL_VALIDATION).Font.Underline = xlUnderlineStyleNone
        .Range(UI_CELL_VALIDATION).Font.ColorIndex = xlAutomatic
    End With

    ClearValidationLink

    On Error Resume Next
    ThisWorkbook.Worksheets(WS_VALIDATION).Cells.Clear
    On Error GoTo 0

End Sub

Public Sub ImportValidation_BackToImport()
    ThisWorkbook.Worksheets(WS_IMPORT_UI).Activate
    ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_FILE).Select
End Sub

' ===============================
' Template Builder
' ===============================
Private Sub BuildImportInstructionsSheet(ByVal ws As Worksheet)

    Dim r As Long

    ws.Cells.Clear

    ws.Range("A1").value = "Sales Import Instructions"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A3:D3").value = Array("Field", "Required?", "Auto / Default", "Explanation")
    With ws.Range("A3:D3")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
    End With

    r = 4
    WriteInstructionRow ws, r, "Sales_Order_No", "Optional", "Auto-generated if blank", "Internal sales order number. If multiple lines belong to one order, use the same Sales_Order_No or same External_Order_No."
    r = r + 1
    WriteInstructionRow ws, r, "Sales_Date", "Required", "", "Must be a valid date."
    r = r + 1
    WriteInstructionRow ws, r, "Customer_ID", "Optional", "Matched by Customer_Name if blank", "SimpleERP customer ID."
    r = r + 1
    WriteInstructionRow ws, r, "Customer_Name", "Required", "", "Customer name for this order."
    r = r + 1
    WriteInstructionRow ws, r, "SKU", "Required", "", "Must exist in Products_DB."
    r = r + 1
    WriteInstructionRow ws, r, "Product_Name", "Optional", "Recommended", "If filled, it must match the product master name for that SKU. This helps catch wrong SKU input."
    r = r + 1
    WriteInstructionRow ws, r, "Qty", "Required", "", "Must be greater than 0 and cannot exceed current stock."
    r = r + 1
    WriteInstructionRow ws, r, "Unit_Price", "Optional", "Uses Selling_Price if blank", "Selling price per unit."
    r = r + 1
    WriteInstructionRow ws, r, "Discount_Type", "Optional", "Blank / PERCENT / AMOUNT", "Discount type."
    r = r + 1
    WriteInstructionRow ws, r, "Discount_Value", "Optional", "Default = 0", "If Discount_Type = PERCENT, valid range is 0 to 100. If AMOUNT, cannot exceed subtotal."
    r = r + 1
    WriteInstructionRow ws, r, "Notes", "Optional", "", "Any notes for this line."
    r = r + 1
    WriteInstructionRow ws, r, "Source", "Optional", "Default = IMPORT", "Examples: SHOPIFY / MANUAL / AMAZON."
    r = r + 1
    WriteInstructionRow ws, r, "External_Order_No", "Optional", "Used to group lines if Sales_Order_No blank", "External platform order number."

    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 34
    ws.Columns("D").ColumnWidth = 60
    ws.Range("A3:D" & r).WrapText = True
    ws.Range("A3:D" & r).Borders.LineStyle = xlContinuous

End Sub

Private Sub WriteInstructionRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal fld As String, ByVal req As String, ByVal autoTxt As String, ByVal explainTxt As String)

    ws.Cells(rowNum, 1).value = fld
    ws.Cells(rowNum, 2).value = req
    ws.Cells(rowNum, 3).value = autoTxt
    ws.Cells(rowNum, 4).value = explainTxt

End Sub

Private Sub BuildImportDataSheet(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long

    ws.Cells.Clear

    headers = Array( _
        "Sales_Order_No", _
        "Sales_Date *", _
        "Customer_ID", _
        "Customer_Name *", _
        "SKU *", _
        "Product_Name", _
        "Qty *", _
        "Unit_Price (Auto if blank)", _
        "Discount_Type", _
        "Discount_Value", _
        "Notes", _
        "Source", _
        "External_Order_No")

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = headers(i)
    Next i

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(217, 225, 242)

    For i = 1 To 13
        If InStr(1, ws.Cells(1, i).value, "*", vbTextCompare) > 0 Then
            ws.Cells(1, i).Font.Color = RGB(192, 0, 0)
        End If
    Next i

    ws.Range("A2").value = "Optional"
    ws.Range("B2").value = "Required"
    ws.Range("C2").value = "Optional"
    ws.Range("D2").value = "Required"
    ws.Range("E2").value = "Required"
    ws.Range("F2").value = "Optional / Check"
    ws.Range("G2").value = "Required"
    ws.Range("H2").value = "Optional / Auto"
    ws.Range("I2").value = "Optional"
    ws.Range("J2").value = "Optional"
    ws.Range("K2").value = "Optional"
    ws.Range("L2").value = "Optional"
    ws.Range("M2").value = "Optional"

    ws.Rows(2).Font.Italic = True
    ws.Rows(2).Font.Color = RGB(90, 90, 90)
    ws.Range("A1:M2").Borders.LineStyle = xlContinuous

    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 24
    ws.Columns("G").ColumnWidth = 10
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 14
    ws.Columns("K").ColumnWidth = 18
    ws.Columns("L").ColumnWidth = 14
    ws.Columns("M").ColumnWidth = 18

End Sub

' ===============================
' Validation Core
' ===============================
Private Function ValidateImportFile(ByVal writeReport As Boolean) As Boolean

    Dim filePath As String
    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, i As Long

    Dim hdr As Object
    Dim results As Collection
    Dim res As Object

    Dim salesTbl As ListObject, prodTbl As ListObject, custTbl As ListObject
    Dim dictProducts As Object, dictCustomers As Object, dictExistingOrders As Object
    Dim dictGenerated As Object, dictTempStock As Object

    Dim rowHasData As Boolean
    Dim errCount As Long, validCount As Long
    Dim nextSONum As Long
    Dim errMsg As String
    Dim skuText As String
    Dim orderText As String

    On Error GoTo ErrHandler

    ProgressStart "Validating Sales Import", "Loading master data..."

    filePath = GetImportFilePath()
    If filePath = "" Then
        ProgressEnd
        MsgBox "Please select file first.", vbExclamation
        ValidateImportFile = False
        Exit Function
    End If

    Set salesTbl = GetTableSafe(WS_SALES_DB, TBL_SALES)
    Set prodTbl = GetTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)
    Set custTbl = GetTableSafe(WS_CUSTOMERS_DB, TBL_CUSTOMERS)

    If salesTbl Is Nothing Or prodTbl Is Nothing Then
        ProgressEnd
        MsgBox "Sales_DB or Products_DB table not found.", vbCritical
        ValidateImportFile = False
        Exit Function
    End If

    Set dictProducts = BuildProductDict(prodTbl)
    Set dictCustomers = BuildCustomerDict(custTbl)
    Set dictExistingOrders = BuildExistingSalesOrderDict(salesTbl)
    Set dictGenerated = CreateObject("Scripting.Dictionary")
    Set dictTempStock = BuildTempStockDict(prodTbl)
    Set results = New Collection

    nextSONum = GetNextSalesOrderSeq(salesTbl)

    ProgressStep "Opening import file..."

    On Error Resume Next
    Set wb = Workbooks.Open(filePath)
    On Error GoTo ErrHandler

    If wb Is Nothing Then
        ProgressEnd
        MsgBox "Cannot open import file. Please make sure the file exists and is not corrupted.", vbCritical
        ValidateImportFile = False
        Exit Function
    End If

    Set ws = GetImportDataSheet(wb)

    If ws Is Nothing Then
        wb.Close False
        Set wb = Nothing
        ProgressEnd
        MsgBox "Cannot find data sheet in import file.", vbCritical
        ValidateImportFile = False
        Exit Function
    End If

    ProgressStep "Reading headers..."
    Set hdr = BuildHeaderMap(ws)

    If Not CheckRequiredHeaders(hdr, results) Then
        wb.Close False
        Set wb = Nothing

        If writeReport Then WriteValidationReport results
        UpdateValidationUI False, 0, results.Count

        ProgressEnd
        ValidateImportFile = False
        Exit Function
    End If

    If WorksheetFunction.CountA(ws.Cells) = 0 Then
        wb.Close False
        Set wb = Nothing
        ProgressEnd
        MsgBox "Import file is empty.", vbExclamation
        ValidateImportFile = False
        Exit Function
    End If

    lastRow = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row

    For i = 3 To lastRow

        skuText = Trim$(CStr(GetCellString(ws, i, hdr, "SKU")))
        orderText = Trim$(CStr(GetCellString(ws, i, hdr, "Sales_Order_No")))

        If orderText = "" Then
            orderText = Trim$(CStr(GetCellString(ws, i, hdr, "External_Order_No")))
        End If

        ProgressUpdate "Validating rows", i - 2, lastRow - 2, _
            "Order: " & orderText & " | SKU: " & skuText

        rowHasData = RowHasAnyData(ws, i, 13)
        If Not rowHasData Then GoTo nextRow

        Set res = ValidateSingleImportRow( _
            ws, i, hdr, dictProducts, dictCustomers, dictExistingOrders, _
            dictGenerated, dictTempStock, nextSONum)

        results.Add res

        If res("Status") = "VALID" Then
            validCount = validCount + 1
        Else
            errCount = errCount + 1
        End If

nextRow:
    Next i

    wb.Close False
    Set wb = Nothing

    ProgressStep "Writing validation report..."

    If writeReport Then WriteValidationReport results

    UpdateValidationUI (errCount = 0), validCount, errCount

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    ValidateImportFile = (errCount = 0)
    Exit Function

ErrHandler:
    errMsg = Err.Description

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0

    ProgressEnd
    MsgBox "Validation error: " & errMsg, vbCritical
    ValidateImportFile = False

End Function

Private Sub UpdateValidationUI(ByVal isValid As Boolean, ByVal validCount As Long, ByVal errCount As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)

    ClearValidationLink

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

Private Function ValidateSingleImportRow( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal hdr As Object, _
    ByVal dictProducts As Object, _
    ByVal dictCustomers As Object, _
    ByVal dictExistingOrders As Object, _
    ByVal dictGenerated As Object, _
    ByVal dictTempStock As Object, _
    ByRef nextSONum As Long) As Object

    Dim res As Object
    Dim salesOrderNo As String, salesDate As Variant, customerID As String, customerName As String
    Dim sku As String, inputProductName As String, productName As String, productID As String
    Dim qty As Double, unitPrice As Double, discountType As String, discountValue As Double
    Dim notesText As String, sourceText As String, extOrderNo As String
    Dim subtotal As Double, lineTotal As Double
    Dim unitCost As Double, sellingPrice As Double
    Dim msg As String
    Dim prod As Object

    Set res = CreateObject("Scripting.Dictionary")

    salesOrderNo = GetCellString(ws, rowNum, hdr, "Sales_Order_No")
    salesDate = GetCellValue(ws, rowNum, hdr, "Sales_Date")
    customerID = GetCellString(ws, rowNum, hdr, "Customer_ID")
    customerName = GetCellString(ws, rowNum, hdr, "Customer_Name")
    sku = UCase$(Trim$(GetCellString(ws, rowNum, hdr, "SKU")))
    inputProductName = Trim$(GetCellString(ws, rowNum, hdr, "Product_Name"))
    qty = NzDbl(GetCellValue(ws, rowNum, hdr, "Qty"))
    unitPrice = NzDbl(GetCellValue(ws, rowNum, hdr, "Unit_Price"))
    discountType = UCase$(Trim$(GetCellString(ws, rowNum, hdr, "Discount_Type")))
    discountValue = NzDbl(GetCellValue(ws, rowNum, hdr, "Discount_Value"))
    notesText = GetCellString(ws, rowNum, hdr, "Notes")
    sourceText = GetCellString(ws, rowNum, hdr, "Source")
    extOrderNo = GetCellString(ws, rowNum, hdr, "External_Order_No")

    If sourceText = "" Then sourceText = "IMPORT"

    If Not IsDate(salesDate) Then msg = AppendMsg(msg, "Sales_Date invalid")
    If customerName = "" Then msg = AppendMsg(msg, "Customer_Name required")
    If sku = "" Then msg = AppendMsg(msg, "SKU required")
    If qty <= 0 Then msg = AppendMsg(msg, "Qty must be > 0")

    salesOrderNo = ResolveSalesOrderNo(salesOrderNo, extOrderNo, dictGenerated, nextSONum)

    If dictExistingOrders.Exists(UCase$(salesOrderNo)) Then
        msg = AppendMsg(msg, "Sales_Order_No already exists in Sales_DB")
    End If

    If sku <> "" Then
        If dictProducts.Exists(UCase$(sku)) Then

            Set prod = dictProducts(UCase$(sku))

            productID = prod("Product_ID")
            productName = prod("Product_Name")
            unitCost = prod("Unit_Cost")
            sellingPrice = prod("Selling_Price")

            If productName = "" Then
                msg = AppendMsg(msg, "Product master missing Product_Name")
            End If

            If inputProductName <> "" Then
                If StrComp(inputProductName, productName, vbTextCompare) <> 0 Then
                    msg = AppendMsg(msg, "Product_Name does not match SKU master")
                End If
            End If

            If unitPrice <= 0 Then
                unitPrice = sellingPrice
            End If

            If unitPrice < 0 Then
                msg = AppendMsg(msg, "Unit_Price cannot be negative")
            End If

            If dictTempStock.Exists(UCase$(sku)) Then
                If dictTempStock(UCase$(sku)) < qty Then
                    msg = AppendMsg(msg, "Insufficient stock for SKU")
                Else
                    dictTempStock(UCase$(sku)) = dictTempStock(UCase$(sku)) - qty
                End If
            Else
                msg = AppendMsg(msg, "SKU stock record missing")
            End If

        Else
            msg = AppendMsg(msg, "SKU not found in Products_DB")
        End If
    End If

    If discountType <> "" Then
        If discountType <> "PERCENT" And discountType <> "AMOUNT" Then
            msg = AppendMsg(msg, "Discount_Type must be blank / PERCENT / AMOUNT")
        End If
    End If

    subtotal = qty * unitPrice
    lineTotal = subtotal

    If discountType = "PERCENT" Then
        If discountValue < 0 Or discountValue > 100 Then
            msg = AppendMsg(msg, "Discount_Value invalid for PERCENT")
        Else
            lineTotal = subtotal * (1 - discountValue / 100)
        End If
    ElseIf discountType = "AMOUNT" Then
        If discountValue < 0 Or discountValue > subtotal Then
            msg = AppendMsg(msg, "Discount_Value invalid for AMOUNT")
        Else
            lineTotal = subtotal - discountValue
        End If
    Else
        discountValue = 0
    End If

    If customerID = "" Then
        If dictCustomers.Exists(UCase$(customerName)) Then
            customerID = dictCustomers(UCase$(customerName))
        End If
    End If

    res("Row_No") = rowNum
    If msg = "" Then
        res("Status") = "VALID"
        res("Message") = "OK"
    Else
        res("Status") = "ERROR"
        res("Message") = msg
    End If

    res("Sales_Order_No") = salesOrderNo
    If IsDate(salesDate) Then
        res("Sales_Date") = CDate(salesDate)
    Else
        res("Sales_Date") = Empty
    End If
    res("Customer_ID") = customerID
    res("Customer_Name") = customerName
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
    res("Source") = sourceText
    res("External_Order_No") = extOrderNo
    res("Unit_Cost") = unitCost

    Set ValidateSingleImportRow = res

End Function

Private Function ResolveSalesOrderNo( _
    ByVal salesOrderNo As String, _
    ByVal externalOrderNo As String, _
    ByVal dictGenerated As Object, _
    ByRef nextSONum As Long) As String

    Dim key As String

    salesOrderNo = Trim$(salesOrderNo)
    externalOrderNo = Trim$(externalOrderNo)

    If salesOrderNo <> "" Then
        ResolveSalesOrderNo = salesOrderNo
        Exit Function
    End If

    If externalOrderNo <> "" Then
        key = UCase$(externalOrderNo)
        If dictGenerated.Exists(key) Then
            ResolveSalesOrderNo = dictGenerated(key)
        Else
            ResolveSalesOrderNo = GenerateSalesOrderNo(nextSONum)
            dictGenerated.Add key, ResolveSalesOrderNo
            nextSONum = nextSONum + 1
        End If
    Else
        ResolveSalesOrderNo = GenerateSalesOrderNo(nextSONum)
        nextSONum = nextSONum + 1
    End If

End Function

Private Function GenerateSalesOrderNo(ByVal seqNum As Long) As String
    GenerateSalesOrderNo = "S" & Format$(seqNum, "000000")
End Function

' ===============================
' Import Execution
' ===============================
Private Sub RunValidatedImport(ByVal runOption As String)

    Dim wsVal As Worksheet
    Dim salesTbl As ListObject, prodTbl As ListObject, logTbl As ListObject
    Dim prodRows As Object
    Dim r As Long, lastRow As Long
    Dim importCount As Long

    Dim soNo As String, salesDate As Date, customerID As String, customerName As String
    Dim productID As String, sku As String, productName As String
    Dim qty As Double, unitPrice As Double, discountType As String, discountValue As Double
    Dim subtotal As Double, lineTotal As Double, notesText As String, unitCost As Double
    Dim newStock As Double

    On Error GoTo ErrHandler

    Set wsVal = ThisWorkbook.Worksheets(WS_VALIDATION)
    Set salesTbl = GetTableSafe(WS_SALES_DB, TBL_SALES)
    Set prodTbl = GetTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)
    Set logTbl = GetTableSafe(WS_INVENTORY_LOG, TBL_LOG)

    If wsVal Is Nothing Then
        MsgBox "Validation report not found. Please validate first.", vbExclamation
        Exit Sub
    End If

    If salesTbl Is Nothing Or prodTbl Is Nothing Or logTbl Is Nothing Then
        MsgBox "Required tables not found.", vbCritical
        Exit Sub
    End If

    lastRow = wsVal.Cells.Find(What:="*", After:=wsVal.Range("A1"), LookIn:=xlFormulas, _
                               LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                               MatchCase:=False).Row

    If lastRow < 4 Then
        MsgBox "No validation data found.", vbExclamation
        Exit Sub
    End If

If runOption = RUN_OPTION_STRICT Then
    If HasValidationErrors(wsVal) Then
        ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT BLOCKED - VALIDATION FAILED"
        MsgBox "Validation still contains errors. Please fix them first.", vbExclamation
        Exit Sub
    End If
End If

    Set prodRows = BuildProductRowDict(prodTbl)

    ProgressStart "Running Sales Import", "Writing sales and inventory records..."

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = 4 To lastRow

        soNo = Trim$(CStr(wsVal.Cells(r, 3).value))
        sku = Trim$(CStr(wsVal.Cells(r, 8).value))

        ProgressUpdate "Importing rows", r - 3, lastRow - 3, _
            "Order: " & soNo & " | SKU: " & sku

        If UCase$(Trim$(CStr(wsVal.Cells(r, 2).value))) <> "VALID" Then GoTo nextRow

        salesDate = wsVal.Cells(r, 4).value
        customerID = Trim$(CStr(wsVal.Cells(r, 5).value))
        customerName = Trim$(CStr(wsVal.Cells(r, 6).value))
        productID = Trim$(CStr(wsVal.Cells(r, 7).value))
        productName = Trim$(CStr(wsVal.Cells(r, 9).value))
        qty = NzDbl(wsVal.Cells(r, 10).value)
        unitPrice = NzDbl(wsVal.Cells(r, 11).value)
        discountType = Trim$(CStr(wsVal.Cells(r, 12).value))
        discountValue = NzDbl(wsVal.Cells(r, 13).value)
        subtotal = NzDbl(wsVal.Cells(r, 14).value)
        lineTotal = NzDbl(wsVal.Cells(r, 15).value)
        notesText = Trim$(CStr(wsVal.Cells(r, 16).value))
        unitCost = NzDbl(wsVal.Cells(r, 20).value)

        WriteSalesRow salesTbl, soNo, salesDate, customerID, customerName, _
                      productID, sku, productName, qty, unitPrice, _
                      discountType, discountValue, subtotal, lineTotal, notesText

        newStock = DeductProductStock(prodTbl, prodRows, sku, qty)

        WriteInventoryLogRow logTbl, salesDate, soNo, productID, sku, productName, _
                             qty, newStock, unitCost, customerID, customerName, notesText

        importCount = importCount + 1

nextRow:
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT SUCCESS - " & importCount & " row(s)"
    MsgBox "Import completed: " & importCount & " row(s).", vbInformation
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ProgressEnd
    MsgBox "Import error: " & Err.Description, vbCritical

End Sub

Private Sub WriteSalesRow( _
    ByVal salesTbl As ListObject, _
    ByVal soNo As String, _
    ByVal salesDate As Date, _
    ByVal customerID As String, _
    ByVal customerName As String, _
    ByVal productID As String, _
    ByVal sku As String, _
    ByVal productName As String, _
    ByVal qty As Double, _
    ByVal unitPrice As Double, _
    ByVal discountType As String, _
    ByVal discountValue As Double, _
    ByVal subtotal As Double, _
    ByVal lineTotal As Double, _
    ByVal notesText As String)

    Dim lr As ListRow
    Set lr = salesTbl.ListRows.Add

    SetCellByHeader lr.Range, salesTbl, "Sales_Order_No", soNo
    SetCellByHeader lr.Range, salesTbl, "Sales_Date", salesDate
    SetCellByHeader lr.Range, salesTbl, "Customer_ID", customerID
    SetCellByHeader lr.Range, salesTbl, "Customer_Name", customerName
    SetCellByHeader lr.Range, salesTbl, "Product_ID", productID
    SetCellByHeader lr.Range, salesTbl, "SKU", sku
    SetCellByHeader lr.Range, salesTbl, "Product_Name", productName
    SetCellByHeader lr.Range, salesTbl, "Qty", qty
    SetCellByHeader lr.Range, salesTbl, "Unit_Price", unitPrice
    SetCellByHeader lr.Range, salesTbl, "Discount_Type", discountType
    SetCellByHeader lr.Range, salesTbl, "Discount_Value", discountValue
    SetCellByHeader lr.Range, salesTbl, "Line_Subtotal", subtotal
    SetCellByHeader lr.Range, salesTbl, "Line_Total", lineTotal
    SetCellByHeader lr.Range, salesTbl, "Invoice_Status", "Not Invoiced"
    SetCellByHeader lr.Range, salesTbl, "Payment_Status", "Unpaid"
    SetCellByHeader lr.Range, salesTbl, "Notes", notesText
    SetCellByHeader lr.Range, salesTbl, "Created_At", Now
    SetCellByHeader lr.Range, salesTbl, "Updated_At", Now

End Sub

Private Function DeductProductStock( _
    ByVal prodTbl As ListObject, _
    ByVal prodRows As Object, _
    ByVal sku As String, _
    ByVal qty As Double) As Double

    Dim rr As Range
    Dim colStock As Long

    colStock = GetHeaderColumn(prodTbl, "Current_Stock")

    If prodRows.Exists(UCase$(sku)) Then
        Set rr = prodRows(UCase$(sku))
        rr.Cells(1, colStock).value = NzDbl(rr.Cells(1, colStock).value) - qty
        DeductProductStock = NzDbl(rr.Cells(1, colStock).value)
    Else
        DeductProductStock = 0
    End If

End Function

Private Sub WriteInventoryLogRow( _
    ByVal logTbl As ListObject, _
    ByVal salesDate As Date, _
    ByVal soNo As String, _
    ByVal productID As String, _
    ByVal sku As String, _
    ByVal productName As String, _
    ByVal qty As Double, _
    ByVal balanceAfter As Double, _
    ByVal unitCost As Double, _
    ByVal customerID As String, _
    ByVal customerName As String, _
    ByVal notesText As String)

    Dim lr As ListRow
    Set lr = logTbl.ListRows.Add

    SetCellByHeader lr.Range, logTbl, "Log_ID", NextLogID(logTbl)
    SetCellByHeader lr.Range, logTbl, "Log_Date", salesDate
    SetCellByHeader lr.Range, logTbl, "Tran_Type", "SALE"
    SetCellByHeader lr.Range, logTbl, "Ref_No", soNo
    SetCellByHeader lr.Range, logTbl, "Product_ID", productID
    SetCellByHeader lr.Range, logTbl, "SKU", sku
    SetCellByHeader lr.Range, logTbl, "Product_Name", productName
    SetCellByHeader lr.Range, logTbl, "Qty_Change", -qty
    SetCellByHeader lr.Range, logTbl, "Balance_After", balanceAfter
    SetCellByHeader lr.Range, logTbl, "Unit_Cost", unitCost
    SetCellByHeader lr.Range, logTbl, "Total_Cost", unitCost * qty
    SetCellByHeader lr.Range, logTbl, "Party_Type", "CUSTOMER"
    SetCellByHeader lr.Range, logTbl, "Party_ID", customerID
    SetCellByHeader lr.Range, logTbl, "Party_Name", customerName
    SetCellByHeader lr.Range, logTbl, "Notes", notesText
    SetCellByHeader lr.Range, logTbl, "Created_At", Now

End Sub

' ===============================
' Validation Report
' ===============================
Private Sub WriteValidationReport(ByVal results As Collection)

    Dim ws As Worksheet
    Dim i As Long, r As Long
    Dim res As Object

    Set ws = GetOrCreateSheet(WS_VALIDATION)
    ws.Cells.Clear

    AddValidationBackLink ws

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

Private Sub AddValidationBackLink(ByVal ws As Worksheet)

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

    ws.Range("A2").value = "Validation Details"
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Size = 14

End Sub

Private Sub ClearValidationSheetNav()

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_VALIDATION)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    On Error Resume Next
    ws.Hyperlinks.Delete
    On Error GoTo 0

End Sub

Private Function HasValidationErrors(ByVal ws As Worksheet) As Boolean

    Dim lastRow As Long, i As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 4 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(i, 2).value))) = "ERROR" Then
            HasValidationErrors = True
            Exit Function
        End If
    Next i

End Function

' ===============================
' Helpers
' ===============================
Private Sub ClearValidationLink()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)

    On Error Resume Next
    ws.Range(UI_CELL_VALIDATION).Hyperlinks.Delete
    On Error GoTo 0
End Sub

Private Function GetImportFilePath() As String
    If Trim$(gImportFilePath) <> "" Then
        GetImportFilePath = gImportFilePath
    Else
        GetImportFilePath = Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_FILE).value))
    End If
End Function

Private Function GetImportDataSheet(ByVal wb As Workbook) As Worksheet

    On Error Resume Next
    Set GetImportDataSheet = wb.Worksheets("Data")
    On Error GoTo 0

    If GetImportDataSheet Is Nothing Then
        Set GetImportDataSheet = wb.Worksheets(1)
    End If

End Function

Private Function NormalizeHeaderName(ByVal txt As String) As String

    txt = UCase$(Trim$(txt))
    txt = Replace(txt, "*", "")
    txt = Replace(txt, "(AUTO IF BLANK)", "")
    txt = Replace(txt, "(AUTO)", "")
    txt = Trim$(txt)

    NormalizeHeaderName = txt

End Function

Private Function BuildHeaderMap(ByVal ws As Worksheet) As Object

    Dim dict As Object
    Dim lastCol As Long, c As Long
    Dim h As String

    Set dict = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        h = NormalizeHeaderName(CStr(ws.Cells(1, c).value))
        If h <> "" Then dict(h) = c
    Next c

    Set BuildHeaderMap = dict

End Function

Private Function CheckRequiredHeaders(ByVal hdr As Object, ByVal results As Collection) As Boolean

    Dim req As Variant, i As Long
    Dim res As Object

    req = Array("SALES_DATE", "CUSTOMER_NAME", "SKU", "QTY")

    For i = LBound(req) To UBound(req)
        If Not hdr.Exists(req(i)) Then
            Set res = CreateObject("Scripting.Dictionary")
            FillEmptyValidationRow res
            res("Row_No") = 0
            res("Status") = "ERROR"
            res("Message") = "Missing required header: " & req(i)
            results.Add res
        End If
    Next i

    CheckRequiredHeaders = (results.Count = 0)

End Function

Private Sub FillEmptyValidationRow(ByVal res As Object)

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

Private Function RowHasAnyData(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal maxCol As Long) As Boolean

    Dim c As Long
    For c = 1 To maxCol
        If Trim$(CStr(ws.Cells(rowNum, c).value)) <> "" Then
            RowHasAnyData = True
            Exit Function
        End If
    Next c

End Function

Private Function GetCellValue(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object, ByVal fieldName As String) As Variant

    Dim k As String
    k = UCase$(fieldName)

    If hdr.Exists(k) Then
        GetCellValue = ws.Cells(rowNum, hdr(k)).value
    Else
        GetCellValue = Empty
    End If

End Function

Private Function GetCellString(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object, ByVal fieldName As String) As String
    GetCellString = Trim$(CStr(GetCellValue(ws, rowNum, hdr, fieldName)))
End Function

Private Function AppendMsg(ByVal baseMsg As String, ByVal addMsg As String) As String

    If baseMsg = "" Then
        AppendMsg = addMsg
    Else
        AppendMsg = baseMsg & " | " & addMsg
    End If

End Function

Private Function NzDbl(ByVal v As Variant) As Double

    If IsError(v) Then
        NzDbl = 0
    ElseIf IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0
    End If

End Function

Private Function GetTableSafe(ByVal wsName As String, ByVal preferredTableName As String) As ListObject

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set GetTableSafe = ws.ListObjects(preferredTableName)
    On Error GoTo 0

    If GetTableSafe Is Nothing Then
        If ws.ListObjects.Count > 0 Then
            Set GetTableSafe = ws.ListObjects(1)
        End If
    End If

End Function

Private Function GetHeaderColumn(ByVal lo As ListObject, ByVal headerName As String) As Long

    Dim i As Long

    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(i).name), Trim$(headerName), vbTextCompare) = 0 Then
            GetHeaderColumn = i
            Exit Function
        End If
    Next i

End Function

Private Sub SetCellByHeader(ByVal rowRange As Range, ByVal lo As ListObject, ByVal headerName As String, ByVal newValue As Variant)

    Dim c As Long
    c = GetHeaderColumn(lo, headerName)
    If c > 0 Then rowRange.Cells(1, c).value = newValue

End Sub

Private Function BuildProductDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object, prod As Object
    Dim i As Long
    Dim cPID As Long, cSKU As Long, cPN As Long, cUC As Long, cSP As Long, cCS As Long
    Dim keySKU As String

    Set dict = CreateObject("Scripting.Dictionary")

    cPID = GetHeaderColumn(prodTbl, "Product_ID")
    cSKU = GetHeaderColumn(prodTbl, "SKU")
    cPN = GetHeaderColumn(prodTbl, "Product_Name")
    cUC = GetHeaderColumn(prodTbl, "Unit_Cost")
    cSP = GetHeaderColumn(prodTbl, "Selling_Price")
    cCS = GetHeaderColumn(prodTbl, "Current_Stock")

    For i = 1 To prodTbl.ListRows.Count

        keySKU = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))

        If keySKU <> "" Then
            Set prod = CreateObject("Scripting.Dictionary")
            prod("Product_ID") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cPID).value))
            prod("Product_Name") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cPN).value))
            prod("Unit_Cost") = NzDbl(prodTbl.DataBodyRange.Cells(i, cUC).value)
            prod("Selling_Price") = NzDbl(prodTbl.DataBodyRange.Cells(i, cSP).value)
            prod("Current_Stock") = NzDbl(prodTbl.DataBodyRange.Cells(i, cCS).value)

            If Not dict.Exists(keySKU) Then
                dict.Add keySKU, prod
            End If
        End If

    Next i

    Set BuildProductDict = dict

End Function

Private Function BuildProductRowDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cSKU As Long
    Dim keySKU As String

    Set dict = CreateObject("Scripting.Dictionary")
    cSKU = GetHeaderColumn(prodTbl, "SKU")

    For i = 1 To prodTbl.ListRows.Count
        keySKU = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))
        If keySKU <> "" Then
            If Not dict.Exists(keySKU) Then
                dict.Add keySKU, prodTbl.DataBodyRange.Rows(i)
            End If
        End If
    Next i

    Set BuildProductRowDict = dict

End Function

Private Function BuildTempStockDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cSKU As Long, cCS As Long
    Dim keySKU As String

    Set dict = CreateObject("Scripting.Dictionary")
    cSKU = GetHeaderColumn(prodTbl, "SKU")
    cCS = GetHeaderColumn(prodTbl, "Current_Stock")

    For i = 1 To prodTbl.ListRows.Count
        keySKU = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))
        If keySKU <> "" Then
            dict(keySKU) = NzDbl(prodTbl.DataBodyRange.Cells(i, cCS).value)
        End If
    Next i

    Set BuildTempStockDict = dict

End Function

Private Function BuildCustomerDict(ByVal custTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cID As Long, cName As Long
    Dim keyName As String

    Set dict = CreateObject("Scripting.Dictionary")

    If custTbl Is Nothing Then
        Set BuildCustomerDict = dict
        Exit Function
    End If

    If custTbl.DataBodyRange Is Nothing Then
        Set BuildCustomerDict = dict
        Exit Function
    End If

    cID = GetHeaderColumn(custTbl, "Customer_ID")
    cName = GetHeaderColumn(custTbl, "Customer_Name")

    If cID = 0 Or cName = 0 Then
        Set BuildCustomerDict = dict
        Exit Function
    End If

    For i = 1 To custTbl.ListRows.Count
        keyName = UCase$(Trim$(CStr(custTbl.DataBodyRange.Cells(i, cName).value)))
        If keyName <> "" Then
            dict(keyName) = Trim$(CStr(custTbl.DataBodyRange.Cells(i, cID).value))
        End If
    Next i

    Set BuildCustomerDict = dict

End Function

Private Function BuildExistingSalesOrderDict(ByVal salesTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cSO As Long
    Dim keySO As String

    Set dict = CreateObject("Scripting.Dictionary")

    cSO = GetHeaderColumn(salesTbl, "Sales_Order_No")
    If cSO = 0 Then
        Set BuildExistingSalesOrderDict = dict
        Exit Function
    End If

    If salesTbl.DataBodyRange Is Nothing Then
        Set BuildExistingSalesOrderDict = dict
        Exit Function
    End If

    For i = 1 To salesTbl.ListRows.Count
        keySO = UCase$(Trim$(CStr(salesTbl.DataBodyRange.Cells(i, cSO).value)))
        If keySO <> "" Then dict(keySO) = True
    Next i

    Set BuildExistingSalesOrderDict = dict

End Function

Private Function GetNextSalesOrderSeq(ByVal salesTbl As ListObject) As Long

    Dim i As Long, cSO As Long
    Dim soNo As String, n As Long, maxN As Long
    Dim digitsOnly As String

    cSO = GetHeaderColumn(salesTbl, "Sales_Order_No")
    If cSO = 0 Then
        GetNextSalesOrderSeq = 1
        Exit Function
    End If

    If salesTbl.DataBodyRange Is Nothing Then
        GetNextSalesOrderSeq = 1
        Exit Function
    End If

    For i = 1 To salesTbl.ListRows.Count
        soNo = Trim$(CStr(salesTbl.DataBodyRange.Cells(i, cSO).value))
        digitsOnly = ExtractDigits(soNo)
        If digitsOnly <> "" Then
            n = CLng(digitsOnly)
            If n > maxN Then maxN = n
        End If
    Next i

    GetNextSalesOrderSeq = maxN + 1

End Function

Private Function ExtractDigits(ByVal txt As String) As String

    Dim i As Long, ch As String

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch Like "#" Then ExtractDigits = ExtractDigits & ch
    Next i

End Function

Private Function NextLogID(ByVal logTbl As ListObject) As String

    Dim c As Long, i As Long
    Dim txt As String, n As Long, maxN As Long

    c = GetHeaderColumn(logTbl, "Log_ID")
    If c = 0 Or logTbl.DataBodyRange Is Nothing Then
        NextLogID = "L000001"
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

    NextLogID = "L" & Format$(maxN + 1, "000000")

End Function

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetOrCreateSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheet Is Nothing Then
        Set GetOrCreateSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheet.name = sheetName
    End If

End Function


Private Sub GetSalesValidationCounts(ByRef errCount As Long, ByRef validCount As Long)

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
