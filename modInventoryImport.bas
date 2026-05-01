Attribute VB_Name = "modInventoryImport"
Option Explicit

Private gInventoryImportFilePath As String

Private Const WS_IMPORT_UI As String = "Import_UI"
Private Const WS_VALIDATION As String = "Import_Validation"
Private Const WS_PRODUCTS_DB As String = "Products_DB"
Private Const WS_INVENTORY_LOG As String = "Inventory_Log"

Private Const TBL_PRODUCTS As String = "tblProducts"
Private Const TBL_LOG As String = "tblInventoryLog"

Private Const UI_CELL_TYPE As String = "B20"
Private Const UI_CELL_FILE As String = "B21"
Private Const UI_CELL_VALIDATION As String = "B22"
Private Const UI_CELL_RESULT As String = "B23"
Private Const UI_CELL_OPTION As String = "B24"

Private Const IMPORT_TYPE_STANDARD As String = "Inventory Import"
Private Const IMPORT_TYPE_SHOPIFY As String = "Shopify Inventory Import"

Private Const RUN_OPTION_STRICT As String = "STOP_ON_ERROR"
Private Const RUN_OPTION_VALID_ONLY As String = "IMPORT_VALID_ONLY"

Private gSelectedShopifyInventoryCol As Long

Public Sub InventoryImport_DownloadTemplate()

    Dim importType As String
    importType = Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_TYPE).value))

    Select Case importType
        Case IMPORT_TYPE_STANDARD
            DownloadStandardInventoryTemplate

        Case IMPORT_TYPE_SHOPIFY
            MsgBox "Shopify Inventory Import does not require a template." & vbCrLf & _
                   "Please export the inventory CSV directly from Shopify.", vbInformation

        Case Else
            MsgBox "Please select a valid Inventory Import Type.", vbExclamation
    End Select

End Sub

Public Sub InventoryImport_SelectFile()

    Dim fd As Object
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)
    Set fd = Application.FileDialog(3)

    With fd
        .Title = "Select Inventory Import File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel and CSV Files", "*.xlsx;*.xls;*.csv"
        .Filters.Add "All Files", "*.*"

        If .Show <> -1 Then Exit Sub

        gInventoryImportFilePath = .SelectedItems(1)
    End With

    ws.Range(UI_CELL_FILE).value = gInventoryImportFilePath
    ws.Range(UI_CELL_VALIDATION).ClearContents
    ws.Range(UI_CELL_RESULT).ClearContents

    InventoryImport_ClearValidationLink

End Sub

Public Sub InventoryImport_Validate()

    Dim ok As Boolean

    ok = ValidateInventoryImportFile(True)

    If ok Then
        MsgBox "Inventory validation passed.", vbInformation
    Else
        MsgBox "Inventory validation finished with errors. Click Last Validation to view details.", vbExclamation
    End If

End Sub

Public Sub InventoryImport_Run()

    Dim ok As Boolean
    Dim runOption As String
    Dim validCount As Long
    Dim errCount As Long

    ok = ValidateInventoryImportFile(True)

    runOption = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_OPTION).value)))
    If runOption = "" Then runOption = RUN_OPTION_STRICT

    GetInventoryValidationCounts errCount, validCount

    If runOption = RUN_OPTION_STRICT Then
        If Not ok Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT BLOCKED - VALIDATION FAILED"
            MsgBox "Inventory import stopped because validation failed.", vbExclamation
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

    RunValidatedInventoryImport runOption

End Sub

Public Sub InventoryImport_Clear()

    gInventoryImportFilePath = ""

    With ThisWorkbook.Worksheets(WS_IMPORT_UI)
        .Range(UI_CELL_FILE).ClearContents
        .Range(UI_CELL_VALIDATION).ClearContents
        .Range(UI_CELL_RESULT).ClearContents
        .Range(UI_CELL_VALIDATION).Interior.Pattern = xlNone
        .Range(UI_CELL_VALIDATION).Font.Underline = xlUnderlineStyleNone
        .Range(UI_CELL_VALIDATION).Font.ColorIndex = xlAutomatic
    End With

    InventoryImport_ClearValidationLink

End Sub

Private Sub DownloadStandardInventoryTemplate()

    Dim wb As Workbook
    Dim wsInst As Worksheet
    Dim wsData As Worksheet
    Dim savePath As Variant
    Dim prodTbl As ListObject

    On Error GoTo ErrHandler

    Set prodTbl = GetTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)

    If prodTbl Is Nothing Then
        MsgBox "Products_DB table not found.", vbCritical
        Exit Sub
    End If

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="Inventory_Import_Template.xlsx", _
        FileFilter:="Excel Files (*.xlsx), *.xlsx")

    If VarType(savePath) = vbBoolean Then Exit Sub

    ProgressStart "Building Inventory Template", "Preparing workbook..."

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = Workbooks.Add

    ProgressStep "Building instruction sheet..."
    Set wsInst = wb.Sheets(1)
    wsInst.name = "Instructions"
    BuildInventoryInstructionsSheet wsInst

    ProgressStep "Building inventory template rows..."
    Set wsData = wb.Sheets.Add(After:=wsInst)
    wsData.name = "Data"
    BuildInventoryDataSheet wsData, prodTbl

    ProgressStep "Saving template..."
    wb.SaveAs CStr(savePath), FileFormat:=xlOpenXMLWorkbook
    wb.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    MsgBox "Inventory import template downloaded.", vbInformation
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ProgressEnd
    MsgBox "Download template error " & Err.Number & ": " & Err.Description, vbCritical

End Sub

Private Sub BuildInventoryInstructionsSheet(ByVal ws As Worksheet)

    ws.Cells.Clear

    ws.Range("A1").value = "Inventory Import Instructions"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A3:D3").value = Array("Field", "Required?", "Auto / Default", "Explanation")
    ws.Range("A3:D3").Font.Bold = True
    ws.Range("A3:D3").Interior.Color = RGB(217, 225, 242)

    ws.Range("A4:D4").value = Array("SKU", "Required", "", "Must exist in Products_DB.")
    ws.Range("A5:D5").value = Array("Product_Name", "Optional", "From Products_DB", "For reference only.")
    ws.Range("A6:D6").value = Array("Current_Stock", "Optional", "From Products_DB", "For reference only.")
    ws.Range("A7:D7").value = Array("New_Stock", "Required", "", "The new SimpleERP stock after import.")
    ws.Range("A8:D8").value = Array("Notes", "Optional", "", "Additional note for inventory log.")

    ws.Columns("A").ColumnWidth = 20
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 24
    ws.Columns("D").ColumnWidth = 60
    ws.Range("A3:D8").Borders.LineStyle = xlContinuous
    ws.Range("A3:D8").WrapText = True

End Sub

Private Sub BuildInventoryDataSheet(ByVal ws As Worksheet, ByVal prodTbl As ListObject)

    Dim cSKU As Long, cName As Long, cStock As Long
    Dim i As Long, r As Long
    Dim skuText As String

    ws.Cells.Clear

    ws.Range("A1:E1").value = Array("SKU *", "Product_Name", "Current_Stock", "New_Stock *", "Notes")
    ws.Range("A2:E2").value = Array("Required", "Reference", "Reference", "Required", "Optional")

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(217, 225, 242)
    ws.Range("A1").Font.Color = RGB(192, 0, 0)
    ws.Range("D1").Font.Color = RGB(192, 0, 0)
    ws.Rows(2).Font.Italic = True
    ws.Rows(2).Font.Color = RGB(90, 90, 90)

    cSKU = GetHeaderColumn(prodTbl, "SKU")
    cName = GetHeaderColumn(prodTbl, "Product_Name")
    cStock = GetHeaderColumn(prodTbl, "Current_Stock")

    If cSKU = 0 Or cName = 0 Or cStock = 0 Then
        MsgBox "Products_DB is missing SKU, Product_Name, or Current_Stock column.", vbCritical
        Exit Sub
    End If

    r = 3

    If Not prodTbl.DataBodyRange Is Nothing Then
        For i = 1 To prodTbl.ListRows.Count

            skuText = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value))
            ProgressUpdate "Building template rows", i, prodTbl.ListRows.Count, "SKU: " & skuText

            ws.Cells(r, 1).value = prodTbl.DataBodyRange.Cells(i, cSKU).value
            ws.Cells(r, 2).value = prodTbl.DataBodyRange.Cells(i, cName).value
            ws.Cells(r, 3).value = prodTbl.DataBodyRange.Cells(i, cStock).value
            ws.Cells(r, 4).value = ""
            ws.Cells(r, 5).value = ""

            r = r + 1
        Next i
    End If

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 35

    ws.Range("A1:E" & Application.Max(2, r - 1)).Borders.LineStyle = xlContinuous
    ws.Range("C:D").NumberFormat = "#,##0.00"

End Sub

Private Function ValidateInventoryImportFile(ByVal writeReport As Boolean) As Boolean

    Dim importType As String
    Dim filePath As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim hdr As Object
    Dim prodTbl As ListObject
    Dim productDict As Object
    Dim fileSKUDict As Object
    Dim results As Collection
    Dim res As Object
    Dim lastRow As Long
    Dim maxCol As Long
    Dim i As Long
    Dim validCount As Long
    Dim errCount As Long
    Dim skuText As String

    On Error GoTo ErrHandler

    ProgressStart "Validating Inventory Import", "Loading product master..."
    
    ' reset location selection each new validation
    gSelectedShopifyInventoryCol = 0

    importType = Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_TYPE).value))
    filePath = GetInventoryImportFilePath()

    If filePath = "" Then
        ProgressEnd
        MsgBox "Please select inventory import file first.", vbExclamation
        ValidateInventoryImportFile = False
        Exit Function
    End If

    Set prodTbl = GetTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)

    If prodTbl Is Nothing Then
        ProgressEnd
        MsgBox "Products_DB table not found.", vbCritical
        ValidateInventoryImportFile = False
        Exit Function
    End If

    Set productDict = BuildProductDict(prodTbl)
    Set fileSKUDict = CreateObject("Scripting.Dictionary")
    Set results = New Collection

    ProgressStep "Opening import file..."
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)

    If importType = IMPORT_TYPE_STANDARD Then
        Set ws = GetDataSheet(wb)
    ElseIf importType = IMPORT_TYPE_SHOPIFY Then
        Set ws = wb.Worksheets(1)
    Else
        wb.Close False
        Set wb = Nothing
        ProgressEnd
        MsgBox "Invalid Inventory Import Type.", vbExclamation
        ValidateInventoryImportFile = False
        Exit Function
    End If

    If ws Is Nothing Then
        wb.Close False
        Set wb = Nothing
        ProgressEnd
        MsgBox "Cannot find import data sheet.", vbCritical
        ValidateInventoryImportFile = False
        Exit Function
    End If

    If Application.WorksheetFunction.CountA(ws.Rows(1)) = 0 Then
        wb.Close False
        Set wb = Nothing
        ProgressEnd
        MsgBox "Import file header row is empty.", vbExclamation
        ValidateInventoryImportFile = False
        Exit Function
    End If

    ProgressStep "Reading headers..."
    Set hdr = BuildHeaderMap(ws)

    If Not CheckRequiredInventoryHeaders(importType, hdr, results) Then
        wb.Close False
        Set wb = Nothing

        If writeReport Then WriteInventoryValidationReport results
        UpdateInventoryValidationUI False, 0, results.Count

        ProgressEnd
        ValidateInventoryImportFile = False
        Exit Function
    End If

    maxCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    lastRow = GetLastUsedRow(ws)

    If lastRow < 2 Then
        wb.Close False
        Set wb = Nothing
        ProgressEnd
        MsgBox "Import file has no data rows.", vbExclamation
        ValidateInventoryImportFile = False
        Exit Function
    End If

    If importType = IMPORT_TYPE_STANDARD Then

        For i = 3 To lastRow
            skuText = Trim$(CStr(ws.Cells(i, 1).value))
            ProgressUpdate "Validating rows", i - 2, lastRow - 2, "SKU: " & skuText

            If Not RowHasAnyData(ws, i, maxCol) Then GoTo NextStandard

            Set res = ValidateOneStandardInventoryRow(ws, i, hdr, productDict, fileSKUDict)
            results.Add res

            If UCase$(CStr(res("Status"))) = "VALID" Then
                validCount = validCount + 1
            Else
                errCount = errCount + 1
            End If

NextStandard:
        Next i

    ElseIf importType = IMPORT_TYPE_SHOPIFY Then

        For i = 2 To lastRow
            skuText = ""
            If hdr.Exists("SKU") Then skuText = Trim$(CStr(ws.Cells(i, hdr("SKU")).value))

            ProgressUpdate "Validating Shopify rows", i - 1, lastRow - 1, "SKU: " & skuText

            If Not RowHasAnyData(ws, i, maxCol) Then GoTo NextShopify

            Set res = ValidateOneShopifyInventoryRow(ws, i, hdr, productDict, fileSKUDict)
            results.Add res

            If UCase$(CStr(res("Status"))) = "VALID" Then
                validCount = validCount + 1
            Else
                errCount = errCount + 1
            End If

NextShopify:
        Next i

    End If

    wb.Close False
    Set wb = Nothing

    ProgressStep "Writing validation report..."
    If writeReport Then WriteInventoryValidationReport results
    UpdateInventoryValidationUI (errCount = 0), validCount, errCount

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    ValidateInventoryImportFile = (errCount = 0)
    Exit Function

ErrHandler:
    Dim errNum As Long
    Dim errDesc As String

    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0

    ProgressEnd

    MsgBox "Inventory validation error " & errNum & ": " & errDesc, vbCritical
    ValidateInventoryImportFile = False

End Function

Private Function ValidateOneStandardInventoryRow( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal hdr As Object, _
    ByVal productDict As Object, _
    ByVal fileSKUDict As Object) As Object

    Dim res As Object
    Dim sku As String
    Dim productName As String
    Dim oldStock As Double
    Dim newStock As Double
    Dim deltaQty As Double
    Dim notesText As String
    Dim msg As String
    Dim productRec As Object
    Dim rawNewStock As Variant

    Set res = CreateObject("Scripting.Dictionary")

    sku = UCase$(Trim$(GetCellString(ws, rowNum, hdr, "SKU")))
    rawNewStock = GetCellValue(ws, rowNum, hdr, "NEW_STOCK")
    notesText = Trim$(GetCellString(ws, rowNum, hdr, "NOTES"))

    If sku = "" Then
        msg = AppendMsg(msg, "SKU required")
    Else
        If fileSKUDict.Exists(sku) Then
            msg = AppendMsg(msg, "Duplicate SKU within import file")
        Else
            fileSKUDict.Add sku, True
        End If

        If productDict.Exists(sku) Then
            Set productRec = productDict(sku)
            productName = SafeDictValue(productRec, "Product_Name")
            oldStock = NzDbl(SafeDictValue(productRec, "Current_Stock"))
        Else
            msg = AppendMsg(msg, "SKU not found in Products_DB")
        End If
    End If

    If Trim$(CStr(rawNewStock)) = "" Then
        msg = AppendMsg(msg, "New_Stock required")
        newStock = 0
    ElseIf Not IsNumeric(rawNewStock) Then
        msg = AppendMsg(msg, "New_Stock must be numeric")
        newStock = 0
    Else
        newStock = CDbl(rawNewStock)
        If newStock < 0 Then msg = AppendMsg(msg, "New_Stock cannot be negative")
    End If

    deltaQty = newStock - oldStock

    FillInventoryResult res, rowNum, msg, sku, productName, oldStock, newStock, deltaQty, "STOCK_LOAD", notesText

    Set ValidateOneStandardInventoryRow = res

End Function

Private Function ValidateOneShopifyInventoryRow( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal hdr As Object, _
    ByVal productDict As Object, _
    ByVal fileSKUDict As Object) As Object

    Dim res As Object
    Dim sku As String
    Dim productName As String
    Dim oldStock As Double
    Dim newStock As Double
    Dim deltaQty As Double
    Dim locationText As String
    Dim notesText As String
    Dim msg As String
    Dim productRec As Object
    Dim availableRaw As Variant
    Dim availableCol As Long

    Set res = CreateObject("Scripting.Dictionary")

    sku = UCase$(Trim$(GetCellString(ws, rowNum, hdr, "SKU")))

    availableCol = GetShopifyAvailableColumn(hdr)

    If availableCol > 0 Then
        availableRaw = ws.Cells(rowNum, availableCol).value
        locationText = GetHeaderNameByColumn(ws, availableCol)
    Else
        availableRaw = Empty
        locationText = ""
    End If

    If sku = "" Then
        msg = AppendMsg(msg, "SKU required")
    Else
        If fileSKUDict.Exists(sku) Then
            msg = AppendMsg(msg, "Duplicate SKU within Shopify inventory file")
        Else
            fileSKUDict.Add sku, True
        End If

        If productDict.Exists(sku) Then
            Set productRec = productDict(sku)
            productName = SafeDictValue(productRec, "Product_Name")
            oldStock = NzDbl(SafeDictValue(productRec, "Current_Stock"))
        Else
            msg = AppendMsg(msg, "SKU not found in Products_DB")
        End If
    End If

    If Trim$(CStr(availableRaw)) = "" Then
        msg = AppendMsg(msg, "Stock value required")
        newStock = 0
    ElseIf Not IsNumeric(availableRaw) Then
        msg = AppendMsg(msg, "Stock value must be numeric")
        newStock = 0
    Else
        newStock = CDbl(availableRaw)
        If newStock < 0 Then msg = AppendMsg(msg, "Stock value cannot be negative")
    End If

    deltaQty = newStock - oldStock

    notesText = "Shopify inventory sync"
    If locationText <> "" Then notesText = notesText & " | Stock Column=" & locationText

    FillInventoryResult res, rowNum, msg, sku, productName, oldStock, newStock, deltaQty, "SHOPIFY_STOCK_SYNC", notesText

    Set ValidateOneShopifyInventoryRow = res

End Function

Private Sub FillInventoryResult( _
    ByVal res As Object, _
    ByVal rowNum As Long, _
    ByVal msg As String, _
    ByVal sku As String, _
    ByVal productName As String, _
    ByVal oldStock As Double, _
    ByVal newStock As Double, _
    ByVal deltaQty As Double, _
    ByVal tranType As String, _
    ByVal notesText As String)

    res("Row_No") = rowNum

    If msg = "" Then
        res("Status") = "VALID"
        res("Message") = "OK"
    Else
        res("Status") = "ERROR"
        res("Message") = msg
    End If

    res("Tran_Type") = tranType
    res("SKU") = sku
    res("Product_Name") = productName
    res("Old_Stock") = oldStock
    res("New_Stock") = newStock
    res("Delta_Qty") = deltaQty
    res("Notes") = notesText

End Sub

Private Sub RunValidatedInventoryImport(ByVal runOption As String)

    Dim wsVal As Worksheet
    Dim prodTbl As ListObject
    Dim logTbl As ListObject
    Dim r As Long
    Dim lastRow As Long
    Dim importCount As Long
    Dim rowObj As Object
    Dim skuText As String

    On Error GoTo ErrHandler

    Set wsVal = ThisWorkbook.Worksheets(WS_VALIDATION)
    Set prodTbl = GetTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)
    Set logTbl = GetTableSafe(WS_INVENTORY_LOG, TBL_LOG)

    If prodTbl Is Nothing Then
        MsgBox "Products_DB table not found.", vbCritical
        Exit Sub
    End If

    If logTbl Is Nothing Then
        MsgBox "Inventory_Log table not found.", vbCritical
        Exit Sub
    End If

    lastRow = wsVal.Cells(wsVal.Rows.Count, 1).End(xlUp).Row

    If lastRow < 4 Then
        MsgBox "No validation data found.", vbExclamation
        Exit Sub
    End If

    If runOption = RUN_OPTION_STRICT Then
        If HasInventoryValidationErrors(wsVal) Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT BLOCKED - VALIDATION FAILED"
            MsgBox "Validation contains errors. Please fix them first.", vbExclamation
            Exit Sub
        End If
    End If

    ProgressStart "Running Inventory Import", "Writing stock balances..."

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = 4 To lastRow

        skuText = Trim$(CStr(wsVal.Cells(r, 4).value))
        ProgressUpdate "Importing rows", r - 3, lastRow - 3, "SKU: " & skuText

        If UCase$(Trim$(CStr(wsVal.Cells(r, 2).value))) <> "VALID" Then GoTo nextRow

        Set rowObj = CreateObject("Scripting.Dictionary")
        rowObj("Tran_Type") = Trim$(CStr(wsVal.Cells(r, 3).value))
        rowObj("SKU") = Trim$(CStr(wsVal.Cells(r, 4).value))
        rowObj("Product_Name") = Trim$(CStr(wsVal.Cells(r, 5).value))
        rowObj("Old_Stock") = NzDbl(wsVal.Cells(r, 6).value)
        rowObj("New_Stock") = NzDbl(wsVal.Cells(r, 7).value)
        rowObj("Delta_Qty") = NzDbl(wsVal.Cells(r, 8).value)
        rowObj("Notes") = Trim$(CStr(wsVal.Cells(r, 9).value))

        UpdateProductStock prodTbl, rowObj("SKU"), rowObj("New_Stock")
        WriteInventoryImportLog prodTbl, logTbl, rowObj

        importCount = importCount + 1

nextRow:
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT SUCCESS - " & importCount & " valid row(s)"
    MsgBox "Inventory import completed: " & importCount & " valid row(s).", vbInformation
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ProgressEnd
    MsgBox "Inventory import error " & Err.Number & ": " & Err.Description, vbCritical

End Sub

Private Sub UpdateProductStock(ByVal prodTbl As ListObject, ByVal sku As String, ByVal newStock As Double)

    Dim rr As Range

    Set rr = FindProductRowBySKU(prodTbl, sku)
    If rr Is Nothing Then Exit Sub

    SetCellInRowByHeader rr, prodTbl, "Current_Stock", newStock
    SetCellInRowByHeader rr, prodTbl, "Updated_At", Now

End Sub

Private Sub WriteInventoryImportLog(ByVal prodTbl As ListObject, ByVal logTbl As ListObject, ByVal rowObj As Object)

    Dim rr As Range
    Dim lr As ListRow
    Dim productID As String
    Dim sku As String
    Dim productName As String
    Dim unitCost As Double
    Dim deltaQty As Double
    Dim newStock As Double
    Dim refNo As String

    sku = rowObj("SKU")
    productName = rowObj("Product_Name")
    deltaQty = rowObj("Delta_Qty")
    newStock = rowObj("New_Stock")

    Set rr = FindProductRowBySKU(prodTbl, sku)

    If Not rr Is Nothing Then
        productID = Trim$(CStr(GetCellFromRowByHeader(rr, prodTbl, "Product_ID")))
        unitCost = NzDbl(GetCellFromRowByHeader(rr, prodTbl, "Unit_Cost"))
    End If

    refNo = "INVIMP-" & Format$(Now, "yyyymmdd-hhnnss")

    Set lr = logTbl.ListRows.Add

    SetCellByHeader lr.Range, logTbl, "Log_ID", NextInventoryLogID(logTbl)
    SetCellByHeader lr.Range, logTbl, "Log_Date", Date
    SetCellByHeader lr.Range, logTbl, "Tran_Type", rowObj("Tran_Type")
    SetCellByHeader lr.Range, logTbl, "Ref_No", refNo
    SetCellByHeader lr.Range, logTbl, "Product_ID", productID
    SetCellByHeader lr.Range, logTbl, "SKU", sku
    SetCellByHeader lr.Range, logTbl, "Product_Name", productName
    SetCellByHeader lr.Range, logTbl, "Qty_Change", deltaQty
    SetCellByHeader lr.Range, logTbl, "Balance_After", newStock
    SetCellByHeader lr.Range, logTbl, "Unit_Cost", unitCost
    SetCellByHeader lr.Range, logTbl, "Total_Cost", unitCost * deltaQty
    SetCellByHeader lr.Range, logTbl, "Party_Type", "SYSTEM"
    SetCellByHeader lr.Range, logTbl, "Party_ID", ""
    SetCellByHeader lr.Range, logTbl, "Party_Name", "Inventory Import"
    SetCellByHeader lr.Range, logTbl, "Notes", rowObj("Notes") & " | Old Stock=" & rowObj("Old_Stock") & " | New Stock=" & rowObj("New_Stock")
    SetCellByHeader lr.Range, logTbl, "Created_At", Now

End Sub

Private Sub WriteInventoryValidationReport(ByVal results As Collection)

    Dim ws As Worksheet
    Dim r As Long
    Dim i As Long
    Dim res As Object

    Set ws = GetOrCreateSheet(WS_VALIDATION)
    ws.Cells.Clear

    AddValidationBackLink ws, "Inventory Import Validation Details"

    ws.Range("A3:J3").value = Array("Row_No", "Status", "Tran_Type", "SKU", "Product_Name", "Old_Stock", "New_Stock", "Delta_Qty", "Notes", "Message")
    ws.Rows(3).Font.Bold = True
    ws.Rows(3).Interior.Color = RGB(217, 225, 242)

    r = 4

    For i = 1 To results.Count
        Set res = results(i)

        ws.Cells(r, 1).value = SafeDictValue(res, "Row_No")
        ws.Cells(r, 2).value = SafeDictValue(res, "Status")
        ws.Cells(r, 3).value = SafeDictValue(res, "Tran_Type")
        ws.Cells(r, 4).value = SafeDictValue(res, "SKU")
        ws.Cells(r, 5).value = SafeDictValue(res, "Product_Name")
        ws.Cells(r, 6).value = SafeDictValue(res, "Old_Stock")
        ws.Cells(r, 7).value = SafeDictValue(res, "New_Stock")
        ws.Cells(r, 8).value = SafeDictValue(res, "Delta_Qty")
        ws.Cells(r, 9).value = SafeDictValue(res, "Notes")
        ws.Cells(r, 10).value = SafeDictValue(res, "Message")

        If UCase$(Trim$(CStr(ws.Cells(r, 2).value))) = "ERROR" Then
            ws.Rows(r).Interior.Color = RGB(255, 230, 230)
        Else
            ws.Rows(r).Interior.Color = RGB(226, 239, 218)
        End If

        r = r + 1
    Next i

    ws.Columns("A:J").AutoFit
    If r > 4 Then ws.Range("F4:H" & r - 1).NumberFormat = "#,##0.00"
    If r > 3 Then ws.Range("A3:J" & r - 1).Borders.LineStyle = xlContinuous

End Sub

Private Sub UpdateInventoryValidationUI(ByVal isValid As Boolean, ByVal validCount As Long, ByVal errCount As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)

    InventoryImport_ClearValidationLink

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

Private Sub InventoryImport_ClearValidationLink()

    On Error Resume Next
    ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_VALIDATION).Hyperlinks.Delete
    On Error GoTo 0

End Sub

Private Sub GetInventoryValidationCounts(ByRef errCount As Long, ByRef validCount As Long)

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

Private Function HasInventoryValidationErrors(ByVal ws As Worksheet) As Boolean

    Dim i As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 4 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(i, 2).value))) = "ERROR" Then
            HasInventoryValidationErrors = True
            Exit Function
        End If
    Next i

End Function

Private Function CheckRequiredInventoryHeaders(ByVal importType As String, ByVal hdr As Object, ByVal results As Collection) As Boolean

    Dim res As Object

    If importType = IMPORT_TYPE_STANDARD Then

        If Not hdr.Exists("SKU") Then
            Set res = CreateObject("Scripting.Dictionary")
            FillEmptyValidationResult res
            res("Row_No") = 0
            res("Status") = "ERROR"
            res("Message") = "Missing required header: SKU"
            results.Add res
        End If

        If Not hdr.Exists("NEW_STOCK") Then
            Set res = CreateObject("Scripting.Dictionary")
            FillEmptyValidationResult res
            res("Row_No") = 0
            res("Status") = "ERROR"
            res("Message") = "Missing required header: NEW_STOCK"
            results.Add res
        End If

    ElseIf importType = IMPORT_TYPE_SHOPIFY Then

        If Not hdr.Exists("SKU") Then
            Set res = CreateObject("Scripting.Dictionary")
            FillEmptyValidationResult res
            res("Row_No") = 0
            res("Status") = "ERROR"
            res("Message") = "Missing required header: SKU"
            results.Add res
        End If

        If GetShopifyAvailableColumn(hdr) = 0 Then
            Set res = CreateObject("Scripting.Dictionary")
            FillEmptyValidationResult res
            res("Row_No") = 0
            res("Status") = "ERROR"
            res("Message") = "Cannot find stock column. Expected Available, On hand, or Shopify location column after SKU."
            results.Add res
        End If

    End If

    CheckRequiredInventoryHeaders = (results.Count = 0)

End Function

Private Sub FillEmptyValidationResult(ByVal res As Object)

    res("Row_No") = 0
    res("Status") = ""
    res("Tran_Type") = ""
    res("SKU") = ""
    res("Product_Name") = ""
    res("Old_Stock") = 0
    res("New_Stock") = 0
    res("Delta_Qty") = 0
    res("Notes") = ""
    res("Message") = ""

End Sub

Private Function GetInventoryImportFilePath() As String

    If Trim$(gInventoryImportFilePath) <> "" Then
        GetInventoryImportFilePath = gInventoryImportFilePath
    Else
        GetInventoryImportFilePath = Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_FILE).value))
    End If

End Function

Private Function BuildProductDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object
    Dim rec As Object
    Dim i As Long
    Dim cSKU As Long
    Dim cName As Long
    Dim cStock As Long
    Dim skuKey As String

    Set dict = CreateObject("Scripting.Dictionary")

    cSKU = GetHeaderColumn(prodTbl, "SKU")
    cName = GetHeaderColumn(prodTbl, "Product_Name")
    cStock = GetHeaderColumn(prodTbl, "Current_Stock")

    If cSKU = 0 Or cName = 0 Or cStock = 0 Then
        Set BuildProductDict = dict
        Exit Function
    End If

    If prodTbl.DataBodyRange Is Nothing Then
        Set BuildProductDict = dict
        Exit Function
    End If

    For i = 1 To prodTbl.ListRows.Count
        skuKey = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))

        If skuKey <> "" Then
            Set rec = CreateObject("Scripting.Dictionary")
            rec.Add "SKU", skuKey
            rec.Add "Product_Name", Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cName).value))
            rec.Add "Current_Stock", NzDbl(prodTbl.DataBodyRange.Cells(i, cStock).value)

            If Not dict.Exists(skuKey) Then
                dict.Add skuKey, rec
            End If
        End If
    Next i

    Set BuildProductDict = dict

End Function

Private Function GetTableSafe(ByVal wsName As String, ByVal tableName As String) As ListObject

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set GetTableSafe = ws.ListObjects(tableName)
    On Error GoTo 0

    If GetTableSafe Is Nothing Then
        If ws.ListObjects.Count > 0 Then
            Set GetTableSafe = ws.ListObjects(1)
        End If
    End If

End Function

Private Function BuildHeaderMap(ByVal ws As Worksheet) As Object

    Dim dict As Object
    Dim lastCol As Long
    Dim c As Long
    Dim h As String

    Set dict = CreateObject("Scripting.Dictionary")

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        h = NormalizeHeader(CStr(ws.Cells(1, c).value))
        If h <> "" Then
            If Not dict.Exists(h) Then dict.Add h, c
        End If
    Next c

    Set BuildHeaderMap = dict

End Function

Private Function NormalizeHeader(ByVal txt As String) As String

    txt = UCase$(Trim$(txt))
    txt = Replace(txt, "*", "")
    txt = Replace(txt, "(NOT EDITABLE)", "")
    txt = Replace(txt, "(NOTEDITABLE)", "")
    txt = Replace(txt, "(CURRENT)", "")
    txt = Replace(txt, "(NEW)", "")
    txt = Replace(txt, " ", "_")
    txt = Replace(txt, "-", "_")
    txt = Replace(txt, "/", "_")
    txt = Replace(txt, ".", "")
    txt = Replace(txt, "(", "")
    txt = Replace(txt, ")", "")

    Do While InStr(txt, "__") > 0
        txt = Replace(txt, "__", "_")
    Loop

    Do While Len(txt) > 0 And Right$(txt, 1) = "_"
        txt = Left$(txt, Len(txt) - 1)
    Loop

    Do While Len(txt) > 0 And Left$(txt, 1) = "_"
        txt = Mid$(txt, 2)
    Loop

    NormalizeHeader = Trim$(txt)

End Function

Private Function GetCellValue(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object, ByVal headerName As String) As Variant

    Dim k As String
    k = NormalizeHeader(headerName)

    If hdr.Exists(k) Then
        GetCellValue = ws.Cells(rowNum, hdr(k)).value
    Else
        GetCellValue = Empty
    End If

End Function

Private Function GetCellString(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object, ByVal headerName As String) As String

    GetCellString = Trim$(CStr(GetCellValue(ws, rowNum, hdr, headerName)))

End Function

Private Function RowHasAnyData(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal maxCol As Long) As Boolean

    Dim c As Long

    For c = 1 To maxCol
        If Trim$(CStr(ws.Cells(rowNum, c).value)) <> "" Then
            RowHasAnyData = True
            Exit Function
        End If
    Next c

End Function

Private Function GetDataSheet(ByVal wb As Workbook) As Worksheet

    On Error Resume Next
    Set GetDataSheet = wb.Worksheets("Data")
    On Error GoTo 0

    If GetDataSheet Is Nothing Then
        Set GetDataSheet = wb.Worksheets(1)
    End If

End Function

Private Function GetLastUsedRow(ByVal ws As Worksheet) As Long

    Dim f As Range

    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                          MatchCase:=False)
    On Error GoTo 0

    If f Is Nothing Then
        GetLastUsedRow = 1
    Else
        GetLastUsedRow = f.Row
    End If

End Function

Private Function FindProductRowBySKU(ByVal prodTbl As ListObject, ByVal sku As String) As Range

    Dim cSKU As Long
    Dim i As Long

    cSKU = GetHeaderColumn(prodTbl, "SKU")

    If cSKU = 0 Or prodTbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To prodTbl.ListRows.Count
        If UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value))) = UCase$(Trim$(sku)) Then
            Set FindProductRowBySKU = prodTbl.DataBodyRange.Rows(i)
            Exit Function
        End If
    Next i

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

Private Function GetCellFromRowByHeader(ByVal rr As Range, ByVal lo As ListObject, ByVal headerName As String) As Variant

    Dim c As Long
    c = GetHeaderColumn(lo, headerName)

    If c > 0 Then
        GetCellFromRowByHeader = rr.Cells(1, c).value
    Else
        GetCellFromRowByHeader = Empty
    End If

End Function

Private Sub SetCellInRowByHeader(ByVal rr As Range, ByVal lo As ListObject, ByVal headerName As String, ByVal newValue As Variant)

    Dim c As Long
    c = GetHeaderColumn(lo, headerName)

    If c > 0 Then rr.Cells(1, c).value = newValue

End Sub

Private Sub SetCellByHeader(ByVal rowRange As Range, ByVal lo As ListObject, ByVal headerName As String, ByVal newValue As Variant)

    Dim c As Long
    c = GetHeaderColumn(lo, headerName)

    If c > 0 Then rowRange.Cells(1, c).value = newValue

End Sub

Private Function NextInventoryLogID(ByVal logTbl As ListObject) As String

    Dim c As Long
    Dim i As Long
    Dim txt As String
    Dim n As Long
    Dim maxN As Long

    c = GetHeaderColumn(logTbl, "Log_ID")

    If c = 0 Or logTbl.DataBodyRange Is Nothing Then
        NextInventoryLogID = "L000001"
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

    NextInventoryLogID = "L" & Format$(maxN + 1, "000000")

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

Private Sub AddValidationBackLink(ByVal ws As Worksheet, ByVal titleText As String)

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

    ws.Range("A2").value = titleText
    ws.Range("A2").Font.Bold = True
    ws.Range("A2").Font.Size = 14

End Sub

Private Function NzDbl(ByVal v As Variant) As Double

    If IsError(v) Then
        NzDbl = 0
    ElseIf IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0
    End If

End Function

Private Function AppendMsg(ByVal baseMsg As String, ByVal addMsg As String) As String

    If baseMsg = "" Then
        AppendMsg = addMsg
    Else
        AppendMsg = baseMsg & " | " & addMsg
    End If

End Function

Private Function SafeDictValue(ByVal dict As Object, ByVal key As String) As Variant

    If dict Is Nothing Then
        SafeDictValue = ""
    ElseIf dict.Exists(key) Then
        SafeDictValue = dict(key)
    Else
        SafeDictValue = ""
    End If

End Function

Private Function GetShopifyAvailableColumn(ByVal hdr As Object) As Long

    Dim preferred As Variant
    Dim exclude As Object
    Dim locationCols As Object
    Dim i As Long
    Dim k As String
    Dim candidateKey As Variant
    Dim candidateCol As Long
    Dim skuCol As Long

    If gSelectedShopifyInventoryCol > 0 Then
        GetShopifyAvailableColumn = gSelectedShopifyInventoryCol
        Exit Function
    End If

    preferred = Array("AVAILABLE", "AVAILABLE_STOCK", "ON_HAND", "ONHAND", "CURRENT", "CURRENT_STOCK")

    For i = LBound(preferred) To UBound(preferred)
        k = NormalizeHeader(CStr(preferred(i)))
        If hdr.Exists(k) Then
            gSelectedShopifyInventoryCol = CLng(hdr(k))
            GetShopifyAvailableColumn = gSelectedShopifyInventoryCol
            Exit Function
        End If
    Next i

    Set exclude = CreateObject("Scripting.Dictionary")
    exclude("HANDLE") = True
    exclude("TITLE") = True
    exclude("OPTION1_NAME") = True
    exclude("OPTION1_VALUE") = True
    exclude("OPTION2_NAME") = True
    exclude("OPTION2_VALUE") = True
    exclude("OPTION3_NAME") = True
    exclude("OPTION3_VALUE") = True
    exclude("SKU") = True
    exclude("HS_CODE") = True
    exclude("COO") = True
    exclude("LOCATION") = True
    exclude("COMMITTED") = True
    exclude("UNAVAILABLE") = True

    Set locationCols = CreateObject("Scripting.Dictionary")

    If hdr.Exists("SKU") Then
        skuCol = CLng(hdr("SKU"))

        For Each candidateKey In hdr.Keys
            candidateCol = CLng(hdr(candidateKey))

            If candidateCol > skuCol Then
                If Not exclude.Exists(UCase$(CStr(candidateKey))) Then
                    locationCols(CStr(candidateKey)) = candidateCol
                End If
            End If
        Next candidateKey
    End If

    If locationCols.Count = 1 Then
        For Each candidateKey In locationCols.Keys
            gSelectedShopifyInventoryCol = CLng(locationCols(candidateKey))
            GetShopifyAvailableColumn = gSelectedShopifyInventoryCol
            Exit Function
        Next candidateKey
    ElseIf locationCols.Count > 1 Then
        frmSelectInventoryLocation.SelectedColumn = 0
        frmSelectInventoryLocation.LoadLocations ActiveWorkbook.Worksheets(1), locationCols
        frmSelectInventoryLocation.Show vbModal

        gSelectedShopifyInventoryCol = frmSelectInventoryLocation.SelectedColumn
        GetShopifyAvailableColumn = gSelectedShopifyInventoryCol
        Exit Function
    End If

End Function

Private Function GetHeaderNameByColumn( _
    ByVal ws As Worksheet, _
    ByVal colNum As Long) As String

    If colNum > 0 Then
        GetHeaderNameByColumn = _
            Trim$(CStr(ws.Cells(1, colNum).value))
    Else
        GetHeaderNameByColumn = ""
    End If

End Function
