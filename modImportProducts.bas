Attribute VB_Name = "modImportProducts"
Option Explicit

Private gProductImportFilePath As String

Private Const WS_IMPORT_UI As String = "Import_UI"
Private Const WS_VALIDATION As String = "Import_Validation"
Private Const WS_PRODUCTS_DB As String = "Products_DB"
Private Const WS_LOG As String = "Inventory_Log"

Private Const TBL_PRODUCTS As String = "tblProducts"
Private Const TBL_LOG As String = "tblInventoryLog"

Private Const UI_CELL_TYPE As String = "B12"
Private Const UI_CELL_FILE As String = "B13"
Private Const UI_CELL_VALIDATION As String = "B14"
Private Const UI_CELL_RESULT As String = "B15"
Private Const UI_CELL_RUN_OPTION As String = "B16"

Private Const RUN_OPTION_STRICT As String = "STOP_ON_ERROR"
Private Const RUN_OPTION_VALID_ONLY As String = "IMPORT_VALID_ONLY"

Public Sub ProductImport_DownloadTemplate()

    Dim wb As Workbook
    Dim wsInst As Worksheet
    Dim wsData As Worksheet
    Dim savePath As Variant

    On Error GoTo ErrHandler

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="Product_Import_Template.xlsx", _
        FileFilter:="Excel Files (*.xlsx), *.xlsx")

    If VarType(savePath) = vbBoolean Then Exit Sub

    ProgressStart "Building Product Template", "Preparing workbook..."

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wb = Workbooks.Add

    ProgressStep "Building instruction sheet..."
    Set wsInst = wb.Sheets(1)
    wsInst.name = "Instructions"

    ProgressStep "Building data sheet..."
    Set wsData = wb.Sheets.Add(After:=wsInst)
    wsData.name = "Data"

    BuildProductImportInstructionsSheet wsInst
    BuildProductImportDataSheet wsData

    ProgressStep "Saving template..."
    wb.SaveAs CStr(savePath), FileFormat:=xlOpenXMLWorkbook
    wb.Close False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    MsgBox "Product import template downloaded.", vbInformation
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    ProgressEnd
    MsgBox "Download Template error: " & Err.Description, vbCritical

End Sub
Public Sub ProductImport_SelectFile()

    Dim fd As Object
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)
    Set fd = Application.FileDialog(3)

    With fd
        .Title = "Select Product Import File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel and CSV Files", "*.xlsx;*.xls;*.csv"
        .Filters.Add "All Files", "*.*"

        If .Show <> -1 Then Exit Sub

        gProductImportFilePath = .SelectedItems(1)
    End With

    ws.Range(UI_CELL_FILE).value = gProductImportFilePath
    ws.Range(UI_CELL_VALIDATION).ClearContents
    ws.Range(UI_CELL_RESULT).ClearContents

    ProductImport_ClearValidationLink

End Sub

Public Sub ProductImport_Validate()

    Dim ok As Boolean

    ok = ValidateProductImportFile(True)

    If ok Then
        MsgBox "Product validation passed.", vbInformation
    Else
        MsgBox "Product validation finished with errors. Click Last Validation to open Import_Validation.", vbExclamation
    End If

End Sub

Public Sub ProductImport_Run()

    Dim ok As Boolean
    Dim runOption As String
    Dim errCount As Long, validCount As Long

    ok = ValidateProductImportFile(True)
    runOption = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RUN_OPTION).value)))

    GetProductValidationCounts errCount, validCount

    If runOption = "" Then runOption = RUN_OPTION_STRICT

    If runOption = RUN_OPTION_STRICT Then
        If Not ok Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT BLOCKED - VALIDATION FAILED"
            MsgBox "Product import stopped because validation failed.", vbExclamation
            Exit Sub
        End If
    ElseIf runOption = RUN_OPTION_VALID_ONLY Then
        If validCount = 0 Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "NO VALID ROWS TO IMPORT"
            MsgBox "There are no VALID rows to import.", vbExclamation
            Exit Sub
        End If
    Else
        MsgBox "Invalid Run Option in B16. Use STOP_ON_ERROR or IMPORT_VALID_ONLY.", vbExclamation
        Exit Sub
    End If

    RunValidatedProductImport runOption

End Sub

Public Sub ProductImport_Clear()

    gProductImportFilePath = ""

    With ThisWorkbook.Worksheets(WS_IMPORT_UI)
        .Range(UI_CELL_FILE).ClearContents
        .Range(UI_CELL_VALIDATION).ClearContents
        .Range(UI_CELL_RESULT).ClearContents
        .Range(UI_CELL_VALIDATION).Interior.Pattern = xlNone
        .Range(UI_CELL_VALIDATION).Font.Underline = xlUnderlineStyleNone
        .Range(UI_CELL_VALIDATION).Font.ColorIndex = xlAutomatic
    End With

    ProductImport_ClearValidationLink

End Sub

Private Sub BuildProductImportInstructionsSheet(ByVal ws As Worksheet)

    Dim r As Long

    ws.Cells.Clear

    ws.Range("A1").value = "Product Import Instructions"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A3:D3").value = Array("Field", "Required?", "Auto / Default", "Explanation")
    With ws.Range("A3:D3")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
    End With

    r = 4
    WriteProductInstructionRow ws, r, "Product_ID", "Optional", "Auto-generated if blank for NEW rows", "If SKU already exists, Product_ID must match existing record when provided."
    r = r + 1
    WriteProductInstructionRow ws, r, "SKU", "Required", "", "Unique product key. Existing SKU means UPDATE. New SKU means NEW."
    r = r + 1
    WriteProductInstructionRow ws, r, "Product_Name", "Required", "", "Main product name."
    r = r + 1
    WriteProductInstructionRow ws, r, "Variant_Desc", "Optional", "", "Variant description such as color or size."
    r = r + 1
    WriteProductInstructionRow ws, r, "Category", "Optional", "", "Product category."
    r = r + 1
    WriteProductInstructionRow ws, r, "Unit_Cost", "Optional", "Default = 0", "Product unit cost."
    r = r + 1
    WriteProductInstructionRow ws, r, "Selling_Price", "Optional", "Default = 0", "Product selling price."
    r = r + 1
    WriteProductInstructionRow ws, r, "Opening_Stock", "Optional", "Default = 0", "For NEW: Current_Stock = Opening_Stock. For UPDATE: stock is adjusted by Opening_Stock delta."
    r = r + 1
    WriteProductInstructionRow ws, r, "Reorder_Level", "Optional", "Default = 0", "Reorder level."
    r = r + 1
    WriteProductInstructionRow ws, r, "Safety_Days_Override", "Optional", "Default = 0", "Safety days override."
    r = r + 1
    WriteProductInstructionRow ws, r, "Lead_Time_Days", "Optional", "Default = 0", "Lead time in days."
    r = r + 1
    WriteProductInstructionRow ws, r, "Active_Status", "Optional", "Default = Active", "Usually Active or Inactive."
    r = r + 1
    WriteProductInstructionRow ws, r, "Notes", "Optional", "", "Any additional notes."

    ws.Columns("A").ColumnWidth = 22
    ws.Columns("B").ColumnWidth = 14
    ws.Columns("C").ColumnWidth = 34
    ws.Columns("D").ColumnWidth = 68
    ws.Range("A3:D" & r).WrapText = True
    ws.Range("A3:D" & r).Borders.LineStyle = xlContinuous

End Sub

Private Sub WriteProductInstructionRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal fld As String, ByVal req As String, ByVal autoTxt As String, ByVal explainTxt As String)

    ws.Cells(rowNum, 1).value = fld
    ws.Cells(rowNum, 2).value = req
    ws.Cells(rowNum, 3).value = autoTxt
    ws.Cells(rowNum, 4).value = explainTxt

End Sub

Private Sub BuildProductImportDataSheet(ByVal ws As Worksheet)

    Dim headers As Variant
    Dim i As Long

    ws.Cells.Clear

    headers = Array( _
        "Product_ID", _
        "SKU *", _
        "Product_Name *", _
        "Variant_Desc", _
        "Category", _
        "Unit_Cost", _
        "Selling_Price", _
        "Opening_Stock", _
        "Reorder_Level", _
        "Safety_Days_Override", _
        "Lead_Time_Days", _
        "Active_Status", _
        "Notes")

    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i + 1).value = headers(i)
    Next i

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).Interior.Color = RGB(217, 225, 242)

    For i = 1 To UBound(headers) + 1
        If InStr(1, CStr(ws.Cells(1, i).value), "*", vbTextCompare) > 0 Then
            ws.Cells(1, i).Font.Color = RGB(192, 0, 0)
        End If
    Next i

    ws.Range("A2").value = "Optional"
    ws.Range("B2").value = "Required"
    ws.Range("C2").value = "Required"
    ws.Range("D2").value = "Optional"
    ws.Range("E2").value = "Optional"
    ws.Range("F2").value = "Optional"
    ws.Range("G2").value = "Optional"
    ws.Range("H2").value = "Optional"
    ws.Range("I2").value = "Optional"
    ws.Range("J2").value = "Optional"
    ws.Range("K2").value = "Optional"
    ws.Range("L2").value = "Optional"
    ws.Range("M2").value = "Optional"

    ws.Rows(2).Font.Italic = True
    ws.Rows(2).Font.Color = RGB(90, 90, 90)
    ws.Range("A1:M2").Borders.LineStyle = xlContinuous

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 20
    ws.Columns("C").ColumnWidth = 28
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 18
    ws.Columns("F").ColumnWidth = 12
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 20
    ws.Columns("K").ColumnWidth = 14
    ws.Columns("L").ColumnWidth = 14
    ws.Columns("M").ColumnWidth = 20

End Sub

Private Function ValidateProductImportFile(ByVal writeReport As Boolean) As Boolean

    Dim filePath As String
    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim hdr As Object
    Dim results As Collection
    Dim res As Object
    Dim prodTbl As ListObject
    Dim dictExistingSKU As Object, dictExistingPID As Object, dictExistingProduct As Object
    Dim dictFileSKU As Object, dictFilePID As Object
    Dim rowHasData As Boolean
    Dim errCount As Long, validCount As Long
    Dim errMsg As String

    On Error GoTo ErrHandler

    ProgressStart "Validating Product Import", "Loading product master..."

    filePath = GetProductImportFilePath()
    If filePath = "" Then
        ProgressEnd
        MsgBox "Please select file first.", vbExclamation
        ValidateProductImportFile = False
        Exit Function
    End If

    Set prodTbl = GetProductTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)
    If prodTbl Is Nothing Then
        ProgressEnd
        MsgBox "Products_DB table not found.", vbCritical
        ValidateProductImportFile = False
        Exit Function
    End If

    Set dictExistingSKU = BuildExistingProductSKUDict(prodTbl)
    Set dictExistingPID = BuildExistingProductIDDict(prodTbl)
    Set dictExistingProduct = BuildExistingProductRecordDict(prodTbl)
    Set dictFileSKU = CreateObject("Scripting.Dictionary")
    Set dictFilePID = CreateObject("Scripting.Dictionary")
    Set results = New Collection

    ProgressStep "Opening import file..."
    Set wb = Workbooks.Open(filePath)

    Set ws = GetImportDataSheetGeneric(wb)
    If ws Is Nothing Then
        wb.Close False
        Set wb = Nothing
        ProgressEnd
        MsgBox "Cannot find data sheet in import file.", vbCritical
        ValidateProductImportFile = False
        Exit Function
    End If

    ProgressStep "Reading headers..."
    Set hdr = BuildHeaderMapGeneric(ws)

    If Not CheckRequiredProductHeaders(hdr, results) Then
        wb.Close False
        Set wb = Nothing

        If writeReport Then WriteProductValidationReport results
        UpdateProductValidationUI False, 0, results.Count

        ProgressEnd
        ValidateProductImportFile = False
        Exit Function
    End If

    If WorksheetFunction.CountA(ws.Cells) = 0 Then
        wb.Close False
        Set wb = Nothing
        ProgressEnd
        MsgBox "Import file is empty.", vbExclamation
        ValidateProductImportFile = False
        Exit Function
    End If

    lastRow = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row

    For i = 3 To lastRow

        ProgressUpdate "Validating rows", i - 2, lastRow - 2, _
            "SKU: " & Trim$(CStr(GetCellStringGeneric(ws, i, hdr, "SKU")))

        rowHasData = RowHasAnyDataGeneric(ws, i, 13)
        If Not rowHasData Then GoTo nextRow

        Set res = ValidateSingleProductImportRow( _
            ws, i, hdr, dictExistingSKU, dictExistingPID, dictExistingProduct, dictFileSKU, dictFilePID)

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

    If writeReport Then WriteProductValidationReport results
    UpdateProductValidationUI (errCount = 0), validCount, errCount

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    ValidateProductImportFile = (errCount = 0)
    Exit Function

ErrHandler:
    errMsg = Err.Description

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0

    ProgressEnd
    MsgBox "Product validation error: " & errMsg, vbCritical
    ValidateProductImportFile = False

End Function

Private Function ValidateSingleProductImportRow( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal hdr As Object, _
    ByVal dictExistingSKU As Object, _
    ByVal dictExistingPID As Object, _
    ByVal dictExistingProduct As Object, _
    ByVal dictFileSKU As Object, _
    ByVal dictFilePID As Object) As Object

    Dim res As Object
    Dim productID As String, sku As String, productName As String
    Dim variantDesc As String, category As String
    Dim unitCost As Double, sellingPrice As Double, openingStock As Double
    Dim reorderLevel As Double, safetyDays As Double, leadTime As Double
    Dim activeStatus As String, notesText As String
    Dim actionType As String
    Dim msg As String
    Dim rec As Object

    Set res = CreateObject("Scripting.Dictionary")

    productID = GetCellStringGeneric(ws, rowNum, hdr, "Product_ID")
    sku = UCase$(Trim$(GetCellStringGeneric(ws, rowNum, hdr, "SKU")))
    productName = Trim$(GetCellStringGeneric(ws, rowNum, hdr, "Product_Name"))
    variantDesc = Trim$(GetCellStringGeneric(ws, rowNum, hdr, "Variant_Desc"))
    category = Trim$(GetCellStringGeneric(ws, rowNum, hdr, "Category"))
    unitCost = NzDblGeneric(GetCellValueGeneric(ws, rowNum, hdr, "Unit_Cost"))
    sellingPrice = NzDblGeneric(GetCellValueGeneric(ws, rowNum, hdr, "Selling_Price"))
    openingStock = NzDblGeneric(GetCellValueGeneric(ws, rowNum, hdr, "Opening_Stock"))
    reorderLevel = NzDblGeneric(GetCellValueGeneric(ws, rowNum, hdr, "Reorder_Level"))
    safetyDays = NzDblGeneric(GetCellValueGeneric(ws, rowNum, hdr, "Safety_Days_Override"))
    leadTime = NzDblGeneric(GetCellValueGeneric(ws, rowNum, hdr, "Lead_Time_Days"))
    activeStatus = Trim$(GetCellStringGeneric(ws, rowNum, hdr, "Active_Status"))
    notesText = Trim$(GetCellStringGeneric(ws, rowNum, hdr, "Notes"))

    If sku = "" Then msg = AppendMsgGeneric(msg, "SKU required")
    If productName = "" Then msg = AppendMsgGeneric(msg, "Product_Name required")

    If unitCost < 0 Then msg = AppendMsgGeneric(msg, "Unit_Cost cannot be negative")
    If sellingPrice < 0 Then msg = AppendMsgGeneric(msg, "Selling_Price cannot be negative")
    If openingStock < 0 Then msg = AppendMsgGeneric(msg, "Opening_Stock cannot be negative")
    If reorderLevel < 0 Then msg = AppendMsgGeneric(msg, "Reorder_Level cannot be negative")
    If safetyDays < 0 Then msg = AppendMsgGeneric(msg, "Safety_Days_Override cannot be negative")
    If leadTime < 0 Then msg = AppendMsgGeneric(msg, "Lead_Time_Days cannot be negative")

    If activeStatus = "" Then activeStatus = "Active"

    If sku <> "" Then
        If dictFileSKU.Exists(UCase$(sku)) Then
            msg = AppendMsgGeneric(msg, "Duplicate SKU within import file")
        Else
            dictFileSKU.Add UCase$(sku), True
        End If
    End If

    If productID <> "" Then
        If dictFilePID.Exists(UCase$(productID)) Then
            msg = AppendMsgGeneric(msg, "Duplicate Product_ID within import file")
        Else
            dictFilePID.Add UCase$(productID), True
        End If
    End If

    If sku <> "" And dictExistingSKU.Exists(UCase$(sku)) Then
        actionType = "UPDATE"

        If dictExistingProduct.Exists(UCase$(sku)) Then
            Set rec = dictExistingProduct(UCase$(sku))

            If productID <> "" Then
                If UCase$(productID) <> UCase$(rec("Product_ID")) Then
                    msg = AppendMsgGeneric(msg, "Product_ID does not match existing SKU record")
                End If
            Else
                productID = rec("Product_ID")
            End If
        End If

    Else
        actionType = "NEW"

        If productID <> "" Then
            If dictExistingPID.Exists(UCase$(productID)) Then
                msg = AppendMsgGeneric(msg, "Product_ID already exists in Products_DB")
            End If
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

    res("Action_Type") = actionType
    res("Product_ID") = productID
    res("SKU") = sku
    res("Product_Name") = productName
    res("Variant_Desc") = variantDesc
    res("Category") = category
    res("Unit_Cost") = unitCost
    res("Selling_Price") = sellingPrice
    res("Opening_Stock") = openingStock
    res("Reorder_Level") = reorderLevel
    res("Safety_Days_Override") = safetyDays
    res("Lead_Time_Days") = leadTime
    res("Active_Status") = activeStatus
    res("Notes") = notesText

    Set ValidateSingleProductImportRow = res

End Function

Private Sub RunValidatedProductImport(ByVal runOption As String)

    Dim wsVal As Worksheet
    Dim prodTbl As ListObject
    Dim logTbl As ListObject
    Dim r As Long, lastRow As Long
    Dim importCount As Long
    Dim nextPIDNum As Long
    Dim rowObj As Object
    Dim actionType As String
    Dim skuText As String

    On Error GoTo ErrHandler

    Set wsVal = ThisWorkbook.Worksheets(WS_VALIDATION)
    Set prodTbl = GetProductTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)
    Set logTbl = GetProductTableSafe(WS_LOG, TBL_LOG)

    If prodTbl Is Nothing Then
        MsgBox "Products_DB table not found.", vbCritical
        Exit Sub
    End If

    If logTbl Is Nothing Then
        MsgBox "Inventory_Log table not found.", vbCritical
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
        If HasProductValidationErrors(wsVal) Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT BLOCKED - VALIDATION FAILED"
            MsgBox "Validation still contains errors. Please fix them first.", vbExclamation
            Exit Sub
        End If
    End If

    nextPIDNum = GetNextProductIDSeq(prodTbl)

    ProgressStart "Running Product Import", "Writing product records..."

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = 4 To lastRow

        skuText = Trim$(CStr(wsVal.Cells(r, 5).value))
        ProgressUpdate "Importing rows", r - 3, lastRow - 3, "SKU: " & skuText

        If UCase$(Trim$(CStr(wsVal.Cells(r, 2).value))) <> "VALID" Then
            If runOption = RUN_OPTION_VALID_ONLY Then GoTo nextRow
            GoTo nextRow
        End If

        Set rowObj = CreateObject("Scripting.Dictionary")
        rowObj("Action_Type") = Trim$(CStr(wsVal.Cells(r, 3).value))
        rowObj("Product_ID") = Trim$(CStr(wsVal.Cells(r, 4).value))
        rowObj("SKU") = Trim$(CStr(wsVal.Cells(r, 5).value))
        rowObj("Product_Name") = Trim$(CStr(wsVal.Cells(r, 6).value))
        rowObj("Variant_Desc") = Trim$(CStr(wsVal.Cells(r, 7).value))
        rowObj("Category") = Trim$(CStr(wsVal.Cells(r, 8).value))
        rowObj("Unit_Cost") = NzDblGeneric(wsVal.Cells(r, 9).value)
        rowObj("Selling_Price") = NzDblGeneric(wsVal.Cells(r, 10).value)
        rowObj("Opening_Stock") = NzDblGeneric(wsVal.Cells(r, 11).value)
        rowObj("Reorder_Level") = NzDblGeneric(wsVal.Cells(r, 12).value)
        rowObj("Safety_Days_Override") = NzDblGeneric(wsVal.Cells(r, 13).value)
        rowObj("Lead_Time_Days") = NzDblGeneric(wsVal.Cells(r, 14).value)
        rowObj("Active_Status") = Trim$(CStr(wsVal.Cells(r, 15).value))
        rowObj("Notes") = Trim$(CStr(wsVal.Cells(r, 16).value))

        actionType = UCase$(Trim$(rowObj("Action_Type")))

        If actionType = "NEW" Then

            If rowObj("Product_ID") = "" Then
                rowObj("Product_ID") = GenerateProductID(nextPIDNum)
                nextPIDNum = nextPIDNum + 1
            End If

            rowObj("Current_Stock") = rowObj("Opening_Stock")
            WriteNewProductRow prodTbl, rowObj

            If rowObj("Opening_Stock") <> 0 Then
                WriteProductOpeningLog logTbl, rowObj("Product_ID"), rowObj("SKU"), rowObj("Product_Name"), _
                    rowObj("Opening_Stock"), rowObj("Current_Stock"), "OPENING", _
                    "Initial opening stock from Product Import"
            End If

        ElseIf actionType = "UPDATE" Then

            UpdateExistingProductRow prodTbl, logTbl, rowObj

        End If

        importCount = importCount + 1

nextRow:
    Next r

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT SUCCESS - " & importCount & " valid row(s)"
    MsgBox "Product import completed: " & importCount & " valid row(s).", vbInformation
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ProgressEnd
    MsgBox "Product import error: " & Err.Description, vbCritical

End Sub

Private Sub WriteNewProductRow(ByVal prodTbl As ListObject, ByVal rowObj As Object)

    Dim lr As ListRow
    Set lr = prodTbl.ListRows.Add

    SetCellByHeaderGeneric lr.Range, prodTbl, "Product_ID", rowObj("Product_ID")
    SetCellByHeaderGeneric lr.Range, prodTbl, "SKU", rowObj("SKU")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Product_Name", rowObj("Product_Name")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Variant_Desc", rowObj("Variant_Desc")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Category", rowObj("Category")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Unit_Cost", rowObj("Unit_Cost")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Selling_Price", rowObj("Selling_Price")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Opening_Stock", rowObj("Opening_Stock")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Current_Stock", rowObj("Current_Stock")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Reorder_Level", rowObj("Reorder_Level")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Safety_Days_Override", rowObj("Safety_Days_Override")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Lead_Time_Days", rowObj("Lead_Time_Days")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Active_Status", rowObj("Active_Status")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Notes", rowObj("Notes")
    SetCellByHeaderGeneric lr.Range, prodTbl, "Created_At", Now
    SetCellByHeaderGeneric lr.Range, prodTbl, "Updated_At", Now

End Sub

Private Sub UpdateExistingProductRow(ByVal prodTbl As ListObject, ByVal logTbl As ListObject, ByVal rowObj As Object)

    Dim rr As Range
    Dim openingOld As Double, currentOld As Double, currentNew As Double, deltaOpen As Double

    Set rr = FindProductRowBySKU(prodTbl, rowObj("SKU"))
    If rr Is Nothing Then Exit Sub

    openingOld = NzDblGeneric(GetCellFromRowByHeader(rr, prodTbl, "Opening_Stock"))
    currentOld = NzDblGeneric(GetCellFromRowByHeader(rr, prodTbl, "Current_Stock"))

    deltaOpen = rowObj("Opening_Stock") - openingOld
    currentNew = currentOld + deltaOpen
    If currentNew < 0 Then currentNew = 0

    SetCellInRowByHeader rr, prodTbl, "Product_ID", rowObj("Product_ID")
    SetCellInRowByHeader rr, prodTbl, "Product_Name", rowObj("Product_Name")
    SetCellInRowByHeader rr, prodTbl, "Variant_Desc", rowObj("Variant_Desc")
    SetCellInRowByHeader rr, prodTbl, "Category", rowObj("Category")
    SetCellInRowByHeader rr, prodTbl, "Unit_Cost", rowObj("Unit_Cost")
    SetCellInRowByHeader rr, prodTbl, "Selling_Price", rowObj("Selling_Price")
    SetCellInRowByHeader rr, prodTbl, "Opening_Stock", rowObj("Opening_Stock")
    SetCellInRowByHeader rr, prodTbl, "Current_Stock", currentNew
    SetCellInRowByHeader rr, prodTbl, "Reorder_Level", rowObj("Reorder_Level")
    SetCellInRowByHeader rr, prodTbl, "Safety_Days_Override", rowObj("Safety_Days_Override")
    SetCellInRowByHeader rr, prodTbl, "Lead_Time_Days", rowObj("Lead_Time_Days")
    SetCellInRowByHeader rr, prodTbl, "Active_Status", rowObj("Active_Status")
    SetCellInRowByHeader rr, prodTbl, "Notes", rowObj("Notes")
    SetCellInRowByHeader rr, prodTbl, "Updated_At", Now

    If deltaOpen <> 0 Then
        WriteProductOpeningLog logTbl, rowObj("Product_ID"), rowObj("SKU"), rowObj("Product_Name"), deltaOpen, currentNew, "OPENING_ADJ", _
            "Opening stock adjusted by Product Import: " & openingOld & " -> " & rowObj("Opening_Stock")
    End If

End Sub

Private Sub WriteProductOpeningLog( _
    ByVal logTbl As ListObject, _
    ByVal productID As String, _
    ByVal sku As String, _
    ByVal productName As String, _
    ByVal qtyChange As Double, _
    ByVal balanceAfter As Double, _
    ByVal tranType As String, _
    ByVal noteText As String)

    Dim lr As ListRow
    Set lr = logTbl.ListRows.Add

    SetCellByHeaderGeneric lr.Range, logTbl, "Log_ID", NextInventoryLogID(logTbl)
    SetCellByHeaderGeneric lr.Range, logTbl, "Log_Date", Date
    SetCellByHeaderGeneric lr.Range, logTbl, "Tran_Type", tranType
    SetCellByHeaderGeneric lr.Range, logTbl, "Ref_No", "PRODIMP-" & Format$(Now, "yyyymmdd-hhnnss")
    SetCellByHeaderGeneric lr.Range, logTbl, "Product_ID", productID
    SetCellByHeaderGeneric lr.Range, logTbl, "SKU", sku
    SetCellByHeaderGeneric lr.Range, logTbl, "Product_Name", productName
    SetCellByHeaderGeneric lr.Range, logTbl, "Qty_Change", qtyChange
    SetCellByHeaderGeneric lr.Range, logTbl, "Balance_After", balanceAfter
    SetCellByHeaderGeneric lr.Range, logTbl, "Unit_Cost", 0
    SetCellByHeaderGeneric lr.Range, logTbl, "Total_Cost", 0
    SetCellByHeaderGeneric lr.Range, logTbl, "Party_Type", "SYSTEM"
    SetCellByHeaderGeneric lr.Range, logTbl, "Party_ID", ""
    SetCellByHeaderGeneric lr.Range, logTbl, "Party_Name", "Product Import"
    SetCellByHeaderGeneric lr.Range, logTbl, "Notes", noteText
    SetCellByHeaderGeneric lr.Range, logTbl, "Created_At", Now

End Sub

Private Sub UpdateProductValidationUI(ByVal isValid As Boolean, ByVal validCount As Long, ByVal errCount As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)

    ProductImport_ClearValidationLink

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

Private Sub ProductImport_ClearValidationLink()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)

    On Error Resume Next
    ws.Range(UI_CELL_VALIDATION).Hyperlinks.Delete
    On Error GoTo 0

End Sub

Private Sub WriteProductValidationReport(ByVal results As Collection)

    Dim ws As Worksheet
    Dim i As Long, r As Long
    Dim res As Object

    Set ws = GetOrCreateSheetGeneric(WS_VALIDATION)
    ws.Cells.Clear

    AddValidationBackLinkGeneric ws, "Product Validation Details"

    ws.Range("A3:Q3").value = Array( _
        "Row_No", "Status", "Action_Type", "Product_ID", "SKU", "Product_Name", "Variant_Desc", _
        "Category", "Unit_Cost", "Selling_Price", "Opening_Stock", _
        "Reorder_Level", "Safety_Days_Override", "Lead_Time_Days", "Active_Status", "Notes", "Message")

    ws.Rows(3).Font.Bold = True
    ws.Rows(3).Interior.Color = RGB(217, 225, 242)

    r = 4
    For i = 1 To results.Count
        Set res = results(i)

        ws.Cells(r, 1).value = res("Row_No")
        ws.Cells(r, 2).value = res("Status")
        ws.Cells(r, 3).value = res("Action_Type")
        ws.Cells(r, 4).value = res("Product_ID")
        ws.Cells(r, 5).value = res("SKU")
        ws.Cells(r, 6).value = res("Product_Name")
        ws.Cells(r, 7).value = res("Variant_Desc")
        ws.Cells(r, 8).value = res("Category")
        ws.Cells(r, 9).value = res("Unit_Cost")
        ws.Cells(r, 10).value = res("Selling_Price")
        ws.Cells(r, 11).value = res("Opening_Stock")
        ws.Cells(r, 12).value = res("Reorder_Level")
        ws.Cells(r, 13).value = res("Safety_Days_Override")
        ws.Cells(r, 14).value = res("Lead_Time_Days")
        ws.Cells(r, 15).value = res("Active_Status")
        ws.Cells(r, 16).value = res("Notes")
        ws.Cells(r, 17).value = res("Message")

        If UCase$(Trim$(res("Status"))) = "ERROR" Then
            ws.Rows(r).Interior.Color = RGB(255, 230, 230)
        ElseIf UCase$(Trim$(res("Action_Type"))) = "UPDATE" Then
            ws.Rows(r).Interior.Color = RGB(255, 242, 204)
        ElseIf UCase$(Trim$(res("Action_Type"))) = "NEW" Then
            ws.Rows(r).Interior.Color = RGB(226, 239, 218)
        End If

        r = r + 1
    Next i

    ws.Columns("A:Q").AutoFit
If r > 4 Then
    ws.Range("I4:N" & r - 1).NumberFormat = "#,##0.00"
End If

If r > 3 Then
    ws.Range("A3:Q" & r - 1).Borders.LineStyle = xlContinuous
End If

End Sub

Private Function HasProductValidationErrors(ByVal ws As Worksheet) As Boolean

    Dim lastRow As Long, i As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 4 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(i, 2).value))) = "ERROR" Then
            HasProductValidationErrors = True
            Exit Function
        End If
    Next i

End Function

Private Sub GetProductValidationCounts(ByRef errCount As Long, ByRef validCount As Long)

    Dim ws As Worksheet
    Dim lastRow As Long, i As Long
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

Private Function GetProductImportFilePath() As String
    If Trim$(gProductImportFilePath) <> "" Then
        GetProductImportFilePath = gProductImportFilePath
    Else
        GetProductImportFilePath = Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_FILE).value))
    End If
End Function

Private Function GetProductTableSafe(ByVal wsName As String, ByVal preferredTableName As String) As ListObject

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set GetProductTableSafe = ws.ListObjects(preferredTableName)
    On Error GoTo 0

    If GetProductTableSafe Is Nothing Then
        If ws.ListObjects.Count > 0 Then
            Set GetProductTableSafe = ws.ListObjects(1)
        End If
    End If

End Function

Private Function BuildExistingProductSKUDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cSKU As Long
    Dim keySKU As String

    Set dict = CreateObject("Scripting.Dictionary")
    cSKU = GetHeaderColumnGeneric(prodTbl, "SKU")

    If cSKU = 0 Or prodTbl.DataBodyRange Is Nothing Then
        Set BuildExistingProductSKUDict = dict
        Exit Function
    End If

    For i = 1 To prodTbl.ListRows.Count
        keySKU = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))
        If keySKU <> "" Then dict(keySKU) = True
    Next i

    Set BuildExistingProductSKUDict = dict

End Function

Private Function BuildExistingProductIDDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cPID As Long
    Dim keyPID As String

    Set dict = CreateObject("Scripting.Dictionary")
    cPID = GetHeaderColumnGeneric(prodTbl, "Product_ID")

    If cPID = 0 Or prodTbl.DataBodyRange Is Nothing Then
        Set BuildExistingProductIDDict = dict
        Exit Function
    End If

    For i = 1 To prodTbl.ListRows.Count
        keyPID = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cPID).value)))
        If keyPID <> "" Then dict(keyPID) = True
    Next i

    Set BuildExistingProductIDDict = dict

End Function

Private Function BuildExistingProductRecordDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object, rec As Object
    Dim i As Long
    Dim cPID As Long, cSKU As Long, cPN As Long, cVD As Long, cCAT As Long
    Dim cUC As Long, cSP As Long, cOS As Long, cCS As Long, cRL As Long
    Dim cSD As Long, cLT As Long, cAS As Long, cNOTES As Long
    Dim sku As String

    Set dict = CreateObject("Scripting.Dictionary")

    cPID = GetHeaderColumnGeneric(prodTbl, "Product_ID")
    cSKU = GetHeaderColumnGeneric(prodTbl, "SKU")
    cPN = GetHeaderColumnGeneric(prodTbl, "Product_Name")
    cVD = GetHeaderColumnGeneric(prodTbl, "Variant_Desc")
    cCAT = GetHeaderColumnGeneric(prodTbl, "Category")
    cUC = GetHeaderColumnGeneric(prodTbl, "Unit_Cost")
    cSP = GetHeaderColumnGeneric(prodTbl, "Selling_Price")
    cOS = GetHeaderColumnGeneric(prodTbl, "Opening_Stock")
    cCS = GetHeaderColumnGeneric(prodTbl, "Current_Stock")
    cRL = GetHeaderColumnGeneric(prodTbl, "Reorder_Level")
    cSD = GetHeaderColumnGeneric(prodTbl, "Safety_Days_Override")
    cLT = GetHeaderColumnGeneric(prodTbl, "Lead_Time_Days")
    cAS = GetHeaderColumnGeneric(prodTbl, "Active_Status")
    cNOTES = GetHeaderColumnGeneric(prodTbl, "Notes")

    If cSKU = 0 Or prodTbl.DataBodyRange Is Nothing Then
        Set BuildExistingProductRecordDict = dict
        Exit Function
    End If

    For i = 1 To prodTbl.ListRows.Count
        sku = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))
        If sku <> "" Then
            Set rec = CreateObject("Scripting.Dictionary")
            rec("Product_ID") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cPID).value))
            rec("SKU") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value))
            rec("Product_Name") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cPN).value))
            rec("Variant_Desc") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cVD).value))
            rec("Category") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cCAT).value))
            rec("Unit_Cost") = NzDblGeneric(prodTbl.DataBodyRange.Cells(i, cUC).value)
            rec("Selling_Price") = NzDblGeneric(prodTbl.DataBodyRange.Cells(i, cSP).value)
            rec("Opening_Stock") = NzDblGeneric(prodTbl.DataBodyRange.Cells(i, cOS).value)
            rec("Current_Stock") = NzDblGeneric(prodTbl.DataBodyRange.Cells(i, cCS).value)
            rec("Reorder_Level") = NzDblGeneric(prodTbl.DataBodyRange.Cells(i, cRL).value)
            rec("Safety_Days_Override") = NzDblGeneric(prodTbl.DataBodyRange.Cells(i, cSD).value)
            rec("Lead_Time_Days") = NzDblGeneric(prodTbl.DataBodyRange.Cells(i, cLT).value)
            rec("Active_Status") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cAS).value))
            rec("Notes") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cNOTES).value))
            Set dict(sku) = rec
        End If
    Next i

    Set BuildExistingProductRecordDict = dict

End Function

Private Function GetNextProductIDSeq(ByVal prodTbl As ListObject) As Long

    Dim i As Long, cPID As Long
    Dim pid As String, digitsOnly As String
    Dim n As Long, maxN As Long

    cPID = GetHeaderColumnGeneric(prodTbl, "Product_ID")
    If cPID = 0 Or prodTbl.DataBodyRange Is Nothing Then
        GetNextProductIDSeq = 1
        Exit Function
    End If

    For i = 1 To prodTbl.ListRows.Count
        pid = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cPID).value))
        digitsOnly = ExtractDigitsGeneric(pid)
        If digitsOnly <> "" Then
            n = CLng(digitsOnly)
            If n > maxN Then maxN = n
        End If
    Next i

    GetNextProductIDSeq = maxN + 1

End Function

Private Function GenerateProductID(ByVal seqNum As Long) As String
    GenerateProductID = "P" & Format$(seqNum, "00000")
End Function

Private Function CheckRequiredProductHeaders(ByVal hdr As Object, ByVal results As Collection) As Boolean

    Dim req As Variant, i As Long
    Dim res As Object

    req = Array("SKU", "PRODUCT_NAME")

    For i = LBound(req) To UBound(req)
        If Not hdr.Exists(req(i)) Then
            Set res = CreateObject("Scripting.Dictionary")
            FillEmptyProductValidationRow res
            res("Row_No") = 0
            res("Status") = "ERROR"
            res("Message") = "Missing required header: " & req(i)
            results.Add res
        End If
    Next i

    CheckRequiredProductHeaders = (results.Count = 0)

End Function

Private Sub FillEmptyProductValidationRow(ByVal res As Object)

    res("Row_No") = 0
    res("Status") = ""
    res("Action_Type") = ""
    res("Product_ID") = ""
    res("SKU") = ""
    res("Product_Name") = ""
    res("Variant_Desc") = ""
    res("Category") = ""
    res("Unit_Cost") = 0
    res("Selling_Price") = 0
    res("Opening_Stock") = 0
    res("Reorder_Level") = 0
    res("Safety_Days_Override") = 0
    res("Lead_Time_Days") = 0
    res("Active_Status") = ""
    res("Notes") = ""
    res("Message") = ""

End Sub

Private Function GetImportDataSheetGeneric(ByVal wb As Workbook) As Worksheet

    On Error Resume Next
    Set GetImportDataSheetGeneric = wb.Worksheets("Data")
    On Error GoTo 0

    If GetImportDataSheetGeneric Is Nothing Then
        Set GetImportDataSheetGeneric = wb.Worksheets(1)
    End If

End Function

Private Function BuildHeaderMapGeneric(ByVal ws As Worksheet) As Object

    Dim dict As Object
    Dim lastCol As Long, c As Long
    Dim h As String

    Set dict = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        h = NormalizeHeaderGeneric(CStr(ws.Cells(1, c).value))
        If h <> "" Then dict(h) = c
    Next c

    Set BuildHeaderMapGeneric = dict

End Function

Private Function NormalizeHeaderGeneric(ByVal txt As String) As String

    txt = UCase$(Trim$(txt))
    txt = Replace(txt, "*", "")
    txt = Trim$(txt)
    txt = Replace(txt, " ", "_")

    Do While Right$(txt, 1) = "_"
        txt = Left$(txt, Len(txt) - 1)
    Loop

    Do While InStr(txt, "__") > 0
        txt = Replace(txt, "__", "_")
    Loop

    NormalizeHeaderGeneric = txt

End Function

Private Function RowHasAnyDataGeneric(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal maxCol As Long) As Boolean

    Dim c As Long
    For c = 1 To maxCol
        If Trim$(CStr(ws.Cells(rowNum, c).value)) <> "" Then
            RowHasAnyDataGeneric = True
            Exit Function
        End If
    Next c

End Function

Private Function GetCellValueGeneric(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object, ByVal fieldName As String) As Variant

    Dim k As String
    k = UCase$(fieldName)

    If hdr.Exists(k) Then
        GetCellValueGeneric = ws.Cells(rowNum, hdr(k)).value
    Else
        GetCellValueGeneric = Empty
    End If

End Function

Private Function GetCellStringGeneric(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object, ByVal fieldName As String) As String
    GetCellStringGeneric = Trim$(CStr(GetCellValueGeneric(ws, rowNum, hdr, fieldName)))
End Function

Private Function AppendMsgGeneric(ByVal baseMsg As String, ByVal addMsg As String) As String

    If baseMsg = "" Then
        AppendMsgGeneric = addMsg
    Else
        AppendMsgGeneric = baseMsg & " | " & addMsg
    End If

End Function

Private Function NzDblGeneric(ByVal v As Variant) As Double

    If IsError(v) Then
        NzDblGeneric = 0
    ElseIf IsNumeric(v) Then
        NzDblGeneric = CDbl(v)
    Else
        NzDblGeneric = 0
    End If

End Function

Private Function GetHeaderColumnGeneric(ByVal lo As ListObject, ByVal headerName As String) As Long

    Dim i As Long

    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(i).name), Trim$(headerName), vbTextCompare) = 0 Then
            GetHeaderColumnGeneric = i
            Exit Function
        End If
    Next i

End Function

Private Sub SetCellByHeaderGeneric(ByVal rowRange As Range, ByVal lo As ListObject, ByVal headerName As String, ByVal newValue As Variant)

    Dim c As Long
    c = GetHeaderColumnGeneric(lo, headerName)
    If c > 0 Then rowRange.Cells(1, c).value = newValue

End Sub

Private Function ExtractDigitsGeneric(ByVal txt As String) As String

    Dim i As Long, ch As String

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch Like "#" Then ExtractDigitsGeneric = ExtractDigitsGeneric & ch
    Next i

End Function

Private Function GetOrCreateSheetGeneric(ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetOrCreateSheetGeneric = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateSheetGeneric Is Nothing Then
        Set GetOrCreateSheetGeneric = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateSheetGeneric.name = sheetName
    End If

End Function

Private Sub AddValidationBackLinkGeneric(ByVal ws As Worksheet, ByVal titleText As String)

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

Private Function FindProductRowBySKU(ByVal prodTbl As ListObject, ByVal sku As String) As Range

    Dim cSKU As Long
    Dim i As Long
    Dim keySKU As String

    cSKU = GetHeaderColumnGeneric(prodTbl, "SKU")
    If cSKU = 0 Or prodTbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To prodTbl.ListRows.Count
        keySKU = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))
        If keySKU = UCase$(Trim$(sku)) Then
            Set FindProductRowBySKU = prodTbl.DataBodyRange.Rows(i)
            Exit Function
        End If
    Next i

End Function

Private Function GetCellFromRowByHeader(ByVal rr As Range, ByVal prodTbl As ListObject, ByVal headerName As String) As Variant

    Dim c As Long
    c = GetHeaderColumnGeneric(prodTbl, headerName)

    If c > 0 Then
        GetCellFromRowByHeader = rr.Cells(1, c).value
    Else
        GetCellFromRowByHeader = Empty
    End If

End Function

Private Sub SetCellInRowByHeader(ByVal rr As Range, ByVal prodTbl As ListObject, ByVal headerName As String, ByVal newValue As Variant)

    Dim c As Long
    c = GetHeaderColumnGeneric(prodTbl, headerName)

    If c > 0 Then
        rr.Cells(1, c).value = newValue
    End If

End Sub

Private Function NextInventoryLogID(ByVal logTbl As ListObject) As String

    Dim c As Long, i As Long
    Dim txt As String, n As Long, maxN As Long

    c = GetHeaderColumnGeneric(logTbl, "Log_ID")
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

