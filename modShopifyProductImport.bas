Attribute VB_Name = "modShopifyProductImport"
Option Explicit

Private gShopifyProductFilePath As String

Private Const WS_IMPORT_UI As String = "Import_UI"
Private Const WS_VALIDATION As String = "Import_Validation"
Private Const WS_PRODUCTS_DB As String = "Products_DB"

Private Const TBL_PRODUCTS As String = "tblProducts"

Private Const UI_CELL_FILE As String = "B13"
Private Const UI_CELL_VALIDATION As String = "B14"
Private Const UI_CELL_RESULT As String = "B15"
Private Const UI_CELL_RUN_OPTION As String = "B16"

Private Const RUN_OPTION_STRICT As String = "STOP_ON_ERROR"
Private Const RUN_OPTION_VALID_ONLY As String = "IMPORT_VALID_ONLY"

' ===============================
' Public Buttons
' ===============================
Public Sub ShopifyProductImport_SelectFile()

    Dim fd As Object
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)
    Set fd = Application.FileDialog(3)

    With fd
        .Title = "Select Shopify Product CSV File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .Filters.Add "All Files", "*.*"

        If .Show <> -1 Then Exit Sub

        gShopifyProductFilePath = .SelectedItems(1)
    End With

    ws.Range(UI_CELL_FILE).value = gShopifyProductFilePath
    ws.Range(UI_CELL_VALIDATION).ClearContents
    ws.Range(UI_CELL_RESULT).ClearContents

    ShopifyProduct_ClearValidationLink

End Sub

Public Sub ShopifyProductImport_Validate()

    Dim ok As Boolean

    ok = ValidateShopifyProductImportFile(True)

    If ok Then
        MsgBox "Shopify product validation passed.", vbInformation
    Else
        MsgBox "Shopify product validation finished with errors. Click Last Validation to open Import_Validation.", vbExclamation
    End If

End Sub

Public Sub ShopifyProductImport_Run()

    Dim ok As Boolean
    Dim runOption As String
    Dim errCount As Long, validCount As Long

    ok = ValidateShopifyProductImportFile(True)

    runOption = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RUN_OPTION).value)))
    If runOption = "" Then runOption = RUN_OPTION_STRICT

    GetShopifyProductValidationCounts errCount, validCount

    If runOption = RUN_OPTION_STRICT Then
        If Not ok Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT BLOCKED - VALIDATION FAILED"
            MsgBox "Shopify product import stopped because validation failed.", vbExclamation
            Exit Sub
        End If

    ElseIf runOption = RUN_OPTION_VALID_ONLY Then
        If validCount = 0 Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "NO VALID ROWS TO IMPORT"
            MsgBox "There are no VALID rows to import.", vbExclamation
            Exit Sub
        End If

    Else
        MsgBox "Invalid Import Option in B16. Use STOP_ON_ERROR or IMPORT_VALID_ONLY.", vbExclamation
        Exit Sub
    End If

    RunValidatedShopifyProductImport runOption

End Sub

Public Sub ShopifyProductImport_Clear()

    gShopifyProductFilePath = ""

    With ThisWorkbook.Worksheets(WS_IMPORT_UI)
        .Range(UI_CELL_FILE).ClearContents
        .Range(UI_CELL_VALIDATION).ClearContents
        .Range(UI_CELL_RESULT).ClearContents
        .Range(UI_CELL_VALIDATION).Interior.Pattern = xlNone
        .Range(UI_CELL_VALIDATION).Font.Underline = xlUnderlineStyleNone
        .Range(UI_CELL_VALIDATION).Font.ColorIndex = xlAutomatic
    End With

    ShopifyProduct_ClearValidationLink

End Sub

' ===============================
' Validation Core
' ===============================
Private Function ValidateShopifyProductImportFile(ByVal writeReport As Boolean) As Boolean

    Dim filePath As String
    Dim wb As Workbook, ws As Worksheet
    Dim lastRow As Long, i As Long
    Dim hdr As Object
    Dim results As Collection
    Dim res As Object
    Dim prodTbl As ListObject
    Dim dictExistingSKU As Object, dictExistingPID As Object, dictExistingProduct As Object
    Dim dictFileSKU As Object
    Dim dictHandleCache As Object
    Dim errCount As Long, validCount As Long
    Dim errMsg As String
    Dim skuText As String

    On Error GoTo ErrHandler

    ProgressStart "Validating Shopify Products", "Loading product master..."

    filePath = GetShopifyProductImportFilePath()
    If filePath = "" Then
        ProgressEnd
        MsgBox "Please select Shopify product file first.", vbExclamation
        ValidateShopifyProductImportFile = False
        Exit Function
    End If

    Set prodTbl = GetShopifyProductTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)
    If prodTbl Is Nothing Then
        ProgressEnd
        MsgBox "Products_DB table not found.", vbCritical
        ValidateShopifyProductImportFile = False
        Exit Function
    End If

    Set dictExistingSKU = BuildExistingShopifyProductSKUDict(prodTbl)
    Set dictExistingPID = BuildExistingShopifyProductIDDict(prodTbl)
    Set dictExistingProduct = BuildExistingShopifyProductRecordDict(prodTbl)
    Set dictFileSKU = CreateObject("Scripting.Dictionary")
    Set dictHandleCache = CreateObject("Scripting.Dictionary")
    Set results = New Collection

    ProgressStep "Opening Shopify product file..."
    Set wb = Workbooks.Open(filePath)
    Set ws = wb.Worksheets(1)

    ProgressStep "Reading headers..."
    Set hdr = BuildShopifyProductHeaderMap(ws)

    If Not CheckRequiredShopifyProductHeaders(hdr, results) Then
        wb.Close False
        Set wb = Nothing

        If writeReport Then WriteShopifyProductValidationReport results
        UpdateShopifyProductValidationUI False, 0, results.Count

        ProgressEnd
        ValidateShopifyProductImportFile = False
        Exit Function
    End If

    If WorksheetFunction.CountA(ws.Cells) = 0 Then
        wb.Close False
        Set wb = Nothing
        ProgressEnd
        MsgBox "Shopify product import file is empty.", vbExclamation
        ValidateShopifyProductImportFile = False
        Exit Function
    End If

    lastRow = ws.Cells.Find(What:="*", After:=ws.Range("A1"), LookIn:=xlFormulas, _
                            LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row

    For i = 2 To lastRow

        skuText = ""
        If hdr.Exists("VARIANTSKU") Then skuText = Trim$(CStr(ws.Cells(i, hdr("VARIANTSKU")).value))

        ProgressUpdate "Validating Shopify rows", i - 1, lastRow - 1, "SKU: " & skuText

        If Not ShopifyProductRowHasAnyData(ws, i, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column) Then GoTo nextRow
        If ShouldSkipShopifyProductRow(ws, i, hdr) Then GoTo nextRow

        Set res = ValidateSingleShopifyProductRow(ws, i, hdr, dictExistingSKU, dictExistingPID, dictExistingProduct, dictFileSKU, dictHandleCache)
        results.Add res

        If UCase$(Trim$(CStr(res("Status")))) = "VALID" Then
            validCount = validCount + 1
        Else
            errCount = errCount + 1
        End If

nextRow:
    Next i

    wb.Close False
    Set wb = Nothing

    ProgressStep "Writing validation report..."

    If writeReport Then WriteShopifyProductValidationReport results
    UpdateShopifyProductValidationUI (errCount = 0), validCount, errCount

    ProgressUpdate "Finalizing", 1, 1, "Complete"
    Application.Wait Now + TimeValue("0:00:01")
    ProgressEnd

    ValidateShopifyProductImportFile = (errCount = 0)
    Exit Function

ErrHandler:
    errMsg = Err.Description

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close False
    On Error GoTo 0

    ProgressEnd
    MsgBox "Shopify product validation error: " & errMsg, vbCritical
    ValidateShopifyProductImportFile = False

End Function

Private Function ValidateSingleShopifyProductRow( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal hdr As Object, _
    ByVal dictExistingSKU As Object, _
    ByVal dictExistingPID As Object, _
    ByVal dictExistingProduct As Object, _
    ByVal dictFileSKU As Object, _
    ByVal dictHandleCache As Object) As Object

    Dim res As Object
    Dim handleText As String
    Dim sku As String, productName As String, variantDesc As String, category As String
    Dim unitCost As Double, sellingPrice As Double
    Dim activeStatus As String, notesText As String
    Dim actionType As String, productID As String
    Dim msg As String
    Dim vendorText As String, tagsText As String, statusText As String
    Dim rec As Object
    Dim handleInfo As Object

    Set res = CreateObject("Scripting.Dictionary")

    handleText = UCase$(Trim$(GetShopifyCellString(ws, rowNum, hdr, "HANDLE")))

    productName = Trim$(GetShopifyCellString(ws, rowNum, hdr, "TITLE"))
    category = Trim$(GetShopifyCellString(ws, rowNum, hdr, "TYPE"))
    vendorText = Trim$(GetShopifyCellString(ws, rowNum, hdr, "VENDOR"))
    tagsText = Trim$(GetShopifyCellString(ws, rowNum, hdr, "TAGS"))
    statusText = Trim$(GetShopifyCellString(ws, rowNum, hdr, "STATUS"))

    If handleText <> "" Then

        If Not dictHandleCache.Exists(handleText) Then
            Set handleInfo = CreateObject("Scripting.Dictionary")
            handleInfo("TITLE") = ""
            handleInfo("TYPE") = ""
            handleInfo("VENDOR") = ""
            handleInfo("TAGS") = ""
            handleInfo("STATUS") = ""
            dictHandleCache.Add handleText, handleInfo
        End If

        Set handleInfo = dictHandleCache(handleText)

        If productName <> "" Then
            handleInfo("TITLE") = productName
        Else
            productName = Trim$(CStr(handleInfo("TITLE")))
        End If

        If category <> "" Then
            handleInfo("TYPE") = category
        Else
            category = Trim$(CStr(handleInfo("TYPE")))
        End If

        If vendorText <> "" Then
            handleInfo("VENDOR") = vendorText
        Else
            vendorText = Trim$(CStr(handleInfo("VENDOR")))
        End If

        If tagsText <> "" Then
            handleInfo("TAGS") = tagsText
        Else
            tagsText = Trim$(CStr(handleInfo("TAGS")))
        End If

        If statusText <> "" Then
            handleInfo("STATUS") = statusText
        Else
            statusText = Trim$(CStr(handleInfo("STATUS")))
        End If

    End If

    sku = UCase$(Trim$(GetShopifyCellString(ws, rowNum, hdr, "VARIANTSKU")))
    unitCost = ShopifyNzDbl(GetShopifyCellValue(ws, rowNum, hdr, "COSTPERITEM"))
    sellingPrice = ShopifyNzDbl(GetShopifyCellValue(ws, rowNum, hdr, "VARIANTPRICE"))

    variantDesc = BuildVariantDesc( _
        Trim$(GetShopifyCellString(ws, rowNum, hdr, "OPTION1VALUE")), _
        Trim$(GetShopifyCellString(ws, rowNum, hdr, "OPTION2VALUE")), _
        Trim$(GetShopifyCellString(ws, rowNum, hdr, "OPTION3VALUE")))

    activeStatus = MapShopifyStatusToERP(statusText)
    If activeStatus = "" Then activeStatus = "Active"

    notesText = BuildShopifyProductNotes(vendorText, tagsText)

    If sku = "" Then msg = AppendShopifyProductMsg(msg, "Variant SKU required")
    If productName = "" Then msg = AppendShopifyProductMsg(msg, "Title required")

    If unitCost < 0 Then msg = AppendShopifyProductMsg(msg, "Cost per item cannot be negative")
    If sellingPrice < 0 Then msg = AppendShopifyProductMsg(msg, "Variant Price cannot be negative")

    If sku <> "" Then
        If dictFileSKU.Exists(UCase$(sku)) Then
            msg = AppendShopifyProductMsg(msg, "Duplicate SKU within Shopify file")
        Else
            dictFileSKU.Add UCase$(sku), True
        End If
    End If

    If sku <> "" And dictExistingSKU.Exists(UCase$(sku)) Then
        actionType = "UPDATE"

        If dictExistingProduct.Exists(UCase$(sku)) Then
            Set rec = dictExistingProduct(UCase$(sku))
            productID = rec("Product_ID")
        End If
    Else
        actionType = "NEW"
        productID = ""
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
    res("Opening_Stock") = 0
    res("Reorder_Level") = 0
    res("Safety_Days_Override") = 0
    res("Lead_Time_Days") = 0
    res("Active_Status") = activeStatus
    res("Notes") = notesText

    Set ValidateSingleShopifyProductRow = res

End Function

Private Function ShouldSkipShopifyProductRow(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object) As Boolean

    Dim handleText As String
    Dim titleText As String
    Dim skuText As String
    Dim option1Text As String
    Dim priceText As String

    handleText = Trim$(GetShopifyCellString(ws, rowNum, hdr, "HANDLE"))
    titleText = Trim$(GetShopifyCellString(ws, rowNum, hdr, "TITLE"))
    skuText = Trim$(GetShopifyCellString(ws, rowNum, hdr, "VARIANTSKU"))
    option1Text = Trim$(GetShopifyCellString(ws, rowNum, hdr, "OPTION1VALUE"))
    priceText = Trim$(CStr(GetShopifyCellValue(ws, rowNum, hdr, "VARIANTPRICE")))

    If handleText = "" And titleText = "" And skuText = "" And option1Text = "" And priceText = "" Then
        ShouldSkipShopifyProductRow = True
        Exit Function
    End If

End Function

' ===============================
' Run Import
' ===============================
Private Sub RunValidatedShopifyProductImport(ByVal runOption As String)

    Dim wsVal As Worksheet
    Dim prodTbl As ListObject
    Dim r As Long, lastRow As Long
    Dim importCount As Long
    Dim nextPIDNum As Long
    Dim rowObj As Object
    Dim actionType As String
    Dim skuText As String

    On Error GoTo ErrHandler

    Set wsVal = ThisWorkbook.Worksheets(WS_VALIDATION)
    Set prodTbl = GetShopifyProductTableSafe(WS_PRODUCTS_DB, TBL_PRODUCTS)

    If prodTbl Is Nothing Then
        MsgBox "Products_DB table not found.", vbCritical
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
        If HasShopifyProductValidationErrors(wsVal) Then
            ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_RESULT).value = "IMPORT BLOCKED - VALIDATION FAILED"
            MsgBox "Validation still contains errors. Please fix them first.", vbExclamation
            Exit Sub
        End If
    End If

    nextPIDNum = GetNextShopifyProductIDSeq(prodTbl)

    ProgressStart "Running Shopify Product Import", "Writing product records..."

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    For r = 4 To lastRow

        skuText = Trim$(CStr(wsVal.Cells(r, 5).value))
        ProgressUpdate "Importing Shopify rows", r - 3, lastRow - 3, "SKU: " & skuText

        If UCase$(Trim$(CStr(wsVal.Cells(r, 2).value))) <> "VALID" Then GoTo nextRow

        Set rowObj = CreateObject("Scripting.Dictionary")
        rowObj("Action_Type") = Trim$(CStr(wsVal.Cells(r, 3).value))
        rowObj("Product_ID") = Trim$(CStr(wsVal.Cells(r, 4).value))
        rowObj("SKU") = Trim$(CStr(wsVal.Cells(r, 5).value))
        rowObj("Product_Name") = Trim$(CStr(wsVal.Cells(r, 6).value))
        rowObj("Variant_Desc") = Trim$(CStr(wsVal.Cells(r, 7).value))
        rowObj("Category") = Trim$(CStr(wsVal.Cells(r, 8).value))
        rowObj("Unit_Cost") = ShopifyNzDbl(wsVal.Cells(r, 9).value)
        rowObj("Selling_Price") = ShopifyNzDbl(wsVal.Cells(r, 10).value)
        rowObj("Opening_Stock") = 0
        rowObj("Current_Stock") = 0
        rowObj("Reorder_Level") = 0
        rowObj("Safety_Days_Override") = 0
        rowObj("Lead_Time_Days") = 0
        rowObj("Active_Status") = Trim$(CStr(wsVal.Cells(r, 15).value))
        rowObj("Notes") = Trim$(CStr(wsVal.Cells(r, 16).value))

        actionType = UCase$(Trim$(rowObj("Action_Type")))

        If actionType = "NEW" Then

            If rowObj("Product_ID") = "" Then
                rowObj("Product_ID") = GenerateShopifyProductID(nextPIDNum)
                nextPIDNum = nextPIDNum + 1
            End If

            WriteNewShopifyProductRow prodTbl, rowObj

        ElseIf actionType = "UPDATE" Then

            UpdateExistingShopifyProductRow prodTbl, rowObj

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
    MsgBox "Shopify product import completed: " & importCount & " valid row(s).", vbInformation
    Exit Sub

ErrHandler:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    ProgressEnd
    MsgBox "Shopify product import error: " & Err.Description, vbCritical

End Sub

Private Sub WriteNewShopifyProductRow(ByVal prodTbl As ListObject, ByVal rowObj As Object)

    Dim lr As ListRow
    Set lr = prodTbl.ListRows.Add

    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Product_ID", rowObj("Product_ID")
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "SKU", rowObj("SKU")
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Product_Name", rowObj("Product_Name")
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Variant_Desc", rowObj("Variant_Desc")
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Category", rowObj("Category")
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Unit_Cost", rowObj("Unit_Cost")
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Selling_Price", rowObj("Selling_Price")
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Opening_Stock", 0
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Current_Stock", 0
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Reorder_Level", 0
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Safety_Days_Override", 0
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Lead_Time_Days", 0
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Active_Status", rowObj("Active_Status")
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Notes", rowObj("Notes")
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Created_At", Now
    SetCellByHeaderShopifyProduct lr.Range, prodTbl, "Updated_At", Now

End Sub

Private Sub UpdateExistingShopifyProductRow(ByVal prodTbl As ListObject, ByVal rowObj As Object)

    Dim rr As Range

    Set rr = FindShopifyProductRowBySKU(prodTbl, rowObj("SKU"))
    If rr Is Nothing Then Exit Sub

    SetCellInShopifyProductRow rr, prodTbl, "Product_Name", rowObj("Product_Name")
    SetCellInShopifyProductRow rr, prodTbl, "Variant_Desc", rowObj("Variant_Desc")
    SetCellInShopifyProductRow rr, prodTbl, "Category", rowObj("Category")
    SetCellInShopifyProductRow rr, prodTbl, "Unit_Cost", rowObj("Unit_Cost")
    SetCellInShopifyProductRow rr, prodTbl, "Selling_Price", rowObj("Selling_Price")
    SetCellInShopifyProductRow rr, prodTbl, "Active_Status", rowObj("Active_Status")
    SetCellInShopifyProductRow rr, prodTbl, "Notes", rowObj("Notes")
    SetCellInShopifyProductRow rr, prodTbl, "Updated_At", Now

End Sub

' ===============================
' Validation UI / Report
' ===============================
Private Sub UpdateShopifyProductValidationUI(ByVal isValid As Boolean, ByVal validCount As Long, ByVal errCount As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)

    ShopifyProduct_ClearValidationLink

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

Private Sub ShopifyProduct_ClearValidationLink()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_IMPORT_UI)

    On Error Resume Next
    ws.Range(UI_CELL_VALIDATION).Hyperlinks.Delete
    On Error GoTo 0

End Sub

Private Sub WriteShopifyProductValidationReport(ByVal results As Collection)

    Dim ws As Worksheet
    Dim i As Long, r As Long
    Dim res As Object

    Set ws = GetOrCreateShopifyProductSheet(WS_VALIDATION)
    ws.Cells.Clear

    AddShopifyProductValidationBackLink ws, "Shopify Product Validation Details"

    ws.Range("A3:Q3").value = Array( _
        "Row_No", "Status", "Action_Type", "Product_ID", "SKU", "Product_Name", "Variant_Desc", _
        "Category", "Unit_Cost", "Selling_Price", "Opening_Stock", "Reorder_Level", _
        "Safety_Days_Override", "Lead_Time_Days", "Active_Status", "Notes", "Message")

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

        If UCase$(Trim$(CStr(res("Status")))) = "ERROR" Then
            ws.Rows(r).Interior.Color = RGB(255, 230, 230)
        ElseIf UCase$(Trim$(CStr(res("Action_Type")))) = "UPDATE" Then
            ws.Rows(r).Interior.Color = RGB(255, 242, 204)
        ElseIf UCase$(Trim$(CStr(res("Action_Type")))) = "NEW" Then
            ws.Rows(r).Interior.Color = RGB(226, 239, 218)
        End If

        r = r + 1
    Next i

    ws.Columns("A:Q").AutoFit
    ws.Range("I4:N" & r - 1).NumberFormat = "#,##0.00"
    ws.Range("A3:Q" & r - 1).Borders.LineStyle = xlContinuous
    
    If r > 4 Then ws.Range("I4:N" & r - 1).NumberFormat = "#,##0.00"
If r > 3 Then ws.Range("A3:Q" & r - 1).Borders.LineStyle = xlContinuous

End Sub

Private Function HasShopifyProductValidationErrors(ByVal ws As Worksheet) As Boolean

    Dim lastRow As Long, i As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For i = 4 To lastRow
        If UCase$(Trim$(CStr(ws.Cells(i, 2).value))) = "ERROR" Then
            HasShopifyProductValidationErrors = True
            Exit Function
        End If
    Next i

End Function

Private Sub GetShopifyProductValidationCounts(ByRef errCount As Long, ByRef validCount As Long)

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

' ===============================
' Helpers
' ===============================
Private Function GetShopifyProductImportFilePath() As String
    If Trim$(gShopifyProductFilePath) <> "" Then
        GetShopifyProductImportFilePath = gShopifyProductFilePath
    Else
        GetShopifyProductImportFilePath = Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(UI_CELL_FILE).value))
    End If
End Function

Private Function CheckRequiredShopifyProductHeaders(ByVal hdr As Object, ByVal results As Collection) As Boolean

    Dim req As Variant, i As Long
    Dim res As Object

    req = Array("TITLE", "VARIANTSKU")

    For i = LBound(req) To UBound(req)
        If Not hdr.Exists(req(i)) Then
            Set res = CreateObject("Scripting.Dictionary")
            FillEmptyShopifyProductValidationRow res
            res("Row_No") = 0
            res("Status") = "ERROR"
            res("Message") = "Missing required Shopify header: " & req(i)
            results.Add res
        End If
    Next i

    CheckRequiredShopifyProductHeaders = (results.Count = 0)

End Function

Private Function BuildShopifyProductHeaderMap(ByVal ws As Worksheet) As Object

    Dim dict As Object
    Dim lastCol As Long, c As Long
    Dim h As String

    Set dict = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        h = NormalizeShopifyProductHeader(CStr(ws.Cells(1, c).value))
        If h <> "" Then dict(h) = c
    Next c

    Set BuildShopifyProductHeaderMap = dict

End Function

Private Function NormalizeShopifyProductHeader(ByVal txt As String) As String

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

    NormalizeShopifyProductHeader = txt

End Function

Private Function GetShopifyCellValue(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object, ByVal normalizedHeader As String) As Variant

    If hdr.Exists(normalizedHeader) Then
        GetShopifyCellValue = ws.Cells(rowNum, hdr(normalizedHeader)).value
    Else
        GetShopifyCellValue = Empty
    End If

End Function

Private Function GetShopifyCellString(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal hdr As Object, ByVal normalizedHeader As String) As String
    GetShopifyCellString = Trim$(CStr(GetShopifyCellValue(ws, rowNum, hdr, normalizedHeader)))
End Function

Private Function ShopifyNzDbl(ByVal v As Variant) As Double

    If IsError(v) Then
        ShopifyNzDbl = 0
    ElseIf IsNumeric(v) Then
        ShopifyNzDbl = CDbl(v)
    Else
        ShopifyNzDbl = 0
    End If

End Function

Private Function ShopifyProductRowHasAnyData(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal maxCol As Long) As Boolean

    Dim c As Long
    For c = 1 To maxCol
        If Trim$(CStr(ws.Cells(rowNum, c).value)) <> "" Then
            ShopifyProductRowHasAnyData = True
            Exit Function
        End If
    Next c

End Function

Private Function BuildVariantDesc(ByVal v1 As String, ByVal v2 As String, ByVal v3 As String) As String

    Dim tmp As String

    If v1 <> "" And UCase$(v1) <> "DEFAULT TITLE" Then tmp = v1
    If v2 <> "" Then
        If tmp <> "" Then tmp = tmp & " / "
        tmp = tmp & v2
    End If
    If v3 <> "" Then
        If tmp <> "" Then tmp = tmp & " / "
        tmp = tmp & v3
    End If

    BuildVariantDesc = tmp

End Function

Private Function BuildShopifyProductNotes(ByVal vendorText As String, ByVal tagsText As String) As String

    Dim s As String

    If vendorText <> "" Then s = "Vendor=" & vendorText
    If tagsText <> "" Then
        If s <> "" Then s = s & " | "
        s = s & "Tags=" & tagsText
    End If

    BuildShopifyProductNotes = s

End Function

Private Function MapShopifyStatusToERP(ByVal statusText As String) As String

    Select Case UCase$(Trim$(statusText))
        Case "ACTIVE"
            MapShopifyStatusToERP = "Active"
        Case "ARCHIVED", "DRAFT"
            MapShopifyStatusToERP = "Inactive"
        Case Else
            MapShopifyStatusToERP = "Active"
    End Select

End Function

Private Function AppendShopifyProductMsg(ByVal baseMsg As String, ByVal addMsg As String) As String

    If baseMsg = "" Then
        AppendShopifyProductMsg = addMsg
    Else
        AppendShopifyProductMsg = baseMsg & " | " & addMsg
    End If

End Function

Private Function GetShopifyProductTableSafe(ByVal wsName As String, ByVal preferredTableName As String) As ListObject

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(wsName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    On Error Resume Next
    Set GetShopifyProductTableSafe = ws.ListObjects(preferredTableName)
    On Error GoTo 0

    If GetShopifyProductTableSafe Is Nothing Then
        If ws.ListObjects.Count > 0 Then
            Set GetShopifyProductTableSafe = ws.ListObjects(1)
        End If
    End If

End Function

Private Function BuildExistingShopifyProductSKUDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cSKU As Long
    Dim keySKU As String

    Set dict = CreateObject("Scripting.Dictionary")
    cSKU = GetHeaderColumnShopifyProduct(prodTbl, "SKU")

    If cSKU = 0 Or prodTbl.DataBodyRange Is Nothing Then
        Set BuildExistingShopifyProductSKUDict = dict
        Exit Function
    End If

    For i = 1 To prodTbl.ListRows.Count
        keySKU = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))
        If keySKU <> "" Then dict(keySKU) = True
    Next i

    Set BuildExistingShopifyProductSKUDict = dict

End Function

Private Function BuildExistingShopifyProductIDDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object
    Dim i As Long, cPID As Long
    Dim keyPID As String

    Set dict = CreateObject("Scripting.Dictionary")
    cPID = GetHeaderColumnShopifyProduct(prodTbl, "Product_ID")

    If cPID = 0 Or prodTbl.DataBodyRange Is Nothing Then
        Set BuildExistingShopifyProductIDDict = dict
        Exit Function
    End If

    For i = 1 To prodTbl.ListRows.Count
        keyPID = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cPID).value)))
        If keyPID <> "" Then dict(keyPID) = True
    Next i

    Set BuildExistingShopifyProductIDDict = dict

End Function

Private Function BuildExistingShopifyProductRecordDict(ByVal prodTbl As ListObject) As Object

    Dim dict As Object, rec As Object
    Dim i As Long
    Dim cPID As Long, cSKU As Long, cPN As Long, cVD As Long, cCAT As Long
    Dim cUC As Long, cSP As Long, cOS As Long, cCS As Long, cRL As Long
    Dim cSD As Long, cLT As Long, cAS As Long, cNOTES As Long
    Dim sku As String

    Set dict = CreateObject("Scripting.Dictionary")

    cPID = GetHeaderColumnShopifyProduct(prodTbl, "Product_ID")
    cSKU = GetHeaderColumnShopifyProduct(prodTbl, "SKU")
    cPN = GetHeaderColumnShopifyProduct(prodTbl, "Product_Name")
    cVD = GetHeaderColumnShopifyProduct(prodTbl, "Variant_Desc")
    cCAT = GetHeaderColumnShopifyProduct(prodTbl, "Category")
    cUC = GetHeaderColumnShopifyProduct(prodTbl, "Unit_Cost")
    cSP = GetHeaderColumnShopifyProduct(prodTbl, "Selling_Price")
    cOS = GetHeaderColumnShopifyProduct(prodTbl, "Opening_Stock")
    cCS = GetHeaderColumnShopifyProduct(prodTbl, "Current_Stock")
    cRL = GetHeaderColumnShopifyProduct(prodTbl, "Reorder_Level")
    cSD = GetHeaderColumnShopifyProduct(prodTbl, "Safety_Days_Override")
    cLT = GetHeaderColumnShopifyProduct(prodTbl, "Lead_Time_Days")
    cAS = GetHeaderColumnShopifyProduct(prodTbl, "Active_Status")
    cNOTES = GetHeaderColumnShopifyProduct(prodTbl, "Notes")

    If cSKU = 0 Or prodTbl.DataBodyRange Is Nothing Then
        Set BuildExistingShopifyProductRecordDict = dict
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
            rec("Unit_Cost") = ShopifyNzDbl(prodTbl.DataBodyRange.Cells(i, cUC).value)
            rec("Selling_Price") = ShopifyNzDbl(prodTbl.DataBodyRange.Cells(i, cSP).value)
            rec("Opening_Stock") = ShopifyNzDbl(prodTbl.DataBodyRange.Cells(i, cOS).value)
            rec("Current_Stock") = ShopifyNzDbl(prodTbl.DataBodyRange.Cells(i, cCS).value)
            rec("Reorder_Level") = ShopifyNzDbl(prodTbl.DataBodyRange.Cells(i, cRL).value)
            rec("Safety_Days_Override") = ShopifyNzDbl(prodTbl.DataBodyRange.Cells(i, cSD).value)
            rec("Lead_Time_Days") = ShopifyNzDbl(prodTbl.DataBodyRange.Cells(i, cLT).value)
            rec("Active_Status") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cAS).value))
            rec("Notes") = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cNOTES).value))
            Set dict(sku) = rec
        End If
    Next i

    Set BuildExistingShopifyProductRecordDict = dict

End Function

Private Function GetNextShopifyProductIDSeq(ByVal prodTbl As ListObject) As Long

    Dim i As Long, cPID As Long
    Dim pid As String, digitsOnly As String
    Dim n As Long, maxN As Long

    cPID = GetHeaderColumnShopifyProduct(prodTbl, "Product_ID")
    If cPID = 0 Or prodTbl.DataBodyRange Is Nothing Then
        GetNextShopifyProductIDSeq = 1
        Exit Function
    End If

    For i = 1 To prodTbl.ListRows.Count
        pid = Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cPID).value))
        digitsOnly = ExtractShopifyProductDigits(pid)
        If digitsOnly <> "" Then
            n = CLng(digitsOnly)
            If n > maxN Then maxN = n
        End If
    Next i

    GetNextShopifyProductIDSeq = maxN + 1

End Function

Private Function GenerateShopifyProductID(ByVal seqNum As Long) As String
    GenerateShopifyProductID = "P" & Format$(seqNum, "00000")
End Function

Private Function GetHeaderColumnShopifyProduct(ByVal lo As ListObject, ByVal headerName As String) As Long

    Dim i As Long

    For i = 1 To lo.ListColumns.Count
        If StrComp(Trim$(lo.ListColumns(i).name), Trim$(headerName), vbTextCompare) = 0 Then
            GetHeaderColumnShopifyProduct = i
            Exit Function
        End If
    Next i

End Function

Private Sub SetCellByHeaderShopifyProduct(ByVal rowRange As Range, ByVal lo As ListObject, ByVal headerName As String, ByVal newValue As Variant)

    Dim c As Long
    c = GetHeaderColumnShopifyProduct(lo, headerName)
    If c > 0 Then rowRange.Cells(1, c).value = newValue

End Sub

Private Function ExtractShopifyProductDigits(ByVal txt As String) As String

    Dim i As Long, ch As String

    For i = 1 To Len(txt)
        ch = Mid$(txt, i, 1)
        If ch Like "#" Then ExtractShopifyProductDigits = ExtractShopifyProductDigits & ch
    Next i

End Function

Private Function GetOrCreateShopifyProductSheet(ByVal sheetName As String) As Worksheet

    On Error Resume Next
    Set GetOrCreateShopifyProductSheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If GetOrCreateShopifyProductSheet Is Nothing Then
        Set GetOrCreateShopifyProductSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetOrCreateShopifyProductSheet.name = sheetName
    End If

End Function

Private Sub AddShopifyProductValidationBackLink(ByVal ws As Worksheet, ByVal titleText As String)

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

Private Function FindShopifyProductRowBySKU(ByVal prodTbl As ListObject, ByVal sku As String) As Range

    Dim cSKU As Long
    Dim i As Long
    Dim keySKU As String

    cSKU = GetHeaderColumnShopifyProduct(prodTbl, "SKU")
    If cSKU = 0 Or prodTbl.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To prodTbl.ListRows.Count
        keySKU = UCase$(Trim$(CStr(prodTbl.DataBodyRange.Cells(i, cSKU).value)))
        If keySKU = UCase$(Trim$(sku)) Then
            Set FindShopifyProductRowBySKU = prodTbl.DataBodyRange.Rows(i)
            Exit Function
        End If
    Next i

End Function

Private Sub SetCellInShopifyProductRow(ByVal rr As Range, ByVal prodTbl As ListObject, ByVal headerName As String, ByVal newValue As Variant)

    Dim c As Long
    c = GetHeaderColumnShopifyProduct(prodTbl, headerName)

    If c > 0 Then
        rr.Cells(1, c).value = newValue
    End If

End Sub

Private Sub FillEmptyShopifyProductValidationRow(ByVal res As Object)

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

