Attribute VB_Name = "modPurchaseImport"
Option Explicit

Private Const SHEET_IMPORT_UI As String = "Import_UI"
Private Const SHEET_VALIDATION As String = "Import_Validation"

Private Const SHEET_PURCHASE As String = "Purchase_DB"
Private Const SHEET_PRODUCTS As String = "Products_DB"
Private Const SHEET_SUPPLIERS As String = "Suppliers_DB"

Private Const TABLE_PURCHASE As String = "tblPurchase"
Private Const TABLE_PRODUCTS As String = "tblProducts"
Private Const TABLE_SUPPLIERS As String = "tblSuppliers"

Private Const CELL_PO_IMPORT_TYPE As String = "B28"
Private Const CELL_PO_IMPORT_FILE As String = "B29"
Private Const CELL_PO_LAST_VALIDATION As String = "B30"
Private Const CELL_PO_LAST_RESULT As String = "B31"
Private Const CELL_PO_IMPORT_OPTION As String = "B32"

Private Const IMPORT_OPTION_VALID_ONLY As String = "IMPORT_VALID_ONLY"
Private Const IMPORT_OPTION_STOP_ON_ERROR As String = "STOP_ON_ERROR"

Private Const DUP_ASK As String = "ASK"
Private Const DUP_SKIP_ALL As String = "SKIP_ALL"
Private Const DUP_REPLACE_ALL As String = "REPLACE_ALL"
Private Const DUP_STOP As String = "STOP"

Private Const H_EXTERNAL_PO As String = "External_PO_No"
Private Const H_PO_NO As String = "Purchase_Order_No"
Private Const H_PO_DATE As String = "Purchase_Date"
Private Const H_SUPPLIER_NAME As String = "Supplier_Name"
Private Const H_SKU As String = "SKU"
Private Const H_QTY As String = "Qty"
Private Const H_RECEIVED_QTY As String = "Received_Qty"
Private Const H_UNIT_COST As String = "Unit_Cost"
Private Const H_LINE_TOTAL As String = "Line_Total"
Private Const H_AMOUNT_PAID As String = "Amount_Paid"
Private Const H_CREATED_AT As String = "Created_At"
Private Const H_NOTES As String = "Notes"

'==================================================
' PUBLIC BUTTON ACTIONS
'==================================================
Public Sub POImport_DownloadTemplate()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsSup As Worksheet
    Dim tblSup As ListObject
    Dim tblTemplate As ListObject
    Dim savePath As Variant
    Dim lastSupplierRow As Long
    Dim i As Long
    Dim colSupplierName As Long

    On Error GoTo ErrHandler

    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.name = "PO_Import_Template"

    ws.Range("A1").value = H_EXTERNAL_PO
    ws.Range("B1").value = H_PO_NO
    ws.Range("C1").value = H_PO_DATE
    ws.Range("D1").value = H_SUPPLIER_NAME
    ws.Range("E1").value = H_SKU
    ws.Range("F1").value = H_QTY
    ws.Range("G1").value = H_RECEIVED_QTY
    ws.Range("H1").value = H_UNIT_COST
    ws.Range("I1").value = H_LINE_TOTAL
    ws.Range("J1").value = H_AMOUNT_PAID
    ws.Range("K1").value = H_CREATED_AT
    ws.Range("L1").value = H_NOTES

    Set tblTemplate = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:L2"), , xlYes)
    tblTemplate.name = "tblPOImportTemplate"
    tblTemplate.TableStyle = "TableStyleMedium2"

    'Important: use normal cell formula, not structured table formula.
    'Excel Table will auto-fill this formula when users add new rows.
    ws.Range("I2").Formula = "=IF(OR(F2="""",H2=""""),"""",F2*H2)"

    ws.Columns("A:L").ColumnWidth = 20
    ws.Columns("C:C").NumberFormat = "yyyy-mm-dd"
    ws.Columns("K:K").NumberFormat = "yyyy-mm-dd"
    ws.Columns("F:J").NumberFormat = "0.00"

    'Template notes - keep outside import columns
    ws.Range("N1").value = "Notes:"
    ws.Range("N2").value = "1. Purchase_Order_No can be blank. If blank, system will auto-generate internal PO No."
    ws.Range("N3").value = "2. External_PO_No is used to group lines from the same external/vendor PO."
    ws.Range("N4").value = "3. Supplier_Name must already exist in Suppliers_DB."
    ws.Range("N5").value = "4. SKU must already exist in Products_DB."
    ws.Range("N6").value = "5. PO Import does NOT change inventory. Use Inventory Import to set current stock."
    ws.Range("N7").value = "6. Received_Qty can be used to import historical/partial/closed PO status."
    ws.Range("N8").value = "7. Line_Total is for checking only. System will calculate final Line_Total by Qty × Unit_Cost."
    ws.Range("N9").value = "8. Amount_Paid is used to calculate Payment_Status: Unpaid / Partial / Paid."
    ws.Range("N10").value = "9. Created_At can be blank. If blank, it will default to Purchase_Date."

    ws.Columns("N:N").ColumnWidth = 95
    ws.Columns("N:N").WrapText = True

    With ws.Range("N1:N10")
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlTop
    End With
    ws.Range("N1").Font.Bold = True

    Set wsSup = wb.Worksheets.Add(After:=ws)
    wsSup.name = "Supplier_List"
    wsSup.Visible = xlSheetVeryHidden
    wsSup.Range("A1").value = "Supplier_Name"

    On Error Resume Next
    Set tblSup = ThisWorkbook.Worksheets(SHEET_SUPPLIERS).ListObjects(TABLE_SUPPLIERS)
    On Error GoTo ErrHandler

    If Not tblSup Is Nothing Then
        colSupplierName = GetCol(tblSup, "Supplier_Name")
        If Not tblSup.DataBodyRange Is Nothing Then
            For i = 1 To tblSup.ListRows.Count
                wsSup.Cells(i + 1, 1).value = tblSup.DataBodyRange.Cells(i, colSupplierName).value
            Next i
        End If
    End If

    lastSupplierRow = wsSup.Cells(wsSup.Rows.Count, 1).End(xlUp).Row

    If lastSupplierRow >= 2 Then
        wb.Names.Add name:="SupplierNames", RefersTo:="=Supplier_List!$A$2:$A$" & lastSupplierRow

        With ws.Range("D2:D1000").Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=SupplierNames"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = False
        End With
    End If

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="PO_Import_Template.xlsx", _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx")

    If savePath = False Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=CStr(savePath), FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox "PO Import template created successfully.", vbInformation, "PO Import"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "Error creating PO Import template: " & Err.Description, vbCritical, "PO Import"

End Sub

Public Sub POImport_SelectFile()

    Dim fd As FileDialog
    Dim selectedPath As String
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_IMPORT_UI)
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select PO Import File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel / CSV Files", "*.xlsx;*.xls;*.csv"

        If .Show <> -1 Then Exit Sub
        selectedPath = .SelectedItems(1)
    End With

    ws.Range(CELL_PO_IMPORT_FILE).value = selectedPath
    ws.Range(CELL_PO_LAST_VALIDATION).value = ""
    ws.Range(CELL_PO_LAST_RESULT).value = ""
    ws.Range(CELL_PO_IMPORT_OPTION).value = IMPORT_OPTION_VALID_ONLY

End Sub

Public Sub POImport_Clear()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_IMPORT_UI)

    ws.Range(CELL_PO_IMPORT_TYPE).value = "PO Import"
    ws.Range(CELL_PO_IMPORT_FILE).value = ""
    ws.Range(CELL_PO_LAST_VALIDATION).value = ""
    ws.Range(CELL_PO_LAST_RESULT).value = ""
    ws.Range(CELL_PO_IMPORT_OPTION).value = IMPORT_OPTION_VALID_ONLY
    ws.Range(CELL_PO_LAST_VALIDATION).Interior.Pattern = xlNone
    ws.Range(CELL_PO_LAST_RESULT).Interior.Pattern = xlNone

End Sub

Public Sub POImport_Validate()

    Dim filePath As String
    Dim errorCount As Long
    Dim warningCount As Long
    Dim validCount As Long

    filePath = Trim$(CStr(ThisWorkbook.Worksheets(SHEET_IMPORT_UI).Range(CELL_PO_IMPORT_FILE).value))

    If filePath = "" Then
        MsgBox "Please select a PO import file first.", vbExclamation, "PO Import"
        Exit Sub
    End If

    ValidatePOImportFile filePath, errorCount, warningCount, validCount
    WritePOValidationStatus errorCount, warningCount, validCount

End Sub

Public Sub POImport_RunImport()

    Dim filePath As String
    Dim importOption As String
    Dim errorCount As Long
    Dim warningCount As Long
    Dim validCount As Long
    Dim importedCount As Long

    filePath = Trim$(CStr(ThisWorkbook.Worksheets(SHEET_IMPORT_UI).Range(CELL_PO_IMPORT_FILE).value))
    importOption = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(SHEET_IMPORT_UI).Range(CELL_PO_IMPORT_OPTION).value)))

    If filePath = "" Then
        MsgBox "Please select a PO import file first.", vbExclamation, "PO Import"
        Exit Sub
    End If

    ValidatePOImportFile filePath, errorCount, warningCount, validCount
    WritePOValidationStatus errorCount, warningCount, validCount

    If importOption = IMPORT_OPTION_STOP_ON_ERROR And errorCount > 0 Then
        With ThisWorkbook.Worksheets(SHEET_IMPORT_UI)
            .Range(CELL_PO_LAST_RESULT).value = "IMPORT STOPPED - validation error(s) found"
            .Range(CELL_PO_LAST_RESULT).Interior.Color = RGB(255, 199, 206)
        End With
        MsgBox "Import stopped because validation errors were found.", vbExclamation, "PO Import"
        Exit Sub
    End If

    If validCount = 0 Then
        MsgBox "No valid PO rows to import.", vbExclamation, "PO Import"
        Exit Sub
    End If

    importedCount = ImportValidPORows(filePath)

    With ThisWorkbook.Worksheets(SHEET_IMPORT_UI)
        If importedCount >= 0 Then
            .Range(CELL_PO_LAST_RESULT).value = "IMPORT SUCCESS - " & importedCount & " valid row(s)"
            .Range(CELL_PO_LAST_RESULT).Interior.Color = RGB(226, 239, 218)
        Else
            .Range(CELL_PO_LAST_RESULT).value = "IMPORT STOPPED BY USER"
            .Range(CELL_PO_LAST_RESULT).Interior.Color = RGB(255, 235, 156)
        End If
    End With

    If importedCount >= 0 Then
        MsgBox "PO Import completed. Imported " & importedCount & " valid row(s).", vbInformation, "PO Import"
    Else
        MsgBox "PO Import stopped by user.", vbExclamation, "PO Import"
    End If

End Sub

'==================================================
' VALIDATION
'==================================================
Private Sub ValidatePOImportFile(ByVal filePath As String, _
                                 ByRef errorCount As Long, _
                                 ByRef warningCount As Long, _
                                 ByRef validCount As Long)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim rowStatus As String
    Dim msg As String

    Dim dictSup As Object
    Dim dictProd As Object
    Dim dictDupFile As Object
    Dim dictDupDB As Object

    On Error GoTo ErrHandler

    ClearValidationReport

    Set dictSup = BuildSupplierDict()
    Set dictProd = BuildProductDict()
    Set dictDupFile = CreateObject("Scripting.Dictionary")
    Set dictDupDB = BuildExistingPOKeyDict()

    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws = wb.Worksheets(1)

    lastRow = GetLastUsedRowInCols(ws, 1, 12)

    ProgressStart "Validating PO Import", "Checking PO import file..."

    For r = 2 To lastRow

        If IsPOImportInstructionRow(ws, r) Then Exit For
        If IsPOImportBlankRow(ws, r) Then GoTo NextValidationRow

        ProgressUpdate "Validating row", r - 1, lastRow - 1, "Row " & r

        ValidatePOImportRow ws, r, dictSup, dictProd, dictDupFile, dictDupDB, rowStatus, msg

        If rowStatus = "ERROR" Then
            errorCount = errorCount + 1
        ElseIf rowStatus = "WARNING" Then
            warningCount = warningCount + 1
            validCount = validCount + 1
        ElseIf rowStatus = "VALID" Then
            validCount = validCount + 1
        End If

        If rowStatus <> "VALID" Then
            AppendValidationReport ws, r, rowStatus, msg
        End If

NextValidationRow:
    Next r

    ProgressEnd

    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    On Error Resume Next
    ProgressEnd
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Validation error: " & Err.Description, vbCritical, "PO Import"

End Sub

Private Sub ValidatePOImportRow(ByVal ws As Worksheet, _
                                ByVal r As Long, _
                                ByVal dictSup As Object, _
                                ByVal dictProd As Object, _
                                ByVal dictDupFile As Object, _
                                ByVal dictDupDB As Object, _
                                ByRef rowStatus As String, _
                                ByRef msg As String)

    Dim externalPO As String
    Dim poNo As String
    Dim supplierName As String
    Dim sku As String
    Dim qty As Double
    Dim receivedQty As Double
    Dim unitCost As Double
    Dim lineTotalInput As Double
    Dim expectedLineTotal As Double
    Dim amountPaid As Double
    Dim poDate As Variant
    Dim createdAt As Variant
    Dim dupKey As String

    rowStatus = "VALID"
    msg = ""

    externalPO = Trim$(CStr(GetCellByHeader(ws, r, H_EXTERNAL_PO)))
    poNo = Trim$(CStr(GetCellByHeader(ws, r, H_PO_NO)))
    poDate = GetCellByHeader(ws, r, H_PO_DATE)
    supplierName = Trim$(CStr(GetCellByHeader(ws, r, H_SUPPLIER_NAME)))
    sku = Trim$(CStr(GetCellByHeader(ws, r, H_SKU)))
    qty = NzNumber(GetCellByHeader(ws, r, H_QTY))
    receivedQty = NzNumber(GetCellByHeader(ws, r, H_RECEIVED_QTY))
    unitCost = NzNumber(GetCellByHeader(ws, r, H_UNIT_COST))
    lineTotalInput = NzNumber(GetCellByHeader(ws, r, H_LINE_TOTAL))
    amountPaid = NzNumber(GetCellByHeader(ws, r, H_AMOUNT_PAID))
    createdAt = GetCellByHeader(ws, r, H_CREATED_AT)

    expectedLineTotal = Round(qty * unitCost, 2)

    If supplierName = "" Then AddError msg, "Missing Supplier_Name."
    If supplierName <> "" Then
        If Not dictSup.Exists(UCase$(supplierName)) Then
            AddError msg, "Supplier not found. Please create it in Supplier UI or import suppliers first."
        End If
    End If

    If sku = "" Then AddError msg, "Missing SKU."
    If sku <> "" Then
        If Not dictProd.Exists(UCase$(sku)) Then
            AddError msg, "SKU not found in Products_DB."
        End If
    End If

    If qty <= 0 Then AddError msg, "Qty must be greater than 0."
    If receivedQty < 0 Then AddError msg, "Received_Qty cannot be negative."
    If receivedQty > qty Then AddError msg, "Received_Qty cannot be greater than Qty."
    If unitCost < 0 Then AddError msg, "Unit_Cost cannot be negative."
    If amountPaid < 0 Then AddError msg, "Amount_Paid cannot be negative."

    If qty > 0 And unitCost >= 0 Then
        If amountPaid > expectedLineTotal Then
            AddError msg, "Amount_Paid cannot be greater than Line_Total."
        End If
    End If

    If Trim$(CStr(poDate)) <> "" Then
        If Not IsDate(poDate) Then AddError msg, "Invalid Purchase_Date."
    End If

    If Trim$(CStr(createdAt)) <> "" Then
        If Not IsDate(createdAt) Then AddError msg, "Invalid Created_At."
    End If

    dupKey = BuildPOImportDuplicateKey(externalPO, poNo, sku)

    If dupKey <> "" Then
        If dictDupFile.Exists(dupKey) Then
            AddError msg, "Duplicate PO key inside import file."
        Else
            dictDupFile.Add dupKey, True
        End If

        If dictDupDB.Exists(dupKey) Then
            AddWarning msg, "Duplicate PO key already exists in Purchase_DB. You can choose Skip or Replace during import."
        End If
    End If

    If HasErrorMessage(msg) Then
        rowStatus = "ERROR"
        Exit Sub
    End If

    If Trim$(CStr(poDate)) = "" Then AddWarning msg, "Purchase_Date blank. Default to today."
    If Trim$(CStr(createdAt)) = "" Then AddWarning msg, "Created_At blank. Default to Purchase_Date."
    If unitCost = 0 Then AddWarning msg, "Unit_Cost is 0."

    If lineTotalInput <> 0 Then
        If Abs(lineTotalInput - expectedLineTotal) > 0.01 Then
            AddWarning msg, "Line_Total does not match Qty ¡Á Unit_Cost. System will use Qty ¡Á Unit_Cost."
        End If
    End If

    If msg <> "" Then rowStatus = "WARNING"

End Sub

'==================================================
' IMPORT
'==================================================
Private Function ImportValidPORows(ByVal filePath As String) As Long

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim dictSup As Object
    Dim dictProd As Object
    Dim dictDupFile As Object
    Dim dictDupDB As Object
    Dim dictExternalToPO As Object

    Dim rowStatus As String
    Dim msg As String
    Dim duplicateMode As String
    Dim dupKey As String
    Dim duplicateAction As String

    On Error GoTo ErrHandler

    Set dictSup = BuildSupplierDict()
    Set dictProd = BuildProductDict()
    Set dictDupFile = CreateObject("Scripting.Dictionary")
    Set dictDupDB = BuildExistingPOKeyDict()
    Set dictExternalToPO = CreateObject("Scripting.Dictionary")

    duplicateMode = DUP_ASK

    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws = wb.Worksheets(1)

    lastRow = GetLastUsedRowInCols(ws, 1, 12)

    ProgressStart "Importing PO", "Writing valid PO rows..."

    For r = 2 To lastRow

        If IsPOImportInstructionRow(ws, r) Then Exit For
        If IsPOImportBlankRow(ws, r) Then GoTo NextImportRow

        ProgressUpdate "Importing row", r - 1, lastRow - 1, "Row " & r

        ValidatePOImportRow ws, r, dictSup, dictProd, dictDupFile, dictDupDB, rowStatus, msg

        If rowStatus = "VALID" Or rowStatus = "WARNING" Then

            dupKey = GetPOImportRowDuplicateKey(ws, r)

            If dupKey <> "" And dictDupDB.Exists(dupKey) Then

                duplicateAction = ResolveDuplicatePOAction(ws, r, duplicateMode)

                If duplicateAction = DUP_STOP Then
                    ImportValidPORows = -1
                    GoTo CleanExit
                ElseIf duplicateAction = DUP_SKIP_ALL Then
                    duplicateMode = DUP_SKIP_ALL
                    GoTo NextImportRow
                ElseIf duplicateAction = DUP_REPLACE_ALL Then
                    duplicateMode = DUP_REPLACE_ALL
                    DeleteExistingPOKeyRows dupKey
                    dictDupDB.Remove dupKey
                ElseIf duplicateAction = "SKIP_THIS" Then
                    GoTo NextImportRow
                ElseIf duplicateAction = "REPLACE_THIS" Then
                    DeleteExistingPOKeyRows dupKey
                    dictDupDB.Remove dupKey
                End If

            End If

            WritePOImportRow ws, r, dictSup, dictProd, dictExternalToPO
            ImportValidPORows = ImportValidPORows + 1

            dupKey = GetPOImportRowDuplicateKey(ws, r)
            If dupKey <> "" Then
                If Not dictDupDB.Exists(dupKey) Then dictDupDB.Add dupKey, True
            End If

        End If

NextImportRow:
    Next r

CleanExit:
    ProgressEnd

    wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    Exit Function

ErrHandler:
    On Error Resume Next
    ProgressEnd
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    Application.ScreenUpdating = True
    MsgBox "Import error: " & Err.Description, vbCritical, "PO Import"

End Function

Private Sub WritePOImportRow(ByVal ws As Worksheet, _
                             ByVal r As Long, _
                             ByVal dictSup As Object, _
                             ByVal dictProd As Object, _
                             ByVal dictExternalToPO As Object)

    Dim tbl As ListObject
    Dim newRow As ListRow

    Dim externalPO As String
    Dim poNo As String
    Dim purchaseDate As Date
    Dim supplierName As String
    Dim supplierData As Variant
    Dim supplierID As String

    Dim sku As String
    Dim productData As Variant
    Dim productID As String
    Dim productName As String

    Dim qty As Double
    Dim receivedQty As Double
    Dim remainingQty As Double
    Dim unitCost As Double
    Dim lineTotal As Double
    Dim amountPaid As Double
    Dim balanceDue As Double

    Dim paymentStatus As String
    Dim poStatus As String
    Dim lineStatus As String
    Dim createdAt As Date
    Dim notes As String

    Set tbl = ThisWorkbook.Worksheets(SHEET_PURCHASE).ListObjects(TABLE_PURCHASE)

    externalPO = Trim$(CStr(GetCellByHeader(ws, r, H_EXTERNAL_PO)))
    poNo = Trim$(CStr(GetCellByHeader(ws, r, H_PO_NO)))
    supplierName = Trim$(CStr(GetCellByHeader(ws, r, H_SUPPLIER_NAME)))
    sku = Trim$(CStr(GetCellByHeader(ws, r, H_SKU)))

    If poNo = "" Then
        If externalPO <> "" Then
            If dictExternalToPO.Exists(UCase$(externalPO)) Then
                poNo = dictExternalToPO(UCase$(externalPO))
            Else
                poNo = GenerateNextPONo()
                dictExternalToPO.Add UCase$(externalPO), poNo
            End If
        Else
            poNo = GenerateNextPONo()
        End If
    End If

    If IsDate(GetCellByHeader(ws, r, H_PO_DATE)) Then
        purchaseDate = CDate(GetCellByHeader(ws, r, H_PO_DATE))
    Else
        purchaseDate = Date
    End If

    If IsDate(GetCellByHeader(ws, r, H_CREATED_AT)) Then
        createdAt = CDate(GetCellByHeader(ws, r, H_CREATED_AT))
    Else
        createdAt = purchaseDate
    End If

    supplierData = dictSup(UCase$(supplierName))
    supplierID = CStr(supplierData(0))
    supplierName = CStr(supplierData(1))

    productData = dictProd(UCase$(sku))
    productID = CStr(productData(0))
    productName = CStr(productData(1))

    qty = NzNumber(GetCellByHeader(ws, r, H_QTY))
    receivedQty = NzNumber(GetCellByHeader(ws, r, H_RECEIVED_QTY))
    remainingQty = qty - receivedQty

    unitCost = NzNumber(GetCellByHeader(ws, r, H_UNIT_COST))
    lineTotal = Round(qty * unitCost, 2)

    amountPaid = NzNumber(GetCellByHeader(ws, r, H_AMOUNT_PAID))
    balanceDue = Round(lineTotal - amountPaid, 2)

    paymentStatus = GetPaymentStatus(amountPaid, lineTotal)
    lineStatus = GetLineStatus(receivedQty, qty)
    poStatus = lineStatus

    notes = Trim$(CStr(GetCellByHeader(ws, r, H_NOTES)))

    Set newRow = tbl.ListRows.Add

    With newRow.Range
        .Cells(1, GetCol(tbl, "Purchase_Order_No")).value = poNo
        WriteOptionalValue newRow.Range, tbl, "External_PO_No", externalPO
        .Cells(1, GetCol(tbl, "Purchase_Date")).value = purchaseDate
        .Cells(1, GetCol(tbl, "Supplier_ID")).value = supplierID
        .Cells(1, GetCol(tbl, "Supplier_Name")).value = supplierName
        .Cells(1, GetCol(tbl, "Product_ID")).value = productID
        .Cells(1, GetCol(tbl, "SKU")).value = sku
        .Cells(1, GetCol(tbl, "Product_Name")).value = productName
        .Cells(1, GetCol(tbl, "Qty")).value = qty
        .Cells(1, GetCol(tbl, "Received_Qty")).value = receivedQty
        .Cells(1, GetCol(tbl, "Remaining_Qty")).value = remainingQty
        .Cells(1, GetCol(tbl, "Unit_Cost")).value = unitCost
        .Cells(1, GetCol(tbl, "Line_Total")).value = lineTotal
        .Cells(1, GetCol(tbl, "Amount_Paid")).value = amountPaid
        .Cells(1, GetCol(tbl, "Balance_Due")).value = balanceDue
        .Cells(1, GetCol(tbl, "Payment_Status")).value = paymentStatus
        .Cells(1, GetCol(tbl, "Status")).value = poStatus
        .Cells(1, GetCol(tbl, "Line_Status")).value = lineStatus
        .Cells(1, GetCol(tbl, "Notes")).value = notes
        .Cells(1, GetCol(tbl, "Created_At")).value = createdAt
        .Cells(1, GetCol(tbl, "Updated_At")).value = Now
        WriteOptionalValue newRow.Range, tbl, "Source", "PO Import"
    End With

End Sub

'==================================================
' DUPLICATE HANDLING
'==================================================
Private Function ResolveDuplicatePOAction(ByVal ws As Worksheet, _
                                          ByVal rowNo As Long, _
                                          ByVal duplicateMode As String) As String

    Dim externalPO As String
    Dim poNo As String
    Dim sku As String
    Dim displayNo As String
    Dim action As String

    If duplicateMode = DUP_SKIP_ALL Then
        ResolveDuplicatePOAction = DUP_SKIP_ALL
        Exit Function
    End If

    If duplicateMode = DUP_REPLACE_ALL Then
        ResolveDuplicatePOAction = DUP_REPLACE_ALL
        Exit Function
    End If

    externalPO = Trim$(CStr(GetCellByHeader(ws, rowNo, H_EXTERNAL_PO)))
    poNo = Trim$(CStr(GetCellByHeader(ws, rowNo, H_PO_NO)))
    sku = Trim$(CStr(GetCellByHeader(ws, rowNo, H_SKU)))

    If externalPO <> "" Then
        displayNo = externalPO
    ElseIf poNo <> "" Then
        displayNo = poNo
    Else
        displayNo = "Row " & rowNo & " | SKU: " & sku
    End If

    frmDuplicateOrderAction.SelectedAction = ""
    frmDuplicateOrderAction.SetOrderInfo displayNo
    frmDuplicateOrderAction.Show vbModal

    action = UCase$(Trim$(frmDuplicateOrderAction.SelectedAction))

    Select Case action
        Case "SKIP_ONE"
            ResolveDuplicatePOAction = "SKIP_THIS"

        Case "SKIP_ALL"
            ResolveDuplicatePOAction = DUP_SKIP_ALL

        Case "REPLACE_ONE"
            ResolveDuplicatePOAction = "REPLACE_THIS"

        Case "REPLACE_ALL"
            ResolveDuplicatePOAction = DUP_REPLACE_ALL

        Case "STOP", ""
            ResolveDuplicatePOAction = DUP_STOP

        Case Else
            ResolveDuplicatePOAction = DUP_STOP
    End Select

End Function

Private Function GetPOImportRowDuplicateKey(ByVal ws As Worksheet, ByVal rowNo As Long) As String

    GetPOImportRowDuplicateKey = BuildPOImportDuplicateKey( _
        Trim$(CStr(GetCellByHeader(ws, rowNo, H_EXTERNAL_PO))), _
        Trim$(CStr(GetCellByHeader(ws, rowNo, H_PO_NO))), _
        Trim$(CStr(GetCellByHeader(ws, rowNo, H_SKU))) _
    )

End Function

Private Sub DeleteExistingPOKeyRows(ByVal dupKey As String)

    Dim tbl As ListObject
    Dim i As Long
    Dim externalPO As String
    Dim poNo As String
    Dim sku As String
    Dim rowKey As String
    Dim colExternal As Long

    Set tbl = ThisWorkbook.Worksheets(SHEET_PURCHASE).ListObjects(TABLE_PURCHASE)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    colExternal = GetOptionalCol(tbl, "External_PO_No")

    For i = tbl.ListRows.Count To 1 Step -1

        externalPO = ""
        If colExternal > 0 Then
            externalPO = Trim$(CStr(tbl.DataBodyRange.Cells(i, colExternal).value))
        End If

        poNo = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Purchase_Order_No")).value))
        sku = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "SKU")).value))

        rowKey = BuildPOImportDuplicateKey(externalPO, poNo, sku)

        If rowKey = dupKey Then
            tbl.ListRows(i).Delete
        End If

    Next i

End Sub

'==================================================
' DICTIONARIES
'==================================================
Private Function BuildSupplierDict() As Object

    Dim dict As Object
    Dim tbl As ListObject
    Dim i As Long
    Dim supplierID As String
    Dim supplierName As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set tbl = ThisWorkbook.Worksheets(SHEET_SUPPLIERS).ListObjects(TABLE_SUPPLIERS)

    If tbl.DataBodyRange Is Nothing Then
        Set BuildSupplierDict = dict
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        supplierID = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Supplier_ID")).value))
        supplierName = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Supplier_Name")).value))

        If supplierName <> "" Then
            If Not dict.Exists(UCase$(supplierName)) Then
                dict.Add UCase$(supplierName), Array(supplierID, supplierName)
            End If
        End If
    Next i

    Set BuildSupplierDict = dict

End Function

Private Function BuildProductDict() As Object

    Dim dict As Object
    Dim tbl As ListObject
    Dim i As Long
    Dim productID As String
    Dim sku As String
    Dim productName As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set tbl = ThisWorkbook.Worksheets(SHEET_PRODUCTS).ListObjects(TABLE_PRODUCTS)

    If tbl.DataBodyRange Is Nothing Then
        Set BuildProductDict = dict
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        productID = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Product_ID")).value))
        sku = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "SKU")).value))
        productName = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Product_Name")).value))

        If sku <> "" Then
            If Not dict.Exists(UCase$(sku)) Then
                dict.Add UCase$(sku), Array(productID, productName)
            End If
        End If
    Next i

    Set BuildProductDict = dict

End Function

Private Function BuildExistingPOKeyDict() As Object

    Dim dict As Object
    Dim tbl As ListObject
    Dim i As Long
    Dim externalPO As String
    Dim poNo As String
    Dim sku As String
    Dim key As String
    Dim colExternal As Long

    Set dict = CreateObject("Scripting.Dictionary")
    Set tbl = ThisWorkbook.Worksheets(SHEET_PURCHASE).ListObjects(TABLE_PURCHASE)

    If tbl.DataBodyRange Is Nothing Then
        Set BuildExistingPOKeyDict = dict
        Exit Function
    End If

    colExternal = GetOptionalCol(tbl, "External_PO_No")

    For i = 1 To tbl.ListRows.Count
        externalPO = ""

        If colExternal > 0 Then
            externalPO = Trim$(CStr(tbl.DataBodyRange.Cells(i, colExternal).value))
        End If

        poNo = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Purchase_Order_No")).value))
        sku = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "SKU")).value))

        key = BuildPOImportDuplicateKey(externalPO, poNo, sku)

        If key <> "" Then
            If Not dict.Exists(key) Then dict.Add key, True
        End If
    Next i

    Set BuildExistingPOKeyDict = dict

End Function

'==================================================
' HELPERS
'==================================================
Private Function BuildPOImportDuplicateKey(ByVal externalPO As String, _
                                           ByVal poNo As String, _
                                           ByVal sku As String) As String

    externalPO = Trim$(externalPO)
    poNo = Trim$(poNo)
    sku = Trim$(sku)

    If sku = "" Then
        BuildPOImportDuplicateKey = ""
    ElseIf externalPO <> "" Then
        BuildPOImportDuplicateKey = "EXT|" & UCase$(externalPO) & "|" & UCase$(sku)
    ElseIf poNo <> "" Then
        BuildPOImportDuplicateKey = "PO|" & UCase$(poNo) & "|" & UCase$(sku)
    Else
        BuildPOImportDuplicateKey = ""
    End If

End Function

Private Function GetPaymentStatus(ByVal amountPaid As Double, ByVal lineTotal As Double) As String

    If amountPaid <= 0 Then
        GetPaymentStatus = "Unpaid"
    ElseIf amountPaid < lineTotal Then
        GetPaymentStatus = "Partial"
    Else
        GetPaymentStatus = "Paid"
    End If

End Function

Private Function GetLineStatus(ByVal receivedQty As Double, ByVal qty As Double) As String

    If receivedQty <= 0 Then
        GetLineStatus = "Open"
    ElseIf receivedQty < qty Then
        GetLineStatus = "Partial"
    Else
        GetLineStatus = "Closed"
    End If

End Function

Private Function GenerateNextPONo() As String

    Dim tbl As ListObject
    Dim i As Long
    Dim poNo As String
    Dim maxNo As Long
    Dim n As Long

    Set tbl = ThisWorkbook.Worksheets(SHEET_PURCHASE).ListObjects(TABLE_PURCHASE)

    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            poNo = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Purchase_Order_No")).value))

            If UCase$(Left$(poNo, 2)) = "PO" Then
                n = CLng(val(Mid$(poNo, 3)))
                If n > maxNo Then maxNo = n
            End If
        Next i
    End If

    GenerateNextPONo = "PO" & Format$(maxNo + 1, "00000")

End Function

Private Function GetCellByHeader(ByVal ws As Worksheet, ByVal rowNo As Long, ByVal headerName As String) As Variant

    Dim colNo As Long
    colNo = FindHeaderCol(ws, headerName)

    If colNo = 0 Then
        GetCellByHeader = ""
    Else
        GetCellByHeader = ws.Cells(rowNo, colNo).value
    End If

End Function

Private Function FindHeaderCol(ByVal ws As Worksheet, ByVal headerName As String) As Long

    Dim lastCol As Long
    Dim c As Long

    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If StrComp(Trim$(CStr(ws.Cells(1, c).value)), headerName, vbTextCompare) = 0 Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c

    FindHeaderCol = 0

End Function

Private Function IsPOImportInstructionRow(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean

    Dim firstCell As String
    firstCell = Trim$(CStr(ws.Cells(rowNo, 1).value))

    IsPOImportInstructionRow = (UCase$(firstCell) = "NOTES:")

End Function

Private Function IsPOImportBlankRow(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean

    Dim c As Long

    For c = 1 To 12
        If Trim$(CStr(ws.Cells(rowNo, c).value)) <> "" Then
            IsPOImportBlankRow = False
            Exit Function
        End If
    Next c

    IsPOImportBlankRow = True

End Function

Private Function GetLastUsedRowInCols(ByVal ws As Worksheet, ByVal firstCol As Long, ByVal lastCol As Long) As Long

    Dim c As Long
    Dim tempRow As Long
    Dim maxRow As Long

    maxRow = 1

    For c = firstCol To lastCol
        tempRow = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If tempRow > maxRow Then maxRow = tempRow
    Next c

    GetLastUsedRowInCols = maxRow

End Function

Private Sub AddError(ByRef msg As String, ByVal newMsg As String)
    If msg <> "" Then msg = msg & " | "
    msg = msg & "ERROR: " & newMsg
End Sub

Private Sub AddWarning(ByRef msg As String, ByVal newMsg As String)
    If msg <> "" Then msg = msg & " | "
    msg = msg & "WARNING: " & newMsg
End Sub

Private Function HasErrorMessage(ByVal msg As String) As Boolean
    HasErrorMessage = (InStr(1, msg, "ERROR:", vbTextCompare) > 0)
End Function

Private Function NzNumber(ByVal v As Variant) As Double

    If IsError(v) Then
        NzNumber = 0
    ElseIf Trim$(CStr(v)) = "" Then
        NzNumber = 0
    ElseIf IsNumeric(v) Then
        NzNumber = CDbl(v)
    Else
        NzNumber = 0
    End If

End Function

Private Function GetCol(ByVal tbl As ListObject, ByVal colName As String) As Long
    GetCol = tbl.ListColumns(colName).Index
End Function

Private Function GetOptionalCol(ByVal tbl As ListObject, ByVal colName As String) As Long

    Dim lc As ListColumn

    GetOptionalCol = 0

    For Each lc In tbl.ListColumns
        If StrComp(Trim$(lc.name), Trim$(colName), vbTextCompare) = 0 Then
            GetOptionalCol = lc.Index
            Exit Function
        End If
    Next lc

End Function

Private Sub WriteOptionalValue(ByVal rowRange As Range, ByVal tbl As ListObject, ByVal colName As String, ByVal valueToWrite As Variant)

    Dim colNo As Long
    colNo = GetOptionalCol(tbl, colName)

    If colNo > 0 Then
        rowRange.Cells(1, colNo).value = valueToWrite
    End If

End Sub

Private Sub ClearValidationReport()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_VALIDATION)

    ws.Cells.Clear

    ws.Hyperlinks.Add _
        Anchor:=ws.Range("A1"), _
        Address:="", _
        SubAddress:=SHEET_IMPORT_UI & "!A1", _
        TextToDisplay:="Back to Import UI"

    ws.Range("A2").value = "Row"
    ws.Range("B2").value = "Status"
    ws.Range("C2").value = "External_PO_No"
    ws.Range("D2").value = "Purchase_Order_No"
    ws.Range("E2").value = "Supplier_Name"
    ws.Range("F2").value = "SKU"
    ws.Range("G2").value = "Qty"
    ws.Range("H2").value = "Received_Qty"
    ws.Range("I2").value = "Unit_Cost"
    ws.Range("J2").value = "Line_Total"
    ws.Range("K2").value = "Amount_Paid"
    ws.Range("L2").value = "Message"

    With ws.Range("A2:L2")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
    End With

    ws.Columns("A:L").ColumnWidth = 18
    ws.Columns("L:L").ColumnWidth = 80
    ws.Columns("L:L").WrapText = True

End Sub

Private Sub AppendValidationReport(ByVal srcWs As Worksheet, _
                                   ByVal rowNo As Long, _
                                   ByVal statusText As String, _
                                   ByVal messageText As String)

    Dim ws As Worksheet
    Dim nextRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_VALIDATION)

    nextRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If nextRow < 3 Then nextRow = 3

    ws.Cells(nextRow, 1).value = rowNo
    ws.Cells(nextRow, 2).value = statusText
    ws.Cells(nextRow, 3).value = GetCellByHeader(srcWs, rowNo, H_EXTERNAL_PO)
    ws.Cells(nextRow, 4).value = GetCellByHeader(srcWs, rowNo, H_PO_NO)
    ws.Cells(nextRow, 5).value = GetCellByHeader(srcWs, rowNo, H_SUPPLIER_NAME)
    ws.Cells(nextRow, 6).value = GetCellByHeader(srcWs, rowNo, H_SKU)
    ws.Cells(nextRow, 7).value = GetCellByHeader(srcWs, rowNo, H_QTY)
    ws.Cells(nextRow, 8).value = GetCellByHeader(srcWs, rowNo, H_RECEIVED_QTY)
    ws.Cells(nextRow, 9).value = GetCellByHeader(srcWs, rowNo, H_UNIT_COST)
    ws.Cells(nextRow, 10).value = GetCellByHeader(srcWs, rowNo, H_LINE_TOTAL)
    ws.Cells(nextRow, 11).value = GetCellByHeader(srcWs, rowNo, H_AMOUNT_PAID)
    ws.Cells(nextRow, 12).value = messageText

    With ws.Range(ws.Cells(nextRow, 1), ws.Cells(nextRow, 12))
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlTop
    End With

    If statusText = "ERROR" Then
        ws.Cells(nextRow, 2).Interior.Color = RGB(255, 199, 206)
    ElseIf statusText = "WARNING" Then
        ws.Cells(nextRow, 2).Interior.Color = RGB(255, 235, 156)
    End If

End Sub

Private Sub WritePOValidationStatus(ByVal errorCount As Long, _
                                    ByVal warningCount As Long, _
                                    ByVal validCount As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_IMPORT_UI)

    On Error Resume Next
    ws.Range(CELL_PO_LAST_VALIDATION).Hyperlinks.Delete
    On Error GoTo 0

    If errorCount > 0 Then
        ws.Range(CELL_PO_LAST_VALIDATION).value = "ERROR - " & errorCount & " row(s) | " & validCount & " valid row(s) (Click to view)"
        ws.Range(CELL_PO_LAST_VALIDATION).Interior.Color = RGB(255, 199, 206)
        AddValidationHyperlink ws.Range(CELL_PO_LAST_VALIDATION)
    ElseIf warningCount > 0 Then
        ws.Range(CELL_PO_LAST_VALIDATION).value = "WARNING - " & warningCount & " row(s) | " & validCount & " valid row(s) (Click to view)"
        ws.Range(CELL_PO_LAST_VALIDATION).Interior.Color = RGB(255, 235, 156)
        AddValidationHyperlink ws.Range(CELL_PO_LAST_VALIDATION)
    Else
        ws.Range(CELL_PO_LAST_VALIDATION).value = "VALID - " & validCount & " row(s)"
        ws.Range(CELL_PO_LAST_VALIDATION).Interior.Color = RGB(226, 239, 218)
    End If

End Sub

Private Sub AddValidationHyperlink(ByVal targetCell As Range)

    On Error Resume Next
    targetCell.Hyperlinks.Delete
    On Error GoTo 0

    targetCell.Worksheet.Hyperlinks.Add _
        Anchor:=targetCell, _
        Address:="", _
        SubAddress:=SHEET_VALIDATION & "!A1", _
        TextToDisplay:=CStr(targetCell.value)

End Sub

