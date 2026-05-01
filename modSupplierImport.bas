Attribute VB_Name = "modSupplierImport"
Option Explicit

Private Const SHEET_IMPORT_UI As String = "Import_UI"
Private Const SHEET_VALIDATION As String = "Import_Validation"

Private Const SHEET_SUPPLIERS As String = "Suppliers_DB"
Private Const TABLE_SUPPLIERS As String = "tblSuppliers"

Private Const CELL_IMPORT_TYPE As String = "B36"
Private Const CELL_IMPORT_FILE As String = "B37"
Private Const CELL_LAST_VALIDATION As String = "B38"
Private Const CELL_LAST_RESULT As String = "B39"
Private Const CELL_IMPORT_OPTION As String = "B40"

Private Const IMPORT_OPTION_VALID_ONLY As String = "IMPORT_VALID_ONLY"
Private Const IMPORT_OPTION_STOP_ON_ERROR As String = "STOP_ON_ERROR"

Private Const DUP_ASK As String = "ASK"
Private Const DUP_SKIP_ALL As String = "SKIP_ALL"
Private Const DUP_REPLACE_ALL As String = "REPLACE_ALL"
Private Const DUP_STOP As String = "STOP"

Private Const H_SUPPLIER_ID As String = "Supplier_ID"
Private Const H_SUPPLIER_NAME As String = "Supplier_Name"
Private Const H_CONTACT_PERSON As String = "Contact_Person"
Private Const H_EMAIL As String = "Email"
Private Const H_PHONE As String = "Phone"
Private Const H_ADDRESS As String = "Address"
Private Const H_COUNTRY As String = "Country"
Private Const H_DEFAULT_CURRENCY As String = "Default_Currency"
Private Const H_PAYMENT_TERMS As String = "Payment_Terms"
Private Const H_NOTES As String = "Notes"
Private Const H_ACTIVE_STATUS As String = "Active_Status"
Private Const H_CREATED_AT As String = "Created_At"

'==================================================
' PUBLIC ACTIONS
'==================================================
Public Sub SupplierImport_DownloadTemplate()

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim tblTemplate As ListObject
    Dim savePath As Variant

    On Error GoTo ErrHandler

    Set wb = Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.name = "Supplier_Import_Template"

    ws.Range("A1").value = H_SUPPLIER_ID
    ws.Range("B1").value = H_SUPPLIER_NAME
    ws.Range("C1").value = H_CONTACT_PERSON
    ws.Range("D1").value = H_EMAIL
    ws.Range("E1").value = H_PHONE
    ws.Range("F1").value = H_ADDRESS
    ws.Range("G1").value = H_COUNTRY
    ws.Range("H1").value = H_DEFAULT_CURRENCY
    ws.Range("I1").value = H_PAYMENT_TERMS
    ws.Range("J1").value = H_NOTES
    ws.Range("K1").value = H_ACTIVE_STATUS
    ws.Range("L1").value = H_CREATED_AT

    Set tblTemplate = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:L2"), , xlYes)
    tblTemplate.name = "tblSupplierImportTemplate"
    tblTemplate.TableStyle = "TableStyleMedium2"

    ws.Columns("A:L").ColumnWidth = 22
    ws.Columns("L:L").NumberFormat = "yyyy-mm-dd"

    ws.Range("N1").value = "Notes:"
    ws.Range("N2").value = "1. Supplier_ID can be blank. If blank, system will auto-generate Supplier_ID."
    ws.Range("N3").value = "2. Supplier_Name is required and used for duplicate checking."
    ws.Range("N4").value = "3. Active_Status can be blank. If blank, system will default to Active."
    ws.Range("N5").value = "4. Created_At can be blank. If blank, system will default to Today."
    ws.Range("N6").value = "5. Default_Currency should use standard codes, e.g. CAD, USD, CNY."
    ws.Range("N7").value = "6. Payment_Terms examples: PREPAID, NET30, NET60, DEPOSIT30."

    ws.Columns("N:N").ColumnWidth = 95
    ws.Columns("N:N").WrapText = True

    With ws.Range("N1:N7")
        .Interior.Color = RGB(255, 242, 204)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlTop
    End With
    ws.Range("N1").Font.Bold = True

    With ws.Range("H2:H1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="CAD,USD,CNY,EUR,GBP,AUD"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With

    With ws.Range("K2:K1000").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="Active,Inactive"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With

    savePath = Application.GetSaveAsFilename( _
        InitialFileName:="Supplier_Import_Template.xlsx", _
        FileFilter:="Excel Workbook (*.xlsx), *.xlsx")

    If savePath = False Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If

    Application.DisplayAlerts = False
    wb.SaveAs Filename:=CStr(savePath), FileFormat:=xlOpenXMLWorkbook
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True

    MsgBox "Supplier Import template created successfully.", vbInformation, "Supplier Import"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    MsgBox "Error creating Supplier Import template: " & Err.Description, vbCritical, "Supplier Import"

End Sub

Public Sub SupplierImport_SelectFile()

    Dim fd As FileDialog
    Dim selectedPath As String
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(SHEET_IMPORT_UI)
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "Select Supplier Import File"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Excel / CSV Files", "*.xlsx;*.xls;*.csv"

        If .Show <> -1 Then Exit Sub
        selectedPath = .SelectedItems(1)
    End With

    ws.Range(CELL_IMPORT_FILE).value = selectedPath
    ws.Range(CELL_LAST_VALIDATION).value = ""
    ws.Range(CELL_LAST_RESULT).value = ""
    ws.Range(CELL_IMPORT_OPTION).value = IMPORT_OPTION_VALID_ONLY

End Sub

Public Sub SupplierImport_Clear()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_IMPORT_UI)

    ws.Range(CELL_IMPORT_TYPE).value = "Supplier Import"
    ws.Range(CELL_IMPORT_FILE).value = ""
    ws.Range(CELL_LAST_VALIDATION).value = ""
    ws.Range(CELL_LAST_RESULT).value = ""
    ws.Range(CELL_IMPORT_OPTION).value = IMPORT_OPTION_VALID_ONLY
    ws.Range(CELL_LAST_VALIDATION).Interior.Pattern = xlNone
    ws.Range(CELL_LAST_RESULT).Interior.Pattern = xlNone

End Sub

Public Sub SupplierImport_Validate()

    Dim filePath As String
    Dim errorCount As Long
    Dim warningCount As Long
    Dim validCount As Long

    filePath = Trim$(CStr(ThisWorkbook.Worksheets(SHEET_IMPORT_UI).Range(CELL_IMPORT_FILE).value))

    If filePath = "" Then
        MsgBox "Please select a Supplier import file first.", vbExclamation, "Supplier Import"
        Exit Sub
    End If

    ValidateSupplierImportFile filePath, errorCount, warningCount, validCount
    WriteSupplierValidationStatus errorCount, warningCount, validCount

End Sub

Public Sub SupplierImport_RunImport()

    Dim filePath As String
    Dim importOption As String
    Dim errorCount As Long
    Dim warningCount As Long
    Dim validCount As Long
    Dim importedCount As Long

    filePath = Trim$(CStr(ThisWorkbook.Worksheets(SHEET_IMPORT_UI).Range(CELL_IMPORT_FILE).value))
    importOption = UCase$(Trim$(CStr(ThisWorkbook.Worksheets(SHEET_IMPORT_UI).Range(CELL_IMPORT_OPTION).value)))

    If filePath = "" Then
        MsgBox "Please select a Supplier import file first.", vbExclamation, "Supplier Import"
        Exit Sub
    End If

    ValidateSupplierImportFile filePath, errorCount, warningCount, validCount
    WriteSupplierValidationStatus errorCount, warningCount, validCount

    If importOption = IMPORT_OPTION_STOP_ON_ERROR And errorCount > 0 Then
        With ThisWorkbook.Worksheets(SHEET_IMPORT_UI)
            .Range(CELL_LAST_RESULT).value = "IMPORT STOPPED - validation error(s) found"
            .Range(CELL_LAST_RESULT).Interior.Color = RGB(255, 199, 206)
        End With
        MsgBox "Import stopped because validation errors were found.", vbExclamation, "Supplier Import"
        Exit Sub
    End If

    If validCount = 0 Then
        MsgBox "No valid Supplier rows to import.", vbExclamation, "Supplier Import"
        Exit Sub
    End If

    importedCount = ImportValidSupplierRows(filePath)

    With ThisWorkbook.Worksheets(SHEET_IMPORT_UI)
        If importedCount >= 0 Then
            .Range(CELL_LAST_RESULT).value = "IMPORT SUCCESS - " & importedCount & " valid row(s)"
            .Range(CELL_LAST_RESULT).Interior.Color = RGB(226, 239, 218)
        Else
            .Range(CELL_LAST_RESULT).value = "IMPORT STOPPED BY USER"
            .Range(CELL_LAST_RESULT).Interior.Color = RGB(255, 235, 156)
        End If
    End With

    If importedCount >= 0 Then
        MsgBox "Supplier Import completed. Imported " & importedCount & " valid row(s).", vbInformation, "Supplier Import"
    Else
        MsgBox "Supplier Import stopped by user.", vbExclamation, "Supplier Import"
    End If

End Sub

'==================================================
' VALIDATION
'==================================================
Private Sub ValidateSupplierImportFile(ByVal filePath As String, _
                                       ByRef errorCount As Long, _
                                       ByRef warningCount As Long, _
                                       ByRef validCount As Long)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim rowStatus As String
    Dim msg As String

    Dim dictNameFile As Object
    Dim dictExistingNames As Object

    On Error GoTo ErrHandler

    ClearValidationReport

    Set dictNameFile = CreateObject("Scripting.Dictionary")
    Set dictExistingNames = BuildExistingSupplierNameDict()

    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws = wb.Worksheets(1)

    lastRow = GetLastUsedRowInCols(ws, 1, 12)

    ProgressStart "Validating Supplier Import", "Checking supplier import file..."

    For r = 2 To lastRow

        If IsSupplierImportInstructionRow(ws, r) Then Exit For
        If IsSupplierImportBlankRow(ws, r) Then GoTo NextValidationRow

        ProgressUpdate "Validating row", r - 1, lastRow - 1, "Row " & r

        ValidateSupplierImportRow ws, r, dictNameFile, dictExistingNames, rowStatus, msg

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
    MsgBox "Validation error: " & Err.Description, vbCritical, "Supplier Import"

End Sub

Private Sub ValidateSupplierImportRow(ByVal ws As Worksheet, _
                                      ByVal r As Long, _
                                      ByVal dictNameFile As Object, _
                                      ByVal dictExistingNames As Object, _
                                      ByRef rowStatus As String, _
                                      ByRef msg As String)

    Dim supplierName As String
    Dim email As String
    Dim activeStatus As String
    Dim createdAt As Variant
    Dim currencyCode As String
    Dim key As String

    rowStatus = "VALID"
    msg = ""

    supplierName = Trim$(CStr(GetCellByHeader(ws, r, H_SUPPLIER_NAME)))
    email = Trim$(CStr(GetCellByHeader(ws, r, H_EMAIL)))
    activeStatus = Trim$(CStr(GetCellByHeader(ws, r, H_ACTIVE_STATUS)))
    createdAt = GetCellByHeader(ws, r, H_CREATED_AT)
    currencyCode = Trim$(CStr(GetCellByHeader(ws, r, H_DEFAULT_CURRENCY)))

    If supplierName = "" Then AddError msg, "Missing Supplier_Name."

    If supplierName <> "" Then
        key = UCase$(supplierName)

        If dictNameFile.Exists(key) Then
            AddError msg, "Duplicate Supplier_Name inside import file."
        Else
            dictNameFile.Add key, True
        End If

        If dictExistingNames.Exists(key) Then
            AddWarning msg, "Supplier_Name already exists in Suppliers_DB. You can choose Skip or Replace during import."
        End If
    End If

    If email <> "" Then
        If Not IsValidEmail(email) Then AddError msg, "Invalid Email format."
    End If

    If activeStatus <> "" Then
        If UCase$(activeStatus) <> "ACTIVE" And UCase$(activeStatus) <> "INACTIVE" Then
            AddError msg, "Active_Status must be Active or Inactive."
        End If
    End If

    If Trim$(CStr(createdAt)) <> "" Then
        If Not IsDate(createdAt) Then AddError msg, "Invalid Created_At."
    End If

    If currencyCode <> "" Then
        If Len(currencyCode) <> 3 Then AddWarning msg, "Default_Currency should be a 3-letter code, e.g. CAD, USD, CNY."
    End If

    If HasErrorMessage(msg) Then
        rowStatus = "ERROR"
        Exit Sub
    End If

    If activeStatus = "" Then AddWarning msg, "Active_Status blank. Default to Active."
    If Trim$(CStr(createdAt)) = "" Then AddWarning msg, "Created_At blank. Default to Today."

    If msg <> "" Then rowStatus = "WARNING"

End Sub

'==================================================
' IMPORT
'==================================================
Private Function ImportValidSupplierRows(ByVal filePath As String) As Long

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long

    Dim dictNameFile As Object
    Dim dictExistingNames As Object

    Dim rowStatus As String
    Dim msg As String
    Dim duplicateMode As String
    Dim supplierName As String
    Dim key As String
    Dim duplicateAction As String

    On Error GoTo ErrHandler

    Set dictNameFile = CreateObject("Scripting.Dictionary")
    Set dictExistingNames = BuildExistingSupplierNameDict()

    duplicateMode = DUP_ASK

    Application.ScreenUpdating = False
    Set wb = Workbooks.Open(filePath, ReadOnly:=True)
    Set ws = wb.Worksheets(1)

    lastRow = GetLastUsedRowInCols(ws, 1, 12)

    ProgressStart "Importing Suppliers", "Writing valid supplier rows..."

    For r = 2 To lastRow

        If IsSupplierImportInstructionRow(ws, r) Then Exit For
        If IsSupplierImportBlankRow(ws, r) Then GoTo NextImportRow

        ProgressUpdate "Importing row", r - 1, lastRow - 1, "Row " & r

        ValidateSupplierImportRow ws, r, dictNameFile, dictExistingNames, rowStatus, msg

        If rowStatus = "VALID" Or rowStatus = "WARNING" Then

            supplierName = Trim$(CStr(GetCellByHeader(ws, r, H_SUPPLIER_NAME)))
            key = UCase$(supplierName)

            If key <> "" And dictExistingNames.Exists(key) Then

                duplicateAction = ResolveDuplicateSupplierAction(ws, r, duplicateMode)

                If duplicateAction = DUP_STOP Then
                    ImportValidSupplierRows = -1
                    GoTo CleanExit
                ElseIf duplicateAction = DUP_SKIP_ALL Then
                    duplicateMode = DUP_SKIP_ALL
                    GoTo NextImportRow
                ElseIf duplicateAction = DUP_REPLACE_ALL Then
                    duplicateMode = DUP_REPLACE_ALL
                    ReplaceExistingSupplierRow ws, r, supplierName
                ElseIf duplicateAction = "SKIP_THIS" Then
                    GoTo NextImportRow
                ElseIf duplicateAction = "REPLACE_THIS" Then
                    ReplaceExistingSupplierRow ws, r, supplierName
                End If

            Else
                WriteSupplierImportRow ws, r
                If key <> "" And Not dictExistingNames.Exists(key) Then dictExistingNames.Add key, True
            End If

            ImportValidSupplierRows = ImportValidSupplierRows + 1

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
    MsgBox "Import error: " & Err.Description, vbCritical, "Supplier Import"

End Function

Private Sub WriteSupplierImportRow(ByVal ws As Worksheet, ByVal r As Long)

    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim supplierID As String
    Dim supplierName As String
    Dim activeStatus As String
    Dim createdAt As Date

    Set tbl = ThisWorkbook.Worksheets(SHEET_SUPPLIERS).ListObjects(TABLE_SUPPLIERS)

    supplierID = Trim$(CStr(GetCellByHeader(ws, r, H_SUPPLIER_ID)))
    supplierName = Trim$(CStr(GetCellByHeader(ws, r, H_SUPPLIER_NAME)))
    activeStatus = NormalizeActiveStatus(Trim$(CStr(GetCellByHeader(ws, r, H_ACTIVE_STATUS))))

    If supplierID = "" Then supplierID = GenerateNextSupplierID()

    If IsDate(GetCellByHeader(ws, r, H_CREATED_AT)) Then
        createdAt = CDate(GetCellByHeader(ws, r, H_CREATED_AT))
    Else
        createdAt = Date
    End If

    Set newRow = tbl.ListRows.Add

    WriteSupplierRowFromImport newRow.Range, tbl, ws, r, supplierID, supplierName, activeStatus, createdAt

End Sub

Private Sub ReplaceExistingSupplierRow(ByVal ws As Worksheet, ByVal r As Long, ByVal supplierName As String)

    Dim tbl As ListObject
    Dim i As Long
    Dim colName As Long
    Dim existingID As String
    Dim supplierID As String
    Dim activeStatus As String
    Dim createdAt As Date

    Set tbl = ThisWorkbook.Worksheets(SHEET_SUPPLIERS).ListObjects(TABLE_SUPPLIERS)

    If tbl.DataBodyRange Is Nothing Then Exit Sub

    colName = GetCol(tbl, "Supplier_Name")

    For i = 1 To tbl.ListRows.Count

        If StrComp(Trim$(CStr(tbl.DataBodyRange.Cells(i, colName).value)), supplierName, vbTextCompare) = 0 Then

            existingID = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Supplier_ID")).value))
            supplierID = Trim$(CStr(GetCellByHeader(ws, r, H_SUPPLIER_ID)))

            If supplierID = "" Then supplierID = existingID
            If supplierID = "" Then supplierID = GenerateNextSupplierID()

            activeStatus = NormalizeActiveStatus(Trim$(CStr(GetCellByHeader(ws, r, H_ACTIVE_STATUS))))

            If IsDate(GetCellByHeader(ws, r, H_CREATED_AT)) Then
                createdAt = CDate(GetCellByHeader(ws, r, H_CREATED_AT))
            Else
                createdAt = NzDate(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Created_At")).value, Date)
            End If

            WriteSupplierRowFromImport tbl.ListRows(i).Range, tbl, ws, r, supplierID, supplierName, activeStatus, createdAt
            Exit Sub

        End If
    Next i

End Sub

Private Sub WriteSupplierRowFromImport(ByVal rowRange As Range, _
                                       ByVal tbl As ListObject, _
                                       ByVal ws As Worksheet, _
                                       ByVal r As Long, _
                                       ByVal supplierID As String, _
                                       ByVal supplierName As String, _
                                       ByVal activeStatus As String, _
                                       ByVal createdAt As Date)

    rowRange.Cells(1, GetCol(tbl, "Supplier_ID")).value = supplierID
    rowRange.Cells(1, GetCol(tbl, "Supplier_Name")).value = supplierName
    rowRange.Cells(1, GetCol(tbl, "Contact_Person")).value = GetCellByHeader(ws, r, H_CONTACT_PERSON)
    rowRange.Cells(1, GetCol(tbl, "Email")).value = GetCellByHeader(ws, r, H_EMAIL)
    rowRange.Cells(1, GetCol(tbl, "Phone")).value = GetCellByHeader(ws, r, H_PHONE)
    rowRange.Cells(1, GetCol(tbl, "Address")).value = GetCellByHeader(ws, r, H_ADDRESS)
    rowRange.Cells(1, GetCol(tbl, "Country")).value = GetCellByHeader(ws, r, H_COUNTRY)
    rowRange.Cells(1, GetCol(tbl, "Default_Currency")).value = UCase$(Trim$(CStr(GetCellByHeader(ws, r, H_DEFAULT_CURRENCY))))
    rowRange.Cells(1, GetCol(tbl, "Payment_Terms")).value = UCase$(Trim$(CStr(GetCellByHeader(ws, r, H_PAYMENT_TERMS))))
    rowRange.Cells(1, GetCol(tbl, "Notes")).value = GetCellByHeader(ws, r, H_NOTES)
    rowRange.Cells(1, GetCol(tbl, "Active_Status")).value = activeStatus
    rowRange.Cells(1, GetCol(tbl, "Created_At")).value = createdAt
    rowRange.Cells(1, GetCol(tbl, "Updated_At")).value = Now

End Sub

'==================================================
' DUPLICATE HANDLING
'==================================================
Private Function ResolveDuplicateSupplierAction(ByVal ws As Worksheet, _
                                                ByVal rowNo As Long, _
                                                ByVal duplicateMode As String) As String

    Dim supplierName As String
    Dim action As String

    If duplicateMode = DUP_SKIP_ALL Then
        ResolveDuplicateSupplierAction = DUP_SKIP_ALL
        Exit Function
    End If

    If duplicateMode = DUP_REPLACE_ALL Then
        ResolveDuplicateSupplierAction = DUP_REPLACE_ALL
        Exit Function
    End If

    supplierName = Trim$(CStr(GetCellByHeader(ws, rowNo, H_SUPPLIER_NAME)))

    frmDuplicateOrderAction.SelectedAction = ""
    frmDuplicateOrderAction.SetOrderInfo supplierName
    frmDuplicateOrderAction.Show vbModal

    action = UCase$(Trim$(frmDuplicateOrderAction.SelectedAction))

    Select Case action
        Case "SKIP_ONE"
            ResolveDuplicateSupplierAction = "SKIP_THIS"
        Case "SKIP_ALL"
            ResolveDuplicateSupplierAction = DUP_SKIP_ALL
        Case "REPLACE_ONE"
            ResolveDuplicateSupplierAction = "REPLACE_THIS"
        Case "REPLACE_ALL"
            ResolveDuplicateSupplierAction = DUP_REPLACE_ALL
        Case "STOP", ""
            ResolveDuplicateSupplierAction = DUP_STOP
        Case Else
            ResolveDuplicateSupplierAction = DUP_STOP
    End Select

End Function

'==================================================
' DICTIONARIES
'==================================================
Private Function BuildExistingSupplierNameDict() As Object

    Dim dict As Object
    Dim tbl As ListObject
    Dim i As Long
    Dim supplierName As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set tbl = ThisWorkbook.Worksheets(SHEET_SUPPLIERS).ListObjects(TABLE_SUPPLIERS)

    If tbl.DataBodyRange Is Nothing Then
        Set BuildExistingSupplierNameDict = dict
        Exit Function
    End If

    For i = 1 To tbl.ListRows.Count
        supplierName = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Supplier_Name")).value))

        If supplierName <> "" Then
            If Not dict.Exists(UCase$(supplierName)) Then
                dict.Add UCase$(supplierName), True
            End If
        End If
    Next i

    Set BuildExistingSupplierNameDict = dict

End Function

'==================================================
' HELPERS
'==================================================
Private Function GenerateNextSupplierID() As String

    Dim tbl As ListObject
    Dim i As Long
    Dim supplierID As String
    Dim maxNo As Long
    Dim n As Long

    Set tbl = ThisWorkbook.Worksheets(SHEET_SUPPLIERS).ListObjects(TABLE_SUPPLIERS)

    If Not tbl.DataBodyRange Is Nothing Then
        For i = 1 To tbl.ListRows.Count
            supplierID = Trim$(CStr(tbl.DataBodyRange.Cells(i, GetCol(tbl, "Supplier_ID")).value))

            If UCase$(Left$(supplierID, 1)) = "S" Then
                n = CLng(val(Mid$(supplierID, 2)))
                If n > maxNo Then maxNo = n
            End If
        Next i
    End If

    GenerateNextSupplierID = "S" & Format$(maxNo + 1, "00000")

End Function

Private Function NormalizeActiveStatus(ByVal activeStatus As String) As String

    If Trim$(activeStatus) = "" Then
        NormalizeActiveStatus = "Active"
    ElseIf UCase$(activeStatus) = "INACTIVE" Then
        NormalizeActiveStatus = "Inactive"
    Else
        NormalizeActiveStatus = "Active"
    End If

End Function

Private Function IsValidEmail(ByVal emailText As String) As Boolean

    IsValidEmail = (InStr(1, emailText, "@") > 1 And InStrRev(emailText, ".") > InStr(1, emailText, "@") + 1)

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

Private Function IsSupplierImportInstructionRow(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean

    Dim firstCell As String
    firstCell = Trim$(CStr(ws.Cells(rowNo, 1).value))

    IsSupplierImportInstructionRow = (UCase$(firstCell) = "NOTES:")

End Function

Private Function IsSupplierImportBlankRow(ByVal ws As Worksheet, ByVal rowNo As Long) As Boolean

    Dim c As Long

    For c = 1 To 12
        If Trim$(CStr(ws.Cells(rowNo, c).value)) <> "" Then
            IsSupplierImportBlankRow = False
            Exit Function
        End If
    Next c

    IsSupplierImportBlankRow = True

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

Private Function NzDate(ByVal v As Variant, ByVal fallbackDate As Date) As Date

    If IsDate(v) Then
        NzDate = CDate(v)
    Else
        NzDate = fallbackDate
    End If

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

Private Function GetCol(ByVal tbl As ListObject, ByVal colName As String) As Long
    GetCol = tbl.ListColumns(colName).Index
End Function

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
    ws.Range("C2").value = "Supplier_ID"
    ws.Range("D2").value = "Supplier_Name"
    ws.Range("E2").value = "Contact_Person"
    ws.Range("F2").value = "Email"
    ws.Range("G2").value = "Phone"
    ws.Range("H2").value = "Country"
    ws.Range("I2").value = "Default_Currency"
    ws.Range("J2").value = "Payment_Terms"
    ws.Range("K2").value = "Active_Status"
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
    ws.Cells(nextRow, 3).value = GetCellByHeader(srcWs, rowNo, H_SUPPLIER_ID)
    ws.Cells(nextRow, 4).value = GetCellByHeader(srcWs, rowNo, H_SUPPLIER_NAME)
    ws.Cells(nextRow, 5).value = GetCellByHeader(srcWs, rowNo, H_CONTACT_PERSON)
    ws.Cells(nextRow, 6).value = GetCellByHeader(srcWs, rowNo, H_EMAIL)
    ws.Cells(nextRow, 7).value = GetCellByHeader(srcWs, rowNo, H_PHONE)
    ws.Cells(nextRow, 8).value = GetCellByHeader(srcWs, rowNo, H_COUNTRY)
    ws.Cells(nextRow, 9).value = GetCellByHeader(srcWs, rowNo, H_DEFAULT_CURRENCY)
    ws.Cells(nextRow, 10).value = GetCellByHeader(srcWs, rowNo, H_PAYMENT_TERMS)
    ws.Cells(nextRow, 11).value = GetCellByHeader(srcWs, rowNo, H_ACTIVE_STATUS)
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

Private Sub WriteSupplierValidationStatus(ByVal errorCount As Long, _
                                          ByVal warningCount As Long, _
                                          ByVal validCount As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_IMPORT_UI)

    On Error Resume Next
    ws.Range(CELL_LAST_VALIDATION).Hyperlinks.Delete
    On Error GoTo 0

    If errorCount > 0 Then
        ws.Range(CELL_LAST_VALIDATION).value = "ERROR - " & errorCount & " row(s) | " & validCount & " valid row(s) (Click to view)"
        ws.Range(CELL_LAST_VALIDATION).Interior.Color = RGB(255, 199, 206)
        AddValidationHyperlink ws.Range(CELL_LAST_VALIDATION)
    ElseIf warningCount > 0 Then
        ws.Range(CELL_LAST_VALIDATION).value = "WARNING - " & warningCount & " row(s) | " & validCount & " valid row(s) (Click to view)"
        ws.Range(CELL_LAST_VALIDATION).Interior.Color = RGB(255, 235, 156)
        AddValidationHyperlink ws.Range(CELL_LAST_VALIDATION)
    Else
        ws.Range(CELL_LAST_VALIDATION).value = "VALID - " & validCount & " row(s)"
        ws.Range(CELL_LAST_VALIDATION).Interior.Color = RGB(226, 239, 218)
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

