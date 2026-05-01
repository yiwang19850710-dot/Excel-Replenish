Attribute VB_Name = "modSuppliers"
Option Explicit

Private Const SHEET_UI As String = "Suppliers_UI"
Private Const TABLE_SUPPLIERS As String = "tblSuppliers"
Private Const SHEET_PWD As String = ""

' UI Mapping
Private Const CELL_LOOKUP As String = "B3"
Private Const CELL_ID As String = "B5"
Private Const CELL_NAME As String = "B6"
Private Const CELL_CONTACT As String = "B7"
Private Const CELL_EMAIL As String = "B8"
Private Const CELL_PHONE As String = "B9"
Private Const CELL_ADDRESS As String = "B10"
Private Const CELL_COUNTRY As String = "B11"
Private Const CELL_CURRENCY As String = "B12"
Private Const CELL_PAYMENT_TERMS As String = "B13"
Private Const CELL_NOTES As String = "B14"
Private Const CELL_ACTIVE As String = "B15"

'==================================================
' NEW SUPPLIER
'==================================================
Public Sub Supplier_New()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    
    wsUI.Unprotect Password:=SHEET_PWD
    ClearSupplierForm wsUI
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True

End Sub

'==================================================
' SAVE NEW SUPPLIER
'==================================================
Public Sub Supplier_Save()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim supplierID As String
    Dim supplierName As String
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Suppliers_DB")
    Set tbl = wsDB.ListObjects(TABLE_SUPPLIERS)
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    If Trim(CStr(wsUI.Range(CELL_LOOKUP).value)) <> "" Or Trim(CStr(wsUI.Range(CELL_ID).value)) <> "" Then
        MsgBox "This form looks like an existing supplier." & vbCrLf & _
               "Click New Supplier first to create a new record.", vbExclamation
        GoTo SafeExit
    End If
    
    ValidateSupplierForm wsUI
    
    supplierName = Trim(CStr(wsUI.Range(CELL_NAME).value))
    If SupplierNameExists(tbl, supplierName) Then
        MsgBox "Supplier Name already exists in Suppliers_DB.", vbCritical
        GoTo SafeExit
    End If
    
    supplierID = GenerateSupplierID(tbl)
    wsUI.Range(CELL_ID).value = supplierID
    
    Set newRow = tbl.ListRows.Add
    
    With newRow.Range
        .Cells(1, GetCol(tbl, "Supplier_ID")).value = supplierID
        .Cells(1, GetCol(tbl, "Supplier_Name")).value = wsUI.Range(CELL_NAME).value
        .Cells(1, GetCol(tbl, "Contact_Person")).value = wsUI.Range(CELL_CONTACT).value
        .Cells(1, GetCol(tbl, "Email")).value = wsUI.Range(CELL_EMAIL).value
        .Cells(1, GetCol(tbl, "Phone")).value = wsUI.Range(CELL_PHONE).value
        .Cells(1, GetCol(tbl, "Address")).value = wsUI.Range(CELL_ADDRESS).value
        .Cells(1, GetCol(tbl, "Country")).value = wsUI.Range(CELL_COUNTRY).value
        .Cells(1, GetCol(tbl, "Default_Currency")).value = UCase$(Trim$(CStr(wsUI.Range(CELL_CURRENCY).value)))
        .Cells(1, GetCol(tbl, "Payment_Terms")).value = UCase$(Trim$(CStr(wsUI.Range(CELL_PAYMENT_TERMS).value)))
        .Cells(1, GetCol(tbl, "Notes")).value = wsUI.Range(CELL_NOTES).value
        .Cells(1, GetCol(tbl, "Active_Status")).value = wsUI.Range(CELL_ACTIVE).value
        .Cells(1, GetCol(tbl, "Created_At")).value = Date
        .Cells(1, GetCol(tbl, "Updated_At")).value = Date
    End With
    
    MsgBox "Supplier saved successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' LOAD SUPPLIER
'==================================================
Public Sub Supplier_Load()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tbl As ListObject
    Dim lookupValue As String
    Dim rowNum As Long
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Suppliers_DB")
    Set tbl = wsDB.ListObjects(TABLE_SUPPLIERS)
    
    lookupValue = Trim(CStr(wsUI.Range(CELL_LOOKUP).value))
    
    If lookupValue = "" Then
        MsgBox "Please enter Supplier Lookup first.", vbExclamation
        Exit Sub
    End If
    
    rowNum = FindSupplierRow(tbl, lookupValue)
    If rowNum = 0 Then
        MsgBox "Supplier not found.", vbExclamation
        Exit Sub
    End If
    
    wsUI.Unprotect Password:=SHEET_PWD
    ClearSupplierForm wsUI
    wsUI.Range(CELL_LOOKUP).value = lookupValue
    
    wsUI.Range(CELL_ID).value = tbl.ListColumns("Supplier_ID").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_NAME).value = tbl.ListColumns("Supplier_Name").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_CONTACT).value = tbl.ListColumns("Contact_Person").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_EMAIL).value = tbl.ListColumns("Email").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_PHONE).value = tbl.ListColumns("Phone").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_ADDRESS).value = tbl.ListColumns("Address").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_COUNTRY).value = tbl.ListColumns("Country").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_CURRENCY).value = GetOptionalDBValue(tbl, rowNum, "Default_Currency")
    wsUI.Range(CELL_PAYMENT_TERMS).value = GetOptionalDBValue(tbl, rowNum, "Payment_Terms")
    wsUI.Range(CELL_NOTES).value = tbl.ListColumns("Notes").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_ACTIVE).value = tbl.ListColumns("Active_Status").DataBodyRange.Cells(rowNum, 1).value
    
    MsgBox "Supplier loaded successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' UPDATE SUPPLIER
'==================================================
Public Sub Supplier_Update()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tbl As ListObject
    Dim rowNum As Long
    Dim supplierID As String
    Dim supplierName As String
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Suppliers_DB")
    Set tbl = wsDB.ListObjects(TABLE_SUPPLIERS)
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    supplierID = Trim(CStr(wsUI.Range(CELL_ID).value))
    supplierName = Trim(CStr(wsUI.Range(CELL_NAME).value))
    
    If supplierID = "" Then
        MsgBox "Please load an existing supplier before updating.", vbExclamation
        GoTo SafeExit
    End If
    
    ValidateSupplierForm wsUI
    
    rowNum = FindSupplierRowByID(tbl, supplierID)
    If rowNum = 0 Then
        MsgBox "Supplier ID not found in Suppliers_DB.", vbCritical
        GoTo SafeExit
    End If
    
    If SupplierNameUsedByAnother(tbl, supplierName, supplierID) Then
        MsgBox "This Supplier Name is already used by another supplier.", vbCritical
        GoTo SafeExit
    End If
    
    With tbl.DataBodyRange.Rows(rowNum)
        .Cells(1, GetCol(tbl, "Supplier_Name")).value = wsUI.Range(CELL_NAME).value
        .Cells(1, GetCol(tbl, "Contact_Person")).value = wsUI.Range(CELL_CONTACT).value
        .Cells(1, GetCol(tbl, "Email")).value = wsUI.Range(CELL_EMAIL).value
        .Cells(1, GetCol(tbl, "Phone")).value = wsUI.Range(CELL_PHONE).value
        .Cells(1, GetCol(tbl, "Address")).value = wsUI.Range(CELL_ADDRESS).value
        .Cells(1, GetCol(tbl, "Country")).value = wsUI.Range(CELL_COUNTRY).value
        .Cells(1, GetCol(tbl, "Default_Currency")).value = UCase$(Trim$(CStr(wsUI.Range(CELL_CURRENCY).value)))
        .Cells(1, GetCol(tbl, "Payment_Terms")).value = UCase$(Trim$(CStr(wsUI.Range(CELL_PAYMENT_TERMS).value)))
        .Cells(1, GetCol(tbl, "Notes")).value = wsUI.Range(CELL_NOTES).value
        .Cells(1, GetCol(tbl, "Active_Status")).value = wsUI.Range(CELL_ACTIVE).value
        .Cells(1, GetCol(tbl, "Updated_At")).value = Date
    End With
    
    wsUI.Range(CELL_LOOKUP).value = wsUI.Range(CELL_NAME).value
    
    MsgBox "Supplier updated successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' SETUP DROPDOWNS
'==================================================
Public Sub Supplier_SetupDropdowns()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    With wsUI.Range(CELL_CURRENCY).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="CAD,USD,CNY,EUR,GBP,AUD"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With

    With wsUI.Range(CELL_PAYMENT_TERMS).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="PREPAID,NET30,NET60,DEPOSIT30"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With

    With wsUI.Range(CELL_ACTIVE).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="Active,Inactive"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = False
    End With

End Sub

'==================================================
' HELPERS
'==================================================
Private Sub ClearSupplierForm(ByVal wsUI As Worksheet)

    wsUI.Range(CELL_LOOKUP).value = ""
    wsUI.Range(CELL_ID).value = ""
    wsUI.Range(CELL_NAME).value = ""
    wsUI.Range(CELL_CONTACT).value = ""
    wsUI.Range(CELL_EMAIL).value = ""
    wsUI.Range(CELL_PHONE).value = ""
    wsUI.Range(CELL_ADDRESS).value = ""
    wsUI.Range(CELL_COUNTRY).value = ""
    wsUI.Range(CELL_CURRENCY).value = ""
    wsUI.Range(CELL_PAYMENT_TERMS).value = ""
    wsUI.Range(CELL_NOTES).value = ""
    wsUI.Range(CELL_ACTIVE).value = "Active"

End Sub

Private Sub ValidateSupplierForm(ByVal wsUI As Worksheet)

    If Trim(CStr(wsUI.Range(CELL_NAME).value)) = "" Then
        Err.Raise vbObjectError + 1101, , "Please enter Supplier Name."
    End If
    
    If Trim(CStr(wsUI.Range(CELL_ACTIVE).value)) = "" Then
        wsUI.Range(CELL_ACTIVE).value = "Active"
    End If

    If Trim(CStr(wsUI.Range(CELL_CURRENCY).value)) <> "" Then
        If Len(Trim(CStr(wsUI.Range(CELL_CURRENCY).value))) <> 3 Then
            Err.Raise vbObjectError + 1102, , "Default Currency should be a 3-letter code, e.g. CAD, USD, CNY."
        End If
        wsUI.Range(CELL_CURRENCY).value = UCase$(Trim$(CStr(wsUI.Range(CELL_CURRENCY).value)))
    End If

    If Trim(CStr(wsUI.Range(CELL_PAYMENT_TERMS).value)) <> "" Then
        wsUI.Range(CELL_PAYMENT_TERMS).value = UCase$(Trim$(CStr(wsUI.Range(CELL_PAYMENT_TERMS).value)))
    End If

End Sub

Private Function GenerateSupplierID(ByVal tbl As ListObject) As String

    Dim i As Long, s As String, n As Long, maxNum As Long
    
    maxNum = 0
    
    If tbl.ListRows.Count = 0 Then
        GenerateSupplierID = "S00001"
        Exit Function
    End If
    
    For i = 1 To tbl.ListRows.Count
        s = Trim(CStr(tbl.ListColumns("Supplier_ID").DataBodyRange.Cells(i, 1).value))
        If Left$(s, 1) = "S" Then
            If IsNumeric(Mid$(s, 2)) Then
                n = CLng(Mid$(s, 2))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i
    
    GenerateSupplierID = "S" & Format(maxNum + 1, "00000")

End Function

Private Function FindSupplierRow(ByVal tbl As ListObject, ByVal lookupValue As String) As Long

    Dim i As Long
    
    FindSupplierRow = 0
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("Supplier_Name").DataBodyRange.Cells(i, 1).value)) = lookupValue Then
            FindSupplierRow = i
            Exit Function
        End If
        If Trim(CStr(tbl.ListColumns("Supplier_ID").DataBodyRange.Cells(i, 1).value)) = lookupValue Then
            FindSupplierRow = i
            Exit Function
        End If
    Next i

End Function

Private Function FindSupplierRowByID(ByVal tbl As ListObject, ByVal supplierID As String) As Long

    Dim i As Long
    
    FindSupplierRowByID = 0
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("Supplier_ID").DataBodyRange.Cells(i, 1).value)) = supplierID Then
            FindSupplierRowByID = i
            Exit Function
        End If
    Next i

End Function

Private Function SupplierNameExists(ByVal tbl As ListObject, ByVal supplierName As String) As Boolean

    Dim i As Long
    
    SupplierNameExists = False
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("Supplier_Name").DataBodyRange.Cells(i, 1).value)) = supplierName Then
            SupplierNameExists = True
            Exit Function
        End If
    Next i

End Function

Private Function SupplierNameUsedByAnother(ByVal tbl As ListObject, ByVal supplierName As String, ByVal currentSupplierID As String) As Boolean

    Dim i As Long
    
    SupplierNameUsedByAnother = False
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("Supplier_Name").DataBodyRange.Cells(i, 1).value)) = supplierName Then
            If Trim(CStr(tbl.ListColumns("Supplier_ID").DataBodyRange.Cells(i, 1).value)) <> currentSupplierID Then
                SupplierNameUsedByAnother = True
                Exit Function
            End If
        End If
    Next i

End Function

Private Function GetOptionalDBValue(ByVal tbl As ListObject, ByVal rowNum As Long, ByVal colName As String) As Variant

    Dim lc As ListColumn

    For Each lc In tbl.ListColumns
        If StrComp(Trim$(lc.name), Trim$(colName), vbTextCompare) = 0 Then
            GetOptionalDBValue = tbl.DataBodyRange.Cells(rowNum, lc.Index).value
            Exit Function
        End If
    Next lc

    GetOptionalDBValue = ""

End Function

Private Function GetCol(ByVal tbl As ListObject, ByVal colName As String) As Long
    GetCol = tbl.ListColumns(colName).Index
End Function

