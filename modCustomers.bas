Attribute VB_Name = "modCustomers"
Option Explicit

Private Const SHEET_UI As String = "Customers_UI"
Private Const TABLE_CUSTOMERS As String = "tblCustomers"
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
Private Const CELL_NOTES As String = "B12"
Private Const CELL_ACTIVE As String = "B13"

'==================================================
' NEW CUSTOMER
'==================================================
Public Sub Customer_New()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    
    wsUI.Unprotect Password:=SHEET_PWD
    ClearCustomerForm wsUI
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True

End Sub

'==================================================
' SAVE NEW CUSTOMER
'==================================================
Public Sub Customer_Save()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim customerID As String
    Dim customerName As String
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Customers_DB")
    Set tbl = wsDB.ListObjects(TABLE_CUSTOMERS)
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    If Trim(CStr(wsUI.Range(CELL_LOOKUP).value)) <> "" Or Trim(CStr(wsUI.Range(CELL_ID).value)) <> "" Then
        MsgBox "This form looks like an existing customer." & vbCrLf & _
               "Click New Customer first to create a new record.", vbExclamation
        GoTo SafeExit
    End If
    
    ValidateCustomerForm wsUI
    
    customerName = Trim(CStr(wsUI.Range(CELL_NAME).value))
    If CustomerNameExists(tbl, customerName) Then
        MsgBox "Customer Name already exists in Customers_DB.", vbCritical
        GoTo SafeExit
    End If
    
    customerID = GenerateCustomerID(tbl)
    wsUI.Range(CELL_ID).value = customerID
    
    Set newRow = tbl.ListRows.Add
    
    With newRow.Range
        .Cells(1, GetCol(tbl, "Customer_ID")).value = customerID
        .Cells(1, GetCol(tbl, "Customer_Name")).value = wsUI.Range(CELL_NAME).value
        .Cells(1, GetCol(tbl, "Contact_Person")).value = wsUI.Range(CELL_CONTACT).value
        .Cells(1, GetCol(tbl, "Email")).value = wsUI.Range(CELL_EMAIL).value
        .Cells(1, GetCol(tbl, "Phone")).value = wsUI.Range(CELL_PHONE).value
        .Cells(1, GetCol(tbl, "Address")).value = wsUI.Range(CELL_ADDRESS).value
        .Cells(1, GetCol(tbl, "Country")).value = wsUI.Range(CELL_COUNTRY).value
        .Cells(1, GetCol(tbl, "Notes")).value = wsUI.Range(CELL_NOTES).value
        .Cells(1, GetCol(tbl, "Active_Status")).value = wsUI.Range(CELL_ACTIVE).value
        .Cells(1, GetCol(tbl, "Created_At")).value = Date
        .Cells(1, GetCol(tbl, "Updated_At")).value = Date
    End With
    
    MsgBox "Customer saved successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' LOAD CUSTOMER
'==================================================
Public Sub Customer_Load()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tbl As ListObject
    Dim lookupValue As String
    Dim rowNum As Long
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Customers_DB")
    Set tbl = wsDB.ListObjects(TABLE_CUSTOMERS)
    
    lookupValue = Trim(CStr(wsUI.Range(CELL_LOOKUP).value))
    
    If lookupValue = "" Then
        MsgBox "Please enter Customer Lookup first.", vbExclamation
        Exit Sub
    End If
    
    rowNum = FindCustomerRow(tbl, lookupValue)
    If rowNum = 0 Then
        MsgBox "Customer not found.", vbExclamation
        Exit Sub
    End If
    
    wsUI.Unprotect Password:=SHEET_PWD
    ClearCustomerForm wsUI
    wsUI.Range(CELL_LOOKUP).value = lookupValue
    
    wsUI.Range(CELL_ID).value = tbl.ListColumns("Customer_ID").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_NAME).value = tbl.ListColumns("Customer_Name").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_CONTACT).value = tbl.ListColumns("Contact_Person").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_EMAIL).value = tbl.ListColumns("Email").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_PHONE).value = tbl.ListColumns("Phone").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_ADDRESS).value = tbl.ListColumns("Address").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_COUNTRY).value = tbl.ListColumns("Country").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_NOTES).value = tbl.ListColumns("Notes").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_ACTIVE).value = tbl.ListColumns("Active_Status").DataBodyRange.Cells(rowNum, 1).value
    
    MsgBox "Customer loaded successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' UPDATE CUSTOMER
'==================================================
Public Sub Customer_Update()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tbl As ListObject
    Dim rowNum As Long
    Dim customerID As String
    Dim customerName As String
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Customers_DB")
    Set tbl = wsDB.ListObjects(TABLE_CUSTOMERS)
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    customerID = Trim(CStr(wsUI.Range(CELL_ID).value))
    customerName = Trim(CStr(wsUI.Range(CELL_NAME).value))
    
    If customerID = "" Then
        MsgBox "Please load an existing customer before updating.", vbExclamation
        GoTo SafeExit
    End If
    
    ValidateCustomerForm wsUI
    
    rowNum = FindCustomerRowByID(tbl, customerID)
    If rowNum = 0 Then
        MsgBox "Customer ID not found in Customers_DB.", vbCritical
        GoTo SafeExit
    End If
    
    If CustomerNameUsedByAnother(tbl, customerName, customerID) Then
        MsgBox "This Customer Name is already used by another customer.", vbCritical
        GoTo SafeExit
    End If
    
    With tbl.DataBodyRange.Rows(rowNum)
        .Cells(1, GetCol(tbl, "Customer_Name")).value = wsUI.Range(CELL_NAME).value
        .Cells(1, GetCol(tbl, "Contact_Person")).value = wsUI.Range(CELL_CONTACT).value
        .Cells(1, GetCol(tbl, "Email")).value = wsUI.Range(CELL_EMAIL).value
        .Cells(1, GetCol(tbl, "Phone")).value = wsUI.Range(CELL_PHONE).value
        .Cells(1, GetCol(tbl, "Address")).value = wsUI.Range(CELL_ADDRESS).value
        .Cells(1, GetCol(tbl, "Country")).value = wsUI.Range(CELL_COUNTRY).value
        .Cells(1, GetCol(tbl, "Notes")).value = wsUI.Range(CELL_NOTES).value
        .Cells(1, GetCol(tbl, "Active_Status")).value = wsUI.Range(CELL_ACTIVE).value
        .Cells(1, GetCol(tbl, "Updated_At")).value = Date
    End With
    
    wsUI.Range(CELL_LOOKUP).value = wsUI.Range(CELL_NAME).value
    
    MsgBox "Customer updated successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' HELPERS
'==================================================
Private Sub ClearCustomerForm(ByVal wsUI As Worksheet)

    wsUI.Range(CELL_LOOKUP).value = ""
    wsUI.Range(CELL_ID).value = ""
    wsUI.Range(CELL_NAME).value = ""
    wsUI.Range(CELL_CONTACT).value = ""
    wsUI.Range(CELL_EMAIL).value = ""
    wsUI.Range(CELL_PHONE).value = ""
    wsUI.Range(CELL_ADDRESS).value = ""
    wsUI.Range(CELL_COUNTRY).value = ""
    wsUI.Range(CELL_NOTES).value = ""
    wsUI.Range(CELL_ACTIVE).value = "Active"

End Sub

Private Sub ValidateCustomerForm(ByVal wsUI As Worksheet)

    If Trim(CStr(wsUI.Range(CELL_NAME).value)) = "" Then
        Err.Raise vbObjectError + 1201, , "Please enter Customer Name."
    End If
    
    If Trim(CStr(wsUI.Range(CELL_ACTIVE).value)) = "" Then
        wsUI.Range(CELL_ACTIVE).value = "Active"
    End If

End Sub

Private Function GenerateCustomerID(ByVal tbl As ListObject) As String

    Dim i As Long, s As String, n As Long, maxNum As Long
    
    maxNum = 0
    
    If tbl.ListRows.Count = 0 Then
        GenerateCustomerID = "C00001"
        Exit Function
    End If
    
    For i = 1 To tbl.ListRows.Count
        s = Trim(CStr(tbl.ListColumns("Customer_ID").DataBodyRange.Cells(i, 1).value))
        If Left$(s, 1) = "C" Then
            If IsNumeric(Mid$(s, 2)) Then
                n = CLng(Mid$(s, 2))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i
    
    GenerateCustomerID = "C" & Format(maxNum + 1, "00000")

End Function

Private Function FindCustomerRow(ByVal tbl As ListObject, ByVal lookupValue As String) As Long

    Dim i As Long
    
    FindCustomerRow = 0
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("Customer_Name").DataBodyRange.Cells(i, 1).value)) = lookupValue Then
            FindCustomerRow = i
            Exit Function
        End If
        If Trim(CStr(tbl.ListColumns("Customer_ID").DataBodyRange.Cells(i, 1).value)) = lookupValue Then
            FindCustomerRow = i
            Exit Function
        End If
    Next i

End Function

Private Function FindCustomerRowByID(ByVal tbl As ListObject, ByVal customerID As String) As Long

    Dim i As Long
    
    FindCustomerRowByID = 0
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("Customer_ID").DataBodyRange.Cells(i, 1).value)) = customerID Then
            FindCustomerRowByID = i
            Exit Function
        End If
    Next i

End Function

Private Function CustomerNameExists(ByVal tbl As ListObject, ByVal customerName As String) As Boolean

    Dim i As Long
    
    CustomerNameExists = False
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("Customer_Name").DataBodyRange.Cells(i, 1).value)) = customerName Then
            CustomerNameExists = True
            Exit Function
        End If
    Next i

End Function

Private Function CustomerNameUsedByAnother(ByVal tbl As ListObject, ByVal customerName As String, ByVal currentCustomerID As String) As Boolean

    Dim i As Long
    
    CustomerNameUsedByAnother = False
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("Customer_Name").DataBodyRange.Cells(i, 1).value)) = customerName Then
            If Trim(CStr(tbl.ListColumns("Customer_ID").DataBodyRange.Cells(i, 1).value)) <> currentCustomerID Then
                CustomerNameUsedByAnother = True
                Exit Function
            End If
        End If
    Next i

End Function

Private Function GetCol(ByVal tbl As ListObject, ByVal colName As String) As Long
    GetCol = tbl.ListColumns(colName).Index
End Function

