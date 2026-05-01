Attribute VB_Name = "modProducts"
Option Explicit

Private Const SHEET_UI As String = "Products_UI"
Private Const TABLE_PRODUCTS As String = "tblProducts"
Private Const SHEET_PWD As String = ""

' UI Mapping
Private Const CELL_LOOKUP As String = "B3"
Private Const CELL_ID As String = "B5"
Private Const CELL_SKU As String = "B6"
Private Const CELL_NAME As String = "B7"
Private Const CELL_VARIANT As String = "B8"
Private Const CELL_CATEGORY As String = "B9"
Private Const CELL_UNIT_COST As String = "B10"
Private Const CELL_SELLING_PRICE As String = "B11"
Private Const CELL_OPENING_STOCK As String = "B12"
Private Const CELL_REORDER_LEVEL As String = "B13"
Private Const CELL_SAFETY_DAYS As String = "B14"
Private Const CELL_LEAD_TIME As String = "B15"
Private Const CELL_ACTIVE As String = "B16"
Private Const CELL_NOTES As String = "B17"

'==================================================
' NEW PRODUCT
'==================================================
Public Sub Product_New()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    
    wsUI.Unprotect Password:=SHEET_PWD
    ClearProductForm wsUI
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True

End Sub

'==================================================
' SAVE NEW PRODUCT
'==================================================
Public Sub Product_Save()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim productID As String
    Dim sku As String
    Dim openingStock As Double
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Products_DB")
    Set tbl = wsDB.ListObjects(TABLE_PRODUCTS)
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    If Trim(CStr(wsUI.Range(CELL_LOOKUP).value)) <> "" Or Trim(CStr(wsUI.Range(CELL_ID).value)) <> "" Then
        MsgBox "This form looks like an existing product." & vbCrLf & _
               "Click New Product first to create a new item.", vbExclamation
        GoTo SafeExit
    End If
    
    ValidateProductForm wsUI, tbl, False
    
    sku = Trim(CStr(wsUI.Range(CELL_SKU).value))
    If ProductSKUExists(tbl, sku) Then
        MsgBox "SKU already exists in Products_DB.", vbCritical
        GoTo SafeExit
    End If
    
    productID = GenerateProductID(tbl)
    wsUI.Range(CELL_ID).value = productID
    
    openingStock = NzNumber(wsUI.Range(CELL_OPENING_STOCK).value)
    
    Set newRow = tbl.ListRows.Add
    
    With newRow.Range
        .Cells(1, GetCol(tbl, "Product_ID")).value = productID
        .Cells(1, GetCol(tbl, "SKU")).value = wsUI.Range(CELL_SKU).value
        .Cells(1, GetCol(tbl, "Product_Name")).value = wsUI.Range(CELL_NAME).value
        .Cells(1, GetCol(tbl, "Variant_Desc")).value = wsUI.Range(CELL_VARIANT).value
        .Cells(1, GetCol(tbl, "Category")).value = wsUI.Range(CELL_CATEGORY).value
        .Cells(1, GetCol(tbl, "Unit_Cost")).value = NzNumber(wsUI.Range(CELL_UNIT_COST).value)
        .Cells(1, GetCol(tbl, "Selling_Price")).value = NzNumber(wsUI.Range(CELL_SELLING_PRICE).value)
        .Cells(1, GetCol(tbl, "Opening_Stock")).value = openingStock
        .Cells(1, GetCol(tbl, "Current_Stock")).value = openingStock
        .Cells(1, GetCol(tbl, "Reorder_Level")).value = NzNumber(wsUI.Range(CELL_REORDER_LEVEL).value)
        .Cells(1, GetCol(tbl, "Safety_Days_Override")).value = NzNumber(wsUI.Range(CELL_SAFETY_DAYS).value)
        .Cells(1, GetCol(tbl, "Lead_Time_Days")).value = NzNumber(wsUI.Range(CELL_LEAD_TIME).value)
        .Cells(1, GetCol(tbl, "Active_Status")).value = wsUI.Range(CELL_ACTIVE).value
        .Cells(1, GetCol(tbl, "Notes")).value = wsUI.Range(CELL_NOTES).value
        .Cells(1, GetCol(tbl, "Created_At")).value = Date
        .Cells(1, GetCol(tbl, "Updated_At")).value = Date
    End With
    
    MsgBox "Product saved successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' LOAD PRODUCT BY SKU
'==================================================
Public Sub Product_Load()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tbl As ListObject
    Dim sku As String
    Dim rowNum As Long
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Products_DB")
    Set tbl = wsDB.ListObjects(TABLE_PRODUCTS)
    
    sku = Trim(CStr(wsUI.Range(CELL_LOOKUP).value))
    
    If sku = "" Then
        MsgBox "Please enter Product Lookup (SKU) first.", vbExclamation
        Exit Sub
    End If
    
    rowNum = FindProductRowBySKU(tbl, sku)
    If rowNum = 0 Then
        MsgBox "Product SKU not found.", vbExclamation
        Exit Sub
    End If
    
    wsUI.Unprotect Password:=SHEET_PWD
    ClearProductForm wsUI
    wsUI.Range(CELL_LOOKUP).value = sku
    
    wsUI.Range(CELL_ID).value = tbl.ListColumns("Product_ID").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_SKU).value = tbl.ListColumns("SKU").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_NAME).value = tbl.ListColumns("Product_Name").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_VARIANT).value = tbl.ListColumns("Variant_Desc").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_CATEGORY).value = tbl.ListColumns("Category").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_UNIT_COST).value = tbl.ListColumns("Unit_Cost").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_SELLING_PRICE).value = tbl.ListColumns("Selling_Price").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_OPENING_STOCK).value = tbl.ListColumns("Opening_Stock").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_REORDER_LEVEL).value = tbl.ListColumns("Reorder_Level").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_SAFETY_DAYS).value = tbl.ListColumns("Safety_Days_Override").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_LEAD_TIME).value = tbl.ListColumns("Lead_Time_Days").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_ACTIVE).value = tbl.ListColumns("Active_Status").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_NOTES).value = tbl.ListColumns("Notes").DataBodyRange.Cells(rowNum, 1).value
    
    MsgBox "Product loaded successfully.", vbInformation

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume SafeExit

End Sub

'==================================================
' UPDATE PRODUCT
'==================================================
Public Sub Product_Update()

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tbl As ListObject
    Dim rowNum As Long
    Dim productID As String
    Dim sku As String
    
    Dim oldOpeningStock As Double
    Dim oldCurrentStock As Double
    Dim newOpeningStock As Double
    Dim stockDelta As Double
    Dim newCurrentStock As Double
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets("Products_DB")
    Set tbl = wsDB.ListObjects(TABLE_PRODUCTS)
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    productID = Trim(CStr(wsUI.Range(CELL_ID).value))
    sku = Trim(CStr(wsUI.Range(CELL_SKU).value))
    
    If productID = "" Then
        MsgBox "Please load an existing product before updating.", vbExclamation
        GoTo SafeExit
    End If
    
    ValidateProductForm wsUI, tbl, True
    
    rowNum = FindProductRowByID(tbl, productID)
    If rowNum = 0 Then
        MsgBox "Product ID not found in Products_DB.", vbCritical
        GoTo SafeExit
    End If
    
    If SKUUsedByAnotherProduct(tbl, sku, productID) Then
        MsgBox "This SKU is already used by another product.", vbCritical
        GoTo SafeExit
    End If
    
    oldOpeningStock = NzNumber(tbl.ListColumns("Opening_Stock").DataBodyRange.Cells(rowNum, 1).value)
    oldCurrentStock = NzNumber(tbl.ListColumns("Current_Stock").DataBodyRange.Cells(rowNum, 1).value)
    newOpeningStock = NzNumber(wsUI.Range(CELL_OPENING_STOCK).value)
    
    stockDelta = newOpeningStock - oldOpeningStock
    newCurrentStock = oldCurrentStock + stockDelta
    
    If newCurrentStock < 0 Then
        MsgBox "Update would make Current Stock negative." & vbCrLf & _
               "Please review Opening Stock or inventory transactions.", vbCritical
        GoTo SafeExit
    End If
    
    With tbl.DataBodyRange.Rows(rowNum)
        .Cells(1, GetCol(tbl, "SKU")).value = wsUI.Range(CELL_SKU).value
        .Cells(1, GetCol(tbl, "Product_Name")).value = wsUI.Range(CELL_NAME).value
        .Cells(1, GetCol(tbl, "Variant_Desc")).value = wsUI.Range(CELL_VARIANT).value
        .Cells(1, GetCol(tbl, "Category")).value = wsUI.Range(CELL_CATEGORY).value
        .Cells(1, GetCol(tbl, "Unit_Cost")).value = NzNumber(wsUI.Range(CELL_UNIT_COST).value)
        .Cells(1, GetCol(tbl, "Selling_Price")).value = NzNumber(wsUI.Range(CELL_SELLING_PRICE).value)
        .Cells(1, GetCol(tbl, "Opening_Stock")).value = newOpeningStock
        .Cells(1, GetCol(tbl, "Current_Stock")).value = newCurrentStock
        .Cells(1, GetCol(tbl, "Reorder_Level")).value = NzNumber(wsUI.Range(CELL_REORDER_LEVEL).value)
        .Cells(1, GetCol(tbl, "Safety_Days_Override")).value = NzNumber(wsUI.Range(CELL_SAFETY_DAYS).value)
        .Cells(1, GetCol(tbl, "Lead_Time_Days")).value = NzNumber(wsUI.Range(CELL_LEAD_TIME).value)
        .Cells(1, GetCol(tbl, "Active_Status")).value = wsUI.Range(CELL_ACTIVE).value
        .Cells(1, GetCol(tbl, "Notes")).value = wsUI.Range(CELL_NOTES).value
        .Cells(1, GetCol(tbl, "Updated_At")).value = Date
    End With
    
    wsUI.Range(CELL_LOOKUP).value = wsUI.Range(CELL_SKU).value
    
    MsgBox "Product updated successfully.", vbInformation

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
Private Sub ClearProductForm(ByVal wsUI As Worksheet)

    wsUI.Range(CELL_LOOKUP).value = ""
    wsUI.Range(CELL_ID).value = ""
    wsUI.Range(CELL_SKU).value = ""
    wsUI.Range(CELL_NAME).value = ""
    wsUI.Range(CELL_VARIANT).value = ""
    wsUI.Range(CELL_CATEGORY).value = ""
    wsUI.Range(CELL_UNIT_COST).value = ""
    wsUI.Range(CELL_SELLING_PRICE).value = ""
    wsUI.Range(CELL_OPENING_STOCK).value = ""
    wsUI.Range(CELL_REORDER_LEVEL).value = ""
    wsUI.Range(CELL_SAFETY_DAYS).value = ""
    wsUI.Range(CELL_LEAD_TIME).value = ""
    wsUI.Range(CELL_ACTIVE).value = "Active"
    wsUI.Range(CELL_NOTES).value = ""

End Sub

Private Sub ValidateProductForm(ByVal wsUI As Worksheet, ByVal tbl As ListObject, ByVal isUpdate As Boolean)

    If Trim(CStr(wsUI.Range(CELL_SKU).value)) = "" Then
        Err.Raise vbObjectError + 1001, , "Please enter SKU."
    End If
    
    If Trim(CStr(wsUI.Range(CELL_NAME).value)) = "" Then
        Err.Raise vbObjectError + 1002, , "Please enter Product Name."
    End If
    
    If Trim(CStr(wsUI.Range(CELL_ACTIVE).value)) = "" Then
        wsUI.Range(CELL_ACTIVE).value = "Active"
    End If
    
    If Not IsBlankOrNumeric(wsUI.Range(CELL_UNIT_COST).value) Then
        Err.Raise vbObjectError + 1003, , "Unit Cost must be numeric."
    End If
    
    If Not IsBlankOrNumeric(wsUI.Range(CELL_SELLING_PRICE).value) Then
        Err.Raise vbObjectError + 1004, , "Selling Price must be numeric."
    End If
    
    If Not IsBlankOrNumeric(wsUI.Range(CELL_OPENING_STOCK).value) Then
        Err.Raise vbObjectError + 1005, , "Opening Stock must be numeric."
    End If
    
    If Not IsBlankOrNumeric(wsUI.Range(CELL_REORDER_LEVEL).value) Then
        Err.Raise vbObjectError + 1006, , "Reorder Level must be numeric."
    End If
    
    If Not IsBlankOrNumeric(wsUI.Range(CELL_SAFETY_DAYS).value) Then
        Err.Raise vbObjectError + 1007, , "Safety Days Override must be numeric."
    End If
    
    If Not IsBlankOrNumeric(wsUI.Range(CELL_LEAD_TIME).value) Then
        Err.Raise vbObjectError + 1008, , "Lead Time Days must be numeric."
    End If

End Sub

Private Function GenerateProductID(ByVal tbl As ListObject) As String

    Dim i As Long, s As String, n As Long, maxNum As Long
    
    maxNum = 0
    
    If tbl.ListRows.Count = 0 Then
        GenerateProductID = "P00001"
        Exit Function
    End If
    
    For i = 1 To tbl.ListRows.Count
        s = Trim(CStr(tbl.ListColumns("Product_ID").DataBodyRange.Cells(i, 1).value))
        If Left$(s, 1) = "P" Then
            If IsNumeric(Mid$(s, 2)) Then
                n = CLng(Mid$(s, 2))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i
    
    GenerateProductID = "P" & Format(maxNum + 1, "00000")

End Function

Private Function FindProductRowBySKU(ByVal tbl As ListObject, ByVal sku As String) As Long

    Dim i As Long
    FindProductRowBySKU = 0
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("SKU").DataBodyRange.Cells(i, 1).value)) = sku Then
            FindProductRowBySKU = i
            Exit Function
        End If
    Next i

End Function

Private Function FindProductRowByID(ByVal tbl As ListObject, ByVal productID As String) As Long

    Dim i As Long
    FindProductRowByID = 0
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("Product_ID").DataBodyRange.Cells(i, 1).value)) = productID Then
            FindProductRowByID = i
            Exit Function
        End If
    Next i

End Function

Private Function ProductSKUExists(ByVal tbl As ListObject, ByVal sku As String) As Boolean

    ProductSKUExists = (FindProductRowBySKU(tbl, sku) > 0)

End Function

Private Function SKUUsedByAnotherProduct(ByVal tbl As ListObject, ByVal sku As String, ByVal currentProductID As String) As Boolean

    Dim i As Long
    SKUUsedByAnotherProduct = False
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("SKU").DataBodyRange.Cells(i, 1).value)) = sku Then
            If Trim(CStr(tbl.ListColumns("Product_ID").DataBodyRange.Cells(i, 1).value)) <> currentProductID Then
                SKUUsedByAnotherProduct = True
                Exit Function
            End If
        End If
    Next i

End Function

Private Function GetCol(ByVal tbl As ListObject, ByVal colName As String) As Long
    GetCol = tbl.ListColumns(colName).Index
End Function

Private Function NzNumber(ByVal v As Variant) As Double
    If IsError(v) Then
        NzNumber = 0
    ElseIf IsNumeric(v) Then
        NzNumber = CDbl(v)
    ElseIf Trim(CStr(v)) = "" Then
        NzNumber = 0
    Else
        NzNumber = 0
    End If
End Function

Private Function IsBlankOrNumeric(ByVal v As Variant) As Boolean
    If Trim(CStr(v)) = "" Then
        IsBlankOrNumeric = True
    ElseIf IsNumeric(v) Then
        IsBlankOrNumeric = True
    Else
        IsBlankOrNumeric = False
    End If
End Function

