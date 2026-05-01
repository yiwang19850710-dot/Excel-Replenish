Attribute VB_Name = "modAdjustment"
Option Explicit

Private Const SHEET_UI As String = "Adjustment_UI"
Private Const TABLE_ADJ As String = "tblAdjustment"
Private Const TABLE_PRODUCTS As String = "tblProducts"
Private Const TABLE_LOG As String = "tblInventoryLog"
Private Const SHEET_PWD As String = ""

' UI Mapping
Private Const CELL_ID As String = "B3"
Private Const CELL_DATE As String = "B4"
Private Const CELL_SKU As String = "B5"
Private Const CELL_NAME As String = "B6"
Private Const CELL_CURRENT As String = "B7"
Private Const CELL_QTY As String = "B8"
Private Const CELL_NEW As String = "B9"
Private Const CELL_REASON As String = "B10"
Private Const CELL_NOTES As String = "B11"

'==================================================
' NEW ADJUSTMENT
'==================================================
Public Sub Adjustment_New()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    
    wsUI.Unprotect Password:=SHEET_PWD
    ClearAdjustmentForm wsUI
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True

End Sub

'==================================================
' SAVE NEW ADJUSTMENT
'==================================================
Public Sub Adjustment_Save()

    Dim wsUI As Worksheet
    Dim wsAdj As Worksheet
    Dim tblA As ListObject
    Dim adjID As String
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    
    On Error Resume Next
    Set wsAdj = ThisWorkbook.Worksheets("Adjustment_DB")
    On Error GoTo ErrHandler
    
    If wsAdj Is Nothing Then
        MsgBox "Sheet 'Adjustment_DB' was not found.", vbCritical, "Adjustment Save"
        Exit Sub
    End If
    
    On Error Resume Next
    Set tblA = wsAdj.ListObjects(TABLE_ADJ)
    On Error GoTo ErrHandler
    
    If tblA Is Nothing Then
        MsgBox "Table '" & TABLE_ADJ & "' was not found in Adjustment_DB." & vbCrLf & _
               "Please check the table name in Table Design.", vbCritical, "Adjustment Save"
        Exit Sub
    End If
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    If Trim(CStr(wsUI.Range(CELL_ID).value)) <> "" Then
        MsgBox "Adjustment ID should be blank when creating a new adjustment.", vbExclamation, "Create Adjustment"
        GoTo SafeExit
    End If
    
    ValidateAdjustmentReadyToSave
    
    adjID = GenerateAdjustmentID(tblA)
    wsUI.Range(CELL_ID).value = adjID
    
    SaveAdjustmentUsingID adjID
    
    MsgBox "Inventory adjustment saved successfully.", vbInformation, "Adjustment Saved"

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Adjustment Save"
    Resume SafeExit

End Sub

'==================================================
' LOAD ADJUSTMENT
'==================================================
Public Sub Adjustment_Load()

    Dim wsUI As Worksheet
    Dim wsAdj As Worksheet
    Dim tblA As ListObject
    Dim adjID As String
    Dim rowNum As Long
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsAdj = ThisWorkbook.Worksheets("Adjustment_DB")
    Set tblA = wsAdj.ListObjects(TABLE_ADJ)
    
    adjID = Trim(CStr(wsUI.Range(CELL_ID).value))
    
    If adjID = "" Then
        MsgBox "Please enter Adjustment ID first.", vbExclamation, "Load Adjustment"
        Exit Sub
    End If
    
    rowNum = FindAdjustmentRowByID(tblA, adjID)
    If rowNum = 0 Then
        MsgBox "Adjustment ID not found.", vbExclamation, "Load Adjustment"
        Exit Sub
    End If
    
    wsUI.Unprotect Password:=SHEET_PWD
    ClearAdjustmentForm wsUI
    wsUI.Range(CELL_ID).value = adjID
    
    wsUI.Range(CELL_DATE).value = tblA.ListColumns("Adjustment_Date").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_SKU).value = tblA.ListColumns("SKU").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_NAME).value = tblA.ListColumns("Product_Name").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_CURRENT).value = tblA.ListColumns("Current_Stock_Before").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_QTY).value = tblA.ListColumns("Adjustment_Qty").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_NEW).value = tblA.ListColumns("Current_Stock_After").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_REASON).value = tblA.ListColumns("Reason").DataBodyRange.Cells(rowNum, 1).value
    wsUI.Range(CELL_NOTES).value = tblA.ListColumns("Notes").DataBodyRange.Cells(rowNum, 1).value
    
    MsgBox "Adjustment loaded successfully.", vbInformation, "Load Adjustment"

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Adjustment Load"
    Resume SafeExit

End Sub

'==================================================
' UPDATE ADJUSTMENT
'==================================================
Public Sub Adjustment_Update()

    Dim wsUI As Worksheet
    Dim wsAdj As Worksheet
    Dim wsProducts As Worksheet
    Dim wsLog As Worksheet
    
    Dim tblA As ListObject
    Dim tblP As ListObject
    Dim tblL As ListObject
    
    Dim adjID As String
    Dim rowNum As Long
    Dim productRow As Long
    Dim oldSKU As String
    Dim oldQty As Double
    Dim currentStock As Double
    Dim rolledBackStock As Double
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsAdj = ThisWorkbook.Worksheets("Adjustment_DB")
    Set wsProducts = ThisWorkbook.Worksheets("Products_DB")
    Set wsLog = ThisWorkbook.Worksheets("Inventory_Log")
    
    Set tblA = wsAdj.ListObjects(TABLE_ADJ)
    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)
    Set tblL = wsLog.ListObjects(TABLE_LOG)
    
    wsUI.Unprotect Password:=SHEET_PWD
    
    adjID = Trim(CStr(wsUI.Range(CELL_ID).value))
    
    If adjID = "" Then
        MsgBox "Please load an existing Adjustment ID before updating.", vbExclamation, "Update Adjustment"
        GoTo SafeExit
    End If
    
    rowNum = FindAdjustmentRowByID(tblA, adjID)
    If rowNum = 0 Then
        MsgBox "Adjustment ID not found in Adjustment_DB.", vbExclamation, "Update Adjustment"
        GoTo SafeExit
    End If
    
    ValidateAdjustmentReadyToSave
    
    If MsgBox("Update this adjustment record?" & vbCrLf & _
              "Adjustment ID: " & adjID, vbQuestion + vbYesNo, "Confirm Update") <> vbYes Then
        GoTo SafeExit
    End If
    
    ' rollback old adjustment
    oldSKU = Trim(CStr(tblA.ListColumns("SKU").DataBodyRange.Cells(rowNum, 1).value))
    oldQty = NzNumber(tblA.ListColumns("Adjustment_Qty").DataBodyRange.Cells(rowNum, 1).value)
    
    productRow = FindProductRowBySKU(tblP, oldSKU)
    If productRow > 0 Then
        currentStock = NzNumber(tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value)
        rolledBackStock = currentStock - oldQty
        
        If rolledBackStock < 0 Then
            MsgBox "This adjustment cannot be updated because rolling back the previous adjustment would make Current Stock negative.", vbExclamation, "Update Adjustment Not Allowed"
            GoTo SafeExit
        End If
        
        tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value = rolledBackStock
    End If
    
    ' delete old inventory log
    For i = tblL.ListRows.Count To 1 Step -1
        If Trim(CStr(tblL.ListColumns("Tran_Type").DataBodyRange.Cells(i, 1).value)) = "ADJUSTMENT" _
           And Trim(CStr(tblL.ListColumns("Ref_No").DataBodyRange.Cells(i, 1).value)) = adjID Then
            tblL.ListRows(i).Delete
        End If
    Next i
    
    ' delete old adjustment row
    tblA.ListRows(rowNum).Delete
    
    ' save current UI using same ID
    SaveAdjustmentUsingID adjID
    
    MsgBox "Adjustment updated successfully.", vbInformation, "Update Adjustment"

SafeExit:
    wsUI.Protect Password:=SHEET_PWD, UserInterfaceOnly:=True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Adjustment Update"
    Resume SafeExit

End Sub

'==================================================
' SAVE LOGIC
'==================================================
Private Sub SaveAdjustmentUsingID(ByVal adjID As String)

    Dim wsUI As Worksheet
    Dim wsAdj As Worksheet
    Dim wsProducts As Worksheet
    Dim wsLog As Worksheet
    
    Dim tblA As ListObject
    Dim tblP As ListObject
    Dim tblL As ListObject
    
    Dim adjDate As Variant
    Dim sku As String
    Dim productName As String
    Dim reason As String
    Dim notes As String
    
    Dim adjQty As Double
    Dim currentBefore As Double
    Dim currentAfter As Double
    
    Dim productRow As Long
    Dim productID As String
    Dim newRow As ListRow
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsAdj = ThisWorkbook.Worksheets("Adjustment_DB")
    Set wsProducts = ThisWorkbook.Worksheets("Products_DB")
    Set wsLog = ThisWorkbook.Worksheets("Inventory_Log")
    
    Set tblA = wsAdj.ListObjects(TABLE_ADJ)
    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)
    Set tblL = wsLog.ListObjects(TABLE_LOG)
    
    adjDate = wsUI.Range(CELL_DATE).value
    sku = Trim(CStr(wsUI.Range(CELL_SKU).value))
    productName = Trim(CStr(wsUI.Range(CELL_NAME).value))
    adjQty = NzNumber(wsUI.Range(CELL_QTY).value)
    reason = Trim(CStr(wsUI.Range(CELL_REASON).value))
    notes = Trim(CStr(wsUI.Range(CELL_NOTES).value))
    
    productRow = FindProductRowBySKU(tblP, sku)
    If productRow = 0 Then
        Err.Raise vbObjectError + 1501, , "Product SKU not found in Products_DB."
    End If
    
    productID = Trim(CStr(tblP.ListColumns("Product_ID").DataBodyRange.Cells(productRow, 1).value))
    currentBefore = NzNumber(tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value)
    currentAfter = currentBefore + adjQty
    
    If currentAfter < 0 Then
        Err.Raise vbObjectError + 1502, , "Adjustment would make Current Stock negative."
    End If
    
    If productName = "" Then
        productName = Trim(CStr(tblP.ListColumns("Product_Name").DataBodyRange.Cells(productRow, 1).value))
        wsUI.Range(CELL_NAME).value = productName
    End If
    
    wsUI.Range(CELL_CURRENT).value = currentBefore
    wsUI.Range(CELL_NEW).value = currentAfter
    
    Set newRow = tblA.ListRows.Add
    
    With newRow.Range
        .Cells(1, GetCol(tblA, "Adjustment_ID")).value = adjID
        .Cells(1, GetCol(tblA, "Adjustment_Date")).value = adjDate
        .Cells(1, GetCol(tblA, "Product_ID")).value = productID
        .Cells(1, GetCol(tblA, "SKU")).value = sku
        .Cells(1, GetCol(tblA, "Product_Name")).value = productName
        .Cells(1, GetCol(tblA, "Current_Stock_Before")).value = currentBefore
        .Cells(1, GetCol(tblA, "Adjustment_Qty")).value = adjQty
        .Cells(1, GetCol(tblA, "Current_Stock_After")).value = currentAfter
        .Cells(1, GetCol(tblA, "Reason")).value = reason
        .Cells(1, GetCol(tblA, "Notes")).value = notes
        .Cells(1, GetCol(tblA, "Created_At")).value = Date
        .Cells(1, GetCol(tblA, "Updated_At")).value = Date
    End With
    
    tblP.ListColumns("Current_Stock").DataBodyRange.Cells(productRow, 1).value = currentAfter
    
    With tblL.ListRows.Add
        .Range(1, 1).value = GenerateLogID(tblL)
        .Range(1, 2).value = adjDate
        .Range(1, 3).value = "ADJUSTMENT"
        .Range(1, 4).value = adjID
        .Range(1, 5).value = productID
        .Range(1, 6).value = sku
        .Range(1, 7).value = productName
        .Range(1, 8).value = adjQty
        .Range(1, 9).value = currentAfter
        .Range(1, 10).value = 0
        .Range(1, 11).value = 0
        .Range(1, 12).value = "SYSTEM"
        .Range(1, 13).value = ""
        .Range(1, 14).value = ""
        .Range(1, 15).value = reason & IIf(notes <> "", " | " & notes, "")
        .Range(1, 16).value = Date
    End With

End Sub

'==================================================
' VALIDATION
'==================================================
Private Sub ValidateAdjustmentReadyToSave()

    Dim wsUI As Worksheet
    Dim wsProducts As Worksheet
    Dim tblP As ListObject
    Dim sku As String
    
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsProducts = ThisWorkbook.Worksheets("Products_DB")
    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)
    
    If Trim(CStr(wsUI.Range(CELL_DATE).value)) = "" Then
        Err.Raise vbObjectError + 1511, , "Please enter Adjustment Date."
    End If
    
    If Not IsDate(wsUI.Range(CELL_DATE).value) Then
        Err.Raise vbObjectError + 1512, , "Adjustment Date is invalid."
    End If
    
    sku = Trim(CStr(wsUI.Range(CELL_SKU).value))
    If sku = "" Then
        Err.Raise vbObjectError + 1513, , "Please enter Product SKU."
    End If
    
    If FindProductRowBySKU(tblP, sku) = 0 Then
        Err.Raise vbObjectError + 1514, , "Product SKU not found in Products_DB."
    End If
    
    If Not IsNumeric(wsUI.Range(CELL_QTY).value) Then
        Err.Raise vbObjectError + 1515, , "Adjustment Qty must be numeric."
    End If
    
    If NzNumber(wsUI.Range(CELL_QTY).value) = 0 Then
        Err.Raise vbObjectError + 1516, , "Adjustment Qty cannot be 0."
    End If
    
    If Trim(CStr(wsUI.Range(CELL_REASON).value)) = "" Then
        Err.Raise vbObjectError + 1517, , "Please enter Reason."
    End If

End Sub

'==================================================
' HELPERS
'==================================================
Private Sub RestoreAdjustmentFormulas(ByVal wsUI As Worksheet)

    wsUI.Range(CELL_NAME).Formula = "=IFERROR(XLOOKUP(" & CELL_SKU & ",tblProducts[SKU],tblProducts[Product_Name]),"""")"
    wsUI.Range(CELL_CURRENT).Formula = "=IFERROR(XLOOKUP(" & CELL_SKU & ",tblProducts[SKU],tblProducts[Current_Stock]),"""")"
    wsUI.Range(CELL_NEW).Formula = "=IFERROR(" & CELL_CURRENT & "+" & CELL_QTY & ","""")"

End Sub

Private Sub ClearAdjustmentForm(ByVal wsUI As Worksheet)

    wsUI.Range(CELL_ID).value = ""
    wsUI.Range(CELL_DATE).value = ""
    wsUI.Range(CELL_SKU).value = ""
    wsUI.Range(CELL_NAME).Formula = ""
    wsUI.Range(CELL_CURRENT).Formula = ""
    wsUI.Range(CELL_QTY).value = ""
    wsUI.Range(CELL_NEW).Formula = ""
    wsUI.Range(CELL_REASON).value = ""
    wsUI.Range(CELL_NOTES).value = ""
    
    RestoreAdjustmentFormulas wsUI

End Sub

Private Function GenerateAdjustmentID(ByVal tbl As ListObject) As String

    Dim i As Long, s As String, n As Long, maxNum As Long
    
    maxNum = 0
    
    If tbl.ListRows.Count = 0 Then
        GenerateAdjustmentID = "A00001"
        Exit Function
    End If
    
    For i = 1 To tbl.ListRows.Count
        s = Trim(CStr(tbl.ListColumns("Adjustment_ID").DataBodyRange.Cells(i, 1).value))
        If Left$(s, 1) = "A" Then
            If IsNumeric(Mid$(s, 2)) Then
                n = CLng(Mid$(s, 2))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i
    
    GenerateAdjustmentID = "A" & Format(maxNum + 1, "00000")

End Function

Private Function FindAdjustmentRowByID(ByVal tbl As ListObject, ByVal adjID As String) As Long

    Dim i As Long
    FindAdjustmentRowByID = 0
    
    For i = 1 To tbl.ListRows.Count
        If Trim(CStr(tbl.ListColumns("Adjustment_ID").DataBodyRange.Cells(i, 1).value)) = adjID Then
            FindAdjustmentRowByID = i
            Exit Function
        End If
    Next i

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

Private Function GenerateLogID(ByVal tbl As ListObject) As String

    Dim maxNum As Long
    Dim i As Long
    Dim s As String
    Dim n As Long
    
    maxNum = 0
    
    If tbl.ListRows.Count = 0 Then
        GenerateLogID = "L00001"
        Exit Function
    End If
    
    For i = 1 To tbl.ListRows.Count
        s = Trim(CStr(tbl.ListColumns("Log_ID").DataBodyRange.Cells(i, 1).value))
        If Left$(s, 1) = "L" Then
            If IsNumeric(Mid$(s, 2)) Then
                n = CLng(Mid$(s, 2))
                If n > maxNum Then maxNum = n
            End If
        End If
    Next i
    
    GenerateLogID = "L" & Format(maxNum + 1, "00000")

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

