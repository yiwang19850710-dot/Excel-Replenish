Attribute VB_Name = "modPOInquiry"
Option Explicit

Private Const SHEET_UI As String = "PO_Inquiry_UI"
Private Const SHEET_PURCHASE As String = "Purchase_DB"
Private Const TABLE_PURCHASE As String = "tblPurchase"
Private Const SHEET_PURCHASE_UI As String = "Purchase_UI"

Private Const CELL_SKU As String = "B3"
Private Const CELL_PRODUCT_NAME As String = "B4"
Private Const CELL_SUPPLIER_NAME As String = "B5"
Private Const CELL_PO_NO As String = "B6"
Private Const CELL_STATUS As String = "B7"
Private Const CELL_LINE_STATUS As String = "B8"

Private Const HEADER_ROW As Long = 11
Private Const FIRST_DATA_ROW As Long = 12
Private Const LAST_OUT_COL As Long = 18

Public Sub POInquiry_Setup()

    Dim ws As Worksheet
    Dim wsSup As Worksheet
    Dim tblSup As ListObject
    Dim i As Long
    Dim colSupplierName As Long
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_UI)

    PreparePOInquiryHeader
    ClearPOInquiryOutput
    ClearSummary

    'Get or create hidden supplier list sheet
    On Error Resume Next
    Set wsSup = ThisWorkbook.Worksheets("PO_Inquiry_Supplier_List")
    On Error GoTo 0

If wsSup Is Nothing Then
    Set wsSup = ThisWorkbook.Worksheets.Add(After:=ws)
    
    On Error Resume Next
    wsSup.name = "PO_Inquiry_Supplier_List"
    If Err.Number <> 0 Then
        wsSup.name = "PO_Inquiry_Supplier_List_" & Format(Now, "hhmmss")
        Err.Clear
    End If
    On Error GoTo 0
End If

    wsSup.Visible = xlSheetVisible
    wsSup.Cells.Clear
    wsSup.Range("A1").value = "Supplier_Name"

    Set tblSup = ThisWorkbook.Worksheets("Suppliers_DB").ListObjects("tblSuppliers")
    colSupplierName = tblSup.ListColumns("Supplier_Name").Index

    If Not tblSup.DataBodyRange Is Nothing Then
        For i = 1 To tblSup.ListRows.Count
            wsSup.Cells(i + 1, 1).value = tblSup.DataBodyRange.Cells(i, colSupplierName).value
        Next i
    End If

    lastRow = wsSup.Cells(wsSup.Rows.Count, 1).End(xlUp).Row

    If lastRow >= 2 Then
        On Error Resume Next
        ThisWorkbook.Names("POInquirySupplierNames").Delete
        On Error GoTo 0

        ThisWorkbook.Names.Add _
            name:="POInquirySupplierNames", _
            RefersTo:="=PO_Inquiry_Supplier_List!$A$2:$A$" & lastRow

        With ws.Range(CELL_SUPPLIER_NAME).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Operator:=xlBetween, Formula1:="=POInquirySupplierNames"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowError = False
        End With
    End If

    wsSup.Visible = xlSheetVeryHidden

    MsgBox "PO Inquiry UI setup completed.", vbInformation, "PO Inquiry"

End Sub

Public Sub POInquiry_Search()
    RenderPOInquiry False
End Sub

Public Sub POInquiry_ShowOpenPO()

    With ThisWorkbook.Worksheets(SHEET_UI)
        .Range(CELL_STATUS).value = "Open"
        .Range(CELL_LINE_STATUS).value = ""
    End With

    RenderPOInquiry False

End Sub

Public Sub POInquiry_ShowPartialPO()

    With ThisWorkbook.Worksheets(SHEET_UI)
        .Range(CELL_STATUS).value = "Partial"
        .Range(CELL_LINE_STATUS).value = ""
    End With

    RenderPOInquiry False

End Sub

Public Sub POInquiry_ShowAll()

    With ThisWorkbook.Worksheets(SHEET_UI)
        .Range(CELL_SKU).value = ""
        .Range(CELL_PRODUCT_NAME).value = ""
        .Range(CELL_SUPPLIER_NAME).value = ""
        .Range(CELL_PO_NO).value = ""
        .Range(CELL_STATUS).value = ""
        .Range(CELL_LINE_STATUS).value = ""
    End With

    RenderPOInquiry True

End Sub

Public Sub POInquiry_Clear()

    With ThisWorkbook.Worksheets(SHEET_UI)
        .Range(CELL_SKU).value = ""
        .Range(CELL_PRODUCT_NAME).value = ""
        .Range(CELL_SUPPLIER_NAME).value = ""
        .Range(CELL_PO_NO).value = ""
        .Range(CELL_STATUS).value = ""
        .Range(CELL_LINE_STATUS).value = ""
    End With

    PreparePOInquiryHeader
    ClearPOInquiryOutput
    ClearSummary

End Sub

Private Sub RenderPOInquiry(Optional ByVal showProgress As Boolean = False)

    Dim wsUI As Worksheet
    Dim wsDB As Worksheet
    Dim tbl As ListObject
    Dim i As Long, outRow As Long, totalRows As Long

    Dim filterSKU As String
    Dim filterProduct As String
    Dim filterSupplier As String
    Dim filterPO As String
    Dim filterStatus As String
    Dim filterLineStatus As String

    Dim sku As String, productName As String, supplierName As String
    Dim poNo As String, statusText As String, lineStatus As String

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsDB = ThisWorkbook.Worksheets(SHEET_PURCHASE)
    Set tbl = wsDB.ListObjects(TABLE_PURCHASE)

    filterSKU = Trim$(CStr(wsUI.Range(CELL_SKU).value))
    filterProduct = Trim$(CStr(wsUI.Range(CELL_PRODUCT_NAME).value))
    filterSupplier = Trim$(CStr(wsUI.Range(CELL_SUPPLIER_NAME).value))
    filterPO = Trim$(CStr(wsUI.Range(CELL_PO_NO).value))
    filterStatus = Trim$(CStr(wsUI.Range(CELL_STATUS).value))
    filterLineStatus = Trim$(CStr(wsUI.Range(CELL_LINE_STATUS).value))

    Application.ScreenUpdating = False

    PreparePOInquiryHeader
    ClearPOInquiryOutput
    ClearSummary

    If tbl.DataBodyRange Is Nothing Then
        wsUI.Cells(FIRST_DATA_ROW, 1).value = "No purchase records found."
        GoTo CleanExit
    End If

    If showProgress Then ProgressStart "Loading PO Inquiry", "Preparing purchase records..."

    totalRows = tbl.ListRows.Count
    outRow = FIRST_DATA_ROW

    For i = 1 To totalRows

        If showProgress Then
            ProgressUpdate "Loading PO records", i, totalRows, "Row " & i
        End If

        poNo = GetText(tbl, i, "Purchase_Order_No")
        supplierName = GetText(tbl, i, "Supplier_Name")
        sku = GetText(tbl, i, "SKU")
        productName = GetText(tbl, i, "Product_Name")
        statusText = GetText(tbl, i, "Status")
        lineStatus = GetText(tbl, i, "Line_Status")

        If MatchesFilter(sku, filterSKU) _
           And MatchesFilter(productName, filterProduct) _
           And MatchesFilter(supplierName, filterSupplier) _
           And MatchesFilter(poNo, filterPO) _
           And MatchesFilter(statusText, filterStatus) _
           And MatchesFilter(lineStatus, filterLineStatus) Then

            WritePOInquiryRow wsUI, tbl, i, outRow
            outRow = outRow + 1
        End If

    Next i

    If outRow = FIRST_DATA_ROW Then
        wsUI.Cells(FIRST_DATA_ROW, 1).value = "No matching PO records found."
    Else
        FormatPOInquiryOutput wsUI, outRow - 1
        SortPOInquiryByDate wsUI, outRow - 1
        UpdateSummary wsUI, outRow - 1
    End If

CleanExit:
    If showProgress Then ProgressEnd
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    If showProgress Then ProgressEnd
    Application.ScreenUpdating = True
    MsgBox "PO Inquiry error: " & Err.Description, vbCritical, "PO Inquiry"

End Sub

Private Sub PreparePOInquiryHeader()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_UI)

    ws.Range("A9").value = "Total Rows"
    ws.Range("C9").value = "Open PO Count"
    ws.Range("E9").value = "Remaining Qty"
    ws.Range("G9").value = "Open Value"

    With ws.Range("A9:H9")
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
        .Borders.LineStyle = xlContinuous
    End With

    ws.Cells(HEADER_ROW, 1).value = "Purchase_Order_No"
    ws.Cells(HEADER_ROW, 2).value = "External_PO_No"
    ws.Cells(HEADER_ROW, 3).value = "Purchase_Date"
    ws.Cells(HEADER_ROW, 4).value = "Supplier_Name"
    ws.Cells(HEADER_ROW, 5).value = "SKU"
    ws.Cells(HEADER_ROW, 6).value = "Product_Name"
    ws.Cells(HEADER_ROW, 7).value = "Qty"
    ws.Cells(HEADER_ROW, 8).value = "Received_Qty"
    ws.Cells(HEADER_ROW, 9).value = "Remaining_Qty"
    ws.Cells(HEADER_ROW, 10).value = "Unit_Cost"
    ws.Cells(HEADER_ROW, 11).value = "Line_Total"
    ws.Cells(HEADER_ROW, 12).value = "Amount_Paid"
    ws.Cells(HEADER_ROW, 13).value = "Balance_Due"
    ws.Cells(HEADER_ROW, 14).value = "Payment_Status"
    ws.Cells(HEADER_ROW, 15).value = "Status"
    ws.Cells(HEADER_ROW, 16).value = "Line_Status"
    ws.Cells(HEADER_ROW, 17).value = "Source"
    ws.Cells(HEADER_ROW, 18).value = "Updated_At"

    With ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, LAST_OUT_COL))
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
    End With

End Sub

Private Sub ClearSummary()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_UI)

    ws.Range("B9").value = ""
    ws.Range("D9").value = ""
    ws.Range("F9").value = ""
    ws.Range("H9").value = ""

End Sub

Private Sub ClearPOInquiryOutput()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_UI)

    ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(ws.Rows.Count, LAST_OUT_COL)).ClearContents
    ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(ws.Rows.Count, LAST_OUT_COL)).Interior.Pattern = xlNone
    ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(ws.Rows.Count, LAST_OUT_COL)).Borders.LineStyle = xlNone

End Sub

Private Sub WritePOInquiryRow(ByVal ws As Worksheet, ByVal tbl As ListObject, ByVal srcRow As Long, ByVal outRow As Long)

    ws.Cells(outRow, 1).value = GetText(tbl, srcRow, "Purchase_Order_No")
    ws.Cells(outRow, 2).value = GetOptionalText(tbl, srcRow, "External_PO_No")
    ws.Cells(outRow, 3).value = GetValue(tbl, srcRow, "Purchase_Date")
    ws.Cells(outRow, 4).value = GetText(tbl, srcRow, "Supplier_Name")
    ws.Cells(outRow, 5).value = GetText(tbl, srcRow, "SKU")
    ws.Cells(outRow, 6).value = GetText(tbl, srcRow, "Product_Name")
    ws.Cells(outRow, 7).value = GetValue(tbl, srcRow, "Qty")
    ws.Cells(outRow, 8).value = GetValue(tbl, srcRow, "Received_Qty")
    ws.Cells(outRow, 9).value = GetValue(tbl, srcRow, "Remaining_Qty")
    ws.Cells(outRow, 10).value = GetValue(tbl, srcRow, "Unit_Cost")
    ws.Cells(outRow, 11).value = GetValue(tbl, srcRow, "Line_Total")
    ws.Cells(outRow, 12).value = GetValue(tbl, srcRow, "Amount_Paid")
    ws.Cells(outRow, 13).value = GetValue(tbl, srcRow, "Balance_Due")
    ws.Cells(outRow, 14).value = GetText(tbl, srcRow, "Payment_Status")
    ws.Cells(outRow, 15).value = GetText(tbl, srcRow, "Status")
    ws.Cells(outRow, 16).value = GetText(tbl, srcRow, "Line_Status")
    ws.Cells(outRow, 17).value = GetOptionalText(tbl, srcRow, "Source")
    ws.Cells(outRow, 18).value = GetValue(tbl, srcRow, "Updated_At")

End Sub

Private Sub FormatPOInquiryOutput(ByVal ws As Worksheet, ByVal lastRow As Long)

    With ws.Range(ws.Cells(FIRST_DATA_ROW, 1), ws.Cells(lastRow, LAST_OUT_COL))
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With

    ws.Range(ws.Cells(FIRST_DATA_ROW, 3), ws.Cells(lastRow, 3)).NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(FIRST_DATA_ROW, 7), ws.Cells(lastRow, 13)).NumberFormat = "#,##0.00"
    ws.Range(ws.Cells(FIRST_DATA_ROW, 18), ws.Cells(lastRow, 18)).NumberFormat = "yyyy-mm-dd"

End Sub

Private Sub SortPOInquiryByDate(ByVal ws As Worksheet, ByVal lastRow As Long)

    If lastRow <= FIRST_DATA_ROW Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(FIRST_DATA_ROW, 3), ws.Cells(lastRow, 3)), _
            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal

        .SetRange ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, LAST_OUT_COL))
        .Header = xlYes
        .Apply
    End With

End Sub

Private Sub UpdateSummary(ByVal ws As Worksheet, ByVal lastRow As Long)

    Dim r As Long
    Dim totalRows As Long
    Dim totalRemainingQty As Double
    Dim totalOpenValue As Double
    Dim dictOpenPO As Object
    Dim poNo As String
    Dim statusText As String
    Dim remainingQty As Double
    Dim unitCost As Double

    Set dictOpenPO = CreateObject("Scripting.Dictionary")

    totalRows = lastRow - FIRST_DATA_ROW + 1

    For r = FIRST_DATA_ROW To lastRow

        statusText = Trim$(CStr(ws.Cells(r, 15).value))
        poNo = Trim$(CStr(ws.Cells(r, 1).value))
        remainingQty = NzDbl(ws.Cells(r, 9).value)
        unitCost = NzDbl(ws.Cells(r, 10).value)

        totalRemainingQty = totalRemainingQty + remainingQty
        totalOpenValue = totalOpenValue + (remainingQty * unitCost)

        If UCase$(statusText) = "OPEN" Or UCase$(statusText) = "PARTIAL" Then
            If poNo <> "" Then
                If Not dictOpenPO.Exists(poNo) Then dictOpenPO.Add poNo, True
            End If
        End If

    Next r

    ws.Range("B9").value = totalRows
    ws.Range("D9").value = dictOpenPO.Count
    ws.Range("F9").value = totalRemainingQty
    ws.Range("H9").value = totalOpenValue

    ws.Range("F9").NumberFormat = "#,##0.00"
    ws.Range("H9").NumberFormat = "#,##0.00"

End Sub

Public Sub POInquiry_LoadSelectedPO()

    Dim ws As Worksheet
    Dim selectedRow As Long
    Dim poNo As String

    Set ws = ThisWorkbook.Worksheets(SHEET_UI)
    selectedRow = ActiveCell.Row

    If selectedRow < FIRST_DATA_ROW Then
        MsgBox "Please select a PO row from the result area first.", vbExclamation, "PO Inquiry"
        Exit Sub
    End If

    poNo = Trim$(CStr(ws.Cells(selectedRow, 1).value))

    If poNo = "" Then
        MsgBox "Selected row does not contain a PO number.", vbExclamation, "PO Inquiry"
        Exit Sub
    End If

    ThisWorkbook.Worksheets(SHEET_PURCHASE_UI).Range("B3").value = poNo
    ThisWorkbook.Worksheets(SHEET_PURCHASE_UI).Activate

    On Error Resume Next
    Application.Run "PurchaseUI_Load"
    If Err.Number <> 0 Then
        Err.Clear
        Application.Run "LoadPurchase"
    End If
    On Error GoTo 0

End Sub

Private Function MatchesFilter(ByVal valueText As String, ByVal filterText As String) As Boolean

    If Trim$(filterText) = "" Then
        MatchesFilter = True
    Else
        MatchesFilter = (InStr(1, valueText, filterText, vbTextCompare) > 0)
    End If

End Function

Private Function GetText(ByVal tbl As ListObject, ByVal rowNo As Long, ByVal colName As String) As String
    GetText = Trim$(CStr(tbl.DataBodyRange.Cells(rowNo, tbl.ListColumns(colName).Index).value))
End Function

Private Function GetValue(ByVal tbl As ListObject, ByVal rowNo As Long, ByVal colName As String) As Variant
    GetValue = tbl.DataBodyRange.Cells(rowNo, tbl.ListColumns(colName).Index).value
End Function

Private Function GetOptionalText(ByVal tbl As ListObject, ByVal rowNo As Long, ByVal colName As String) As String

    Dim lc As ListColumn

    For Each lc In tbl.ListColumns
        If StrComp(lc.name, colName, vbTextCompare) = 0 Then
            GetOptionalText = Trim$(CStr(tbl.DataBodyRange.Cells(rowNo, lc.Index).value))
            Exit Function
        End If
    Next lc

    GetOptionalText = ""

End Function

Private Function NzDbl(ByVal v As Variant) As Double

    If IsError(v) Then
        NzDbl = 0
    ElseIf Trim$(CStr(v)) = "" Then
        NzDbl = 0
    ElseIf IsNumeric(v) Then
        NzDbl = CDbl(v)
    Else
        NzDbl = 0
    End If

End Function

