Attribute VB_Name = "modInventoryUI"
Option Explicit

Private Const SHEET_UI As String = "Inventory_UI"
Private Const SHEET_PRODUCTS As String = "Products_DB"
Private Const SHEET_PURCHASE As String = "Purchase_DB"

Private Const TABLE_PRODUCTS As String = "tblProducts"
Private Const TABLE_PURCHASE As String = "tblPurchase"

Private Const CELL_SEARCH_SKU As String = "B3"
Private Const CELL_SEARCH_NAME As String = "B4"

Private Const HEADER_ROW As Long = 10
Private Const FIRST_DATA_ROW As Long = 11

Private Const COL_OUT_PRODUCT_ID As Long = 1
Private Const COL_OUT_SKU As Long = 2
Private Const COL_OUT_NAME As Long = 3
Private Const COL_OUT_STOCK As Long = 4
Private Const COL_OUT_OPEN_PO_QTY As Long = 5
Private Const COL_OUT_INBOUND_QTY As Long = 6
Private Const COL_OUT_OPEN_PO_COUNT As Long = 7
Private Const COL_OUT_LATEST_PO As Long = 8
Private Const COL_OUT_UNIT_COST As Long = 9
Private Const COL_OUT_INV_VALUE As Long = 10
Private Const COL_OUT_REORDER As Long = 11
Private Const COL_OUT_HEALTH As Long = 12
Private Const COL_OUT_UPDATED As Long = 13

'==================================================
' PUBLIC ACTIONS
'==================================================
Public Sub InventoryUI_Setup()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    PrepareInventoryUIHeader wsUI
    ClearInventoryOutput wsUI

    With wsUI.Range("A1:M40")
        .VerticalAlignment = xlCenter
    End With

End Sub

Public Sub InventoryUI_Search()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    RenderInventoryResults _
        Trim$(CStr(wsUI.Range(CELL_SEARCH_SKU).value)), _
        Trim$(CStr(wsUI.Range(CELL_SEARCH_NAME).value)), _
        False

End Sub

Public Sub InventoryUI_Refresh()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    RenderInventoryResults _
        Trim$(CStr(wsUI.Range(CELL_SEARCH_SKU).value)), _
        Trim$(CStr(wsUI.Range(CELL_SEARCH_NAME).value)), _
        False

End Sub

Public Sub InventoryUI_ShowAll()

    RenderInventoryResults "", "", True

End Sub

Public Sub InventoryUI_ClearSearch()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    wsUI.Range(CELL_SEARCH_SKU).value = ""
    wsUI.Range(CELL_SEARCH_NAME).value = ""

    PrepareInventoryUIHeader wsUI
    ClearInventoryOutput wsUI

End Sub

'==================================================
' CORE RENDER
'==================================================
Private Sub RenderInventoryResults(ByVal skuKeyword As String, _
                                   ByVal nameKeyword As String, _
                                   Optional ByVal showProgress As Boolean = False)

    Dim wsUI As Worksheet
    Dim wsProducts As Worksheet
    Dim wsPurchase As Worksheet

    Dim tblP As ListObject
    Dim tblPO As ListObject

    Dim i As Long
    Dim outRow As Long
    Dim totalRows As Long

    Dim productID As String
    Dim sku As String
    Dim productName As String
    Dim currentStock As Double
    Dim unitCost As Double
    Dim inventoryValue As Double
    Dim reorderLevel As Variant
    Dim updatedAt As Variant
    Dim stockHealth As String

    Dim openPOQty As Double
    Dim inboundQty As Double
    Dim openPOCount As Long
    Dim latestPONo As String
    Dim metrics As Variant

    Dim colProductID As Long
    Dim colSKU As Long
    Dim colName As Long
    Dim colStock As Long
    Dim colUnitCost As Long
    Dim colReorder As Long
    Dim colUpdated As Long

    Dim dictPipeline As Object

    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalc As XlCalculation

    On Error GoTo ErrHandler

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsProducts = ThisWorkbook.Worksheets(SHEET_PRODUCTS)
    Set wsPurchase = ThisWorkbook.Worksheets(SHEET_PURCHASE)

    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)
    Set tblPO = wsPurchase.ListObjects(TABLE_PURCHASE)

    If showProgress Then
        ProgressStart "Loading Inventory", "Preparing inventory list..."
    End If

    PrepareInventoryUIHeader wsUI
    ClearInventoryOutput wsUI

    If tblP.DataBodyRange Is Nothing Then
        If showProgress Then ProgressEnd
        MsgBox "Products_DB is empty.", vbExclamation, "Inventory UI"
        GoTo CleanExit
    End If

    colProductID = GetCol(tblP, "Product_ID")
    colSKU = GetCol(tblP, "SKU")
    colName = GetCol(tblP, "Product_Name")
    colStock = GetCol(tblP, "Current_Stock")
    colUnitCost = GetOptionalCol(tblP, "Unit_Cost")
    colReorder = GetOptionalCol(tblP, "Reorder_Level")
    colUpdated = GetOptionalCol(tblP, "Updated_At")

    If showProgress Then ProgressStep "Building PO pipeline index..."
    Set dictPipeline = BuildPurchasePipelineDict(tblPO)

    outRow = FIRST_DATA_ROW
    totalRows = tblP.ListRows.Count

    For i = 1 To totalRows

        productID = Trim$(CStr(tblP.DataBodyRange.Cells(i, colProductID).value))
        sku = Trim$(CStr(tblP.DataBodyRange.Cells(i, colSKU).value))
        productName = Trim$(CStr(tblP.DataBodyRange.Cells(i, colName).value))

        If showProgress Then
            ProgressUpdate "Loading products", i, totalRows, sku
        End If

        If MatchesInventoryFilter(sku, productName, skuKeyword, nameKeyword) Then

            currentStock = NzNumber(tblP.DataBodyRange.Cells(i, colStock).value)

            If colUnitCost > 0 Then
                unitCost = NzNumber(tblP.DataBodyRange.Cells(i, colUnitCost).value)
            Else
                unitCost = 0
            End If

            inventoryValue = Round(currentStock * unitCost, 2)

            If colReorder > 0 Then
                reorderLevel = tblP.DataBodyRange.Cells(i, colReorder).value
            Else
                reorderLevel = ""
            End If

            If colUpdated > 0 Then
                updatedAt = tblP.DataBodyRange.Cells(i, colUpdated).value
            Else
                updatedAt = ""
            End If

            openPOQty = 0
            inboundQty = 0
            openPOCount = 0
            latestPONo = ""

            If dictPipeline.Exists(sku) Then
                metrics = dictPipeline(sku)
                openPOQty = metrics(0)
                inboundQty = metrics(1)
                openPOCount = CLng(metrics(2))
                latestPONo = CStr(metrics(3))
            End If

            stockHealth = GetInventoryHealthPlaceholder(currentStock, reorderLevel)

            wsUI.Cells(outRow, COL_OUT_PRODUCT_ID).value = productID
            wsUI.Cells(outRow, COL_OUT_SKU).value = sku
            wsUI.Cells(outRow, COL_OUT_NAME).value = productName
            wsUI.Cells(outRow, COL_OUT_STOCK).value = currentStock
            wsUI.Cells(outRow, COL_OUT_OPEN_PO_QTY).value = openPOQty
            wsUI.Cells(outRow, COL_OUT_INBOUND_QTY).value = inboundQty
            wsUI.Cells(outRow, COL_OUT_OPEN_PO_COUNT).value = openPOCount
            WriteLatestPOHyperlink wsUI.Cells(outRow, COL_OUT_LATEST_PO), latestPONo
            wsUI.Cells(outRow, COL_OUT_UNIT_COST).value = unitCost
            wsUI.Cells(outRow, COL_OUT_INV_VALUE).value = inventoryValue
            wsUI.Cells(outRow, COL_OUT_REORDER).value = reorderLevel
            wsUI.Cells(outRow, COL_OUT_HEALTH).value = stockHealth
            wsUI.Cells(outRow, COL_OUT_UPDATED).value = updatedAt

            ApplyInventoryHealthFormat wsUI.Cells(outRow, COL_OUT_HEALTH), stockHealth

            outRow = outRow + 1

        End If
    Next i

    If outRow = FIRST_DATA_ROW Then
        wsUI.Cells(FIRST_DATA_ROW, COL_OUT_NAME).value = "No matching products found."
    Else
        FormatInventoryOutput wsUI, outRow - 1
    End If

CleanExit:
    If showProgress Then ProgressEnd

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalc

    Exit Sub

ErrHandler:
    If showProgress Then ProgressEnd

    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalc

    MsgBox "Error: " & Err.Description, vbCritical, "Inventory UI"

End Sub

'==================================================
' UI PREP
'==================================================
Private Sub PrepareInventoryUIHeader(ByVal wsUI As Worksheet)

    With wsUI
        .Cells(HEADER_ROW, COL_OUT_PRODUCT_ID).value = "Product_ID"
        .Cells(HEADER_ROW, COL_OUT_SKU).value = "SKU"
        .Cells(HEADER_ROW, COL_OUT_NAME).value = "Product_Name"
        .Cells(HEADER_ROW, COL_OUT_STOCK).value = "Current_Stock"
        .Cells(HEADER_ROW, COL_OUT_OPEN_PO_QTY).value = "Open_PO_Qty"
        .Cells(HEADER_ROW, COL_OUT_INBOUND_QTY).value = "Inbound_Qty"
        .Cells(HEADER_ROW, COL_OUT_OPEN_PO_COUNT).value = "Open_PO_Count"
        .Cells(HEADER_ROW, COL_OUT_LATEST_PO).value = "Latest_PO_No"
        .Cells(HEADER_ROW, COL_OUT_UNIT_COST).value = "Unit_Cost"
        .Cells(HEADER_ROW, COL_OUT_INV_VALUE).value = "Inventory_Value"
        .Cells(HEADER_ROW, COL_OUT_REORDER).value = "Reorder_Level"
        .Cells(HEADER_ROW, COL_OUT_HEALTH).value = "Stock_Health"
        .Cells(HEADER_ROW, COL_OUT_UPDATED).value = "Updated_At"

        With .Range(.Cells(HEADER_ROW, COL_OUT_PRODUCT_ID), .Cells(HEADER_ROW, COL_OUT_UPDATED))
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
            .Borders.LineStyle = xlContinuous
        End With
    End With

End Sub

Private Sub ClearInventoryOutput(ByVal wsUI As Worksheet)

    Dim rng As Range

    Set rng = wsUI.Range(wsUI.Cells(FIRST_DATA_ROW, COL_OUT_PRODUCT_ID), wsUI.Cells(wsUI.Rows.Count, COL_OUT_UPDATED))

    On Error Resume Next
    rng.Hyperlinks.Delete
    On Error GoTo 0

    rng.ClearContents
    rng.Interior.Pattern = xlNone
    rng.Borders.LineStyle = xlNone
    rng.Font.Underline = xlUnderlineStyleNone
    rng.Font.ColorIndex = xlAutomatic

End Sub

Private Sub FormatInventoryOutput(ByVal wsUI As Worksheet, ByVal lastRow As Long)

    With wsUI.Range(wsUI.Cells(FIRST_DATA_ROW, COL_OUT_PRODUCT_ID), wsUI.Cells(lastRow, COL_OUT_UPDATED))
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With

    wsUI.Range(wsUI.Cells(FIRST_DATA_ROW, COL_OUT_OPEN_PO_QTY), wsUI.Cells(lastRow, COL_OUT_INBOUND_QTY)).NumberFormat = "#,##0.00"
    wsUI.Range(wsUI.Cells(FIRST_DATA_ROW, COL_OUT_UNIT_COST), wsUI.Cells(lastRow, COL_OUT_UNIT_COST)).NumberFormat = "#,##0.00"
    wsUI.Range(wsUI.Cells(FIRST_DATA_ROW, COL_OUT_INV_VALUE), wsUI.Cells(lastRow, COL_OUT_INV_VALUE)).NumberFormat = "#,##0.00"
    wsUI.Range(wsUI.Cells(FIRST_DATA_ROW, COL_OUT_UPDATED), wsUI.Cells(lastRow, COL_OUT_UPDATED)).NumberFormat = "yyyy-mm-dd"

End Sub

'==================================================
' PURCHASE PIPELINE INDEX
'==================================================
Private Function BuildPurchasePipelineDict(ByVal tblPO As ListObject) As Object

    Dim dict As Object
    Dim dictPOKey As Object

    Dim i As Long
    Dim sku As String
    Dim poNo As String
    Dim poDate As Variant
    Dim qty As Double
    Dim remainingQty As Double
    Dim metrics As Variant
    Dim key As String

    Dim colSKU As Long
    Dim colQty As Long
    Dim colRemaining As Long
    Dim colPONo As Long
    Dim colPODate As Long

    Set dict = CreateObject("Scripting.Dictionary")
    Set dictPOKey = CreateObject("Scripting.Dictionary")

    If tblPO Is Nothing Then
        Set BuildPurchasePipelineDict = dict
        Exit Function
    End If

    If tblPO.DataBodyRange Is Nothing Then
        Set BuildPurchasePipelineDict = dict
        Exit Function
    End If

    colSKU = GetCol(tblPO, "SKU")
    colQty = GetCol(tblPO, "Qty")
    colRemaining = GetCol(tblPO, "Remaining_Qty")
    colPONo = GetCol(tblPO, "Purchase_Order_No")
    colPODate = GetOptionalCol(tblPO, "Purchase_Date")

    For i = 1 To tblPO.ListRows.Count

        sku = Trim$(CStr(tblPO.DataBodyRange.Cells(i, colSKU).value))
        If sku = "" Then GoTo nextRow

        remainingQty = NzNumber(tblPO.DataBodyRange.Cells(i, colRemaining).value)
        If remainingQty <= 0 Then GoTo nextRow

        qty = NzNumber(tblPO.DataBodyRange.Cells(i, colQty).value)
        poNo = Trim$(CStr(tblPO.DataBodyRange.Cells(i, colPONo).value))

        If dict.Exists(sku) Then
            metrics = dict(sku)
        Else
            metrics = Array(0#, 0#, 0&, "", 0#, False)
        End If

        metrics(0) = CDbl(metrics(0)) + qty
        metrics(1) = CDbl(metrics(1)) + remainingQty

        If poNo <> "" Then
            key = sku & Chr(30) & poNo
            If Not dictPOKey.Exists(key) Then
                dictPOKey.Add key, True
                metrics(2) = CLng(metrics(2)) + 1
            End If
        End If

        If colPODate > 0 Then
            poDate = tblPO.DataBodyRange.Cells(i, colPODate).value

            If IsDate(poDate) Then
                If CBool(metrics(5)) = False Then
                    metrics(3) = poNo
                    metrics(4) = CDbl(CDate(poDate))
                    metrics(5) = True
                ElseIf CDbl(CDate(poDate)) >= CDbl(metrics(4)) Then
                    metrics(3) = poNo
                    metrics(4) = CDbl(CDate(poDate))
                End If
            ElseIf CStr(metrics(3)) = "" Then
                metrics(3) = poNo
            End If
        ElseIf CStr(metrics(3)) = "" Then
            metrics(3) = poNo
        End If

        metrics(0) = Round(CDbl(metrics(0)), 2)
        metrics(1) = Round(CDbl(metrics(1)), 2)

        dict(sku) = metrics

nextRow:
    Next i

    Set BuildPurchasePipelineDict = dict

End Function

'==================================================
' FILTER / HEALTH
'==================================================
Private Function MatchesInventoryFilter(ByVal sku As String, _
                                        ByVal productName As String, _
                                        ByVal skuKeyword As String, _
                                        ByVal nameKeyword As String) As Boolean

    Dim okSKU As Boolean
    Dim okName As Boolean

    okSKU = True
    okName = True

    If Trim$(skuKeyword) <> "" Then
        okSKU = (InStr(1, sku, skuKeyword, vbTextCompare) > 0)
    End If

    If Trim$(nameKeyword) <> "" Then
        okName = (InStr(1, productName, nameKeyword, vbTextCompare) > 0)
    End If

    MatchesInventoryFilter = (okSKU And okName)

End Function

Private Function GetInventoryHealthPlaceholder(ByVal currentStock As Double, ByVal reorderLevel As Variant) As String

    If Trim$(CStr(reorderLevel)) = "" Then
        GetInventoryHealthPlaceholder = "-"
        Exit Function
    End If

    If Not IsNumeric(reorderLevel) Then
        GetInventoryHealthPlaceholder = "-"
        Exit Function
    End If

    If currentStock <= 0 Then
        GetInventoryHealthPlaceholder = "Out of Stock"
    ElseIf currentStock <= CDbl(reorderLevel) Then
        GetInventoryHealthPlaceholder = "Low"
    Else
        GetInventoryHealthPlaceholder = "Healthy"
    End If

End Function

Private Sub ApplyInventoryHealthFormat(ByVal targetCell As Range, ByVal stockHealth As String)

    targetCell.Font.Bold = True

    Select Case UCase$(Trim$(stockHealth))
        Case "OUT OF STOCK"
            targetCell.Interior.Color = RGB(255, 199, 206)
        Case "LOW"
            targetCell.Interior.Color = RGB(255, 235, 156)
        Case "HEALTHY"
            targetCell.Interior.Color = RGB(198, 239, 206)
        Case Else
            targetCell.Interior.Pattern = xlNone
    End Select

End Sub

Private Sub WriteLatestPOHyperlink(ByVal targetCell As Range, ByVal latestPONo As String)

    On Error Resume Next
    targetCell.Hyperlinks.Delete
    On Error GoTo 0

    targetCell.value = ""

    If Trim$(latestPONo) = "" Then Exit Sub

    targetCell.Worksheet.Hyperlinks.Add _
        Anchor:=targetCell, _
        Address:="", _
        SubAddress:="Purchase_UI!B3", _
        TextToDisplay:=latestPONo

End Sub

'==================================================
' HELPERS
'==================================================
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

Private Function NzNumber(ByVal v As Variant) As Double

    If IsError(v) Then
        NzNumber = 0
    ElseIf IsNumeric(v) Then
        NzNumber = CDbl(v)
    ElseIf Trim$(CStr(v)) = "" Then
        NzNumber = 0
    Else
        NzNumber = 0
    End If

End Function

