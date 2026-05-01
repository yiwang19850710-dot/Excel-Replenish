Attribute VB_Name = "modInventoryHistory"
Option Explicit

Private Const WS_HISTORY_UI As String = "Inventory_History_UI"
Private Const WS_INVENTORY_LOG As String = "Inventory_Log"
Private Const WS_PRODUCTS_DB As String = "Products_DB"

Private Const CELL_SKU As String = "B4"
Private Const CELL_DATE_FROM As String = "B5"
Private Const CELL_DATE_TO As String = "B6"
Private Const CELL_TYPE As String = "B7"

Private Const CELL_SUM_IN_LABEL As String = "G4"
Private Const CELL_SUM_OUT_LABEL As String = "G5"
Private Const CELL_SUM_NET_LABEL As String = "G6"
Private Const CELL_SUM_CLOSE_LABEL As String = "G7"

Private Const CELL_SUM_IN_VALUE As String = "H4"
Private Const CELL_SUM_OUT_VALUE As String = "H5"
Private Const CELL_SUM_NET_VALUE As String = "H6"
Private Const CELL_SUM_CLOSE_VALUE As String = "H7"

Private Const ROW_DATA_START As Long = 11

Private Const COL_LOG_DATE As Long = 1
Private Const COL_MOVEMENT_TYPE As Long = 2
Private Const COL_REFERENCE As Long = 3
Private Const COL_SKU As Long = 4
Private Const COL_PRODUCT As Long = 5
Private Const COL_QTY_IN As Long = 6
Private Const COL_QTY_OUT As Long = 7
Private Const COL_NET_CHANGE As Long = 8
Private Const COL_RUNNING_BAL As Long = 9
Private Const COL_NOTES As Long = 10

Public Sub InventoryHistory_SetupUI()

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_HISTORY_UI)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    If Trim$(CStr(ws.Range(CELL_TYPE).value)) = "" Then ws.Range(CELL_TYPE).value = "ALL"
    If Trim$(CStr(ws.Range(CELL_DATE_FROM).value)) = "" Then ws.Range(CELL_DATE_FROM).value = Date - 30
    If Trim$(CStr(ws.Range(CELL_DATE_TO).value)) = "" Then ws.Range(CELL_DATE_TO).value = Date

    ws.Range(CELL_SUM_IN_LABEL).value = "Total In"
    ws.Range(CELL_SUM_OUT_LABEL).value = "Total Out"
    ws.Range(CELL_SUM_NET_LABEL).value = "Net Change"
    ws.Range(CELL_SUM_CLOSE_LABEL).value = "Closing Balance"

    On Error Resume Next
    ws.Range(CELL_TYPE).Validation.Delete
    On Error GoTo 0

    With ws.Range(CELL_TYPE).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
             Formula1:="ALL,RECEIVING,SALES,ADJUSTMENT,OPENING,OPENING_ADJ,STOCK_LOAD,SHOPIFY_STOCK_SYNC"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Movement Type"
        .ErrorTitle = "Invalid Type"
        .InputMessage = "Choose movement type."
        .ErrorMessage = "Please choose a valid movement type."
        .ShowInput = True
        .ShowError = True
    End With

    SetupInventoryHistorySKUValidation
    FormatSummaryArea

End Sub

Public Sub SetupInventoryHistorySKUValidation()

    Dim wsUI As Worksheet
    Dim wsP As Worksheet
    Dim lo As ListObject
    Dim refersToText As String

    On Error GoTo SafeExit

    Set wsUI = ThisWorkbook.Worksheets(WS_HISTORY_UI)
    Set wsP = ThisWorkbook.Worksheets(WS_PRODUCTS_DB)

    On Error Resume Next
    Set lo = wsP.ListObjects("tblProducts")
    On Error GoTo SafeExit

    If lo Is Nothing Then Exit Sub

    refersToText = "=" & WS_PRODUCTS_DB & "!" & lo.name & "[SKU]"

    On Error Resume Next
    ThisWorkbook.Names("nmInventoryHistorySKU").Delete
    On Error GoTo SafeExit

    ThisWorkbook.Names.Add name:="nmInventoryHistorySKU", RefersTo:=refersToText

    On Error Resume Next
    wsUI.Range(CELL_SKU).Validation.Delete
    On Error GoTo SafeExit

    With wsUI.Range(CELL_SKU).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=nmInventoryHistorySKU"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "SKU"
        .ErrorTitle = "Invalid SKU"
        .InputMessage = "Choose a SKU from the dropdown list, or leave blank for all SKUs."
        .ErrorMessage = "Please choose a valid SKU."
        .ShowInput = True
        .ShowError = True
    End With

SafeExit:
End Sub

Public Sub InventoryHistory_Search()

    Dim wsUI As Worksheet
    Dim wsLog As Worksheet
    Dim rng As Range
    Dim lastRow As Long

    Dim colLogDate As Long
    Dim colTranType As Long
    Dim colRefNo As Long
    Dim colSKU As Long
    Dim colProduct As Long
    Dim colQtyChange As Long
    Dim colBalanceAfter As Long
    Dim colNotes As Long

    Dim filterSKU As String
    Dim filterType As String
    Dim dateFrom As Date
    Dim dateTo As Date

    Dim i As Long
    Dim writeRow As Long

    Dim logDate As Variant
    Dim tranType As String
    Dim refNo As String
    Dim sku As String
    Dim productName As String
    Dim qtyChange As Double
    Dim balanceAfter As Double
    Dim notesText As String

    Dim qtyIn As Double
    Dim qtyOut As Double
    Dim matchedCount As Long

    Dim totalIn As Double
    Dim totalOut As Double
    Dim netChange As Double
    Dim closingBalance As Variant

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(WS_HISTORY_UI)
    Set wsLog = ThisWorkbook.Worksheets(WS_INVENTORY_LOG)

    If Not ValidateHistoryInputs(wsUI) Then Exit Sub

    filterSKU = Trim$(CStr(wsUI.Range(CELL_SKU).value))
    filterType = UCase$(Trim$(CStr(wsUI.Range(CELL_TYPE).value)))
    dateFrom = CDate(wsUI.Range(CELL_DATE_FROM).value)
    dateTo = CDate(wsUI.Range(CELL_DATE_TO).value)

    If dateTo < dateFrom Then
        MsgBox "Date To cannot be earlier than Date From.", vbExclamation
        Exit Sub
    End If

    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        ClearHistoryResults
        ClearHistorySummary
        MsgBox "Inventory_Log has no data.", vbInformation
        Exit Sub
    End If

    Set rng = wsLog.Range("A1").CurrentRegion

    colLogDate = FindHeaderInRange(rng, "Log_Date")
    colTranType = FindHeaderInRange(rng, "Tran_Type")
    colRefNo = FindHeaderInRange(rng, "Ref_No")
    colSKU = FindHeaderInRange(rng, "SKU")
    colProduct = FindHeaderInRange(rng, "Product_Name")
    colQtyChange = FindHeaderInRange(rng, "Qty_Change")
    colBalanceAfter = FindHeaderInRange(rng, "Balance_After")
    colNotes = FindHeaderInRange(rng, "Notes")

    If colLogDate = 0 Or colTranType = 0 Or colRefNo = 0 Or colSKU = 0 Or _
       colProduct = 0 Or colQtyChange = 0 Or colBalanceAfter = 0 Then
        MsgBox "Inventory_Log is missing one or more required columns.", vbCritical
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ClearHistoryResults
    ClearHistorySummary

    writeRow = ROW_DATA_START
    matchedCount = 0
    totalIn = 0
    totalOut = 0
    netChange = 0
    closingBalance = ""

    For i = 2 To rng.Rows.Count

        logDate = rng.Cells(i, colLogDate).value
        tranType = UCase$(Trim$(CStr(rng.Cells(i, colTranType).value)))
        refNo = Trim$(CStr(rng.Cells(i, colRefNo).value))
        sku = Trim$(CStr(rng.Cells(i, colSKU).value))
        productName = Trim$(CStr(rng.Cells(i, colProduct).value))
        qtyChange = NzNum(rng.Cells(i, colQtyChange).value)
        balanceAfter = NzNum(rng.Cells(i, colBalanceAfter).value)

        If colNotes > 0 Then
            notesText = Trim$(CStr(rng.Cells(i, colNotes).value))
        Else
            notesText = ""
        End If

        If IsDate(logDate) Then
            If CDate(logDate) >= dateFrom And CDate(logDate) <= dateTo Then
                If filterSKU = "" Or StrComp(filterSKU, sku, vbTextCompare) = 0 Then
                    If IsMovementTypeMatch(filterType, tranType) Then

                        qtyIn = 0
                        qtyOut = 0

                        If qtyChange >= 0 Then
                            qtyIn = qtyChange
                        Else
                            qtyOut = Abs(qtyChange)
                        End If

                        wsUI.Cells(writeRow, COL_LOG_DATE).value = CDate(logDate)
                        wsUI.Cells(writeRow, COL_MOVEMENT_TYPE).value = NormalizeTranType(tranType)
                        wsUI.Cells(writeRow, COL_REFERENCE).value = refNo
                        wsUI.Cells(writeRow, COL_SKU).value = sku
                        wsUI.Cells(writeRow, COL_PRODUCT).value = productName
                        wsUI.Cells(writeRow, COL_QTY_IN).value = qtyIn
                        wsUI.Cells(writeRow, COL_QTY_OUT).value = qtyOut
                        wsUI.Cells(writeRow, COL_NET_CHANGE).value = qtyChange
                        wsUI.Cells(writeRow, COL_RUNNING_BAL).value = balanceAfter
                        wsUI.Cells(writeRow, COL_NOTES).value = notesText

                        totalIn = totalIn + qtyIn
                        totalOut = totalOut + qtyOut
                        netChange = netChange + qtyChange
                        closingBalance = balanceAfter

                        writeRow = writeRow + 1
                        matchedCount = matchedCount + 1

                    End If
                End If
            End If
        End If
    Next i

    If matchedCount > 0 Then
        SortHistoryResults writeRow - 1
        closingBalance = wsUI.Cells(writeRow - 1, COL_RUNNING_BAL).value
        FormatHistoryResults writeRow - 1
        WriteHistorySummary totalIn, totalOut, netChange, closingBalance
    Else
        MsgBox "No records found for the selected criteria.", vbInformation
    End If

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Inventory History"
    Resume SafeExit

End Sub

Public Sub InventoryHistory_Clear()

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(WS_HISTORY_UI)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ws.Range(CELL_SKU).ClearContents
    ws.Range(CELL_DATE_FROM).value = Date - 30
    ws.Range(CELL_DATE_TO).value = Date
    ws.Range(CELL_TYPE).value = "ALL"

    ClearHistoryResults
    ClearHistorySummary

    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Private Function ValidateHistoryInputs(ByVal ws As Worksheet) As Boolean

    ValidateHistoryInputs = False

    If Trim$(CStr(ws.Range(CELL_DATE_FROM).value)) = "" Then
        MsgBox "Please enter Date From.", vbExclamation
        Exit Function
    End If

    If Trim$(CStr(ws.Range(CELL_DATE_TO).value)) = "" Then
        MsgBox "Please enter Date To.", vbExclamation
        Exit Function
    End If

    If Not IsDate(ws.Range(CELL_DATE_FROM).value) Then
        MsgBox "Date From is not a valid date.", vbExclamation
        Exit Function
    End If

    If Not IsDate(ws.Range(CELL_DATE_TO).value) Then
        MsgBox "Date To is not a valid date.", vbExclamation
        Exit Function
    End If

    If Trim$(CStr(ws.Range(CELL_TYPE).value)) = "" Then
        ws.Range(CELL_TYPE).value = "ALL"
    End If

    ValidateHistoryInputs = True

End Function

Private Sub ClearHistoryResults()

    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(WS_HISTORY_UI)

    ws.Range(ws.Cells(ROW_DATA_START, COL_LOG_DATE), ws.Cells(ws.Rows.Count, COL_NOTES)).ClearContents
    ws.Range(ws.Cells(ROW_DATA_START, COL_LOG_DATE), ws.Cells(ws.Rows.Count, COL_NOTES)).Interior.Pattern = xlNone
    ws.Range(ws.Cells(ROW_DATA_START, COL_LOG_DATE), ws.Cells(ws.Rows.Count, COL_NOTES)).Borders.LineStyle = xlNone

End Sub

Private Sub FormatHistoryResults(ByVal lastDataRow As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_HISTORY_UI)

    If lastDataRow < ROW_DATA_START Then Exit Sub

    With ws.Range(ws.Cells(ROW_DATA_START, COL_LOG_DATE), ws.Cells(lastDataRow, COL_NOTES))
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With

    ws.Range(ws.Cells(ROW_DATA_START, COL_LOG_DATE), ws.Cells(lastDataRow, COL_LOG_DATE)).NumberFormat = "yyyy-mm-dd"
    ws.Range(ws.Cells(ROW_DATA_START, COL_QTY_IN), ws.Cells(lastDataRow, COL_RUNNING_BAL)).NumberFormat = "#,##0.00"

End Sub

Private Sub SortHistoryResults(ByVal lastDataRow As Long)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_HISTORY_UI)

    If lastDataRow < ROW_DATA_START + 1 Then Exit Sub

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(ws.Cells(ROW_DATA_START, COL_LOG_DATE), ws.Cells(lastDataRow, COL_LOG_DATE)), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(ROW_DATA_START, COL_LOG_DATE), ws.Cells(lastDataRow, COL_NOTES))
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

End Sub

Private Sub WriteHistorySummary(ByVal totalIn As Double, _
                                ByVal totalOut As Double, _
                                ByVal netChange As Double, _
                                ByVal closingBalance As Variant)

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_HISTORY_UI)

    ws.Range(CELL_SUM_IN_VALUE).value = totalIn
    ws.Range(CELL_SUM_OUT_VALUE).value = totalOut
    ws.Range(CELL_SUM_NET_VALUE).value = netChange
    ws.Range(CELL_SUM_CLOSE_VALUE).value = closingBalance

    ws.Range(CELL_SUM_IN_VALUE & ":" & CELL_SUM_CLOSE_VALUE).NumberFormat = "#,##0.00"

End Sub

Private Sub ClearHistorySummary()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_HISTORY_UI)

    ws.Range(CELL_SUM_IN_VALUE).ClearContents
    ws.Range(CELL_SUM_OUT_VALUE).ClearContents
    ws.Range(CELL_SUM_NET_VALUE).ClearContents
    ws.Range(CELL_SUM_CLOSE_VALUE).ClearContents

End Sub

Private Sub FormatSummaryArea()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(WS_HISTORY_UI)

    ws.Range(CELL_SUM_IN_LABEL & ":" & CELL_SUM_CLOSE_LABEL).Font.Bold = True
    ws.Range(CELL_SUM_IN_VALUE & ":" & CELL_SUM_CLOSE_VALUE).NumberFormat = "#,##0.00"

End Sub

Private Function IsMovementTypeMatch(ByVal filterType As String, ByVal logType As String) As Boolean

    filterType = UCase$(Trim$(filterType))
    logType = UCase$(Trim$(logType))

    Select Case filterType
        Case "ALL"
            IsMovementTypeMatch = True

        Case "RECEIVING"
            IsMovementTypeMatch = (logType = "RECEIVE" Or logType = "RECEIVING")

        Case "SALES"
            IsMovementTypeMatch = (logType = "SALE" Or logType = "SALES")

        Case "ADJUSTMENT"
            IsMovementTypeMatch = (logType = "ADJUSTMENT" Or logType = "OPENING_ADJ" Or logType = "STOCK_LOAD" Or logType = "SHOPIFY_STOCK_SYNC")

        Case "OPENING"
            IsMovementTypeMatch = (logType = "OPENING")

        Case "OPENING_ADJ"
            IsMovementTypeMatch = (logType = "OPENING_ADJ")

        Case "STOCK_LOAD"
            IsMovementTypeMatch = (logType = "STOCK_LOAD")

        Case "SHOPIFY_STOCK_SYNC"
            IsMovementTypeMatch = (logType = "SHOPIFY_STOCK_SYNC")

        Case Else
            IsMovementTypeMatch = False
    End Select

End Function

Private Function NormalizeTranType(ByVal logType As String) As String

    logType = UCase$(Trim$(logType))

    Select Case logType
        Case "RECEIVE", "RECEIVING"
            NormalizeTranType = "RECEIVING"
        Case "SALE", "SALES"
            NormalizeTranType = "SALES"
        Case "ADJUSTMENT"
            NormalizeTranType = "ADJUSTMENT"
        Case "OPENING"
            NormalizeTranType = "OPENING"
        Case "OPENING_ADJ"
            NormalizeTranType = "OPENING_ADJ"
        Case "STOCK_LOAD"
            NormalizeTranType = "STOCK_LOAD"
        Case "SHOPIFY_STOCK_SYNC"
            NormalizeTranType = "SHOPIFY_STOCK_SYNC"
        Case Else
            NormalizeTranType = logType
    End Select

End Function

Private Function FindHeaderInRange(ByVal rng As Range, ByVal headerName As String) As Long

    Dim c As Long
    FindHeaderInRange = 0

    If rng Is Nothing Then Exit Function

    For c = 1 To rng.Columns.Count
        If StrComp(Trim$(CStr(rng.Cells(1, c).value)), Trim$(headerName), vbTextCompare) = 0 Then
            FindHeaderInRange = c
            Exit Function
        End If
    Next c

End Function

Private Function NzNum(ByVal v As Variant) As Double

    If IsError(v) Then
        NzNum = 0
    ElseIf IsNumeric(v) Then
        NzNum = CDbl(v)
    Else
        NzNum = 0
    End If

End Function

