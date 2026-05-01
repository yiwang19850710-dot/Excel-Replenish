Attribute VB_Name = "modStockHealth"
Option Explicit

Private Const SHEET_UI As String = "Stock_Health_UI"
Private Const SHEET_PRODUCTS As String = "Products_DB"
Private Const SHEET_SALES As String = "Sales_DB"
Private Const SHEET_PURCHASE As String = "Purchase_DB"
Private Const SHEET_FORECAST As String = "Forecast_DB"
Private Const SHEET_SETUP As String = "Setup_UI"

Private Const TABLE_PRODUCTS As String = "tblProducts"
Private Const TABLE_PURCHASE As String = "tblPurchase"

' ===== UI mapping =====
Private Const CELL_SKU_SELECTOR As String = "B3"
Private Const RANGE_SKU_LIST As String = "B4:B8"

Private Const CELL_PRODUCT_NAME As String = "B10"
Private Const CELL_CURRENT_STOCK As String = "B11"
Private Const CELL_FORECAST_SOURCE As String = "B12"
Private Const CELL_LEAD_TIME As String = "B13"

' ===== Setup_UI mapping =====
Private Const CELL_SETUP_DEFAULT_SAFETY_DAYS As String = "B12"
Private Const CELL_SETUP_HISTORICAL_DAYS As String = "B13"
Private Const CELL_SETUP_DSI_RED_MAX As String = "B14"
Private Const CELL_SETUP_DSI_YELLOW_MAX As String = "B15"
Private Const CELL_SETUP_DSI_GREEN_MAX As String = "B16"

' ===== Grid layout =====
Private Const GRID_START_COL As Long = 2   ' B
Private Const GRID_DAYS As Long = 90
Private Const FIRST_BLOCK_ROW As Long = 15
Private Const BLOCK_HEIGHT As Long = 8
Private Const MAX_SKUS As Long = 5

Private Const NAME_SKU_LIST As String = "nmStockHealthSKU"

'==================================================
' PUBLIC ACTIONS
'==================================================
Public Sub StockHealth_AddSKU()

    Dim wsUI As Worksheet
    Dim tblP As ListObject
    Dim sku As String
    Dim targetCell As Range
    Dim c As Range

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set tblP = ThisWorkbook.Worksheets(SHEET_PRODUCTS).ListObjects(TABLE_PRODUCTS)

    SetupStockHealthSKUValidation

    sku = Trim$(CStr(wsUI.Range(CELL_SKU_SELECTOR).value))
    If sku = "" Then
        MsgBox "Please choose a SKU first.", vbExclamation, "Add SKU"
        Exit Sub
    End If

    If Trim$(CStr(GetProductFieldBySKU(tblP, sku, "Product_Name"))) = "" Then
        MsgBox "SKU not found in Products_DB: " & sku, vbExclamation, "Add SKU"
        Exit Sub
    End If

    For Each c In wsUI.Range(RANGE_SKU_LIST)
        If Trim$(CStr(c.value)) = sku Then
            MsgBox "This SKU is already in the list.", vbInformation, "Add SKU"
            Exit Sub
        End If
    Next c

    Set targetCell = Nothing
    For Each c In wsUI.Range(RANGE_SKU_LIST)
        If Trim$(CStr(c.value)) = "" Then
            Set targetCell = c
            Exit For
        End If
    Next c

    If targetCell Is Nothing Then
        MsgBox "You can select up to " & MAX_SKUS & " SKUs at one time.", vbExclamation, "Add SKU"
        Exit Sub
    End If

    targetCell.value = sku
    wsUI.Range(CELL_SKU_SELECTOR).value = ""

    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Add SKU"

End Sub

Public Sub StockHealth_View()

    Dim wsUI As Worksheet
    Dim wsProducts As Worksheet
    Dim tblP As ListObject
    Dim skuList As Collection
    Dim sku As Variant

    Dim idx As Long
    Dim blockTop As Long

    Dim currentStock As Double
    Dim leadTimeDays As Long
    Dim historicalDays As Long
    Dim forecastSource As String
    Dim productName As String

    Dim dsiRedMax As Double
    Dim dsiYellowMax As Double
    Dim dsiGreenMax As Double

    Dim forecastArr() As Double
    Dim supplyArr() As Double
    Dim dateArr() As Variant
    Dim i As Long

    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldStatusBar As Variant

    On Error GoTo ErrHandler

    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Rendering Stock Health..."

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsProducts = ThisWorkbook.Worksheets(SHEET_PRODUCTS)
    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)

    SetupStockHealthSKUValidation

    Set skuList = GetStockHealthSKUList()
    If skuList.Count = 0 Then
        MsgBox "Please add at least one SKU into the SKU List first.", vbExclamation, "Stock Health"
        GoTo SafeExit
    End If

    historicalDays = CLng(GetSetupNumber(CELL_SETUP_HISTORICAL_DAYS, 30))
    If historicalDays <= 0 Then historicalDays = 30

    dsiRedMax = GetSetupNumber(CELL_SETUP_DSI_RED_MAX, 7)
    dsiYellowMax = GetSetupNumber(CELL_SETUP_DSI_YELLOW_MAX, 15)
    dsiGreenMax = GetSetupNumber(CELL_SETUP_DSI_GREEN_MAX, 45)

    ClearStockHealthOutput wsUI

    idx = 0

    For Each sku In skuList
        idx = idx + 1
        blockTop = FIRST_BLOCK_ROW + (idx - 1) * BLOCK_HEIGHT

        productName = Trim$(CStr(GetProductFieldBySKU(tblP, CStr(sku), "Product_Name")))
        currentStock = NzNumber(GetProductFieldBySKU(tblP, CStr(sku), "Current_Stock"))
        leadTimeDays = CLng(NzNumber(GetProductFieldBySKU(tblP, CStr(sku), "Lead_Time_Days")))
        If leadTimeDays <= 0 Then leadTimeDays = 30

        ReDim forecastArr(1 To 1, 1 To GRID_DAYS)
        ReDim supplyArr(1 To 1, 1 To GRID_DAYS)
        ReDim dateArr(1 To 1, 1 To GRID_DAYS)

        forecastSource = BuildDefaultForecastArray_2D(CStr(sku), historicalDays, forecastArr)
        BuildDefaultSupplyArray_2D CStr(sku), leadTimeDays, supplyArr

        For i = 1 To GRID_DAYS
            dateArr(1, i) = Date + (i - 1)
        Next i

        RenderStockHealthBlock wsUI, blockTop, CStr(sku), productName, currentStock, forecastSource, leadTimeDays, _
                               dateArr, forecastArr, supplyArr, dsiRedMax, dsiYellowMax, dsiGreenMax
    Next sku

    RefreshTopInfoFromFirstSKUInList

SafeExit:
    Application.StatusBar = oldStatusBar
    Application.Calculation = oldCalc
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Stock Health"
    Resume SafeExit

End Sub

Public Sub StockHealth_Clear()

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    wsUI.Range(CELL_SKU_SELECTOR).value = ""
    wsUI.Range(RANGE_SKU_LIST).ClearContents

    wsUI.Range(CELL_PRODUCT_NAME).value = ""
    wsUI.Range(CELL_CURRENT_STOCK).value = ""
    wsUI.Range(CELL_FORECAST_SOURCE).value = ""
    wsUI.Range(CELL_LEAD_TIME).value = ""

    ClearStockHealthOutput wsUI
    SetupStockHealthSKUValidation

End Sub

Public Sub StockHealth_ResetDefaults()
    StockHealth_View
End Sub

Public Sub SetupStockHealthSKUValidation()

    Dim wsUI As Worksheet
    Dim wsProducts As Worksheet
    Dim tblP As ListObject
    Dim refersToText As String

    On Error GoTo SafeExit

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsProducts = ThisWorkbook.Worksheets(SHEET_PRODUCTS)
    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)

    refersToText = "=" & SHEET_PRODUCTS & "!" & tblP.name & "[SKU]"

    On Error Resume Next
    ThisWorkbook.Names(NAME_SKU_LIST).Delete
    On Error GoTo SafeExit

    ThisWorkbook.Names.Add name:=NAME_SKU_LIST, RefersTo:=refersToText

    wsUI.Range(CELL_SKU_SELECTOR).Validation.Delete
    With wsUI.Range(CELL_SKU_SELECTOR).Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=" & NAME_SKU_LIST
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "SKU"
        .ErrorTitle = "Invalid SKU"
        .InputMessage = "Please choose a SKU from the dropdown list."
        .ErrorMessage = "Please choose a valid SKU from the dropdown list."
        .ShowInput = True
        .ShowError = True
    End With

SafeExit:
End Sub

Public Function StockHealth_FillProductInfo() As Boolean

    Dim wsUI As Worksheet
    Dim sku As String

    On Error GoTo FailSafe

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    sku = Trim$(CStr(wsUI.Range(CELL_SKU_SELECTOR).value))

    RefreshTopInfoBySKU sku
    StockHealth_FillProductInfo = True
    Exit Function

FailSafe:
    StockHealth_FillProductInfo = False

End Function

Public Sub RefreshTopInfoBySKU(ByVal sku As String)

    Dim wsUI As Worksheet
    Dim wsProducts As Worksheet
    Dim tblP As ListObject

    Dim productName As String
    Dim currentStock As Double
    Dim leadTimeDays As Long
    Dim historicalDays As Long
    Dim forecastArr() As Double
    Dim forecastSource As String

    On Error GoTo FailSafe

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set wsProducts = ThisWorkbook.Worksheets(SHEET_PRODUCTS)
    Set tblP = wsProducts.ListObjects(TABLE_PRODUCTS)

    wsUI.Range(CELL_PRODUCT_NAME).value = ""
    wsUI.Range(CELL_CURRENT_STOCK).value = ""
    wsUI.Range(CELL_FORECAST_SOURCE).value = ""
    wsUI.Range(CELL_LEAD_TIME).value = ""

    sku = Trim$(sku)
    If sku = "" Then Exit Sub

    productName = Trim$(CStr(GetProductFieldBySKU(tblP, sku, "Product_Name")))
    If productName = "" Then Exit Sub

    currentStock = NzNumber(GetProductFieldBySKU(tblP, sku, "Current_Stock"))
    leadTimeDays = CLng(NzNumber(GetProductFieldBySKU(tblP, sku, "Lead_Time_Days")))
    If leadTimeDays <= 0 Then leadTimeDays = 30

    historicalDays = CLng(GetSetupNumber(CELL_SETUP_HISTORICAL_DAYS, 30))
    If historicalDays <= 0 Then historicalDays = 30

    ReDim forecastArr(1 To 1, 1 To GRID_DAYS)
    forecastSource = BuildDefaultForecastArray_2D(sku, historicalDays, forecastArr)

    wsUI.Range(CELL_PRODUCT_NAME).value = productName
    wsUI.Range(CELL_CURRENT_STOCK).value = currentStock
    wsUI.Range(CELL_FORECAST_SOURCE).value = forecastSource
    wsUI.Range(CELL_LEAD_TIME).value = leadTimeDays

    Exit Sub

FailSafe:
End Sub

Public Sub RefreshTopInfoFromFirstSKUInList()

    Dim wsUI As Worksheet
    Dim c As Range
    Dim sku As String

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    sku = ""
    For Each c In wsUI.Range(RANGE_SKU_LIST)
        If Trim$(CStr(c.value)) <> "" Then
            sku = Trim$(CStr(c.value))
            Exit For
        End If
    Next c

    RefreshTopInfoBySKU sku

End Sub

Public Function StockHealthBlocksExist() As Boolean

    Dim wsUI As Worksheet
    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    StockHealthBlocksExist = (Trim$(CStr(wsUI.Cells(FIRST_BLOCK_ROW + 1, GRID_START_COL).value)) <> "")

End Function

Public Sub StockHealth_RebuildSupplyByLeadTime()

    Dim wsUI As Worksheet
    Dim skuList As Collection
    Dim tblP As ListObject
    Dim sku As Variant

    Dim idx As Long
    Dim blockTop As Long
    Dim leadTimeDays As Long
    Dim supplyArr() As Double
    Dim currentStock As Double
    Dim productName As String
    Dim forecastSource As String

    Dim dsiRedMax As Double
    Dim dsiYellowMax As Double
    Dim dsiGreenMax As Double

    Dim historicalDays As Long
    Dim forecastArr() As Double
    Dim dateArr() As Variant
    Dim i As Long

    Dim oldCalc As XlCalculation
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldStatusBar As Variant

    On Error GoTo ErrHandler

    oldCalc = Application.Calculation
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Rebuilding supply view..."

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set tblP = ThisWorkbook.Worksheets(SHEET_PRODUCTS).ListObjects(TABLE_PRODUCTS)
    Set skuList = GetStockHealthSKUList()

    If skuList.Count = 0 Then GoTo SafeExit

    historicalDays = CLng(GetSetupNumber(CELL_SETUP_HISTORICAL_DAYS, 30))
    dsiRedMax = GetSetupNumber(CELL_SETUP_DSI_RED_MAX, 7)
    dsiYellowMax = GetSetupNumber(CELL_SETUP_DSI_YELLOW_MAX, 15)
    dsiGreenMax = GetSetupNumber(CELL_SETUP_DSI_GREEN_MAX, 45)

    ClearStockHealthOutput wsUI

    idx = 0
    For Each sku In skuList
        idx = idx + 1
        blockTop = FIRST_BLOCK_ROW + (idx - 1) * BLOCK_HEIGHT

        productName = Trim$(CStr(GetProductFieldBySKU(tblP, CStr(sku), "Product_Name")))
        currentStock = NzNumber(GetProductFieldBySKU(tblP, CStr(sku), "Current_Stock"))

        If idx = 1 Then
            leadTimeDays = CLng(NzNumber(wsUI.Range(CELL_LEAD_TIME).value))
        Else
            leadTimeDays = CLng(NzNumber(GetProductFieldBySKU(tblP, CStr(sku), "Lead_Time_Days")))
        End If
        If leadTimeDays <= 0 Then leadTimeDays = 30

        ReDim forecastArr(1 To 1, 1 To GRID_DAYS)
        ReDim supplyArr(1 To 1, 1 To GRID_DAYS)
        ReDim dateArr(1 To 1, 1 To GRID_DAYS)

        forecastSource = BuildDefaultForecastArray_2D(CStr(sku), historicalDays, forecastArr)
        BuildDefaultSupplyArray_2D CStr(sku), leadTimeDays, supplyArr

        For i = 1 To GRID_DAYS
            dateArr(1, i) = Date + (i - 1)
        Next i

        RenderStockHealthBlock wsUI, blockTop, CStr(sku), productName, currentStock, forecastSource, leadTimeDays, _
                               dateArr, forecastArr, supplyArr, dsiRedMax, dsiYellowMax, dsiGreenMax
    Next sku

    RefreshTopInfoFromFirstSKUInList

SafeExit:
    Application.StatusBar = oldStatusBar
    Application.EnableEvents = oldEnableEvents
    Application.Calculation = oldCalc
    Application.ScreenUpdating = oldScreenUpdating
    Exit Sub

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Stock Health"
    Resume SafeExit

End Sub

Public Sub ReapplyAllStockHealthDSIColors()

    Dim wsUI As Worksheet
    Dim skuList As Collection
    Dim idx As Long
    Dim blockTop As Long

    Dim dsiRedMax As Double
    Dim dsiYellowMax As Double
    Dim dsiGreenMax As Double

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)
    Set skuList = GetStockHealthSKUList()

    dsiRedMax = GetSetupNumber(CELL_SETUP_DSI_RED_MAX, 7)
    dsiYellowMax = GetSetupNumber(CELL_SETUP_DSI_YELLOW_MAX, 15)
    dsiGreenMax = GetSetupNumber(CELL_SETUP_DSI_GREEN_MAX, 45)

    For idx = 1 To skuList.Count
        blockTop = FIRST_BLOCK_ROW + (idx - 1) * BLOCK_HEIGHT
        ApplyDSIColorBandOneRow wsUI, blockTop + 5, dsiRedMax, dsiYellowMax, dsiGreenMax
    Next idx

End Sub

'==================================================
' RENDER BLOCK
'==================================================
Private Sub RenderStockHealthBlock(ByVal wsUI As Worksheet, _
                                   ByVal blockTop As Long, _
                                   ByVal sku As String, _
                                   ByVal productName As String, _
                                   ByVal currentStock As Double, _
                                   ByVal forecastSource As String, _
                                   ByVal leadTimeDays As Long, _
                                   ByRef dateArr() As Variant, _
                                   ByRef forecastArr() As Double, _
                                   ByRef supplyArr() As Double, _
                                   ByVal dsiRedMax As Double, _
                                   ByVal dsiYellowMax As Double, _
                                   ByVal dsiGreenMax As Double)

    Dim lastCol As Long
    Dim rowTitle As Long
    Dim rowDate As Long
    Dim rowForecast As Long
    Dim rowSupply As Long
    Dim rowProjected As Long
    Dim rowDSI As Long

    lastCol = GRID_START_COL + GRID_DAYS - 1

    rowTitle = blockTop
    rowDate = blockTop + 1
    rowForecast = blockTop + 2
    rowSupply = blockTop + 3
    rowProjected = blockTop + 4
    rowDSI = blockTop + 5

    wsUI.Cells(rowTitle, 1).value = "SKU: " & sku & _
                                    " | Product: " & productName & _
                                    " | Current Stock: " & Format(currentStock, "#,##0.00") & _
                                    " | Forecast Source: " & forecastSource & _
                                    " | Leadtime: " & leadTimeDays

    With wsUI.Range(wsUI.Cells(rowTitle, 1), wsUI.Cells(rowTitle, lastCol))
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .Borders.LineStyle = xlContinuous
    End With

    wsUI.Cells(rowDate, 1).value = "Date"
    wsUI.Cells(rowForecast, 1).value = "Forecast"
    wsUI.Cells(rowSupply, 1).value = "Supply"
    wsUI.Cells(rowProjected, 1).value = "Projected Stock"
    wsUI.Cells(rowDSI, 1).value = "DSI"

    With wsUI.Range(wsUI.Cells(rowDate, 1), wsUI.Cells(rowDSI, 1))
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .Borders.LineStyle = xlContinuous
    End With

    wsUI.Range(wsUI.Cells(rowDate, GRID_START_COL), wsUI.Cells(rowDate, lastCol)).value = dateArr
    wsUI.Range(wsUI.Cells(rowForecast, GRID_START_COL), wsUI.Cells(rowForecast, lastCol)).value = forecastArr
    wsUI.Range(wsUI.Cells(rowSupply, GRID_START_COL), wsUI.Cells(rowSupply, lastCol)).value = supplyArr

    WriteProjectedStockFormulasFast wsUI, rowForecast, rowSupply, rowProjected, currentStock
    WriteDSIFormulasFast wsUI, rowForecast, rowProjected, rowDSI

    With wsUI.Range(wsUI.Cells(rowDate, GRID_START_COL), wsUI.Cells(rowDSI, lastCol))
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With

    wsUI.Range(wsUI.Cells(rowDate, GRID_START_COL), wsUI.Cells(rowDate, lastCol)).NumberFormat = "mm-dd"
    wsUI.Range(wsUI.Cells(rowForecast, GRID_START_COL), wsUI.Cells(rowDSI, lastCol)).NumberFormat = "#,##0.00"

    wsUI.Range(wsUI.Cells(rowForecast, GRID_START_COL), wsUI.Cells(rowForecast, lastCol)).Interior.Color = RGB(221, 235, 247)
    wsUI.Range(wsUI.Cells(rowSupply, GRID_START_COL), wsUI.Cells(rowSupply, lastCol)).Interior.Color = RGB(226, 239, 218)

    ApplyDSIColorBandOneRow wsUI, rowDSI, dsiRedMax, dsiYellowMax, dsiGreenMax

End Sub

'==================================================
' CORE BUILDERS
'==================================================
Private Function BuildDefaultForecastArray_2D(ByVal sku As String, _
                                              ByVal historicalDays As Long, _
                                              ByRef forecastArr() As Double) As String

    Dim usedForecastDB As Boolean
    Dim fallbackAvg As Double
    Dim i As Long

    usedForecastDB = TryLoadForecastFromForecastDB_2D(sku, forecastArr)

    If usedForecastDB Then
        BuildDefaultForecastArray_2D = "Forecast_DB (Latest Run)"
        Exit Function
    End If

    fallbackAvg = GetAverageDailySales(sku, historicalDays)

    For i = 1 To GRID_DAYS
        forecastArr(1, i) = Round(fallbackAvg, 2)
    Next i

    BuildDefaultForecastArray_2D = "Historical Avg"

End Function

Private Sub BuildDefaultSupplyArray_2D(ByVal sku As String, _
                                       ByVal leadTimeDays As Long, _
                                       ByRef supplyArr() As Double)

    Dim wsPO As Worksheet
    Dim tblPO As ListObject

    Dim i As Long
    Dim poSKU As String
    Dim remainingQty As Double
    Dim poDate As Variant
    Dim etaDate As Date
    Dim dayIndex As Long

    Set wsPO = ThisWorkbook.Worksheets(SHEET_PURCHASE)
    Set tblPO = wsPO.ListObjects(TABLE_PURCHASE)

    If tblPO.DataBodyRange Is Nothing Then Exit Sub

    For i = 1 To tblPO.ListRows.Count

        poSKU = Trim$(CStr(tblPO.ListColumns("SKU").DataBodyRange.Cells(i, 1).value))
        If poSKU = sku Then

            remainingQty = NzNumber(tblPO.ListColumns("Remaining_Qty").DataBodyRange.Cells(i, 1).value)

            If remainingQty > 0 Then
                poDate = tblPO.ListColumns("Purchase_Date").DataBodyRange.Cells(i, 1).value

                If IsDate(poDate) Then
                    etaDate = CDate(poDate) + leadTimeDays
                Else
                    etaDate = Date + leadTimeDays
                End If

                dayIndex = DateDiff("d", Date, etaDate) + 1

                If dayIndex >= 1 And dayIndex <= GRID_DAYS Then
                    supplyArr(1, dayIndex) = Round(supplyArr(1, dayIndex) + remainingQty, 2)
                End If
            End If
        End If
    Next i

End Sub

Private Sub WriteProjectedStockFormulasFast(ByVal wsUI As Worksheet, _
                                            ByVal rowForecast As Long, _
                                            ByVal rowSupply As Long, _
                                            ByVal rowProjected As Long, _
                                            ByVal currentStock As Double)

    Dim i As Long
    Dim c As Long
    Dim prevCol As Long

    For i = 1 To GRID_DAYS
        c = GRID_START_COL + i - 1

        If i = 1 Then
            wsUI.Cells(rowProjected, c).Formula = "=" & Format$(currentStock, "0.00") & "+" & _
                                                  wsUI.Cells(rowSupply, c).Address(False, False) & "-" & _
                                                  wsUI.Cells(rowForecast, c).Address(False, False)
        Else
            prevCol = c - 1
            wsUI.Cells(rowProjected, c).Formula = "=" & _
                                                  wsUI.Cells(rowProjected, prevCol).Address(False, False) & "+" & _
                                                  wsUI.Cells(rowSupply, c).Address(False, False) & "-" & _
                                                  wsUI.Cells(rowForecast, c).Address(False, False)
        End If
    Next i

End Sub

Private Sub WriteDSIFormulasFast(ByVal wsUI As Worksheet, _
                                 ByVal rowForecast As Long, _
                                 ByVal rowProjected As Long, _
                                 ByVal rowDSI As Long)

    Dim i As Long
    Dim c As Long

    For i = 1 To GRID_DAYS
        c = GRID_START_COL + i - 1

        wsUI.Cells(rowDSI, c).Formula = "=IF(" & wsUI.Cells(rowForecast, c).Address(False, False) & "<=0,""""," & _
                                        wsUI.Cells(rowProjected, c).Address(False, False) & "/" & _
                                        wsUI.Cells(rowForecast, c).Address(False, False) & ")"
    Next i

End Sub

'==================================================
' FORECAST SOURCE
'==================================================
Private Function TryLoadForecastFromForecastDB_2D(ByVal sku As String, ByRef forecastArr() As Double) As Boolean

    Dim wsF As Worksheet
    Dim lastRow As Long
    Dim rng As Range

    Dim colSKU As Long
    Dim colRunID As Long
    Dim colDate As Long
    Dim colFinalQty As Long
    Dim colSystemQty As Long
    Dim colRunDate As Long

    Dim latestRunID As String

    Dim i As Long
    Dim rowSKU As String
    Dim rowRunID As String
    Dim rowDate As Variant
    Dim rowQty As Double
    Dim dayIndex As Long
    Dim foundAny As Boolean

    On Error GoTo FailSafe

    Set wsF = ThisWorkbook.Worksheets(SHEET_FORECAST)

    lastRow = wsF.Cells(wsF.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        TryLoadForecastFromForecastDB_2D = False
        Exit Function
    End If

    Set rng = wsF.Range("A1").CurrentRegion

    colSKU = FindHeaderIndexInRange(rng, "SKU")
    colRunID = FindHeaderIndexInRange(rng, "Run_ID")
    colDate = FindHeaderIndexInRange(rng, "Forecast_Date")
    colFinalQty = FindHeaderIndexInRange(rng, "Final_Forecast_Qty")
    colSystemQty = FindHeaderIndexInRange(rng, "System_Forecast_Qty")
    colRunDate = FindHeaderIndexInRange(rng, "Run_Date")

    If colSKU = 0 Or colRunID = 0 Or colDate = 0 Then
        TryLoadForecastFromForecastDB_2D = False
        Exit Function
    End If

    If colFinalQty = 0 And colSystemQty = 0 Then
        TryLoadForecastFromForecastDB_2D = False
        Exit Function
    End If

    latestRunID = GetLatestForecastRunIDBySKU(rng, sku, colSKU, colRunID, colRunDate)
    If Trim$(latestRunID) = "" Then
        TryLoadForecastFromForecastDB_2D = False
        Exit Function
    End If

    foundAny = False

    For i = 2 To rng.Rows.Count
        rowSKU = Trim$(CStr(rng.Cells(i, colSKU).value))
        rowRunID = Trim$(CStr(rng.Cells(i, colRunID).value))

        If rowSKU = sku And rowRunID = latestRunID Then
            rowDate = rng.Cells(i, colDate).value

            If colFinalQty > 0 Then
                rowQty = NzNumber(rng.Cells(i, colFinalQty).value)
            Else
                rowQty = NzNumber(rng.Cells(i, colSystemQty).value)
            End If

            If IsDate(rowDate) Then
                dayIndex = DateDiff("d", Date, CDate(rowDate)) + 1

                If dayIndex >= 1 And dayIndex <= GRID_DAYS Then
                    forecastArr(1, dayIndex) = Round(rowQty, 2)
                    foundAny = True
                End If
            End If
        End If
    Next i

    TryLoadForecastFromForecastDB_2D = foundAny
    Exit Function

FailSafe:
    TryLoadForecastFromForecastDB_2D = False

End Function

Private Function GetLatestForecastRunIDBySKU(ByVal rng As Range, _
                                             ByVal sku As String, _
                                             ByVal colSKU As Long, _
                                             ByVal colRunID As Long, _
                                             ByVal colRunDate As Long) As String

    Dim i As Long
    Dim rowSKU As String
    Dim rowRunID As String
    Dim rowRunDate As Variant

    Dim latestRunDate As Date
    Dim latestRunID As String

    latestRunDate = 0
    latestRunID = ""

    For i = 2 To rng.Rows.Count
        rowSKU = Trim$(CStr(rng.Cells(i, colSKU).value))

        If rowSKU = sku Then
            rowRunID = Trim$(CStr(rng.Cells(i, colRunID).value))

            If colRunDate > 0 Then
                rowRunDate = rng.Cells(i, colRunDate).value
                If IsDate(rowRunDate) Then
                    If CDate(rowRunDate) >= latestRunDate Then
                        latestRunDate = CDate(rowRunDate)
                        latestRunID = rowRunID
                    End If
                Else
                    latestRunID = rowRunID
                End If
            Else
                latestRunID = rowRunID
            End If
        End If
    Next i

    GetLatestForecastRunIDBySKU = latestRunID

End Function

Private Function GetAverageDailySales(ByVal sku As String, ByVal historicalDays As Long) As Double

    Dim wsS As Worksheet
    Dim lastRow As Long
    Dim rng As Range

    Dim colSKU As Long
    Dim colQty As Long
    Dim colDate As Long

    Dim i As Long
    Dim salesDate As Variant
    Dim totalQty As Double

    On Error GoTo FailSafe

    Set wsS = ThisWorkbook.Worksheets(SHEET_SALES)
    lastRow = wsS.Cells(wsS.Rows.Count, 1).End(xlUp).Row

    If lastRow < 2 Then
        GetAverageDailySales = 0
        Exit Function
    End If

    Set rng = wsS.Range("A1").CurrentRegion

    colSKU = FindHeaderIndexInRange(rng, "SKU")
    colQty = FindHeaderIndexInRange(rng, "Qty")
    colDate = FindHeaderIndexInRange(rng, "Sales_Date")

    If colSKU = 0 Or colQty = 0 Then
        GetAverageDailySales = 0
        Exit Function
    End If

    totalQty = 0

    For i = 2 To rng.Rows.Count
        If Trim$(CStr(rng.Cells(i, colSKU).value)) = sku Then

            If colDate > 0 Then
                salesDate = rng.Cells(i, colDate).value
                If IsDate(salesDate) Then
                    If CDate(salesDate) >= Date - historicalDays + 1 And CDate(salesDate) <= Date Then
                        totalQty = totalQty + NzNumber(rng.Cells(i, colQty).value)
                    End If
                End If
            Else
                totalQty = totalQty + NzNumber(rng.Cells(i, colQty).value)
            End If
        End If
    Next i

    If historicalDays <= 0 Then historicalDays = 30
    GetAverageDailySales = Round(totalQty / historicalDays, 2)
    Exit Function

FailSafe:
    GetAverageDailySales = 0

End Function

'==================================================
' FORMAT / CLEAR / LIST
'==================================================
Private Sub ApplyDSIColorBandOneRow(ByVal wsUI As Worksheet, _
                                    ByVal rowDSI As Long, _
                                    ByVal redMax As Double, _
                                    ByVal yellowMax As Double, _
                                    ByVal greenMax As Double)

    Dim i As Long
    Dim c As Long
    Dim dsiVal As Variant
    Dim targetCell As Range

    For i = 1 To GRID_DAYS
        c = GRID_START_COL + i - 1
        Set targetCell = wsUI.Cells(rowDSI, c)

        dsiVal = targetCell.value
        targetCell.Interior.Pattern = xlNone
        targetCell.Font.Bold = True
        targetCell.Font.Color = RGB(0, 0, 0)

        If Trim$(CStr(dsiVal)) <> "" And IsNumeric(dsiVal) Then
            If CDbl(dsiVal) <= redMax Then
                targetCell.Interior.Color = RGB(255, 0, 0)
                targetCell.Font.Color = RGB(255, 255, 255)
            ElseIf CDbl(dsiVal) <= yellowMax Then
                targetCell.Interior.Color = RGB(255, 255, 0)
                targetCell.Font.Color = RGB(0, 0, 0)
            ElseIf CDbl(dsiVal) <= greenMax Then
                targetCell.Interior.Color = RGB(0, 255, 0)
                targetCell.Font.Color = RGB(0, 0, 0)
            Else
                targetCell.Interior.Color = RGB(0, 0, 255)
                targetCell.Font.Color = RGB(255, 255, 255)
            End If
        End If
    Next i

End Sub

Private Sub ClearStockHealthOutput(ByVal wsUI As Worksheet)

    Dim lastCol As Long
    Dim lastRow As Long

    lastCol = GRID_START_COL + GRID_DAYS - 1
    lastRow = FIRST_BLOCK_ROW + (MAX_SKUS * BLOCK_HEIGHT) + 5

    With wsUI.Range(wsUI.Cells(FIRST_BLOCK_ROW, 1), wsUI.Cells(lastRow, lastCol))
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        .Font.ColorIndex = xlAutomatic
        .Font.Bold = False
    End With

End Sub

Private Function GetStockHealthSKUList() As Collection

    Dim wsUI As Worksheet
    Dim c As Range
    Dim result As New Collection

    Set wsUI = ThisWorkbook.Worksheets(SHEET_UI)

    For Each c In wsUI.Range(RANGE_SKU_LIST)
        If Trim$(CStr(c.value)) <> "" Then
            result.Add Trim$(CStr(c.value))
        End If
    Next c

    Set GetStockHealthSKUList = result

End Function

'==================================================
' SETUP / PRODUCTS
'==================================================
Private Function GetSetupNumber(ByVal cellAddress As String, ByVal defaultValue As Double) As Double

    Dim ws As Worksheet
    Dim v As Variant

    Set ws = ThisWorkbook.Worksheets(SHEET_SETUP)
    v = ws.Range(cellAddress).value

    If IsNumeric(v) Then
        GetSetupNumber = CDbl(v)
    Else
        GetSetupNumber = defaultValue
    End If

End Function

Private Function GetProductFieldBySKU(ByVal tblP As ListObject, ByVal sku As String, ByVal returnCol As String) As Variant

    Dim i As Long

    GetProductFieldBySKU = ""

    If tblP.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To tblP.ListRows.Count
        If Trim$(CStr(tblP.ListColumns("SKU").DataBodyRange.Cells(i, 1).value)) = sku Then
            GetProductFieldBySKU = tblP.ListColumns(returnCol).DataBodyRange.Cells(i, 1).value
            Exit Function
        End If
    Next i

End Function

'==================================================
' HELPERS
'==================================================
Private Function FindHeaderIndexInRange(ByVal rng As Range, ByVal headerName As String) As Long

    Dim c As Long
    FindHeaderIndexInRange = 0

    For c = 1 To rng.Columns.Count
        If StrComp(Trim$(CStr(rng.Cells(1, c).value)), Trim$(headerName), vbTextCompare) = 0 Then
            FindHeaderIndexInRange = c
            Exit Function
        End If
    Next c

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

