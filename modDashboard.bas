Attribute VB_Name = "modDashboard"
Option Explicit

' =========================================================
' SimpleERP Dashboard V3.2
' - Fix Report Date note display
' - Try to auto-run Stock Health "View" logic
' - Clean Stock Health UI before loading SKU
' =========================================================

Public Sub BuildDashboardV32()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim rr As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.name = "Dashboard"
    End If

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ws.Cells.Clear

    On Error Resume Next
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    On Error GoTo 0

    ' -------------------------
    ' Base formatting
    ' -------------------------
    ws.Cells.Font.name = "Calibri"
    ws.Cells.Font.Size = 11

    ws.Columns("A").ColumnWidth = 16
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 16
    ws.Columns("E").ColumnWidth = 16
    ws.Columns("F").ColumnWidth = 14
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 18
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 12

    ws.Rows("1:50").RowHeight = 22
    ws.Rows("8:16").RowHeight = 24
    ws.Rows("4:6").RowHeight = 24

    ' -------------------------
    ' Header
    ' -------------------------
    With ws.Range("A1:J1")
        .Merge
        .value = "DASHBOARD"
        .Font.Bold = True
        .Font.Size = 22
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

With ws.Range("A2:J2")
    .Merge
    .value = "Your business at a glance. Your actions start here."
    .Font.Size = 12
    .HorizontalAlignment = xlLeft
End With

With ws.Range("A3:J3")
    .Merge
    .value = ""
End With

    ' -------------------------
    ' Report Date
    ' -------------------------
    ws.Range("A5").value = "Report Date"
    ws.Range("A5").Font.Bold = True

    With ws.Range("B5")
        .NumberFormat = "yyyy-mm-dd"
        .Interior.Color = RGB(255, 255, 204)
        .Borders.LineStyle = xlContinuous
        .Font.Bold = True
    End With

    ' Separate note box so text is fully visible
With ws.Range("D4:H5")
    .Merge
    .value = "Controls Sales MTD and Receipts MTD only." & vbCrLf & _
             "Does not change current stock or low stock."
        .Font.Italic = True
        .Font.Color = RGB(120, 120, 120)
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(250, 250, 250)
    End With

    CreateNavButton ws, "Refresh Dashboard", "RefreshDashboardV32", 4, 760, RGB(68, 114, 196), 160, 28

    ' -------------------------
    ' Section Titles
    ' -------------------------
    FormatSectionTitle ws, "A8:I8", "KEY METRICS"
    FormatSectionTitle ws, "A19:D19", "TODAY'S ACTIONS"
    FormatSectionTitle ws, "F19:J19", "LOW STOCK DETAILS"
    FormatSectionTitle ws, "A31:J31", "QUICK NAVIGATION"

    ' -------------------------
    ' Metric cards
    ' -------------------------
    DrawMetricCard ws, "Sales MTD", "A9:C9", "A10:C12", RGB(242, 242, 242)
    DrawMetricCard ws, "Receipts MTD", "D9:F9", "D10:F12", RGB(242, 242, 242)
    DrawMetricCard ws, "Customer A/R", "G9:I9", "G10:I12", RGB(255, 242, 204)

    DrawMetricCard ws, "Supplier A/P", "A14:C14", "A15:C17", RGB(255, 242, 204)
    DrawMetricCard ws, "Inventory Units", "D14:F14", "D15:F17", RGB(226, 239, 218)
    DrawMetricCard ws, "Low Stock SKU", "G14:I14", "G15:I17", RGB(255, 199, 206)

    ' -------------------------
    ' TODAY'S ACTIONS
    ' -------------------------
    ws.Range("A20").value = "Low Stock Items"
    ws.Range("A21").value = "Open Purchase Orders"
    ws.Range("A22").value = "Unpaid Invoices"

    ws.Range("A20:C22").UnMerge
    ws.Range("A20:C20").Merge
    ws.Range("A21:C21").Merge
    ws.Range("A22:C22").Merge

    With ws.Range("A20:C22")
        .Font.Bold = True
        .Borders.LineStyle = xlContinuous
        .WrapText = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("D20:D22")
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 12
        .Interior.Color = RGB(250, 250, 250)
    End With

    ws.Range("A20:C20").Interior.Color = RGB(255, 199, 206)
    ws.Range("A21:C21").Interior.Color = RGB(255, 235, 156)
    ws.Range("A22:C22").Interior.Color = RGB(189, 215, 238)

    CreateNavButton ws, "Review Low Stock", "GoToLowStock", 24, 45, RGB(192, 0, 0), 155, 26

    With ws.Range("A26:D28")
        .Merge
        .value = "Click 'Review Low Stock' to open the most urgent low stock SKU in Stock Health."
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(250, 250, 250)
        .Font.Italic = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
    End With

    ' -------------------------
    ' Low stock detail block
    ' -------------------------
    ws.Range("F20").value = "SKU"
    ws.Range("G20:H20").Merge
    ws.Range("G20").value = "Product Name"
    ws.Range("I20").value = "Current"
    ws.Range("J20").value = "Reorder"

    With ws.Range("F20:J20")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With

    ws.Range("F21:J26").ClearContents
    On Error Resume Next
    ws.Range("G21:H26").UnMerge
    On Error GoTo 0

    For rr = 21 To 26
        ws.Range("G" & rr & ":H" & rr).Merge
    Next rr

    With ws.Range("F21:J26")
        .Borders.LineStyle = xlContinuous
        .Interior.Color = RGB(255, 255, 255)
    End With

    With ws.Range("F27:J28")
        .Merge
        .value = "Tip: Click any low stock row to open that SKU in Stock Health."
        .Font.Italic = True
        .Interior.Color = RGB(242, 242, 242)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ' -------------------------
    ' Navigation buttons
    ' -------------------------
    CreateNavButton ws, "Products", "GoToSheet_Products", 33, 20, RGB(0, 176, 80)
    CreateNavButton ws, "Customers", "GoToSheet_Customers", 33, 160, RGB(0, 176, 80)
    CreateNavButton ws, "Suppliers", "GoToSheet_Suppliers", 33, 300, RGB(0, 176, 80)

    CreateNavButton ws, "Purchase", "GoToSheet_Purchase", 36, 20, RGB(0, 112, 192)
    CreateNavButton ws, "Receiving", "GoToSheet_Receiving", 36, 160, RGB(0, 112, 192)
    CreateNavButton ws, "Sales", "GoToSheet_Sales", 36, 300, RGB(0, 112, 192)
    CreateNavButton ws, "Invoice", "GoToSheet_Invoice", 36, 440, RGB(0, 112, 192)
    CreateNavButton ws, "Payment", "GoToSheet_Payment", 36, 580, RGB(0, 112, 192)

    CreateNavButton ws, "Inventory", "GoToSheet_Inventory", 39, 20, RGB(237, 125, 49)
    CreateNavButton ws, "Forecast", "GoToSheet_Forecast", 39, 160, RGB(237, 125, 49)
    CreateNavButton ws, "Stock Health", "GoToSheet_StockHealth", 39, 300, RGB(237, 125, 49)

    CreateNavButton ws, "Inventory Adjustment", "GoToSheet_Adjustment", 42, 20, RGB(128, 128, 128), 180, 24
    CreateNavButton ws, "Setup", "GoToSheet_Setup", 42, 220, RGB(128, 128, 128), 120, 24

    ' -------------------------
    ' Footer note
    ' -------------------------
    With ws.Range("A45:J47")
        .Merge
        .value = "Note: Report Date changes only the month used for Sales MTD and Receipts MTD. Inventory metrics and Low Stock are based on current product stock."
        .WrapText = True
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(242, 242, 242)
        .Borders.LineStyle = xlContinuous
        .Font.Italic = True
    End With

ws.Activate
ActiveWindow.DisplayGridlines = False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    RefreshDashboardV32

End Sub

Public Sub RefreshDashboardV32()

    Dim ws As Worksheet
    Dim effectiveDate As Date

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Dashboard")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    Application.ScreenUpdating = False

    effectiveDate = GetEffectiveDate(ws)

    ' Card values
    ws.Range("A10").value = ComputeSalesMTD(effectiveDate)
    ws.Range("D10").value = ComputeCustomerReceiptsMTD(effectiveDate)
    ws.Range("G10").value = ComputeCustomerAR()

    ws.Range("A15").value = ComputeSupplierAP()
    ws.Range("D15").value = ComputeInventoryUnits()
    ws.Range("G15").value = ComputeLowStockCount()

    ws.Range("A10").NumberFormat = "$#,##0.00"
    ws.Range("D10").NumberFormat = "$#,##0.00"
    ws.Range("G10").NumberFormat = "$#,##0.00"
    ws.Range("A15").NumberFormat = "$#,##0.00"
    ws.Range("D15").NumberFormat = "#,##0"
    ws.Range("G15").NumberFormat = "#,##0"

    ' Actions
    ws.Range("D20").value = ComputeLowStockCount()
    ws.Range("D21").value = ComputeOpenPOCountDistinct()
    ws.Range("D22").value = ComputeUnpaidInvoiceCount()

    ' Low stock list
    LoadLowStockDetails ws

    Application.ScreenUpdating = True

End Sub

' =========================================================
' Layout Helpers
' =========================================================

Private Sub FormatSectionTitle(ByVal ws As Worksheet, ByVal rngAddress As String, ByVal titleText As String)
    With ws.Range(rngAddress)
        .Merge
        .value = titleText
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub DrawMetricCard(ByVal ws As Worksheet, _
                           ByVal titleText As String, _
                           ByVal titleRangeAddress As String, _
                           ByVal valueRangeAddress As String, _
                           ByVal fillColor As Long)

    With ws.Range(titleRangeAddress)
        .Merge
        .value = titleText
        .Font.Bold = True
        .Font.Size = 11
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = fillColor
        .Borders.LineStyle = xlContinuous
    End With

    With ws.Range(valueRangeAddress)
        .Merge
        .value = ""
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = fillColor
        .Borders.LineStyle = xlContinuous
    End With
End Sub

Private Sub CreateNavButton(ByVal ws As Worksheet, _
                            ByVal btnText As String, _
                            ByVal macroName As String, _
                            ByVal topRow As Long, _
                            ByVal leftPos As Double, _
                            ByVal fillColor As Long, _
                            Optional ByVal btnWidth As Double = 120, _
                            Optional ByVal btnHeight As Double = 24)

    Dim btn As Shape
    Dim topPos As Double

    topPos = ws.Rows(topRow).Top

    Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, leftPos, topPos, btnWidth, btnHeight)

    With btn
        .name = "btn_" & Replace(Replace(btnText, " ", "_"), "/", "_")
        .OnAction = macroName
        .Fill.ForeColor.RGB = fillColor
        .Line.ForeColor.RGB = fillColor
        .TextFrame2.TextRange.Text = btnText
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

' =========================================================
' Effective Date
' =========================================================

Private Function GetEffectiveDate(ByVal ws As Worksheet) As Date
    If IsDate(ws.Range("B5").value) Then
        GetEffectiveDate = CDate(ws.Range("B5").value)
    Else
        GetEffectiveDate = Date
    End If
End Function

' =========================================================
' Core Metrics
' =========================================================

Private Function ComputeSalesMTD(ByVal dt As Date) As Double
    Dim lo As ListObject, i As Long
    Dim idxDate As Long, idxTotal As Long
    Dim saleDate As Variant, totalVal As Variant

    Set lo = FindTableByName("tblSales")
    If lo Is Nothing Then Exit Function
    idxDate = GetColumnIndex(lo, "Sales_Date")
    idxTotal = GetColumnIndex(lo, "Line_Total")
    If idxDate = 0 Or idxTotal = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.DataBodyRange.Rows.Count
        saleDate = lo.DataBodyRange.Cells(i, idxDate).value
        totalVal = lo.DataBodyRange.Cells(i, idxTotal).value
        If IsDate(saleDate) Then
            If Year(CDate(saleDate)) = Year(dt) And Month(CDate(saleDate)) = Month(dt) Then
                If IsNumeric(totalVal) Then ComputeSalesMTD = ComputeSalesMTD + CDbl(totalVal)
            End If
        End If
    Next i
End Function

Private Function ComputeCustomerReceiptsMTD(ByVal dt As Date) As Double
    Dim lo As ListObject, i As Long
    Dim idxDate As Long, idxAmt As Long, idxPartyType As Long
    Dim payDate As Variant, amtVal As Variant, partyType As String

    Set lo = FindTableByName("tblPayment")
    If lo Is Nothing Then Exit Function
    idxDate = GetColumnIndex(lo, "Payment_Date")
    idxAmt = GetColumnIndex(lo, "Amount")
    idxPartyType = GetColumnIndex(lo, "Party_Type")
    If idxDate = 0 Or idxAmt = 0 Or idxPartyType = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.DataBodyRange.Rows.Count
        payDate = lo.DataBodyRange.Cells(i, idxDate).value
        amtVal = lo.DataBodyRange.Cells(i, idxAmt).value
        partyType = UCase(Trim(CStr(lo.DataBodyRange.Cells(i, idxPartyType).value)))

        If IsDate(payDate) Then
            If Year(CDate(payDate)) = Year(dt) And Month(CDate(payDate)) = Month(dt) Then
                If partyType = "CUSTOMER" Then
                    If IsNumeric(amtVal) Then ComputeCustomerReceiptsMTD = ComputeCustomerReceiptsMTD + CDbl(amtVal)
                End If
            End If
        End If
    Next i
End Function

Private Function ComputeCustomerAR() As Double
    Dim lo As ListObject, i As Long, idxBal As Long, balVal As Variant

    Set lo = FindTableByName("tblInvoices")
    If lo Is Nothing Then Exit Function
    idxBal = GetColumnIndex(lo, "Balance_Due")
    If idxBal = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.DataBodyRange.Rows.Count
        balVal = lo.DataBodyRange.Cells(i, idxBal).value
        If IsNumeric(balVal) Then
            If CDbl(balVal) > 0 Then ComputeCustomerAR = ComputeCustomerAR + CDbl(balVal)
        End If
    Next i
End Function

Private Function ComputeSupplierAP() As Double
    Dim lo As ListObject, i As Long, idxBal As Long, balVal As Variant

    Set lo = FindTableByName("tblPurchase")
    If lo Is Nothing Then Exit Function
    idxBal = GetColumnIndex(lo, "Balance_Due")
    If idxBal = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.DataBodyRange.Rows.Count
        balVal = lo.DataBodyRange.Cells(i, idxBal).value
        If IsNumeric(balVal) Then
            If CDbl(balVal) > 0 Then ComputeSupplierAP = ComputeSupplierAP + CDbl(balVal)
        End If
    Next i
End Function

Private Function ComputeInventoryUnits() As Double
    Dim lo As ListObject, i As Long
    Dim idxStock As Long, idxStatus As Long
    Dim stockVal As Variant, statusVal As String

    Set lo = FindTableByName("tblProducts")
    If lo Is Nothing Then Exit Function
    idxStock = GetColumnIndex(lo, "Current_Stock")
    idxStatus = GetColumnIndex(lo, "Active_Status")
    If idxStock = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.DataBodyRange.Rows.Count
        stockVal = lo.DataBodyRange.Cells(i, idxStock).value
        If idxStatus > 0 Then
            statusVal = UCase(Trim(CStr(lo.DataBodyRange.Cells(i, idxStatus).value)))
        Else
            statusVal = "ACTIVE"
        End If

        If statusVal = "ACTIVE" Then
            If IsNumeric(stockVal) Then ComputeInventoryUnits = ComputeInventoryUnits + CDbl(stockVal)
        End If
    Next i
End Function

Private Function ComputeLowStockCount() As Long
    Dim lo As ListObject, i As Long
    Dim idxStock As Long, idxReorder As Long, idxStatus As Long
    Dim stockVal As Double, reorderVal As Double, statusVal As String

    Set lo = FindTableByName("tblProducts")
    If lo Is Nothing Then Exit Function
    idxStock = GetColumnIndex(lo, "Current_Stock")
    idxReorder = GetColumnIndex(lo, "Reorder_Level")
    idxStatus = GetColumnIndex(lo, "Active_Status")
    If idxStock = 0 Or idxReorder = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    For i = 1 To lo.DataBodyRange.Rows.Count
        If idxStatus > 0 Then
            statusVal = UCase(Trim(CStr(lo.DataBodyRange.Cells(i, idxStatus).value)))
        Else
            statusVal = "ACTIVE"
        End If

        If statusVal = "ACTIVE" Then
            stockVal = NzNumber(lo.DataBodyRange.Cells(i, idxStock).value)
            reorderVal = NzNumber(lo.DataBodyRange.Cells(i, idxReorder).value)
            If stockVal <= reorderVal Then ComputeLowStockCount = ComputeLowStockCount + 1
        End If
    Next i
End Function

Private Function ComputeOpenPOCountDistinct() As Long
    Dim lo As ListObject, i As Long
    Dim idxPO As Long, idxRemain As Long, idxLineStatus As Long
    Dim poNo As String, lineStatus As String
    Dim remainVal As Double
    Dim dict As Object

    Set lo = FindTableByName("tblPurchase")
    If lo Is Nothing Then Exit Function
    idxPO = GetColumnIndex(lo, "Purchase_Order_No")
    idxRemain = GetColumnIndex(lo, "Remaining_Qty")
    idxLineStatus = GetColumnIndex(lo, "Line_Status")
    If idxPO = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To lo.DataBodyRange.Rows.Count
        poNo = Trim(CStr(lo.DataBodyRange.Cells(i, idxPO).value))
        If poNo <> "" Then
            If idxRemain > 0 Then remainVal = NzNumber(lo.DataBodyRange.Cells(i, idxRemain).value) Else remainVal = 0
            If idxLineStatus > 0 Then lineStatus = UCase(Trim(CStr(lo.DataBodyRange.Cells(i, idxLineStatus).value))) Else lineStatus = ""

            If remainVal > 0 Or lineStatus = "OPEN" Or lineStatus = "PARTIAL" Then
                If Not dict.Exists(poNo) Then dict.Add poNo, True
            End If
        End If
    Next i

    ComputeOpenPOCountDistinct = dict.Count
End Function

Private Function ComputeUnpaidInvoiceCount() As Long
    Dim lo As ListObject, i As Long
    Dim idxInvNo As Long, idxBal As Long, idxStatus As Long
    Dim balVal As Double, statusVal As String, invNo As String
    Dim dict As Object

    Set lo = FindTableByName("tblInvoices")
    If lo Is Nothing Then Exit Function
    idxInvNo = GetColumnIndex(lo, "Invoice_No")
    idxBal = GetColumnIndex(lo, "Balance_Due")
    idxStatus = GetColumnIndex(lo, "Invoice_Status")
    If idxInvNo = 0 Or idxBal = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Set dict = CreateObject("Scripting.Dictionary")

    For i = 1 To lo.DataBodyRange.Rows.Count
        invNo = Trim(CStr(lo.DataBodyRange.Cells(i, idxInvNo).value))
        balVal = NzNumber(lo.DataBodyRange.Cells(i, idxBal).value)
        If idxStatus > 0 Then statusVal = UCase(Trim(CStr(lo.DataBodyRange.Cells(i, idxStatus).value))) Else statusVal = ""

        If invNo <> "" Then
            If balVal > 0 Or statusVal <> "PAID" Then
                If Not dict.Exists(invNo) Then dict.Add invNo, True
            End If
        End If
    Next i

    ComputeUnpaidInvoiceCount = dict.Count
End Function

' =========================================================
' Low Stock Detail Loader + Clickable Rows
' =========================================================

Private Sub LoadLowStockDetails(ByVal ws As Worksheet)

    Dim lo As ListObject
    Dim idxSKU As Long, idxName As Long, idxStock As Long, idxReorder As Long, idxStatus As Long
    Dim i As Long, r As Long, rr As Long
    Dim stockVal As Double, reorderVal As Double
    Dim statusVal As String
    Dim arr() As Variant
    Dim cnt As Long, maxRows As Long
    Dim rowRange As Range

    maxRows = 6

    ws.Range("F21:J26").ClearContents
    DeleteLowStockRowButtons ws

    On Error Resume Next
    ws.Range("G21:H26").UnMerge
    On Error GoTo 0

    For rr = 21 To 26
        ws.Range("G" & rr & ":H" & rr).Merge
    Next rr

    Set lo = FindTableByName("tblProducts")
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub

    idxSKU = GetColumnIndex(lo, "SKU")
    idxName = GetColumnIndex(lo, "Product_Name")
    idxStock = GetColumnIndex(lo, "Current_Stock")
    idxReorder = GetColumnIndex(lo, "Reorder_Level")
    idxStatus = GetColumnIndex(lo, "Active_Status")

    If idxSKU = 0 Or idxName = 0 Or idxStock = 0 Or idxReorder = 0 Then Exit Sub

    ReDim arr(1 To lo.DataBodyRange.Rows.Count, 1 To 5)
    cnt = 0

    For i = 1 To lo.DataBodyRange.Rows.Count
        If idxStatus > 0 Then
            statusVal = UCase(Trim(CStr(lo.DataBodyRange.Cells(i, idxStatus).value)))
        Else
            statusVal = "ACTIVE"
        End If

        If statusVal = "ACTIVE" Then
            stockVal = NzNumber(lo.DataBodyRange.Cells(i, idxStock).value)
            reorderVal = NzNumber(lo.DataBodyRange.Cells(i, idxReorder).value)

            If stockVal <= reorderVal Then
                cnt = cnt + 1
                arr(cnt, 1) = Trim(CStr(lo.DataBodyRange.Cells(i, idxSKU).value))
                arr(cnt, 2) = Trim(CStr(lo.DataBodyRange.Cells(i, idxName).value))
                arr(cnt, 3) = stockVal
                arr(cnt, 4) = reorderVal
                arr(cnt, 5) = stockVal - reorderVal
            End If
        End If
    Next i

    If cnt = 0 Then
        ws.Range("F21").value = "No low stock items."
        Exit Sub
    End If

    SortLowStockArray arr, cnt

    For r = 1 To WorksheetFunction.Min(cnt, maxRows)
        ws.Cells(20 + r, 6).value = arr(r, 1)
        ws.Cells(20 + r, 7).value = arr(r, 2)
        ws.Cells(20 + r, 9).value = arr(r, 3)
        ws.Cells(20 + r, 10).value = arr(r, 4)

        ws.Cells(20 + r, 6).Font.Bold = True
        ws.Cells(20 + r, 6).Font.Color = RGB(0, 102, 204)

        Set rowRange = ws.Range("F" & (20 + r) & ":J" & (20 + r))
        CreateLowStockRowButton ws, rowRange, arr(r, 1), r
    Next r

    ws.Range("I21:J26").NumberFormat = "#,##0"

End Sub

Private Sub CreateLowStockRowButton(ByVal ws As Worksheet, ByVal targetRange As Range, ByVal sku As String, ByVal idx As Long)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, targetRange.Left, targetRange.Top, targetRange.Width, targetRange.Height)

    With shp
        .name = "lsrow_" & idx
        .OnAction = "DashboardLowStockRowClick"
        .AlternativeText = sku
        .Fill.Visible = msoTrue
        .Fill.Transparency = 1#
        .Line.Visible = msoFalse
        .Placement = xlMoveAndSize
    End With
End Sub

Private Sub DeleteLowStockRowButtons(ByVal ws As Worksheet)
    Dim shp As Shape, i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Set shp = ws.Shapes(i)
        If LCase(Left(shp.name, 6)) = "lsrow_" Then shp.Delete
    Next i
End Sub

Public Sub DashboardLowStockRowClick()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim sku As String

    Set ws = ThisWorkbook.Worksheets("Dashboard")
    Set shp = ws.Shapes(Application.Caller)

    sku = Trim(CStr(shp.AlternativeText))
    If sku = "" Then Exit Sub

    OpenStockHealthForSKU sku

End Sub

Private Function GetFirstLowStockSKU() As String

    Dim lo As ListObject, i As Long
    Dim idxSKU As Long, idxStock As Long, idxReorder As Long, idxStatus As Long
    Dim stockVal As Double, reorderVal As Double
    Dim statusVal As String, bestSKU As String
    Dim bestGap As Double, thisGap As Double

    Set lo = FindTableByName("tblProducts")
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    idxSKU = GetColumnIndex(lo, "SKU")
    idxStock = GetColumnIndex(lo, "Current_Stock")
    idxReorder = GetColumnIndex(lo, "Reorder_Level")
    idxStatus = GetColumnIndex(lo, "Active_Status")

    If idxSKU = 0 Or idxStock = 0 Or idxReorder = 0 Then Exit Function

    bestGap = 999999999

    For i = 1 To lo.DataBodyRange.Rows.Count
        If idxStatus > 0 Then
            statusVal = UCase(Trim(CStr(lo.DataBodyRange.Cells(i, idxStatus).value)))
        Else
            statusVal = "ACTIVE"
        End If

        If statusVal = "ACTIVE" Then
            stockVal = NzNumber(lo.DataBodyRange.Cells(i, idxStock).value)
            reorderVal = NzNumber(lo.DataBodyRange.Cells(i, idxReorder).value)

            If stockVal <= reorderVal Then
                thisGap = stockVal - reorderVal
                If thisGap < bestGap Then
                    bestGap = thisGap
                    bestSKU = Trim(CStr(lo.DataBodyRange.Cells(i, idxSKU).value))
                End If
            End If
        End If
    Next i

    GetFirstLowStockSKU = bestSKU

End Function

' =========================================================
' Stock Health open / clear / render
' =========================================================

Private Sub ClearStockHealthUIState(ByVal wsSH As Worksheet)

    On Error Resume Next

    wsSH.Range("B3:B9").ClearContents
    wsSH.Range("B10:B13").ClearContents

    ' clear generated display area only
    wsSH.Range("A15:CM200").ClearContents

    On Error GoTo 0

End Sub

Private Sub TryRunStockHealthViewMacros()

    If TryRunMacroByName("StockHealth_View") Then Exit Sub
    If TryRunMacroByName("StockHealthView") Then Exit Sub
    If TryRunMacroByName("ViewStockHealth") Then Exit Sub
    If TryRunMacroByName("StockHealth_ShowSelected") Then Exit Sub
    If TryRunMacroByName("StockHealth_RenderView") Then Exit Sub
    If TryRunMacroByName("StockHealth_ViewSelected") Then Exit Sub
    If TryRunMacroByName("StockHealth_ViewSelectedSKU") Then Exit Sub

End Sub

Private Function TryRunMacroByName(ByVal macroName As String) As Boolean
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.name & "'!" & macroName
    TryRunMacroByName = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

Private Sub OpenStockHealthForSKU(ByVal sku As String)

    Dim wsSH As Worksheet

    On Error Resume Next
    Set wsSH = ThisWorkbook.Worksheets("Stock_Health_UI")
    On Error GoTo 0

    If wsSH Is Nothing Then
        MsgBox "Stock_Health_UI sheet not found.", vbExclamation, "Low Stock"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Loading Stock Health..."

    ' 1) clear old content first
    ClearStockHealthUIState wsSH

    ' 2) prepare single SKU before showing page
    wsSH.Range("B3").value = sku
    wsSH.Range("B5").value = sku
    wsSH.Range("B6:B9").ClearContents

    ' make sure first SKU in list is this one
    wsSH.Range("B5").value = sku

    Application.EnableEvents = True

    ' 3) open target page
    wsSH.Activate

    ' 4) fill top info first
    On Error Resume Next
    StockHealth_FillProductInfo
    wsSH.Calculate
    On Error GoTo 0

    ' 5) run exact view macro
    On Error Resume Next
    Application.Run "'" & ThisWorkbook.name & "'!StockHealth_View"
    On Error GoTo 0

    ' 6) final color refresh
    On Error Resume Next
    wsSH.Calculate
    ReapplyAllStockHealthDSIColors
    On Error GoTo 0

SafeExit:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub

Public Sub GoToLowStock()

    Dim lowSku As String

    lowSku = GetFirstLowStockSKU()

    If Trim(lowSku) = "" Then
        MsgBox "No low stock SKU found.", vbInformation, "Review Low Stock"
        Exit Sub
    End If

    OpenStockHealthForSKU lowSku

End Sub

Private Sub SortLowStockArray(ByRef arr() As Variant, ByVal cnt As Long)
    Dim i As Long, j As Long, k As Long
    Dim tmp As Variant

    For i = 1 To cnt - 1
        For j = i + 1 To cnt
            If CDbl(arr(j, 5)) < CDbl(arr(i, 5)) Then
                For k = 1 To 5
                    tmp = arr(i, k)
                    arr(i, k) = arr(j, k)
                    arr(j, k) = tmp
                Next k
            End If
        Next j
    Next i
End Sub

' =========================================================
' Table / Column Helpers
' =========================================================

Private Function FindTableByName(ByVal tableName As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If LCase(lo.name) = LCase(tableName) Then
                Set FindTableByName = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

Private Function GetColumnIndex(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    On Error Resume Next
    For Each lc In lo.ListColumns
        If LCase(Trim(lc.name)) = LCase(Trim(colName)) Then
            GetColumnIndex = lc.Index
            Exit Function
        End If
    Next lc
    On Error GoTo 0
End Function

Private Function NzNumber(ByVal v As Variant) As Double
    If IsNumeric(v) Then
        NzNumber = CDbl(v)
    Else
        NzNumber = 0
    End If
End Function

' =========================================================
' Navigation
' =========================================================

Public Sub GoToSheet_Products()
    SafeGoToSheets Array("Products_UI", "Products_DB")
End Sub

Public Sub GoToSheet_Customers()
    SafeGoToSheets Array("Customers_UI", "Customers_DB")
End Sub

Public Sub GoToSheet_Suppliers()
    SafeGoToSheets Array("Suppliers_UI", "Suppliers_DB")
End Sub

Public Sub GoToSheet_Purchase()
    SafeGoToSheets Array("Purchase_UI", "Purchase_DB")
End Sub

Public Sub GoToSheet_Receiving()
    SafeGoToSheets Array("Receiving_UI", "Receiving_DB")
End Sub

Public Sub GoToSheet_Sales()
    SafeGoToSheets Array("Sales_UI", "Sales_DB")
End Sub

Public Sub GoToSheet_Invoice()
    SafeGoToSheets Array("Invoice_UI", "Invoice_DB")
End Sub

Public Sub GoToSheet_Payment()
    SafeGoToSheets Array("Payment_UI", "Payment_DB")
End Sub

Public Sub GoToSheet_Inventory()
    SafeGoToSheets Array("Inventory_UI", "Inventory_Log", "Products_DB")
End Sub

Public Sub GoToSheet_Forecast()
    SafeGoToSheets Array("Forecast_UI", "Forecast_DB")
End Sub

Public Sub GoToSheet_StockHealth()
    SafeGoToSheets Array("Stock_Health_UI", "Stock Health")
End Sub

Public Sub GoToSheet_Adjustment()
    SafeGoToSheets Array("Inventory_Adjustment_UI", "Adjustment_UI", "Adjustment_DB")
End Sub

Public Sub GoToSheet_Setup()
    SafeGoToSheets Array("Setup_UI", "Settings")
End Sub

Private Sub SafeGoToSheets(ByVal sheetNames As Variant)
    Dim i As Long
    For i = LBound(sheetNames) To UBound(sheetNames)
        If WorksheetExists(CStr(sheetNames(i))) Then
            ThisWorkbook.Worksheets(CStr(sheetNames(i))).Activate
            Exit Sub
        End If
    Next i
    MsgBox "Target sheet not found.", vbExclamation, "Navigation"
End Sub

Private Function WorksheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0
End Function

