Attribute VB_Name = "modInventoryHistoryPrint"
Option Explicit

Private Const WS_HISTORY_UI As String = "Inventory_History_UI"
Private Const WS_HISTORY_PRINT As String = "Inventory_History_Print"

Private Const UI_CELL_SKU As String = "B4"
Private Const UI_CELL_DATE_FROM As String = "B5"
Private Const UI_CELL_DATE_TO As String = "B6"
Private Const UI_CELL_TYPE As String = "B7"

Private Const UI_SUM_IN As String = "H4"
Private Const UI_SUM_OUT As String = "H5"
Private Const UI_SUM_NET As String = "H6"
Private Const UI_SUM_CLOSE As String = "H7"

Private Const UI_ROW_HEADER As Long = 10
Private Const UI_ROW_DATA_START As Long = 11
Private Const UI_COL_LAST As Long = 10   ' A:J

Public Sub InventoryHistory_Preview()
    If BuildInventoryHistoryPrintPage() Then
        ThisWorkbook.Worksheets(WS_HISTORY_PRINT).PrintPreview
    End If
End Sub

Public Sub InventoryHistory_Print()
    If BuildInventoryHistoryPrintPage() Then
        ThisWorkbook.Worksheets(WS_HISTORY_PRINT).PrintOut
    End If
End Sub

Public Function BuildInventoryHistoryPrintPage() As Boolean

    Dim wsUI As Worksheet
    Dim wsP As Worksheet
    Dim lastUIRow As Long
    Dim srcRow As Long
    Dim outRow As Long
    Dim hasData As Boolean

    On Error GoTo ErrHandler

    Set wsUI = ThisWorkbook.Worksheets(WS_HISTORY_UI)
    Set wsP = ThisWorkbook.Worksheets(WS_HISTORY_PRINT)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    PrepareInventoryHistoryPrintSheet wsP

    ' ---------------------------------
    ' Header / report info
    ' ---------------------------------
    InsertCompanyLogo wsP
    With wsP.Range("A1:J1")
        .Merge
        .value = "Inventory History Report"
        .Font.Bold = True
        .Font.Size = 18
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    wsP.Range("A3").value = "SKU"
    wsP.Range("B3").value = NzTxt(wsUI.Range(UI_CELL_SKU).value, "ALL")

    wsP.Range("D3").value = "Date From"
    wsP.Range("E3").value = wsUI.Range(UI_CELL_DATE_FROM).value

    wsP.Range("G3").value = "Date To"
    wsP.Range("H3").value = wsUI.Range(UI_CELL_DATE_TO).value

    wsP.Range("A4").value = "Movement Type"
    wsP.Range("B4").value = NzTxt(wsUI.Range(UI_CELL_TYPE).value, "ALL")

    wsP.Range("D4").value = "Print Time"
    wsP.Range("E4").value = Now

    wsP.Range("A3:H4").Font.Bold = True
    wsP.Range("E3:H3").NumberFormat = "yyyy-mm-dd"
    wsP.Range("E4").NumberFormat = "yyyy-mm-dd hh:mm"

    ' ---------------------------------
    ' Summary
    ' ---------------------------------
    With wsP.Range("A6:J6")
        .Merge
        .value = "Summary"
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
    End With

    wsP.Range("A7").value = "Total In"
    wsP.Range("B7").value = wsUI.Range(UI_SUM_IN).value

    wsP.Range("D7").value = "Total Out"
    wsP.Range("E7").value = wsUI.Range(UI_SUM_OUT).value

    wsP.Range("G7").value = "Net Change"
    wsP.Range("H7").value = wsUI.Range(UI_SUM_NET).value

    wsP.Range("A8").value = "Closing Balance"
    wsP.Range("B8").value = wsUI.Range(UI_SUM_CLOSE).value

    wsP.Range("A7:H8").Font.Bold = True
    wsP.Range("B7:H8").NumberFormat = "#,##0.00"

    ' ---------------------------------
    ' Table header
    ' ---------------------------------
    wsP.Range("A10").value = "Log Date"
    wsP.Range("B10").value = "Movement Type"
    wsP.Range("C10").value = "Reference No"
    wsP.Range("D10").value = "SKU"
    wsP.Range("E10").value = "Product Name"
    wsP.Range("F10").value = "Qty In"
    wsP.Range("G10").value = "Qty Out"
    wsP.Range("H10").value = "Net Change"
    wsP.Range("I10").value = "Running Balance"
    wsP.Range("J10").value = "Notes"

    With wsP.Range("A10:J10")
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
        .Borders.LineStyle = xlContinuous
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' ---------------------------------
    ' Copy current result set from UI
    ' ---------------------------------
    lastUIRow = wsUI.Cells(wsUI.Rows.Count, 1).End(xlUp).Row
    If lastUIRow < UI_ROW_DATA_START Then
        MsgBox "No history result available to print. Please run a search first.", vbInformation
        GoTo SafeExit
    End If

    outRow = 11
    hasData = False

    For srcRow = UI_ROW_DATA_START To lastUIRow
        If Trim$(CStr(wsUI.Cells(srcRow, 1).value)) <> "" Then
            wsP.Range("A" & outRow & ":J" & outRow).value = wsUI.Range("A" & srcRow & ":J" & srcRow).value
            outRow = outRow + 1
            hasData = True
        End If
    Next srcRow

    If Not hasData Then
        MsgBox "No history result available to print. Please run a search first.", vbInformation
        GoTo SafeExit
    End If

    ' ---------------------------------
    ' Format copied data
    ' ---------------------------------
    With wsP.Range("A11:J" & outRow - 1)
        .Borders.LineStyle = xlContinuous
        .VerticalAlignment = xlCenter
    End With

    wsP.Range("A11:A" & outRow - 1).NumberFormat = "yyyy-mm-dd"
    wsP.Range("F11:I" & outRow - 1).NumberFormat = "#,##0.00"

    ' ---------------------------------
    ' Columns / wrapping
    ' ---------------------------------
    wsP.Columns("A").ColumnWidth = 12
    wsP.Columns("B").ColumnWidth = 16
    wsP.Columns("C").ColumnWidth = 14
    wsP.Columns("D").ColumnWidth = 14
    wsP.Columns("E").ColumnWidth = 24
    wsP.Columns("F").ColumnWidth = 10
    wsP.Columns("G").ColumnWidth = 10
    wsP.Columns("H").ColumnWidth = 12
    wsP.Columns("I").ColumnWidth = 14
    wsP.Columns("J").ColumnWidth = 24

    wsP.Range("J11:J" & outRow - 1).WrapText = True

    ' ---------------------------------
    ' Page setup
    ' ---------------------------------
    With wsP.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .PrintTitleRows = "$10:$10"
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.5)
        .BottomMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True
    End With

    BuildInventoryHistoryPrintPage = True

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Function

ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Inventory History Print"
    Resume SafeExit

End Function

Private Sub PrepareInventoryHistoryPrintSheet(ByVal ws As Worksheet)

    ws.Cells.Clear
    ws.Cells.Font.name = "Calibri"
    ws.Cells.Font.Size = 10

    On Error Resume Next
    ws.DrawingObjects.Delete
    On Error GoTo 0

    ws.Rows.RowHeight = 20
End Sub

Private Function NzTxt(ByVal v As Variant, Optional ByVal fallback As String = "") As String
    If IsError(v) Then
        NzTxt = fallback
    ElseIf Trim$(CStr(v)) = "" Then
        NzTxt = fallback
    Else
        NzTxt = Trim$(CStr(v))
    End If
End Function


Private Sub InsertCompanyLogo(ByVal wsTarget As Worksheet)

    Dim wsSetup As Worksheet
    Dim shp As Shape
    Dim newShp As Shape

    On Error Resume Next
    Set wsSetup = ThisWorkbook.Worksheets("Setup_UI")
    Set shp = wsSetup.Shapes("CompanyLogo")
    On Error GoTo 0

    If shp Is Nothing Then Exit Sub

    ' 删除旧的打印页Logo
    Dim s As Shape
    For Each s In wsTarget.Shapes
        If s.name Like "PrintLogo*" Then
            s.Delete
        End If
    Next s

    ' 复制Logo
    shp.Copy
    wsTarget.Paste

    Set newShp = wsTarget.Shapes(wsTarget.Shapes.Count)
    newShp.name = "PrintLogo"

    ' 位置（左上角）
    With newShp
        .Top = wsTarget.Range("A1").Top
        .Left = wsTarget.Range("A1").Left
        .LockAspectRatio = msoTrue
        .Height = 40   ' 可调：40~60
    End With

End Sub
