Attribute VB_Name = "modForecastOverride"
Option Explicit

' =========================================================
' Forecast Override
' =========================================================

Public Sub SaveForecastOverride()
    Dim ws As Worksheet
    Dim viewStartDate As Date
    Dim viewDays As Long
    Dim lastRow As Long
    Dim r As Long
    Dim d As Long
    Dim colNum As Long

    Dim sku As String
    Dim targetDate As Date
    Dim gridValue As Variant

    Dim savedCount As Long

    Set ws = GetWorksheet(WS_FORECAST_UI)
    If ws Is Nothing Then Exit Sub

    viewStartDate = GetUIDate(ws, CELL_VIEW_START_DATE, 1)
    If viewStartDate = 0 Then
        MsgBox "Please enter a valid View Start Date.", vbExclamation
        Exit Sub
    End If

    viewDays = CLng(GetUIDouble(ws, CELL_VIEW_DAYS, 1, DEFAULT_VIEW_DAYS))
    If viewDays <= 0 Then
        MsgBox "View Days must be greater than 0.", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, COL_VIEW_SKU).End(xlUp).Row
    If lastRow < ROW_VIEW_DATA_START Then
        MsgBox "No forecast rows found in Daily View.", vbExclamation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo SafeExit

    savedCount = 0

    For r = ROW_VIEW_DATA_START To lastRow
        sku = Trim$(CStr(ws.Cells(r, COL_VIEW_SKU).value))
        If sku <> "" Then

            For d = 0 To viewDays - 1
                colNum = COL_VIEW_DATE_START + d
                targetDate = viewStartDate + d
                gridValue = ws.Cells(r, colNum).value

                If SaveSingleOverrideValue(sku, targetDate, gridValue) Then
                    savedCount = savedCount + 1
                End If
            Next d

        End If
    Next r

    MsgBox savedCount & " override value(s) saved.", vbInformation

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function SaveSingleOverrideValue(ByVal sku As String, ByVal targetDate As Date, ByVal gridValue As Variant) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow

    Dim skuCol As Long
    Dim dateCol As Long
    Dim sysCol As Long
    Dim overCol As Long
    Dim finalCol As Long
    Dim updatedCol As Long

    Dim rowSKU As String
    Dim rowDate As Variant
    Dim systemQty As Double
    Dim newOverride As Variant
    Dim currentOverride As Variant
    Dim currentFinal As Variant
    Dim changed As Boolean

    Set lo = GetTable(WS_FORECAST_DB, TBL_FORECAST)
    If lo Is Nothing Then Exit Function

    skuCol = GetHeaderColumn(lo, "SKU")
    dateCol = GetHeaderColumn(lo, "Forecast_Date")
    sysCol = GetHeaderColumn(lo, "System_Forecast_Qty")
    overCol = GetHeaderColumn(lo, "Override_Forecast_Qty")
    finalCol = GetHeaderColumn(lo, "Final_Forecast_Qty")
    updatedCol = GetHeaderColumn(lo, "Updated_At")

    If skuCol = 0 Or dateCol = 0 Or sysCol = 0 Or overCol = 0 Or finalCol = 0 Then Exit Function

    For Each lr In lo.ListRows
        rowSKU = Trim$(CStr(lr.Range.Cells(1, skuCol).value))
        If StrComp(rowSKU, Trim$(sku), vbTextCompare) = 0 Then

            rowDate = lr.Range.Cells(1, dateCol).value
            If IsDate(rowDate) Then
                If CLng(CDate(rowDate)) = CLng(targetDate) Then

                    systemQty = NzDbl(lr.Range.Cells(1, sysCol).value, 0)
                    currentOverride = lr.Range.Cells(1, overCol).value
                    currentFinal = lr.Range.Cells(1, finalCol).value

                    ' 如果格子是空，表示清除 override
                    If Trim$(CStr(gridValue)) = "" Then
                        newOverride = ""
                    ElseIf IsNumeric(gridValue) Then
                        newOverride = CDbl(gridValue)
                    Else
                        ' 非数字则忽略
                        Exit Function
                    End If

                    changed = False

                    If Trim$(CStr(newOverride)) = "" Then
                        ' 清空 override -> final 恢复 system
                        If Trim$(CStr(currentOverride)) <> "" Or NzDbl(currentFinal, -999999) <> systemQty Then
                            lr.Range.Cells(1, overCol).value = ""
                            lr.Range.Cells(1, finalCol).value = systemQty
                            changed = True
                        End If
                    Else
                        ' 保存 override
                        If NzDbl(currentOverride, -999999) <> CDbl(newOverride) Or NzDbl(currentFinal, -999999) <> CDbl(newOverride) Then
                            lr.Range.Cells(1, overCol).value = CDbl(newOverride)
                            lr.Range.Cells(1, finalCol).value = CDbl(newOverride)
                            changed = True
                        End If
                    End If

                    If changed Then
                        If updatedCol > 0 Then
                            lr.Range.Cells(1, updatedCol).value = Now
                        End If
                        SaveSingleOverrideValue = True
                    End If

                    Exit Function
                End If
            End If
        End If
    Next lr
End Function

Public Sub ClearForecastOverride()
    Dim ws As Worksheet
    Dim viewStartDate As Date
    Dim viewDays As Long
    Dim lastRow As Long
    Dim r As Long
    Dim d As Long
    Dim colNum As Long

    Dim sku As String
    Dim targetDate As Date
    Dim clearedCount As Long

    Set ws = GetWorksheet(WS_FORECAST_UI)
    If ws Is Nothing Then Exit Sub

    viewStartDate = GetUIDate(ws, CELL_VIEW_START_DATE, 1)
    If viewStartDate = 0 Then
        MsgBox "Please enter a valid View Start Date.", vbExclamation
        Exit Sub
    End If

    viewDays = CLng(GetUIDouble(ws, CELL_VIEW_DAYS, 1, DEFAULT_VIEW_DAYS))
    If viewDays <= 0 Then
        MsgBox "View Days must be greater than 0.", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, COL_VIEW_SKU).End(xlUp).Row
    If lastRow < ROW_VIEW_DATA_START Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error GoTo SafeExit

    clearedCount = 0

    For r = ROW_VIEW_DATA_START To lastRow
        sku = Trim$(CStr(ws.Cells(r, COL_VIEW_SKU).value))
        If sku <> "" Then

            For d = 0 To viewDays - 1
                colNum = COL_VIEW_DATE_START + d
                targetDate = viewStartDate + d

                If ClearSingleOverrideValue(sku, targetDate) Then
                    clearedCount = clearedCount + 1
                End If
            Next d

        End If
    Next r

    ' 清完后重新加载视图，让 UI 看到 final 已恢复
    LoadForecastView

    MsgBox clearedCount & " override value(s) cleared.", vbInformation

SafeExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function ClearSingleOverrideValue(ByVal sku As String, ByVal targetDate As Date) As Boolean
    Dim lo As ListObject
    Dim lr As ListRow

    Dim skuCol As Long
    Dim dateCol As Long
    Dim sysCol As Long
    Dim overCol As Long
    Dim finalCol As Long
    Dim updatedCol As Long

    Dim rowSKU As String
    Dim rowDate As Variant
    Dim systemQty As Double

    Set lo = GetTable(WS_FORECAST_DB, TBL_FORECAST)
    If lo Is Nothing Then Exit Function

    skuCol = GetHeaderColumn(lo, "SKU")
    dateCol = GetHeaderColumn(lo, "Forecast_Date")
    sysCol = GetHeaderColumn(lo, "System_Forecast_Qty")
    overCol = GetHeaderColumn(lo, "Override_Forecast_Qty")
    finalCol = GetHeaderColumn(lo, "Final_Forecast_Qty")
    updatedCol = GetHeaderColumn(lo, "Updated_At")

    If skuCol = 0 Or dateCol = 0 Or sysCol = 0 Or overCol = 0 Or finalCol = 0 Then Exit Function

    For Each lr In lo.ListRows
        rowSKU = Trim$(CStr(lr.Range.Cells(1, skuCol).value))
        If StrComp(rowSKU, Trim$(sku), vbTextCompare) = 0 Then

            rowDate = lr.Range.Cells(1, dateCol).value
            If IsDate(rowDate) Then
                If CLng(CDate(rowDate)) = CLng(targetDate) Then

                    If Trim$(CStr(lr.Range.Cells(1, overCol).value)) <> "" Then
                        systemQty = NzDbl(lr.Range.Cells(1, sysCol).value, 0)

                        lr.Range.Cells(1, overCol).value = ""
                        lr.Range.Cells(1, finalCol).value = systemQty

                        If updatedCol > 0 Then
                            lr.Range.Cells(1, updatedCol).value = Now
                        End If

                        ClearSingleOverrideValue = True
                    End If

                    Exit Function
                End If
            End If
        End If
    Next lr
End Function

Public Sub RecalcDailyViewFinals()
    ' 当前 V1 里，这个过程可作为占位。
    ' 如果以后你想在用户修改单元格后实时重算 UI，可在这里扩展。
    LoadForecastView
End Sub
