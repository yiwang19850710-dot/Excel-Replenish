Attribute VB_Name = "modImportDispatcher"
Option Explicit

Private Const WS_IMPORT_UI As String = "Import_UI"
Private Const CELL_IMPORT_TYPE As String = "B5"

Private Const IMPORT_TYPE_ORDER As String = "Order Import"
Private Const IMPORT_TYPE_SHOPIFY As String = "Shopify Order Import"
Private Const IMPORT_TYPE_AMAZON As String = "Amazon Order Import"

Public Sub ImportUI_DownloadTemplate()

    Select Case GetCurrentImportType()
        Case IMPORT_TYPE_ORDER
            ImportSales_DownloadTemplate

        Case IMPORT_TYPE_SHOPIFY
            MsgBox "Shopify Order Import does not require template download." & vbCrLf & _
                   "Please export the Shopify order CSV directly from Shopify.", vbInformation

        Case IMPORT_TYPE_AMAZON
            MsgBox "Amazon Order Import is not ready yet.", vbInformation

        Case Else
            MsgBox "Please select a valid Import Type in B5.", vbExclamation
    End Select

End Sub

Public Sub ImportUI_SelectFile()

    Select Case GetCurrentImportType()
        Case IMPORT_TYPE_ORDER
            ImportSales_SelectFile

        Case IMPORT_TYPE_SHOPIFY
            ShopifyImport_SelectFile

        Case IMPORT_TYPE_AMAZON
            MsgBox "Amazon Order Import is not ready yet.", vbInformation

        Case Else
            MsgBox "Please select a valid Import Type in B5.", vbExclamation
    End Select

End Sub

Public Sub ImportUI_Validate()

    Select Case GetCurrentImportType()
        Case IMPORT_TYPE_ORDER
            ImportSales_Validate

        Case IMPORT_TYPE_SHOPIFY
            ShopifyImport_Validate

        Case IMPORT_TYPE_AMAZON
            MsgBox "Amazon Order Import is not ready yet.", vbInformation

        Case Else
            MsgBox "Please select a valid Import Type in B5.", vbExclamation
    End Select

End Sub

Public Sub ImportUI_RunImport()

    Select Case GetCurrentImportType()
        Case IMPORT_TYPE_ORDER
            ImportSales_Run

        Case IMPORT_TYPE_SHOPIFY
            ShopifyImport_Run

        Case IMPORT_TYPE_AMAZON
            MsgBox "Amazon Order Import is not ready yet.", vbInformation

        Case Else
            MsgBox "Please select a valid Import Type in B5.", vbExclamation
    End Select

End Sub

Public Sub ImportUI_Clear()

    Select Case GetCurrentImportType()
        Case IMPORT_TYPE_ORDER
            ImportSales_Clear

        Case IMPORT_TYPE_SHOPIFY
            ShopifyImport_Clear

        Case IMPORT_TYPE_AMAZON
            ClearImportUIOnly
            MsgBox "Amazon Order Import is not ready yet.", vbInformation

        Case Else
            MsgBox "Please select a valid Import Type in B5.", vbExclamation
    End Select

End Sub

Public Function GetCurrentImportType() As String
    GetCurrentImportType = Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(CELL_IMPORT_TYPE).value))
End Function

Private Sub ClearImportUIOnly()

    With ThisWorkbook.Worksheets(WS_IMPORT_UI)
        .Range("B6").ClearContents
        .Range("B7").ClearContents
        .Range("B8").ClearContents
        .Range("B7").Interior.Pattern = xlNone
        .Range("B7").Font.Underline = xlUnderlineStyleNone
        .Range("B7").Font.ColorIndex = xlAutomatic
    End With

End Sub

