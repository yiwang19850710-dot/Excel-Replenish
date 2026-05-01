Attribute VB_Name = "modProductImportDispatcher"
Option Explicit

Private Const WS_IMPORT_UI As String = "Import_UI"
Private Const CELL_PRODUCT_IMPORT_TYPE As String = "B12"

Private Const PRODUCT_IMPORT_STANDARD As String = "Product Import"
Private Const PRODUCT_IMPORT_SHOPIFY As String = "Shopify Product Import"

Public Sub ProductUI_DownloadTemplate()

    Select Case GetCurrentProductImportType()
        Case PRODUCT_IMPORT_STANDARD
            ProductImport_DownloadTemplate

        Case PRODUCT_IMPORT_SHOPIFY
            MsgBox "Shopify Product Import does not require template download." & vbCrLf & _
                   "Please export the Shopify product CSV directly from Shopify.", vbInformation

        Case Else
            MsgBox "Please select a valid Product Import Type in B12.", vbExclamation
    End Select

End Sub

Public Sub ProductUI_SelectFile()

    Select Case GetCurrentProductImportType()
        Case PRODUCT_IMPORT_STANDARD
            ProductImport_SelectFile

        Case PRODUCT_IMPORT_SHOPIFY
            ShopifyProductImport_SelectFile

        Case Else
            MsgBox "Please select a valid Product Import Type in B12.", vbExclamation
    End Select

End Sub

Public Sub ProductUI_Validate()

    Select Case GetCurrentProductImportType()
        Case PRODUCT_IMPORT_STANDARD
            ProductImport_Validate

        Case PRODUCT_IMPORT_SHOPIFY
            ShopifyProductImport_Validate

        Case Else
            MsgBox "Please select a valid Product Import Type in B12.", vbExclamation
    End Select

End Sub

Public Sub ProductUI_RunImport()

    Select Case GetCurrentProductImportType()
        Case PRODUCT_IMPORT_STANDARD
            ProductImport_Run

        Case PRODUCT_IMPORT_SHOPIFY
            ShopifyProductImport_Run

        Case Else
            MsgBox "Please select a valid Product Import Type in B12.", vbExclamation
    End Select

End Sub

Public Sub ProductUI_Clear()

    Select Case GetCurrentProductImportType()
        Case PRODUCT_IMPORT_STANDARD
            ProductImport_Clear

        Case PRODUCT_IMPORT_SHOPIFY
            ShopifyProductImport_Clear

        Case Else
            MsgBox "Please select a valid Product Import Type in B12.", vbExclamation
    End Select

End Sub

Public Function GetCurrentProductImportType() As String
    GetCurrentProductImportType = Trim$(CStr(ThisWorkbook.Worksheets(WS_IMPORT_UI).Range(CELL_PRODUCT_IMPORT_TYPE).value))
End Function

