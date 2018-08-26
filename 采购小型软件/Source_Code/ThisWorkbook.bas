VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    
End Sub

Private Sub Workbook_Open()
    Application.EnableEvents = False
    Call sub_WorkBookInitialization
    
    Call fHideSheet(shtSysConf)
 
    Application.EnableEvents = True
End Sub

Sub sub_WorkBookInitialization()
'    Call fHideSheet(shtVendorMaster)
'    Call fShowSheet(shtPurchaseODRaw)
'    Call fHideSheet(shtVendorPrice)
'    Call fHideSheet(shtPurchaseODByProduct)
'    Call fHideSheet(shtPurchaseODByVendor)
'    Call fHideSheet(shtProdMasterExtracted)

    Call fClearConditionFormatAndAdd(shtVendorMaster, Vendor.VendorName, True)
    Call fClearConditionFormatAndAdd(shtPurchaseODRaw, PODRaw.ProdName, True)
    Call fClearConditionFormatAndAdd(shtVendorPrice, Array(VendorPrice.ProdName, VendorPrice.VendorName), True)
    
    'Call fClearConditionFormatAndAdd(shtPurchaseODByProduct, Array(PODByProduct.ProdName, PODByProduct.VendorName), True)
    
    Call fSetValidationForNumberForSheetColumns(shtPurchaseODRaw, PODRaw.PurchaseQty, 0, 999999)
    Call fSetValidationForNumberForSheetColumns(shtVendorPrice, VendorPrice.Price, 0, 999999)
    
    Call fExtractProductFromPriceConfigSheet
    
    Dim sTargetCol As String
    Dim lMaxRow As Long
    Dim sValidationListAddr As String

    sTargetCol = fNum2Letter(PODRaw.ProdName)
    lMaxRow = fGetValidMaxRow(shtPurchaseODRaw) + 100000
    If lMaxRow > Rows.Count Then lMaxRow = 100000
    sValidationListAddr = "=" & shtProdMasterExtracted.Columns("A").Address(external:=True)
    Call fSetValidationListForRange(fGetRangeByStartEndPos(shtPurchaseODRaw, shtPurchaseODRaw.DataFromRow, PODRaw.ProdName, lMaxRow, PODRaw.ProdName) _
                                  , sValidationListAddr)
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    subRefreshRibbon
End Sub
