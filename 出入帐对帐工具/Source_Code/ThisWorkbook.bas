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

Private Sub Workbook_Open()
  
    Call sub_WorkBookInitialization
 
End Sub

Sub sub_WorkBookInitialization()
    'shtBillIn
    '=========================================================
    Call fDeleteAllConditionFormatFromSheet(shtBillIn)
    Call fSetConditionFormatForOddEvenLine(shtBillIn, , , , Array(BillIn.FromCompany), True)
    Call fSetConditionFormatForBorders(shtBillIn, , , , Array(BillIn.FromCompany), True)
    
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBillIn, 2, BillIn.Amount, Rows.Count, BillIn.Amount), 0, 999999999)
    
    'shtBillOut
    '=========================================================
    Call fDeleteAllConditionFormatFromSheet(shtBillOut)
    Call fSetConditionFormatForOddEvenLine(shtBillOut, , , , Array(BillOut.toCompany), True)
    Call fSetConditionFormatForBorders(shtBillOut, , , , Array(BillOut.toCompany), True)
    
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBillOut, 2, BillOut.Amount, Rows.Count, BillOut.Amount), 0, 999999999)
    
    'shtBusinessDetails
    '=========================================================
    Call fDeleteAllConditionFormatFromSheet(shtBusinessDetails)
   ' Call fSetConditionFormatForOddEvenLine(shtBusinessDetails, , shtBusinessDetails.DataStartRow, , Array(BuzDetail.VendorName, BuzDetail.PLATFORM), True)
    Call fSetConditionFormatForBorders(shtBusinessDetails, , shtBusinessDetails.DataStartRow, , Array(BuzDetail.VendorName, BuzDetail.PLATFORM), True)
    
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Point_Qty, Rows.Count, BuzDetail.Point_Qty), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Point_Price, Rows.Count, BuzDetail.Point_Price), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Point_CurrDayPrice, Rows.Count, BuzDetail.Point_CurrDayPrice), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Point_DaysNum, Rows.Count, BuzDetail.Point_DaysNum), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Point_Amt, Rows.Count, BuzDetail.Point_Amt), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.DownLoad_Qty, Rows.Count, BuzDetail.DownLoad_Qty), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.DownLoad_Price, Rows.Count, BuzDetail.DownLoad_Price), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.DownLoad_Amt, Rows.Count, BuzDetail.DownLoad_Amt), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Credit_Qty, Rows.Count, BuzDetail.Credit_Qty), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Credit_Price, Rows.Count, BuzDetail.Credit_Price), 0, 999999999)
    Call fSetValidationForNumberRange(fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.Credit_Amt, Rows.Count, BuzDetail.Credit_Amt), 0, 999999999)
End Sub
   
 
