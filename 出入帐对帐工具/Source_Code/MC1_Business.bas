Attribute VB_Name = "MC1_Business"
Option Explicit
Option Base 1

Sub subMain_CalculateBillInOut()
    On Error GoTo error_handling
    
    Call fInitialization

    Dim dblBillInSumm As Double
    Dim dblBillOutSumm As Double
    
    dblBillInSumm = fSumSheetColumn(shtBillIn, BillIn.Amount)
    dblBillOutSumm = fSumSheetColumn(shtBillOut, BillOut.Amount)
    
    shtSummaryAmount.Range("rgSummaryResult").Cells(1, 1) = dblBillInSumm
    shtSummaryAmount.Range("rgSummaryResult").Cells(2, 1) = dblBillOutSumm
    shtSummaryAmount.Range("rgSummaryResult").Cells(3, 1) = dblBillInSumm - dblBillOutSumm
    
    Application.ScreenUpdating = True
    Call fShowAndActiveSheet(shtSummaryAmount)
    Call fGotoCell(shtSummaryAmount.Range("rgSummaryResult"))
    
    fMsgBox "计算完成，结果在表：[" & shtSummaryAmount.Name & "] 中，请检查！", vbInformation
error_handling:
    If gErrNum <> 0 Then GoTo reset_excel_options
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
reset_excel_options:
    Err.Clear
    fClearRefVariables
    fEnableExcelOptionsAll
End Sub

Function fSumSheetColumn(sht As Worksheet, iCol As Long, Optional alDataStartFromRow As Long = 2) As Double
    Dim dblSum As Double
    Dim lMaxRow As Long
    Dim rg As Range
    
    lMaxRow = fGetValidMaxRow(sht)
    
    dblSum = 0
    
    If lMaxRow >= alDataStartFromRow Then
        Set rg = fGetRangeByStartEndPos(sht, alDataStartFromRow, iCol, lMaxRow, iCol)
        dblSum = WorksheetFunction.Sum(rg)
        Set rg = Nothing
    End If
    
    fSumSheetColumn = dblSum
End Function

Sub subMain_SummarizeBusinssDetail()
    On Error GoTo error_handling
    
    Call fInitialization
    
    Dim arrDetails()
    Call fCopyReadWholeSheetData2Array(shtBusinessDetails, arrDetails, , shtBusinessDetails.DataStartRow, BuzDetail.[_last])
    
    If UBound(arrDetails, 1) < LBound(arrDetails, 1) Then
        fErr "明细表中没有数据!"
    End If
    
    Dim lEachRow As Long
    Dim dblSum_Point As Double
    Dim dblSum_DownLoad As Double
    Dim dblSum_Credit As Double
    
    dblSum_Point = 0
    dblSum_DownLoad = 0
    dblSum_Credit = 0
    
    For lEachRow = LBound(arrDetails, 1) To UBound(arrDetails, 1)
        If arrDetails(lEachRow, BuzDetail.Point_Qty) <= 0 Then arrDetails(lEachRow, BuzDetail.Point_Qty) = 0
        If arrDetails(lEachRow, BuzDetail.DownLoad_Qty) <= 0 Then arrDetails(lEachRow, BuzDetail.DownLoad_Qty) = 0
        If arrDetails(lEachRow, BuzDetail.Credit_Qty) <= 0 Then arrDetails(lEachRow, BuzDetail.Credit_Qty) = 0
        If arrDetails(lEachRow, BuzDetail.Point_Price) <= 0 Then arrDetails(lEachRow, BuzDetail.Point_Price) = 0
        If arrDetails(lEachRow, BuzDetail.DownLoad_Price) <= 0 Then arrDetails(lEachRow, BuzDetail.DownLoad_Price) = 0
        If arrDetails(lEachRow, BuzDetail.Credit_Price) <= 0 Then arrDetails(lEachRow, BuzDetail.Credit_Price) = 0
        
        arrDetails(lEachRow, BuzDetail.Point_CurrDayPrice) = arrDetails(lEachRow, BuzDetail.Point_Qty) * arrDetails(lEachRow, BuzDetail.Point_Price)
        arrDetails(lEachRow, BuzDetail.Point_Amt) = arrDetails(lEachRow, BuzDetail.Point_CurrDayPrice) * arrDetails(lEachRow, BuzDetail.Point_DaysNum)
        
        arrDetails(lEachRow, BuzDetail.DownLoad_Amt) = arrDetails(lEachRow, BuzDetail.DownLoad_Qty) * arrDetails(lEachRow, BuzDetail.DownLoad_Price)
        arrDetails(lEachRow, BuzDetail.Credit_Amt) = arrDetails(lEachRow, BuzDetail.Credit_Qty) * arrDetails(lEachRow, BuzDetail.Credit_Price)
        
        dblSum_Point = dblSum_Point + arrDetails(lEachRow, BuzDetail.Point_Amt)
        dblSum_DownLoad = dblSum_DownLoad + arrDetails(lEachRow, BuzDetail.DownLoad_Amt)
        dblSum_Credit = dblSum_Credit + arrDetails(lEachRow, BuzDetail.Credit_Amt)
    Next
    
    shtBusinessDetails.Cells(shtBusinessDetails.DataStartRow, 1).Resize(UBound(arrDetails, 1) - LBound(arrDetails, 1) + 1, BuzDetail.[_last]).Value = arrDetails
    
    shtBusinessSumm.Range("rgSummary").Cells(1, 1) = dblSum_Point
    shtBusinessSumm.Range("rgSummary").Cells(2, 1) = dblSum_DownLoad
    shtBusinessSumm.Range("rgSummary").Cells(3, 1) = dblSum_Credit
    shtBusinessSumm.Range("rgSummary").Cells(3, 1) = dblSum_Point + dblSum_DownLoad + dblSum_Credit
    
    Application.ScreenUpdating = True
    Call fShowSheet(shtBusinessDetails)
    Call fShowAndActiveSheet(shtBusinessSumm)
    Call fGotoCell(shtBusinessSumm.Range("a1"))
    
    fMsgBox "计算完成，结果在表：[" & shtBusinessSumm.Name & "] 中，请检查！", vbInformation
error_handling:
    If gErrNum <> 0 Then
        shtBusinessSumm.Range("rgSummary").ClearContents
        GoTo reset_excel_options
    End If
    
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
reset_excel_options:
    Err.Clear
    Erase arrDetails
    fClearRefVariables
    fEnableExcelOptionsAll
End Sub

Sub subMain_ClearBuzDetails()
    Dim rg As Range
    Dim lMaxRow As Long
    
    If Not fPromptToConfirmToContinue("你确定要删除表[" & shtBusinessDetails.Name & "]中的数据吗? " _
                & vbCr & "这将不可撤销?") Then
        Exit Sub
    End If
    
    lMaxRow = shtBusinessDetails.UsedRange.Row + shtBusinessDetails.UsedRange.Rows.Count - 1
    
    If lMaxRow > shtBusinessDetails.DataStartRow Then
        Set rg = fGetRangeByStartEndPos(shtBusinessDetails, shtBusinessDetails.DataStartRow, BuzDetail.[_first], lMaxRow, BuzDetail.[_last])
        rg.ClearContents
        rg.ClearComments
        'rg.ClearFormats
        rg.ClearHyperlinks
    End If
    
    shtBusinessSumm.Range("rgSummary").ClearContents
    
    Set rg = Nothing
    Call fShowAndActiveSheet(shtBusinessDetails)
    MsgBox "Done!", vbInformation
End Sub
