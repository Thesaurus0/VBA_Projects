Attribute VB_Name = "MC1_Business"
Option Explicit
Option Base 1

Sub subMain_ConsolidateAndGenReports()
    Dim arrRawData()
    Dim arrOutput()
    Dim dictLog As Dictionary
    Dim lEachRow As Long
    'Dim lMaxRow As Long
    Dim i As Integer
    Dim sMsg As String
    Dim sVendorName As String
    Dim dblPrice As Double
    Dim dictUniquePrice As Dictionary
    Dim dictMultiPrice As Dictionary
    Dim sProduct As String
    Dim arrVendorPrices()
    
    On Error GoTo error_handling
    
    Call fInitialization
     
    Call fDeleteRowsFromSheetLeaveHeader(shtPurchaseODByProduct)
    'Call fDeleteRowsFromSheetLeaveHeader(shtPurchaseODByVendor)
    Call fDeleteRowsFromSheetLeaveHeader(shtLog)
    
    Call fCopyReadWholeSheetData2Array(shtPurchaseODRaw, arrRawData, , shtPurchaseODRaw.DataFromRow)
     
'    lMaxRow = ArrLen(arrRawData)
    If fArrayIsEmptyOrNoData(arrRawData) Then fErr "No data was found in sheet " & shtPurchaseODRaw.name
    
    Call fSortDataInSheetSortSheetData(shtVendorPrice, Array(VendorPrice.ProdName, VendorPrice.Price, VendorPrice.VendorName))
    Call fCopyReadWholeSheetData2Array(shtVendorPrice, arrVendorPrices, , shtVendorPrice.DataFromRow)
    Call fReadVendorPrice(arrVendorPrices, dictUniquePrice, dictMultiPrice)
    
    Set dictLog = New Dictionary
    
    Dim iCnt As Long
    Dim sLines As String
    Dim arrLines
    Dim dblQty As Double
    Dim lLineNumInPrice As Long
    
    iCnt = 0
    For lEachRow = LBound(arrRawData, 1) To UBound(arrRawData, 1)
        sProduct = Trim(arrRawData(lEachRow, PODRaw.ProdName))
        
        If dictMultiPrice.Exists(sProduct) Then
            sLines = dictMultiPrice(sProduct)
            iCnt = iCnt + Len(sLines) - Len(Replace(sLines, DELIMITER, ""))
        End If
    Next
    
    ReDim arrOutput(LBound(arrRawData, 1) To UBound(arrRawData, 1) + iCnt, PODByProduct.[_first] To PODByProduct.[_last])
    
    iCnt = 0
    For lEachRow = LBound(arrRawData, 1) To UBound(arrRawData, 1)
        iCnt = iCnt + 1
        sProduct = Trim(arrRawData(lEachRow, PODRaw.ProdName))
        dblQty = val(arrRawData(lEachRow, PODRaw.PurchaseQty))
        
        arrOutput(iCnt, PODByProduct.ProdName) = sProduct
        arrOutput(iCnt, PODByProduct.Qty) = dblQty
        
        If Not dictMultiPrice.Exists(sProduct) Then
            If dictUniquePrice.Exists(sProduct) Then
                lLineNumInPrice = CLng(dictUniquePrice(sProduct))
                arrOutput(iCnt, PODByProduct.VendorName) = Trim(arrVendorPrices(lLineNumInPrice, VendorPrice.VendorName))
                arrOutput(iCnt, PODByProduct.Price) = val(arrVendorPrices(lLineNumInPrice, VendorPrice.Price))
            Else
                arrOutput(iCnt, PODByProduct.VendorName) = ""
                arrOutput(iCnt, PODByProduct.Price) = ""
                arrOutput(iCnt, PODByProduct.Remarks) = "�����Ҳ�����Ӧ��, ����û�м۸�.!"
                dictLog.Add lEachRow + shtPurchaseODRaw.HeaderByRow, "�����Ҳ�����Ӧ��, ����û�м۸�."
            End If
        Else
            arrLines = Split(dictMultiPrice(sProduct), DELIMITER)
            
            For i = LBound(arrLines) To UBound(arrLines)
                If i > LBound(arrLines) Then iCnt = iCnt + 1
                
                arrOutput(iCnt, PODByProduct.ProdName) = sProduct
                arrOutput(iCnt, PODByProduct.Qty) = dblQty
                
                lLineNumInPrice = CLng(arrLines(i))
                
                arrOutput(iCnt, PODByProduct.VendorName) = Trim(arrVendorPrices(lLineNumInPrice, VendorPrice.VendorName))
                arrOutput(iCnt, PODByProduct.Price) = val(arrVendorPrices(lLineNumInPrice, VendorPrice.Price))
                arrOutput(iCnt, PODByProduct.Remarks) = "�����Ӧ���ṩ����ͬ����ͼ۸�!"
            Next
        End If
next_row:
    Next
     
    Call fAppendArray2Sheet(shtPurchaseODByProduct, arrOutput)
    Call fSetConditionFormatForBorders(shtPurchaseODByProduct, , shtPurchaseODByProduct.DataFromRow, , 1)
    Call fSetConditionFormatForOddEvenLine(shtPurchaseODByProduct, , shtPurchaseODByProduct.DataFromRow)
    
    If dictLog.Count > 0 Then
'        Dim arrLog()
'        arrLog = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictLog)
        Call fAppendArray2Sheet(shtLog, fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictLog, "~"))
        Call fAppendArray2Sheet(shtLog, fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictLog, "~"), 3, 2)
        
        Call fSetConditionFormatForBorders(shtLog, , , , 1)
        Call fSetConditionFormatForOddEvenLine(shtLog, , , , 1)
    End If
error_handling:
    Erase arrRawData
    Erase arrVendorPrices
    Erase arrOutput
    
    If gErrNum <> 0 Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    Application.ScreenUpdating = True
    If dictLog.Count > 0 Then
        'Call fShowAndActiveSheet(shtLog)
        Call fShowAndActiveSheet(shtPurchaseODByProduct)
    Else
        Call fShowAndActiveSheet(shtPurchaseODByProduct)
    End If
    
    sMsg = "�������, please check the sheet : " & shtPurchaseODByProduct.name _
         & IIf(dictLog.Count > 0, vbCr & vbCr & "�������쳣����,�ж�������̵ļ۸���ͬ���������, ����.", "")
    
    MsgBox sMsg, IIf(dictLog.Count > 0, vbCritical, vbInformation)
    Set dictLog = Nothing
    
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub


Function fReadVendorPrice(arrVendorPrices(), dictUniquePrice As Dictionary, dictMultiPrice As Dictionary)
    Dim lEachRow As Long
    Dim sProd As String
    Dim sVendor As String
'    Dim sPrevProd As String
    Dim dblPrice As Double
'    Dim dblPrevPrice As Double
    Dim dblFirstPrice As Double
    
    Dim dictFirstPrice As Dictionary
    Dim dictAll As Dictionary
    
    
    Set dictFirstPrice = New Dictionary
    
    For lEachRow = LBound(arrVendorPrices, 1) To UBound(arrVendorPrices, 1)
        sProd = Trim(arrVendorPrices(lEachRow, VendorPrice.ProdName))
        sVendor = Trim(arrVendorPrices(lEachRow, VendorPrice.VendorName))
        
        dblPrice = val(arrVendorPrices(lEachRow, VendorPrice.Price))
        
        If Len(sProd) <= 0 Then
            fErr "��" & lEachRow + shtVendorPrice.HeaderByRow & "�еĲ�Ʒ�ͺ��ǿյ�!"
        End If
        If Len(sVendor) <= 0 Then
            fErr "��" & lEachRow + shtVendorPrice.HeaderByRow & "�еĹ�Ӧ���ǿյ�!"
        End If
        
        If Not dictFirstPrice.Exists(sProd) Then dictFirstPrice.Add sProd, dblPrice
next_row:
    Next
    
    Set dictAll = New Dictionary
    
    For lEachRow = LBound(arrVendorPrices, 1) To UBound(arrVendorPrices, 1)
        sProd = Trim(arrVendorPrices(lEachRow, VendorPrice.ProdName))
        sVendor = Trim(arrVendorPrices(lEachRow, VendorPrice.VendorName))
        
        dblPrice = val(arrVendorPrices(lEachRow, VendorPrice.Price))
        
        dblFirstPrice = dictFirstPrice(sProd)
        
        If Not dictAll.Exists(sProd) Then
            If dblPrice < dblFirstPrice Then fErr "this price on line " & lEachRow + 1 & " should not be larger than the first one " & dblFirstPrice
            
            If dblPrice = dblFirstPrice Then
                dictAll.Add sProd, lEachRow
            End If
        Else
            If dblPrice < dblFirstPrice Then fErr "this price on line " & lEachRow + 1 & " should not be larger than the first one " & dblFirstPrice
            
            If dblPrice = dblFirstPrice Then
                dictAll(sProd) = dictAll(sProd) & DELIMITER & lEachRow
            End If
        End If
    Next
    
    Set dictFirstPrice = Nothing
    
    Dim i As Long
    Dim sLines As String
    
    Set dictUniquePrice = New Dictionary
    Set dictMultiPrice = New Dictionary
    
    For i = 0 To dictAll.Count - 1
        sLines = dictAll.Items(i)
        If InStr(sLines, DELIMITER) > 0 Then
            dictMultiPrice.Add dictAll.Keys(i), dictAll.Items(i)
        Else
            dictUniquePrice.Add dictAll.Keys(i), dictAll.Items(i)
        End If
    Next
    Set dictAll = Nothing
End Function

Function fStrInDelimiteredStr(ByVal asAnswerStr As String, ByVal sEachAns As String, Optional sDeli As String = "|") As Integer
    Dim sAnswerStr As String
    Dim iAt As Integer
    
    sEachAns = sDeli & Trim(sEachAns) & sDeli
    sAnswerStr = sDeli & Trim(asAnswerStr) & sDeli
    
    iAt = 0
    iAt = InStr(sAnswerStr, sEachAns)
    
    Dim sLeft As String
    If iAt > 0 Then
        If iAt = 1 Then GoTo exit_fun
        
        sLeft = Left(sAnswerStr, iAt - 1)
        
        iAt = Len(sLeft) - Len(Replace(sLeft, sDeli, "")) + 1
    End If
    
exit_fun:
    fStrInDelimiteredStr = iAt
End Function

'
'Function fSetRowHeightForExceedingThreshold(sht As Worksheet, lStartRow As Long, lEndRow As Long _
'                                        , Optional dblRowHeightThreshold As Double = 16)
'    Dim lEachRow As Long
'    Dim rgAll As Range
'    Dim rgTarget As Range
'
'    Set rgAll = sht.Rows(lStartRow & ":" & lEndRow)
'
'    For lEachRow = 1 To lEndRow - lStartRow + 1
'
'        If rgAll.Rows(lEachRow).RowHeight < dblRowHeightThreshold Then
'            If rgTarget Is Nothing Then
'                Set rgTarget = rgAll.Rows(lEachRow)
'            Else
'                Set rgTarget = Union(rgTarget, rgAll.Rows(lEachRow))
'            End If
'        End If
'    Next
'
'    If Not rgTarget Is Nothing Then
'        rgTarget.RowHeight = dblRowHeightThreshold
'    End If
'
'    Set rgAll = Nothing
'    Set rgTarget = Nothing
'End Function
'
'Function fSetConditionalFormatForBorders()
'    Dim lMaxRow As Long
'
'    Call fClearSerialNoFromSheets
'
'    Dim sht As Worksheet
'
'    Set sht = shtCabinet
'
'    lMaxRow = fGetValidMaxRow(sht)
'    If lMaxRow >= 7 Then
'        Call fDeleteAllConditionFormatFromSheet(sht)
'    '    Call fSetConditionFormatForOddEvenLine(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
'        Call fSetConditionFormatForBorders(sht, , 7, , 1)
'        sht.Cells.WrapText = True
'        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.AutoFit
'        'fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, fLetter2Num("K")).EntireColumn.AutoFit
'    End If
'
'    Set sht = shtCabinetFrame
'    lMaxRow = fGetValidMaxRow(sht)
'    If lMaxRow >= 7 Then
'        Call fDeleteAllConditionFormatFromSheet(sht)
'    '    Call fSetConditionFormatForOddEvenLine(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
'        Call fSetConditionFormatForBorders(sht, , 7, , 1)
'        sht.Cells.WrapText = True
'        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.AutoFit
'        'sht.Columns.AutoFit
'    End If
'
'
'    Set sht = shtDoor
'    lMaxRow = fGetValidMaxRow(sht)
'    If lMaxRow >= 7 Then
'        Call fDeleteAllConditionFormatFromSheet(sht)
'    '    Call fSetConditionFormatForOddEvenLine(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
'        Call fSetConditionFormatForBorders(sht, , 7, , 1)
'        sht.Cells.WrapText = True
'        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.AutoFit
'        'sht.Columns.AutoFit
'    End If
'
'
'    Set sht = shtHardwares
'    lMaxRow = fGetValidMaxRow(sht)
'    If lMaxRow >= 7 Then
'        Call fDeleteAllConditionFormatFromSheet(sht)
'    '    Call fSetConditionFormatForOddEvenLine(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
'        Call fSetConditionFormatForBorders(sht, , 7, , 1)
'        sht.Cells.WrapText = True
'        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.AutoFit
'
'       ' sht.Columns.AutoFit
'    End If
'
'    Set sht = Nothing
'End Function
'
'Private Function fSetConditionFormatForBorders(ByRef shtParam As Worksheet, Optional lMaxCol As Long = 0 _
'                                            , Optional lRowFrom As Long = 2, Optional lRowTo As Long = 0 _
'                                            , Optional arrKeyColsNotBlank _
'                                            , Optional bExtendToMore10ThousRows As Boolean = False)
''arrKeyColsNotBlank
''    1. singlecol: 1
''    1. array(1,2,3)
'    If lMaxCol = 0 Then lMaxCol = fGetValidMaxCol(shtParam)
'    If lRowTo = 0 Then lRowTo = fGetValidMaxRow(shtParam)
'
'    If lMaxCol <= 0 Then Exit Function
'    If bExtendToMore10ThousRows Then lRowTo = lRowTo + 100000
'
'    If lRowTo < lRowFrom Then Exit Function
'
'    Dim rngCondFormat As Range
'    Set rngCondFormat = fGetRangeByStartEndPos(shtParam, lRowFrom, 1, lRowTo, lMaxCol)
'
'    Dim sAddr As String
'    Dim sKeyColsFormula As String
'    Dim sFormula As String
'    Dim lColor As Long
'    Dim i As Integer
'    Dim sColLetter As String
'    Dim aFormatCondition As FormatCondition
'
'    If Not IsMissing(arrKeyColsNotBlank) Then
'        If IsArray(arrKeyColsNotBlank) Then
'            For i = LBound(arrKeyColsNotBlank) To UBound(arrKeyColsNotBlank)
'                sColLetter = fNum2Letter(arrKeyColsNotBlank(i))
'                sKeyColsFormula = sKeyColsFormula & "," & "len(trim($" & sColLetter & lRowFrom & ")) > 0"
'            Next
'            If Len(sKeyColsFormula) > 0 Then sKeyColsFormula = Right(sKeyColsFormula, Len(sKeyColsFormula) - 1)
'        Else
'            sColLetter = fNum2Letter(arrKeyColsNotBlank)
'            sKeyColsFormula = "len(trim($" & sColLetter & lRowFrom & ")) > 0"
'            sKeyColsFormula = sKeyColsFormula
'        End If
'    Else
'        sKeyColsFormula = ""
'    End If
'
'    sFormula = "=And( " & sKeyColsFormula & ")"
'
'    Set aFormatCondition = rngCondFormat.FormatConditions.Add(Type:=xlExpression, Formula1:=sFormula)
'    aFormatCondition.SetFirstPriority
'    aFormatCondition.StopIfTrue = False
'
'    aFormatCondition.Borders(xlLeft).Weight = xlThin
'    aFormatCondition.Borders(xlRight).Weight = xlThin
'    aFormatCondition.Borders(xlTop).Weight = xlThin
'    aFormatCondition.Borders(xlBottom).Weight = xlThin
'
'    aFormatCondition.Borders(xlLeft).ThemeColor = 2
'    aFormatCondition.Borders(xlRight).ThemeColor = 2
'    aFormatCondition.Borders(xlTop).ThemeColor = 2
'    aFormatCondition.Borders(xlBottom).ThemeColor = 2
'
'    aFormatCondition.Borders(xlLeft).TintAndShade = 0.499984740745262 '  0.249946592608417
'    aFormatCondition.Borders(xlRight).TintAndShade = 0.499984740745262
'    aFormatCondition.Borders(xlTop).TintAndShade = 0.499984740745262
'    aFormatCondition.Borders(xlBottom).TintAndShade = 0.499984740745262
'
''    aFormatCondition.Borders(xlLeft).Color = -16776961
''    aFormatCondition.Borders(xlRight).Color = -16776961
''    aFormatCondition.Borders(xlTop).Color = -16776961
''    aFormatCondition.Borders(xlBottom).Color = -16776961
'
'    Set aFormatCondition = Nothing
'End Function

Function subMain_ClearBuzDetails()
    If MsgBox("���ڵ����ݽ��ᱻ���, ��ȷ��Ҫ������? ", vbYesNoCancel + vbCritical + vbDefaultButton3) <> vbYes Then Exit Function
    
    Call fDeleteRowsFromSheetLeaveHeader(shtPurchaseODRaw, shtPurchaseODRaw.HeaderByRow)
    Call fDeleteRowsFromSheetLeaveHeader(shtPurchaseODByProduct, shtPurchaseODByProduct.HeaderByRow)
    'Call fDeleteRowsFromSheetLeaveHeader(shtPurchaseODByVendor, shtPurchaseODByProduct.HeaderByRow)
    Call fDeleteRowsFromSheetLeaveHeader(shtLog)
    
'    Dim lMaxRow As Long
'
'    lMaxRow = fGetValidMaxRow(shtPurchaseODRaw)
'
'    If lMaxRow > shtPurchaseODRaw.HeaderByRow Then
'        fGetRangeByStartEndPos(shtPurchaseODRaw, shtPurchaseODRaw.DataFromRow, Product.PurchaseQty, lMaxRow, Product.PurchaseQty).ClearContents
'    End If
End Function
'
'Private Function fClearSerialNoFromSheets()
'    Dim rgFound As Range
'    Dim lMaxRow As Long
'
'    lMaxRow = fGetValidMaxRow(shtCabinet)
'    If lMaxRow >= 7 Then
'        fGetRangeByStartEndPos(shtCabinet, 7, fLetter2Num("C"), lMaxRow, fLetter2Num("C")).HorizontalAlignment = xlLeft
'    End If
'
'    lMaxRow = fGetValidMaxRow(shtCabinetFrame)
'    If lMaxRow >= 7 Then
'        Set rgFound = fFindInWorksheet(shtCabinetFrame.Columns("C"), "�ϼ�")
'        fGetRangeByStartEndPos(shtCabinetFrame, rgFound.Row - 2, 1, lMaxRow, 1).ClearContents
'        fGetRangeByStartEndPos(shtCabinetFrame, 7, fLetter2Num("C"), lMaxRow, fLetter2Num("C")).HorizontalAlignment = xlLeft
'    End If
'
'    lMaxRow = fGetValidMaxRow(shtDoor)
'    If lMaxRow >= 7 Then
'        Set rgFound = fFindInWorksheet(shtDoor.Columns("C"), "�ϼ�")
'        fGetRangeByStartEndPos(shtDoor, rgFound.Row - 2, 1, lMaxRow, 1).ClearContents
'    End If
'
'    lMaxRow = fGetValidMaxRow(shtHardwares)
'    If lMaxRow >= 7 Then
'        Set rgFound = fFindInWorksheet(shtHardwares.Columns("C"), "�ϼ�")
'        fGetRangeByStartEndPos(shtHardwares, rgFound.Row - 2, 1, lMaxRow, 1).ClearContents
'    End If
'
'    Set rgFound = Nothing
'End Function
'
'
'
Function fExtractProductFromPriceConfigSheet()
    shtProdMasterExtracted.Cells.Delete
    
    Dim dictProd As Dictionary
    Dim arrPrices()
    Dim lEachRow As Long
    Dim sProd As String
    
    Set dictProd = New Dictionary
    
    Call fCopyReadWholeSheetData2Array(shtVendorPrice, arrPrices, , shtVendorPrice.DataFromRow)
    
    For lEachRow = LBound(arrPrices, 1) To UBound(arrPrices, 1)
        sProd = Trim(arrPrices(lEachRow, VendorPrice.ProdName))
        
        If Not dictProd.Exists(sProd) Then dictProd.Add sProd, ""
    Next
    Erase arrPrices
    
    Dim arrData()
    arrData = fTranspose1DimenArrayTo2DimenArrayVertically(dictProd.Keys)
    
    shtProdMasterExtracted.Range("A1").Resize(ArrLen(arrData), 1).value = arrData
    
    Erase arrData
    Set dictProd = Nothing
End Function
