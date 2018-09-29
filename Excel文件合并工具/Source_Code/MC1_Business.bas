Attribute VB_Name = "MC1_Business"
Option Explicit
Option Base 1

'Private Const chapter = "�½�"
'Private Const criteria_item = "ִ��Ҫ��"
'Private Const FEASIBLE_TO_PROCESS = "�Ƿ��ִ��"
'Private Const PROCESS_ON_THE_WAY = "�Ƿ���ִ��"
'Private Const REASON_WHY_NOT = "δ��ִ�еľ���ԭ��"
'Private Const YOUR_ACTION = "����Ӧ�Բ���"

'Type typeCols
'    chapter As Long
'    CriteriaItem As Long
'    Feasible As Long
'    processOnTheWay As Long
'    Reason As Long
'    Action As Long
'End Type

Sub subMain_MergeExcelFilesFor3Folders()
    Dim arrOutput()
    Dim dictLog As Dictionary
    Dim lEachRow As Long
    Dim i As Integer
    Dim sMsg As String
    Dim sFolder As String
    Dim sOutputFolder As String
    Dim sProduct As String
'    Dim arrHeader()
'    Dim dictNotInProcess As Dictionary
    
    Call fInitialization
    
'    arrHeader = Array(chapter, criteria_item, FEASIBLE_TO_PROCESS, PROCESS_ON_THE_WAY, REASON_WHY_NOT, YOUR_ACTION)
    
    On Error GoTo error_handling
    
    Set dictLog = New Dictionary
    
    sFolder = fSelectFolderDialog(ThisWorkbook.Path) '
    If Len(sFolder) <= 0 Then fErr '
    
    sOutputFolder = sFolder & "Output" '& Format(Now(), "yyyymmddHHMMSS")
    
    If fFolderExists(sOutputFolder) Then
        If fPromptToConfirmToContinue("The output folder already exists, all files will be deleted first, are you sure to continue?" _
                 & vbCr & sOutputFolder) Then
            fDeleteAllFilesFromFolder(
        Else
            fErr
        End If
    End If
    

    Call fDeleteRowsFromSheetLeaveHeader(shtLog)
    Call fDeleteRowsFromSheetLeaveHeader(shtReportDetails)
    
    fGetFSO
    Dim sOriginFolder As String
    Dim sFirstResponseFolder As String
    Dim sSecondResponseFolder As String
    sOriginFolder = sFolder & "ԭ��\"
    sFirstResponseFolder = sFolder & "��һ�η���\"
    sSecondResponseFolder = sFolder & "�ڶ��η���\"
    
    If Not (gFSO.FolderExists(sOriginFolder) And gFSO.FolderExists(sFirstResponseFolder) And gFSO.FolderExists(sSecondResponseFolder)) Then
        fErr "There must be 3 folders there, but any one was not found:" & vbCr & vbCr _
             & sOriginFolder
    End If
    
    Dim arrOrigin()
    Dim arrFirst()
    Dim arrSecond()
    arrOrigin = fGetAllExcelFileListFromSubFolders(sOriginFolder)
    
    If ArrLen(arrOrigin) <= 0 Then fErr "No file was found in folder " & vbCr & sOriginFolder
    
    arrFirst = fGetAllExcelFileListFromSubFolders(sFirstResponseFolder)
    arrSecond = fGetAllExcelFileListFromSubFolders(sSecondResponseFolder)
    
    Dim i As Long
    
    For i = LBound(arrOrigin) To UBound(arrOrigin)
        sFile = arrOrigin(i)
        sFileNetName = fGetFileNetName(sFile)
        
    Next
    
'    Dim sFile As String
'    Dim wb As Workbook
'    Dim shtInput As Worksheet
'    Dim bAlreadyOpened As Boolean
'    Dim arrColIndex()
'    Dim lHeaderAtRow  As Long
'    Dim dblFeasibleRate As Double
''    Dim colIndex As typeCols
'    Dim sFileNetName As String
'    Dim dblInprocessRate As Double
'    Dim j As Long
'    Dim sNotInProc As String
'    Dim sNotInProcReason As String
'    Dim dictAllItems As Dictionary
'
'    For i = LBound(arrFiles) To UBound(arrFiles)
'        sFile = arrFiles(i)
'        sFileNetName = fGetFileNetName(sFile)
'
'        If Left(sFileNetName, 1) = "~" Then GoTo next_file
'        If Left(sFileNetName, 1) = "$" Then GoTo next_file
'
''        Set dictNotInProcess = New Dictionary
'
'        Set wb = fOpenWorkbook(sFile, bAlreadyOpened, False, , shtInput)
'
'        Call fFindAllColumnsIndexByColNames(shtInput.Rows("1:10"), arrHeader, arrColIndex, lHeaderAtRow)
'
'        colIndex.chapter = arrColIndex(LBound(arrColIndex))
'        colIndex.CriteriaItem = arrColIndex(LBound(arrColIndex) + 1)
'        colIndex.Feasible = arrColIndex(LBound(arrColIndex) + 2)
'        colIndex.processOnTheWay = arrColIndex(LBound(arrColIndex) + 3)
'        colIndex.Reason = arrColIndex(LBound(arrColIndex) + 4)
'        colIndex.Action = arrColIndex(LBound(arrColIndex) + 5)
'
'        If shtInput.Cells(lHeaderAtRow, colIndex.chapter).MergeCells Then lHeaderAtRow = shtInput.Cells(lHeaderAtRow, colIndex.chapter).MergeArea.Row + shtInput.Cells(lHeaderAtRow, colIndex.chapter).MergeArea.Rows.Count - 1
'        arrMaster = fGetRangeByStartEndPos(shtInput, 1, 1, fGetValidMaxRow(shtInput), fGetValidMaxCol(shtInput)).value
'
'        Call fFillArrayByMergedCells(arrMaster, colIndex.Reason, shtInput, lHeaderAtRow + 1)
'        Call fFillArrayByMergedCells(arrMaster, colIndex.Action, shtInput, lHeaderAtRow + 1)
'        dblFeasibleRate = fFillArrayByMergedCellsForSelf(sFileNetName, arrMaster, shtInput, lHeaderAtRow + 1, colIndex _
'        , dictLog, dictNotInProcess, dblInprocessRate, dictAllItems)
'
'        If Not bAlreadyOpened Then Call fCloseWorkBookWithoutSave(wb)
'
'        If dictNotInProcess.Count > 0 Then
'            ReDim arrOutput(1 To dictNotInProcess.Count, 1 To 8)
'            'Set dictNotInProcess = fConsolidateAndCalculate(arrMaster, colIndex)
'            Erase arrMaster
'
'            For j = 0 To dictNotInProcess.Count - 1
'                arrOutput(j + 1, 1) = sFileNetName
'                arrOutput(j + 1, 2) = dblFeasibleRate
'                arrOutput(j + 1, 3) = dblInprocessRate
'
'                sNotInProc = dictNotInProcess.Keys(j)
'                sNotInProcReason = dictNotInProcess.Items(j)
'                arrOutput(j + 1, 4) = Split(sNotInProc, DELIMITER)(0)
'                arrOutput(j + 1, 5) = Split(sNotInProc, DELIMITER)(1)
'
'                arrOutput(j + 1, 6) = Split(sNotInProcReason, DELIMITER)(0)
'                arrOutput(j + 1, 7) = Split(sNotInProcReason, DELIMITER)(1)
'            Next
'        Else
'            ReDim arrOutput(1 To 1, 1 To 8)
'            Erase arrMaster
'
'            arrOutput(1, 1) = sFileNetName
'            arrOutput(1, 2) = dblFeasibleRate
'            arrOutput(1, 3) = 1
'        End If
'
'        Call fAppendArray2Sheet(shtReportDetails, arrOutput)
'        Call fAppendArray2Sheet(shtAllItems, fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictAllItems))
'next_file:
'        Set dictAllItems = Nothing
'    Next

'    Call fAppendArray2Sheet(shtPurchaseODByProduct, arrOutput)
    Call fSetConditionFormatForBorders(shtReportDetails, , 2, , 1)
    Call fSetConditionFormatForOddEvenLine(shtReportDetails, , 2, , 1)
     
'
    If dictLog.Count > 0 Then
        Dim arrLog()
        arrLog = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictLog)
        Call fAppendArray2Sheet(shtLog, arrLog)
        
'        arrLog = fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictLog)
'        'Call fAppendArray2Sheet(shtLog, fConvertDictionaryDelimiteredItemsTo2DimenArrayForPaste(dictLog, "~"), 3, 2)
'        shtLog.Cells(2, 5).Resize(ArrLen(arrLog, 1), ArrLen(arrLog, 2)).value = arrLog

        Call fSetConditionFormatForBorders(shtLog, , , , 1)
        Call fSetConditionFormatForOddEvenLine(shtLog, , , , 1)
    End If
error_handling:
'    Set dictNotInProcess = Nothing
    Erase arrFiles
'    Erase arrVendorPrices
    Erase arrOutput
    
    If gErrNum <> 0 Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    Application.ScreenUpdating = True
    If dictLog.Count > 0 Then
        'Call fShowAndActiveSheet(shtLog)
        Call fShowAndActiveSheet(shtLog)
    Else
        Call fShowAndActiveSheet(shtLog)
    End If
    
    sMsg = "�������, please check the sheet : " & shtReportDetails.name _
         & IIf(dictLog.Count > 0, vbCr & vbCr & "�������쳣����, ����.", "")
    
    MsgBox sMsg, IIf(dictLog.Count > 0, vbCritical, vbInformation)
    Set dictLog = Nothing
    
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
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
    
    
    shtProdMasterExtracted.Columns("A").ClearContents
    shtProdMasterExtracted.Columns("A").NumberFormat = "@"
    shtProdMasterExtracted.Range("A1").Resize(ArrLen(arrData), 1).value = arrData
    
    Erase arrData
    Set dictProd = Nothing
End Function
 
Sub subMain_GenSummaryReport()
    Dim arrOutput()
    Dim dictLog As Dictionary
    Dim lEachRow As Long
    Dim i As Long
    Dim sMsg As String
    Dim sFolder As String
    Dim dblPrice As Double
    Dim sProduct As String
    Dim dictNotInProcess As Dictionary
    
    Call fInitialization

    On Error GoTo error_handling
    fGetRangeByStartEndPos(shtReportSummary, 2, 1, 2 + fGetValidMaxRow(shtReportSummary), 10).ClearContents
    
    Set dictLog = New Dictionary
    
    Call fCopyReadWholeSheetData2Array(shtAllItems, arrMaster)
    
    If ArrLen(arrMaster) <= 0 Then fErr "no data found in " & shtAllItems.name
     
    Dim sFile As String
    Dim wb As Workbook
    Dim sFeasible  As String
    Dim j As Long
    Dim dictUniqueFiles As Dictionary
    Dim dictAnyNotFeasible As Dictionary
    Dim dictChapterCnt As Dictionary
    Dim dictItemCnt As Dictionary
    Dim dictItem2Chapter As Dictionary
    
    Dim lFeasibleCnt As Long
    Dim lInProcCnt As Long
    Dim sInProcess As String
    Dim sItem As String
    Dim sChapter As String
    
    Set dictUniqueFiles = New Dictionary
    Set dictAnyNotFeasible = New Dictionary
    Set dictChapterCnt = New Dictionary
    Set dictItemCnt = New Dictionary
    Set dictItem2Chapter = New Dictionary
    
    lFeasibleCnt = 0
    lInProcCnt = 0
    For i = LBound(arrMaster, 1) To UBound(arrMaster, 1)
        sFile = Trim(arrMaster(i, 1))
        sChapter = Trim(arrMaster(i, 2))
        sItem = Trim(arrMaster(i, 3))
        sFeasible = Trim(arrMaster(i, 4))
        sInProcess = Trim(arrMaster(i, 5))
        
        If Not dictUniqueFiles.Exists(sFile) Then
            dictUniqueFiles.Add sFile, ""
        End If
        
        If sFeasible = "��" Then
            lFeasibleCnt = lFeasibleCnt + 1
        Else
            dictAnyNotFeasible(sFile) = ""
        End If
        
        If sInProcess = "��" Then
            lInProcCnt = lInProcCnt + 1
        Else
            dictChapterCnt(sChapter) = val(dictChapterCnt(sChapter)) + 1
            dictItemCnt(sItem) = val(dictItemCnt(sItem)) + 1
        End If
        
        If Not dictItem2Chapter.Exists(sItem) Then
            dictItem2Chapter.Add sItem, sChapter
        Else
            If sChapter <> dictItem2Chapter(sItem) Then dictItem2Chapter(sItem) = dictItem2Chapter(sItem) & DELIMITER & sChapter
        End If
    Next
    
    Dim dictAllFeasible As Dictionary
    Set dictAllFeasible = New Dictionary
    For i = 0 To dictUniqueFiles.Count - 1
        sFile = dictUniqueFiles.Keys(i)
        If Not dictAnyNotFeasible.Exists(sFile) Then
            dictAllFeasible.Add sFile, ""
        End If
    Next
    
    shtReportSummary.Range("A2").value = dictUniqueFiles.Count
    shtReportSummary.Range("B2").value = lFeasibleCnt / ArrLen(arrMaster, 1)
    shtReportSummary.Range("C2").value = lInProcCnt / ArrLen(arrMaster, 1)
    shtReportSummary.Range("D2").value = dictAllFeasible.Count
    shtReportSummary.Range("E2").value = dictUniqueFiles.Count - dictAllFeasible.Count
    
    Set dictUniqueFiles = Nothing
    Set dictAllFeasible = Nothing
    
    '====================================
    Dim arrChapter()
    
    Dim dictUniqueChapterCnt As Dictionary
    Set dictUniqueChapterCnt = New Dictionary
    For i = 0 To dictChapterCnt.Count - 1
        dictUniqueChapterCnt(dictChapterCnt.Items(i)) = ""
    Next
    
    arrChapter = dictUniqueChapterCnt.Keys()
    Set dictUniqueChapterCnt = Nothing
    
    Call fSortArrayDesc(arrChapter)
    
    If ArrLen(arrChapter) > 11 Then
        ReDim Preserve arrChapter(0 To 10)
    End If

    Dim dictWorst As Dictionary
    Set dictWorst = New Dictionary
        
    For j = LBound(arrChapter) To UBound(arrChapter)
        For i = 0 To dictChapterCnt.Count - 1
            If dictChapterCnt.Items(i) = arrChapter(j) Then
                dictWorst(dictChapterCnt.Keys(i)) = dictChapterCnt.Items(i)
            End If
        Next
    Next
    Erase arrChapter
    Set dictChapterCnt = Nothing
    
    Dim arrWorst()
    ReDim arrWorst(1 To dictWorst.Count, 2)
    'Dim sWorst As String
    For i = 0 To dictWorst.Count - 1
        'sWorst = sWorst & vbLf & dictWorst.Keys(i) & " :  " & dictWorst.Items(i) & "��"
        arrWorst(i + 1, 1) = dictWorst.Keys(i)
        arrWorst(i + 1, 2) = dictWorst.Items(i) & "��"
    Next

'    If Len(sWorst) > 0 Then sWorst = Right(sWorst, Len(sWorst) - 1)
'    shtReportSummary.Range("F2").value = sWorst
    shtReportSummary.Range("F2").Resize(dictWorst.Count, 2).value = arrWorst
    Erase arrWorst
    Set dictUniqueChapterCnt = Nothing
    Set dictWorst = Nothing
'===============
    Dim arrItem()
    
    Dim dictUniqueItemCnt As Dictionary
    Set dictUniqueItemCnt = New Dictionary
    For i = 0 To dictItemCnt.Count - 1
        dictUniqueItemCnt(dictItemCnt.Items(i)) = ""
    Next
    
    arrItem = dictUniqueItemCnt.Keys()
    Set dictUniqueItemCnt = Nothing
    
    Call fSortArrayDesc(arrItem)
    
    If ArrLen(arrItem) > 20 Then
        ReDim Preserve arrItem(0 To 19)
    End If

    Set dictWorst = New Dictionary
        
    For j = LBound(arrItem) To UBound(arrItem)
        For i = 0 To dictItemCnt.Count - 1
            If dictItemCnt.Items(i) = arrItem(j) Then
                dictWorst(dictItemCnt.Keys(i)) = dictItemCnt.Items(i)
            End If
        Next
    Next
    Erase arrItem
    Set dictItemCnt = Nothing
    
    ReDim arrWorst(1 To dictWorst.Count, 3)
    
'    sWorst = ""
    For i = 0 To dictWorst.Count - 1
        'sWorst = sWorst & vbLf & dictWorst.Keys(i) & " :  " & dictWorst.Items(i) & "��"
        arrWorst(i + 1, 1) = dictWorst.Keys(i)
        arrWorst(i + 1, 2) = dictItem2Chapter(dictWorst.Keys(i))
        arrWorst(i + 1, 3) = dictWorst.Items(i) & "��"
    Next

    'If Len(sWorst) > 0 Then sWorst = Right(sWorst, Len(sWorst) - 1)
    'shtReportSummary.Range("G2").value = sWorst
    shtReportSummary.Range("H2").Resize(dictWorst.Count, 3).value = arrWorst
    
    Erase arrWorst
    Set dictUniqueItemCnt = Nothing
    Set dictWorst = Nothing
    Set dictItem2Chapter = Nothing
    
    Call fSetConditionFormatForBorders(shtReportSummary, , , , fLetter2Num("H"))
'    If dictLog.Count > 0 Then
'        Dim arrLog()
'        arrLog = fConvertDictionaryDelimiteredKeysTo2DimenArrayForPaste(dictLog)
'        Call fAppendArray2Sheet(shtLog, arrLog)
'
'
'        Call fSetConditionFormatForBorders(shtLog, , , , 1)
'        Call fSetConditionFormatForOddEvenLine(shtLog, , , , 1)
'    End If
error_handling:
    Set dictUniqueFiles = Nothing
    Set dictNotInProcess = Nothing
    
'    Erase arrVendorPrices
    Erase arrOutput
    
    If gErrNum <> 0 Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    Application.ScreenUpdating = True
    Call fShowAndActiveSheet(shtReportSummary)
    
    sMsg = "�������, please check the sheet : " & shtReportSummary.name
    
    MsgBox sMsg, IIf(dictLog.Count > 0, vbCritical, vbInformation)
    Set dictLog = Nothing
    
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub
