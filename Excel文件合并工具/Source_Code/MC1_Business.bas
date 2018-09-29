Attribute VB_Name = "MC1_Business"
Option Explicit
Option Base 1

Dim dictLog As Dictionary
Sub subMain_MergeExcelFilesFor3Folders()
    Dim arrOutput()
    Dim lEachRow As Long
    Dim i As Integer
    Dim sMsg As String
    Dim sTopFolder As String
    Dim sOutputFolder As String
    Dim sProduct As String
    
    Call fInitialization
     
    On Error GoTo error_handling
    
    Set dictLog = New Dictionary
    
    sTopFolder = fSelectFolderDialog(ThisWorkbook.Path) '
    If Len(sTopFolder) <= 0 Then fErr '
    
    sOutputFolder = fGetFileParentFolder(sTopFolder) & "Output\" '& Format(Now(), "yyyymmddHHMMSS")
    
    If fFolderExists(sOutputFolder) Then
        If fPromptToConfirmToContinue("The output folder already exists, all files will be deleted first, are you sure to continue?" _
                 & vbCr & sOutputFolder) Then
            Call fDeleteAllFromFolder(sOutputFolder)
        Else
            fErr
        End If
    Else
        Call fCheckPath(sOutputFolder, True)
    End If
    
    Call fDeleteRowsFromSheetLeaveHeader(shtLog)
    Call fDeleteRowsFromSheetLeaveHeader(shtReportDetails)
    
    fGetFSO
    Dim sOriginFolder As String
    Dim sFirstResponseFolder As String
    Dim sSecondResponseFolder As String
    sOriginFolder = sTopFolder & "原稿\"
    sFirstResponseFolder = sTopFolder & "第一次反馈\"
    sSecondResponseFolder = sTopFolder & "第二次反馈\"
    
    If Not (gFSO.FolderExists(sOriginFolder) And gFSO.FolderExists(sFirstResponseFolder) And gFSO.FolderExists(sSecondResponseFolder)) Then
        fErr "There must be 3 folders there, but any one was not found:" & vbCr & vbCr _
             & sOriginFolder
    End If
    
    Dim arrOrigin
    Dim arrFirst
    Dim arrSecond
    arrOrigin = fGetAllExcelFileListFromSubFolders(sOriginFolder)
    
    If ArrLen(arrOrigin) <= 0 Then fErr "No file was found in folder " & vbCr & sOriginFolder
    
    arrFirst = fGetAllExcelFileListFromSubFolders(sFirstResponseFolder)
    arrSecond = fGetAllExcelFileListFromSubFolders(sSecondResponseFolder)
    
    Dim dictFirst As Dictionary
    Set dictFirst = fReadFile3Letters(arrFirst)
    Erase arrFirst
    Dim dictSecond As Dictionary
    Set dictSecond = fReadFile3Letters(arrSecond)
    Erase arrSecond
    
    Dim sParentFolder As String
    Dim sOrigFile As String
    Dim wbSource As Workbook
    Dim wbOut As Workbook
    Dim sht As Worksheet
    Dim sOutputFile As String
    Dim sFileNetName As String
    Dim j As Long
    Dim sSubPrtFolder As String
    Dim sBaseName As String
    Dim sFileLetter As String
    Dim sOtherFile As String
    
    ReDim arrOutput(LBound(arrOrigin) To UBound(arrOrigin), 1 To 8)
    sOutputFolder = fCheckPath(sOutputFolder)
    For i = LBound(arrOrigin) To UBound(arrOrigin)
        sOrigFile = arrOrigin(i)
        sFileNetName = fGetFileNetName(sOrigFile)
        sBaseName = fGetFileBaseName(sOrigFile)
        
        If Len(sOrigFile) - Len(sOriginFolder) - Len(sBaseName) - 1 > 0 Then
            sSubPrtFolder = Right(sOrigFile, Len(sOrigFile) - Len(sOriginFolder))
            sSubPrtFolder = Left(sSubPrtFolder, Len(sSubPrtFolder) - Len(sBaseName))
            sOutputFile = sOutputFolder & sSubPrtFolder & sFileNetName & "_tmp.xlsx"
        Else
            sOutputFile = sOutputFolder & sFileNetName & "_tmp.xlsx"
        End If
        
        Set wbSource = fOpenWorkbook(sOrigFile, , True, , sht)
                
        If wbSource.Worksheets.Count > 1 Then
            For j = 1 To wbSource.Worksheets.Count
                If j = 1 Then
                    Set wbOut = fCopySingleSheet2NewWorkbookFile(sht, sOutputFile, "原稿" & j)
                Else
                    Call fCopySingleSheet2WorkBook(sht, wbOut, "原稿" & j)
                End If
            Next
        Else
            Set wbOut = fCopySingleSheet2NewWorkbookFile(sht, sOutputFile, "原稿")
        End If
        
        arrOutput(i, 1) = sFileNetName
        
        Call fCloseWorkBookWithoutSave(wbSource)
        Set wbSource = Nothing
        
        'first responce
        sFileLetter = Left(sFileNetName, 3)
        
        If dictFirst.Exists(sFileLetter) Then
            sOtherFile = dictFirst(sFileLetter)
            Set wbSource = fOpenWorkbook(sOtherFile, , True, , sht)

            If wbSource.Worksheets.Count > 1 Then
                For j = 1 To wbSource.Worksheets.Count
                    Call fCopySingleSheet2WorkBook(sht, wbOut, "第一次反馈" & j)
                Next
            Else
                Call fCopySingleSheet2WorkBook(sht, wbOut, "第一次反馈")
            End If
            
            wbOut.Worksheets(wbOut.Worksheets.Count).Move before:=wbOut.Worksheets(1)
            
            Call fCloseWorkBookWithoutSave(wbSource)
            arrOutput(i, 2) = "有第一次反馈文件"
'        Else
'            dictLog.Add dictLog.Count + 1 & DELIMITER & sFileLetter & " cannot find the file of 第一次反馈", ""
        End If
        
        'second response
        If dictSecond.Exists(sFileLetter) Then
            sOtherFile = dictSecond(sFileLetter)
            Set wbSource = fOpenWorkbook(sOtherFile, , True, , sht)
            
            If Not dictFirst.Exists(sFileLetter) Then
                dictLog.Add dictLog.Count + 1 & DELIMITER & sFileLetter & " 没有第一次反馈, 却有第二次所馈", ""
            End If
                    
            If wbSource.Worksheets.Count > 1 Then
                For j = 1 To wbSource.Worksheets.Count
                    Call fCopySingleSheet2WorkBook(sht, wbOut, "第二次反馈" & j)
                Next
            Else
                    Call fCopySingleSheet2WorkBook(sht, wbOut, "第二次反馈")
            End If
            
            wbOut.Worksheets(wbSource.Worksheets.Count + 1).Move after:=wbOut.Worksheets(wbOut.Worksheets.Count)
            
            Call fCloseWorkBookWithoutSave(wbSource)
            arrOutput(i, 3) = "有第二次反馈文件"
'        Else
'            dictLog.Add dictLog.Count + 1 & DELIMITER & sFileLetter & " cannot find the file of 第二次反馈", ""
        End If
        
        Call fSaveAndCloseWorkBook(wbOut)
        
        Name sOutputFile As Left(sOutputFile, Len(sOutputFile) - Len("_tmp.xlsx")) & ".xlsx"
    Next
    Set dictFirst = Nothing
    Set dictSecond = Nothing
     
    Call fAppendArray2Sheet(shtReportDetails, arrOutput)
    Call fSetConditionFormatForBorders(shtReportDetails, , 2, , 1)
    Call fSetConditionFormatForOddEvenLine(shtReportDetails, , 2, , 1)

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
    'Erase arrFiles
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
    
    sMsg = "处理完成, please check the sheet : " & shtReportDetails.name _
         & IIf(dictLog.Count > 0, vbCr & vbCr & "但是有异常数据, 请检查.", "")
    
    MsgBox sMsg, IIf(dictLog.Count > 0, vbCritical, vbInformation)
    Set dictLog = Nothing
    
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub

Function fReadFile3Letters(arrFirst)
    Dim dictOut As New Dictionary
    
    Dim lEach As Long
    Dim sFile As String
    Dim sLetter As String
    
    For lEach = LBound(arrFirst) To UBound(arrFirst)
        sFile = arrFirst(lEach)
        sLetter = Left(fGetFileNetName(sFile), 3)
        'sLetter = Replace(fGetFileNetName(sFile), 3)
        
        If Not dictOut.Exists(sLetter) Then
            dictOut.Add sLetter, sFile
        Else
            dictLog.Add dictLog.Count + 1 & DELIMITER & "duplicate file's letter " & sLetter & " was found : " & sFile, ""
        End If
    Next
    
    Set fReadFile3Letters = dictOut
    Set dictOut = Nothing
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
    If MsgBox("现在的数据将会被清空, 您确定要继续吗? ", vbYesNoCancel + vbCritical + vbDefaultButton3) <> vbYes Then Exit Function
    
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
  
