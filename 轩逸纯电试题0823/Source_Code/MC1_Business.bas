Attribute VB_Name = "MC1_Business"
Option Explicit
Option Base 1

Sub subMain_ConsolidateAndGenReports()
    
    Dim arrRawData()
    Dim arrOutput()
    Dim dictLog As Dictionary
    Dim lEachRow As Long
    Dim lMaxRow As Long
    Dim sAnswerOptions As String
    Dim sAnswer As String
    Dim arrAOptions
    Dim arrAnswer
    Dim sEachAns As String
    Dim i As Integer
    Dim iAnswerAt As Integer
    Dim sMsg As String
    Dim dblTmp As Double
    
   ' On Error GoTo error_handling
    
    Call fInitialization
    
    Call fDeleteRowsFromSheetLeaveHeader(shtOutput)
    Call fDeleteRowsFromSheetLeaveHeader(shtLog)
    
    Call fCopyReadWholeSheetData2Array(shtRawData, arrRawData, , shtRawData.DataFromRow)
     
    If fArrayIsEmptyOrNoData(arrRawData) Then fErr "No data was found in sheet " & shtRawData.name
    
    lMaxRow = ArrLen(arrRawData)
    
    Set dictLog = New Dictionary
    ReDim arrOutput(LBound(arrRawData, 1) To UBound(arrRawData, 1), Rpt.[_first] To Rpt.[_last])
    
    For lEachRow = LBound(arrRawData, 1) To UBound(arrRawData, 1)
        arrOutput(lEachRow, Rpt.QType) = Trim(arrRawData(lEachRow, RawData.QType))
        arrOutput(lEachRow, Rpt.QDesc) = Trim(arrRawData(lEachRow, RawData.QDesc))
        arrOutput(lEachRow, Rpt.Point) = 1
        
        sAnswerOptions = Trim(arrRawData(lEachRow, RawData.AnswerOptions))
        
        If Len(sAnswerOptions) <= 0 Then GoTo next_row
        
        sAnswerOptions = Replace(sAnswerOptions, "丨", "|")
        sAnswerOptions = Replace(sAnswerOptions, " ", "")
        arrAOptions = Split(sAnswerOptions, "|")
                
        If UBound(arrAOptions) >= 0 Then arrOutput(lEachRow, Rpt.AOptionA) = "A. " & arrAOptions(0)
        If UBound(arrAOptions) >= 1 Then arrOutput(lEachRow, Rpt.AOptionB) = "B. " & arrAOptions(1)
        If UBound(arrAOptions) >= 2 Then arrOutput(lEachRow, Rpt.AOptionC) = "C. " & arrAOptions(2)
        If UBound(arrAOptions) >= 3 Then arrOutput(lEachRow, Rpt.AOptionD) = "D. " & arrAOptions(3)
        If UBound(arrAOptions) >= 4 Then arrOutput(lEachRow, Rpt.AOptionE) = "E. " & arrAOptions(4)
        If UBound(arrAOptions) >= 5 Then arrOutput(lEachRow, Rpt.AOptionF) = "F. " & arrAOptions(5)
        
        sAnswer = Trim(arrRawData(lEachRow, RawData.CorrectAnswer))
        
        If Len(sAnswer) <= 0 Then GoTo next_row
        
        sAnswer = Replace(sAnswer, "丨", "|")
        sAnswer = Replace(sAnswer, " ", "")
        arrAnswer = Split(sAnswer, "|")
        
        For i = LBound(arrAnswer) To UBound(arrAnswer)
            sEachAns = Trim(arrAnswer(i))
            dblTmp = val(sEachAns)
            If dblTmp < 1 And dblTmp > 0 Then sEachAns = "0" & sEachAns
            
            iAnswerAt = fStrInDelimiteredStr(sAnswerOptions, sEachAns, "|")
            If iAnswerAt <= 0 Then
                dictLog(lEachRow + shtRawData.HeaderByRow) = dictLog(lEachRow + shtRawData.HeaderByRow) & vbLf & sEachAns & " is not in the answer options: " & sAnswerOptions
            Else
                'arrOutput(lEachRow, Rpt.CorrectAnswer) = arrOutput(lEachRow, Rpt.CorrectAnswer) & "," & fNum2Letter(iAnswerAt)
                arrOutput(lEachRow, Rpt.CorrectAnswer) = arrOutput(lEachRow, Rpt.CorrectAnswer) & fNum2Letter(iAnswerAt)
            End If
        Next
        
        If dictLog.Exists(lEachRow + shtRawData.HeaderByRow) Then
            dictLog(lEachRow + shtRawData.HeaderByRow) = Right(dictLog(lEachRow + shtRawData.HeaderByRow), Len(dictLog(lEachRow + shtRawData.HeaderByRow)) - 1)
        End If
        
'        If Left(arrOutput(lEachRow, Rpt.CorrectAnswer), 1) = "," Then
'            arrOutput(lEachRow, Rpt.CorrectAnswer) = Right(arrOutput(lEachRow, Rpt.CorrectAnswer), Len(arrOutput(lEachRow, Rpt.CorrectAnswer)) - 1)
'        End If
next_row:
    Next
     
    Call fAppendArray2Sheet(shtOutput, arrOutput)
    Call fSetConditionFormatForBorders(shtOutput, , shtOutput.DataFromRow, , Array(Rpt.QType, Rpt.QDesc))
    Call fSetConditionFormatForOddEvenLine(shtOutput, , shtOutput.DataFromRow)
    
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
    Erase arrOutput
    
    If gErrNum <> 0 Then GoTo reset_excel_options
    If fCheckIfUnCapturedExceptionAbnormalError Then GoTo reset_excel_options
    
    Application.ScreenUpdating = True
    If dictLog.Count > 0 Then
        Call fShowAndActiveSheet(shtLog)
    Else
        Call fShowAndActiveSheet(shtOutput)
    End If
    
    sMsg = "Transformation is completed, please check the sheet : " & shtOutput.name _
         & IIf(dictLog.Count > 0, vbCr & vbCr & "but there are exception data whose Answer cannot be identifed among the answer options, you can check the details in the log sheet", "")
    
    MsgBox sMsg, IIf(dictLog.Count > 0, vbCritical, vbInformation)
    Set dictLog = Nothing
    
reset_excel_options:
    Err.Clear
    fClearGlobalVarialesResetOption
End Sub

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
    'Call fDeleteRowsFromSheetLeaveHeader(shtOutput)
    Call fDeleteRowsFromSheetLeaveHeader(shtRawData, shtRawData.HeaderByRow)
    Call fDeleteRowsFromSheetLeaveHeader(shtLog)
     
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
'        Set rgFound = fFindInWorksheet(shtCabinetFrame.Columns("C"), "合计")
'        fGetRangeByStartEndPos(shtCabinetFrame, rgFound.Row - 2, 1, lMaxRow, 1).ClearContents
'        fGetRangeByStartEndPos(shtCabinetFrame, 7, fLetter2Num("C"), lMaxRow, fLetter2Num("C")).HorizontalAlignment = xlLeft
'    End If
'
'    lMaxRow = fGetValidMaxRow(shtDoor)
'    If lMaxRow >= 7 Then
'        Set rgFound = fFindInWorksheet(shtDoor.Columns("C"), "合计")
'        fGetRangeByStartEndPos(shtDoor, rgFound.Row - 2, 1, lMaxRow, 1).ClearContents
'    End If
'
'    lMaxRow = fGetValidMaxRow(shtHardwares)
'    If lMaxRow >= 7 Then
'        Set rgFound = fFindInWorksheet(shtHardwares.Columns("C"), "合计")
'        fGetRangeByStartEndPos(shtHardwares, rgFound.Row - 2, 1, lMaxRow, 1).ClearContents
'    End If
'
'    Set rgFound = Nothing
'End Function
'
'
'
