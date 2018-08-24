Attribute VB_Name = "MC1_Business"
Option Explicit
Option Base 1

Function fSetRowHeightForAllReportSheets()
    Dim lMaxRow As Long
    
    Call fClearSerialNoFromSheets
    
    Dim sht As Worksheet
    
    Set sht = shtCabinet
    
    lMaxRow = fGetValidMaxRow(sht)
    If lMaxRow >= 7 Then
        Call fSetRowHeightForExceedingThreshold(sht, 7, lMaxRow, 16)
    End If
    
    Set sht = shtCabinetFrame
    lMaxRow = fGetValidMaxRow(sht)
    If lMaxRow >= 7 Then
        Call fSetRowHeightForExceedingThreshold(sht, 7, lMaxRow, 20)
    End If
    
    Set sht = shtDoor
    lMaxRow = fGetValidMaxRow(sht)
    If lMaxRow >= 7 Then
        Call fSetRowHeightForExceedingThreshold(sht, 7, lMaxRow, 20)
    End If
    
    Set sht = shtHardwares
    lMaxRow = fGetValidMaxRow(sht)
    If lMaxRow >= 7 Then
        Call fSetRowHeightForExceedingThreshold(sht, 7, lMaxRow, 20)
    End If
    
    Set sht = Nothing
End Function

Function fSetRowHeightForExceedingThreshold(sht As Worksheet, lStartRow As Long, lEndRow As Long _
                                        , Optional dblRowHeightThreshold As Double = 16)
    Dim lEachRow As Long
    Dim rgAll As Range
    Dim rgTarget As Range
    
    Set rgAll = sht.Rows(lStartRow & ":" & lEndRow)
    
    For lEachRow = 1 To lEndRow - lStartRow + 1
        
        If rgAll.Rows(lEachRow).RowHeight < dblRowHeightThreshold Then
            If rgTarget Is Nothing Then
                Set rgTarget = rgAll.Rows(lEachRow)
            Else
                Set rgTarget = Union(rgTarget, rgAll.Rows(lEachRow))
            End If
        End If
    Next
    
    If Not rgTarget Is Nothing Then
        rgTarget.RowHeight = dblRowHeightThreshold
    End If
    
    Set rgAll = Nothing
    Set rgTarget = Nothing
End Function

Function fSetConditionalFormatForBorders()
    Dim lMaxRow As Long
    
    Call fClearSerialNoFromSheets
    
    Dim sht As Worksheet
    
    Set sht = shtCabinet
    
    lMaxRow = fGetValidMaxRow(sht)
    If lMaxRow >= 7 Then
        Call fDeleteAllConditionFormatFromSheet(sht)
    '    Call fSetConditionFormatForOddEvenLine(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
        Call fSetConditionFormatForBorders(sht, , 7, , 1)
        sht.Cells.WrapText = True
        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.AutoFit
        'fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, fLetter2Num("K")).EntireColumn.AutoFit
    End If
    
    Set sht = shtCabinetFrame
    lMaxRow = fGetValidMaxRow(sht)
    If lMaxRow >= 7 Then
        Call fDeleteAllConditionFormatFromSheet(sht)
    '    Call fSetConditionFormatForOddEvenLine(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
        Call fSetConditionFormatForBorders(sht, , 7, , 1)
        sht.Cells.WrapText = True
        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.AutoFit
        'sht.Columns.AutoFit
    End If
    
    
    Set sht = shtDoor
    lMaxRow = fGetValidMaxRow(sht)
    If lMaxRow >= 7 Then
        Call fDeleteAllConditionFormatFromSheet(sht)
    '    Call fSetConditionFormatForOddEvenLine(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
        Call fSetConditionFormatForBorders(sht, , 7, , 1)
        sht.Cells.WrapText = True
        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.AutoFit
        'sht.Columns.AutoFit
    End If
    
    
    Set sht = shtHardwares
    lMaxRow = fGetValidMaxRow(sht)
    If lMaxRow >= 7 Then
        Call fDeleteAllConditionFormatFromSheet(sht)
    '    Call fSetConditionFormatForOddEvenLine(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
        Call fSetConditionFormatForBorders(sht, , 7, , 1)
        sht.Cells.WrapText = True
        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.AutoFit
        
       ' sht.Columns.AutoFit
    End If
    
    Set sht = Nothing
End Function

Private Function fSetConditionFormatForBorders(ByRef shtParam As Worksheet, Optional lMaxCol As Long = 0 _
                                            , Optional lRowFrom As Long = 2, Optional lRowTo As Long = 0 _
                                            , Optional arrKeyColsNotBlank _
                                            , Optional bExtendToMore10ThousRows As Boolean = False)
'arrKeyColsNotBlank
'    1. singlecol: 1
'    1. array(1,2,3)
    If lMaxCol = 0 Then lMaxCol = fGetValidMaxCol(shtParam)
    If lRowTo = 0 Then lRowTo = fGetValidMaxRow(shtParam)

    If lMaxCol <= 0 Then Exit Function
    If bExtendToMore10ThousRows Then lRowTo = lRowTo + 100000

    If lRowTo < lRowFrom Then Exit Function

    Dim rngCondFormat As Range
    Set rngCondFormat = fGetRangeByStartEndPos(shtParam, lRowFrom, 1, lRowTo, lMaxCol)

    Dim sAddr As String
    Dim sKeyColsFormula As String
    Dim sFormula As String
    Dim lColor As Long
    Dim i As Integer
    Dim sColLetter As String
    Dim aFormatCondition As FormatCondition

    If Not IsMissing(arrKeyColsNotBlank) Then
        If IsArray(arrKeyColsNotBlank) Then
            For i = LBound(arrKeyColsNotBlank) To UBound(arrKeyColsNotBlank)
                sColLetter = fNum2Letter(arrKeyColsNotBlank(i))
                sKeyColsFormula = sKeyColsFormula & "," & "len(trim($" & sColLetter & lRowFrom & ")) > 0"
            Next
            If Len(sKeyColsFormula) > 0 Then sKeyColsFormula = Right(sKeyColsFormula, Len(sKeyColsFormula) - 1)
        Else
            sColLetter = fNum2Letter(arrKeyColsNotBlank)
            sKeyColsFormula = "len(trim($" & sColLetter & lRowFrom & ")) > 0"
            sKeyColsFormula = sKeyColsFormula
        End If
    Else
        sKeyColsFormula = ""
    End If

    sFormula = "=And( " & sKeyColsFormula & ")"

    Set aFormatCondition = rngCondFormat.FormatConditions.Add(Type:=xlExpression, Formula1:=sFormula)
    aFormatCondition.SetFirstPriority
    aFormatCondition.StopIfTrue = False

    aFormatCondition.Borders(xlLeft).Weight = xlThin
    aFormatCondition.Borders(xlRight).Weight = xlThin
    aFormatCondition.Borders(xlTop).Weight = xlThin
    aFormatCondition.Borders(xlBottom).Weight = xlThin

    aFormatCondition.Borders(xlLeft).ThemeColor = 2
    aFormatCondition.Borders(xlRight).ThemeColor = 2
    aFormatCondition.Borders(xlTop).ThemeColor = 2
    aFormatCondition.Borders(xlBottom).ThemeColor = 2
    
    aFormatCondition.Borders(xlLeft).TintAndShade = 0.499984740745262 '  0.249946592608417
    aFormatCondition.Borders(xlRight).TintAndShade = 0.499984740745262
    aFormatCondition.Borders(xlTop).TintAndShade = 0.499984740745262
    aFormatCondition.Borders(xlBottom).TintAndShade = 0.499984740745262
        
'    aFormatCondition.Borders(xlLeft).Color = -16776961
'    aFormatCondition.Borders(xlRight).Color = -16776961
'    aFormatCondition.Borders(xlTop).Color = -16776961
'    aFormatCondition.Borders(xlBottom).Color = -16776961

    Set aFormatCondition = Nothing
End Function

Function subMain_ClearBuzDetails()
    Dim lMaxRow As Long
    
    Dim sht As Worksheet
    
    Set sht = shtRawData
    lMaxRow = sht.UsedRange.Row + sht.UsedRange.Rows.Count - 1
    If lMaxRow >= 1 Then
        fGetRangeByStartEndPos(sht, 1, 1, lMaxRow, 1).EntireRow.Delete Shift:=xlUp
    End If
    
    Set sht = shtCabinet
    lMaxRow = sht.UsedRange.Row + sht.UsedRange.Rows.Count - 1
    If lMaxRow >= 7 Then
        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.Delete Shift:=xlUp
    End If
    
    Set sht = shtCabinetFrame
    lMaxRow = sht.UsedRange.Row + sht.UsedRange.Rows.Count - 1
    If lMaxRow >= 7 Then
        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.Delete Shift:=xlUp
    End If
    
    
    Set sht = shtDoor
    lMaxRow = sht.UsedRange.Row + sht.UsedRange.Rows.Count - 1
    If lMaxRow >= 7 Then
        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.Delete Shift:=xlUp
    End If
    
    Set sht = shtHardwares
    lMaxRow = sht.UsedRange.Row + sht.UsedRange.Rows.Count - 1
    If lMaxRow >= 7 Then
        fGetRangeByStartEndPos(sht, 7, 1, lMaxRow, 1).EntireRow.Delete Shift:=xlUp
    End If
    
    Set sht = Nothing
End Function

Private Function fClearSerialNoFromSheets()
    Dim rgFound As Range
    Dim lMaxRow As Long
    
    lMaxRow = fGetValidMaxRow(shtCabinet)
    If lMaxRow >= 7 Then
        fGetRangeByStartEndPos(shtCabinet, 7, fLetter2Num("C"), lMaxRow, fLetter2Num("C")).HorizontalAlignment = xlLeft
    End If
    
    lMaxRow = fGetValidMaxRow(shtCabinetFrame)
    If lMaxRow >= 7 Then
        Set rgFound = fFindInWorksheet(shtCabinetFrame.Columns("C"), "合计")
        fGetRangeByStartEndPos(shtCabinetFrame, rgFound.Row - 2, 1, lMaxRow, 1).ClearContents
        fGetRangeByStartEndPos(shtCabinetFrame, 7, fLetter2Num("C"), lMaxRow, fLetter2Num("C")).HorizontalAlignment = xlLeft
    End If
    
    lMaxRow = fGetValidMaxRow(shtDoor)
    If lMaxRow >= 7 Then
        Set rgFound = fFindInWorksheet(shtDoor.Columns("C"), "合计")
        fGetRangeByStartEndPos(shtDoor, rgFound.Row - 2, 1, lMaxRow, 1).ClearContents
    End If
    
    lMaxRow = fGetValidMaxRow(shtHardwares)
    If lMaxRow >= 7 Then
        Set rgFound = fFindInWorksheet(shtHardwares.Columns("C"), "合计")
        fGetRangeByStartEndPos(shtHardwares, rgFound.Row - 2, 1, lMaxRow, 1).ClearContents
    End If
    
    Set rgFound = Nothing
End Function



