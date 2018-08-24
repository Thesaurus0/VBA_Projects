Attribute VB_Name = "模块2"
Sub subMain_Print()

Dim num0 As Integer
Dim foundrow2 As Integer
Dim foundrow3 As Integer
Dim foundrow4 As Integer
Dim foundrow5 As Integer



    
    Call fShowSheet(shtRawData)
    Call fShowSheet(shtCabinet)
    Call fShowSheet(shtCabinetFrame)
    Call fShowSheet(shtDoor)
    Call fShowSheet(shtHardwares)

Set sh1 = Worksheets("TopSolid原始数据")
Set sh2 = Worksheets("柜体清单")
Set sh3 = Worksheets("柜框清单")
Set sh4 = Worksheets("门板清单")
Set sh5 = Worksheets("五金清单")

foundrow2 = 1
foundrow3 = 1
foundrow4 = 1
foundrow5 = 1


For num0 = 2 To sh2.Range("C65536").End(xlUp).Row Step 1
    If sh2.Cells(num0, "C").value <> "" Then
        foundrow2 = num0
    End If
Next num0



For num0 = 2 To sh3.Range("C65536").End(xlUp).Row Step 1
    If sh3.Cells(num0, "C").value <> "" Then
        foundrow3 = num0
    End If
Next num0



For num0 = 2 To sh4.Range("C65536").End(xlUp).Row Step 1
    If sh4.Cells(num0, "C").value <> "" Then
        foundrow4 = num0
    End If
Next num0


For num0 = 2 To sh5.Range("C65536").End(xlUp).Row Step 1
    If sh5.Cells(num0, "C").value <> "" Then
        foundrow5 = num0
    End If
Next num0

sh2.Activate

sh2.Range(Cells(7, "L"), Cells(foundrow2, "M")).Select
Selection.ClearContents

sh2.Range(Cells(1, "A"), Cells(foundrow2, "O")).Select
Cells.EntireColumn.AutoFit
Call GS


sh3.Activate
sh3.Range(Cells(1, "A"), Cells(foundrow3, "N")).Select
Cells.EntireColumn.AutoFit
Call GS

sh4.Activate
sh4.Range(Cells(1, "A"), Cells(foundrow4, "M")).Select
Cells.EntireColumn.AutoFit
Call GS

sh5.Activate
sh5.Range(Cells(1, "A"), Cells(foundrow5, "N")).Select
Cells.EntireColumn.AutoFit
Call GS





End Sub

Function GS()

    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End Function
