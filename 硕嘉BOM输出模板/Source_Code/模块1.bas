Attribute VB_Name = "模块1"
Sub subMain_ConsolidateAndGenReports()

StartTime = Timer
Application.ScreenUpdating = False


    
    Call fShowSheet(shtRawData)
    Call fShowSheet(shtCabinet)
    Call fShowSheet(shtCabinetFrame)
    Call fShowSheet(shtDoor)
    Call fShowSheet(shtHardwares)

Dim num1 As Integer '循环第一遍
Dim num2 As Integer
Dim num3 As Integer
Dim num4 As Integer
Dim num0 As Integer
Dim c1 As Double    'Integer       '合计
Dim c2 As Double
Dim c3 As Double
Dim c4 As Double
Dim c0 As Integer       '成品数量


Dim name As String


Dim startrow As Integer
Dim endrow As Integer

Dim ct0 As Integer
Dim ct1 As Integer  '计数 sh2 设置初始行位置(封边条统计)
Dim ct2 As Integer  '计数 sh2 设置初始行位置（板件）
Dim ct3 As Integer  '计数 sh2 A列序号
Dim ct4 As Integer
Dim ct5 As Integer
Dim ct6 As Integer
Dim ct7 As Integer
Dim ct8 As Integer
Dim ct9 As Integer

Dim i1 As Integer
Dim i2 As Integer
Dim i3 As Integer
Dim i4 As Integer


Dim foundrow1 As Integer
Dim foundrow2 As Integer
Dim foundrow3 As Integer
Dim foundrow4 As Integer


Set sh1 = Worksheets("TopSolid原始数据")
Set sh2 = Worksheets("柜体清单")
Set sh3 = Worksheets("柜框清单")
Set sh4 = Worksheets("门板清单")
Set sh5 = Worksheets("五金清单")




For num0 = 2 To sh1.Range("D65536").End(xlUp).Row Step 1
    sh1.Cells(num0, "D").Value = Trim(sh1.Cells(num0, "D").Text)
Next num0


sh1.Select
Range("A1:AC65536").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

sh2.Select
Range("A7:O65536").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

sh3.Select
Range("A7:N65536").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

sh4.Select
Range("A7:M65536").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With

sh5.Select
Range("A7:N65536").Select
With Selection
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .WrapText = False
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = False
End With



                                                                                        '成品数量更改

startrow = 2
endrow = 2
For num0 = 2 To sh1.Range("M65536").End(xlUp).Row Step 1
    If sh1.Cells(num0, "M").Value = "成品" Then
        endrow = num0


        If endrow > 2 Then

            For num1 = startrow To endrow Step 1
                sh1.Cells(num1, "H").Value = sh1.Cells(num1, "H").Value * sh1.Cells(startrow, "AB").Value
            Next num1
        End If
        startrow = endrow
    End If
Next num0


For num0 = startrow To sh1.Range("M65536").End(xlUp).Row Step 1
    If sh1.Cells(num0, "M").Value = "成品" Then
        For num1 = startrow To sh1.Range("M65536").End(xlUp).Row Step 1
            sh1.Cells(num1, "H").Value = sh1.Cells(num1, "H").Value * sh1.Cells(startrow, "AB").Value
        Next num1
    End If
Next num0

For num0 = 2 To sh1.Range("M65536").End(xlUp).Row Step 1
    If sh1.Cells(num0, "M").Value = "成品" Then
        sh1.Cells(num0, "H").Value = sh1.Cells(num0, "AB").Value
    End If
Next num0






ct1 = 7
ct2 = 7        '板程序放置行 空行
ct3 = 1        'sh2 A列序号
ct4 = 7


i1 = 7
For num1 = 1 To sh1.Range("M65536").End(xlUp).Row Step 1
    If sh1.Cells(num1, "M").Value = "成品" Then
        c0 = sh1.Cells(num1, "AB").Value
        
        sh2.Cells(3, "C").Value = sh1.Cells(num1, "N").Value    '客户名称
        sh2.Cells(4, "C").Value = sh1.Cells(num1, "O").Value    '订单编号
        sh2.Cells(3, "G").Value = sh1.Cells(num1, "P").Value    '客户地址
        sh2.Cells(4, "G").Value = sh1.Cells(num1, "Q").Value    '制表人
        sh2.Cells(3, "K").Value = sh1.Cells(num1, "R").Value    '联系电话
        sh2.Cells(3, "M").Value = sh1.Cells(num1, "S").Value    '制表日期
        
        sh2.Cells(i1, "B").Value = sh1.Cells(num1, "D").Value       '柜体名称
        sh2.Cells(i1 + 1, "B").Value = sh1.Cells(num1, "C").Value & "=" & sh1.Cells(num1, "AB").Value '柜体规格
        
        sh3.Cells(3, "C").Value = sh1.Cells(num1, "N").Value
        sh3.Cells(4, "C").Value = sh1.Cells(num1, "O").Value
        sh3.Cells(3, "G").Value = sh1.Cells(num1, "P").Value
        sh3.Cells(4, "G").Value = sh1.Cells(num1, "Q").Value
        sh3.Cells(3, "K").Value = sh1.Cells(num1, "R").Value
        sh3.Cells(4, "K").Value = sh1.Cells(num1, "S").Value

        
       

        sh4.Cells(3, "C").Value = sh1.Cells(num1, "N").Value
        sh4.Cells(4, "C").Value = sh1.Cells(num1, "O").Value
        sh4.Cells(3, "G").Value = sh1.Cells(num1, "P").Value
        sh4.Cells(4, "G").Value = sh1.Cells(num1, "Q").Value
        sh4.Cells(3, "L").Value = sh1.Cells(num1, "R").Value
        sh4.Cells(4, "L").Value = sh1.Cells(num1, "S").Value

        sh5.Cells(3, "C").Value = sh1.Cells(num1, "N").Value
        sh5.Cells(4, "C").Value = sh1.Cells(num1, "O").Value
        sh5.Cells(3, "F").Value = sh1.Cells(num1, "P").Value
        sh5.Cells(4, "F").Value = sh1.Cells(num1, "Q").Value
        sh5.Cells(3, "I").Value = sh1.Cells(num1, "R").Value
        sh5.Cells(4, "I").Value = sh1.Cells(num1, "S").Value
    
    
    
                                                                                                            '柜体清单
    
    ElseIf InStr(sh1.Cells(num1, "M").Value, "板程序") > 0 Then
        sh2.Cells(ct2, "C").Value = sh1.Cells(num1, "D").Value          '板件名称
        sh2.Cells(ct2, "D").Value = sh1.Cells(num1, "E").Value          '板件成型长
        sh2.Cells(ct2, "E").Value = sh1.Cells(num1, "F").Value          '板件成型宽
        sh2.Cells(ct2, "F").Value = sh1.Cells(num1, "G").Value          '板件成型厚
        sh2.Cells(ct2, "G").Value = sh1.Cells(num1, "H").Value          '板件数量
        sh2.Cells(ct2, "A").Value = ct3
        
        If sh2.Cells(ct2, "E").Value < 330 Then                         '柜体平方宽不足300按330
            sh2.Cells(ct2, "H").Value = sh2.Cells(ct2, "D").Value * 330 * sh2.Cells(ct2, "G").Value / 1000000
        ElseIf sh2.Cells(ct2, "E").Value > 600 Then                     '柜体平方宽大于600乘以1.2
            sh2.Cells(ct2, "H").Value = 1.2 * sh2.Cells(ct2, "D").Value * sh2.Cells(ct2, "E").Value * sh2.Cells(ct2, "G").Value / 1000000
        Else
            sh2.Cells(ct2, "H").Value = sh2.Cells(ct2, "D").Value * sh2.Cells(ct2, "E").Value * sh2.Cells(ct2, "G").Value / 1000000
        End If
        
        If sh2.Cells(ct2, "H").Value < 0.1 Then                         '平方数不足0.1按0.1
            sh2.Cells(ct2, "H").Value = 0.1
        End If
        sh2.Cells(ct2, "H").Value = Round(sh2.Cells(ct2, "H").Value, 2)
        sh2.Cells(ct2, "I").Value = sh1.Cells(num1, "I").Value       '板件材质
        sh2.Cells(ct2, "J").Value = sh1.Cells(num1, "J").Value
        
        sh2.Cells(ct2, "K").Value = sh1.Cells(num1, "X").Value          '纹理要求
        sh2.Cells(ct2, "L").Value = sh2.Cells(4, "C").Value & "-" & sh1.Cells(num1, "AA").Value & "-" & sh1.Cells(num1, "A").Value & "-" & "A"     '正面条码
        sh2.Cells(ct2, "M").Value = sh2.Cells(4, "C").Value & "-" & sh1.Cells(num1, "AA").Value & "-" & sh1.Cells(num1, "A").Value & "-" & "B"     '反面条码
        
        sh2.Cells(ct2, "N").Value = sh1.Cells(num1, "W").Value          '封边要求

        ct2 = ct2 + 1
        i1 = i1 + 1
        ct3 = ct3 + 1

    End If

Next num1


ct1 = ct2 + 2

For num1 = 1 To sh1.Range("M65536").End(xlUp).Row Step 1
    
    If sh1.Cells(num1, "M").Value = "封边外形" Then
        foundrow1 = -1  'sh3设置初始判断值为否   （判断）
            
        For i1 = ct2 + 2 To ct1 - 1
            If sh2.Cells(i1, "C").Value = sh1.Cells(num1, "I").Value & "封边条" Then
                foundrow1 = i1
                Exit For
            End If
        Next i1
            
        If foundrow1 >= 0 Then
            sh2.Cells(foundrow1, "G").Value = sh1.Cells(num1, "E").Value / 1000 + sh2.Cells(foundrow1, "G").Value

        Else
            sh2.Cells(ct1, "B").Value = "封边条合计"
            sh2.Cells(ct1, "C").Value = sh1.Cells(num1, "I").Value & "封边条"
            sh2.Cells(ct1, "G").Value = sh1.Cells(num1, "E").Value / 1000
                
            ct1 = ct1 + 1
            
        End If
    End If
    
Next num1


ct7 = ct1
For num0 = 7 To ct2 - 1 Step 1          '表2 循环所有板件
    foundrow1 = -1
    For i1 = ct7 To ct1 - 1
        If sh2.Cells(i1, "C").Value = sh2.Cells(num0, "F").Value & sh2.Cells(num0, "I").Value Then
            foundrow1 = i1
            Exit For
        End If
    Next i1
    
    If foundrow1 >= 0 Then
        sh2.Cells(foundrow1, "H").Value = sh2.Cells(foundrow1, "H").Value + sh2.Cells(num0, "H").Value
    
    
    Else
        sh2.Cells(ct1, "C").Value = sh2.Cells(num0, "F").Value & sh2.Cells(num0, "I").Value
        sh2.Cells(ct1, "H").Value = sh2.Cells(num0, "H").Value
        
        ct1 = ct1 + 1
    End If
    

Next num0






                                                                                    '柜框清单




ct6 = 7
startrow = 1
endrow = 1
For num0 = 1 To sh1.Range("M65536").End(xlUp).Row Step 1
    If sh1.Cells(num0, "M").Value = "成品" Then
        sh3.Cells(ct6, "B").Value = sh1.Cells(num0, "D").Value
        endrow = num0
    End If
    If endrow > 2 Then

        ct4 = ct6
        For num2 = startrow To endrow Step 1
            If InStr(sh1.Cells(num2, "M").Value, "背板") > 0 Then
                foundrow2 = -1
                For i2 = ct6 To ct4 - 1
                    If sh3.Cells(i2, "C").Value = sh1.Cells(num2, "G").Value & "mm" & "背板" Then
                        foundrow2 = i2
                        Exit For
                    End If
                Next i2
                If foundrow2 >= 0 Then
                    sh3.Cells(foundrow2, "I").Value = sh3.Cells(foundrow2, "I").Value + Round(sh1.Cells(num2, "E").Value * sh1.Cells(num2, "F").Value * sh1.Cells(num2, "H").Value / 1000000, 2)
                    sh3.Cells(foundrow2, "G").Value = sh3.Cells(foundrow2, "G").Value + sh1.Cells(num2, "H").Value
                Else
                    sh3.Cells(ct4, "C").Value = sh1.Cells(num2, "G").Value & "mm" & "背板"
                    sh3.Cells(ct4, "I").Value = Round(sh1.Cells(num2, "E").Value * sh1.Cells(num2, "F").Value * sh1.Cells(num2, "H").Value / 1000000, 2)
                    sh3.Cells(ct4, "K").Value = sh1.Cells(num2, "I").Value
                    sh3.Cells(ct4, "L").Value = sh1.Cells(num2, "J").Value
                    sh3.Cells(ct4, "G").Value = sh1.Cells(num2, "H").Value
                    ct4 = ct4 + 1
                    
                
                End If
            End If
        Next num2
        
        ct5 = ct4
        For num3 = startrow To endrow Step 1
            If InStr(sh1.Cells(num3, "M").Value, "门板") > 0 Then
                foundrow3 = -1
                For i3 = ct4 To ct5 - 1
                    If sh3.Cells(i3, "C").Value = sh1.Cells(num3, "G").Value & "mm" & "门板" Then
                        foundrow3 = i3
                        Exit For
                    End If
                Next i3
                If foundrow3 >= 0 Then
                    sh3.Cells(foundrow3, "J").Value = sh3.Cells(foundrow3, "J").Value + Round(sh1.Cells(num3, "E").Value * sh1.Cells(num3, "F").Value * sh1.Cells(num3, "H").Value / 1000000, 2)
                    sh3.Cells(foundrow3, "G").Value = sh3.Cells(foundrow3, "G").Value + sh1.Cells(num3, "H").Value
                Else
                    sh3.Cells(ct5, "C").Value = sh1.Cells(num3, "G").Value & "mm" & "门板"
                    sh3.Cells(ct5, "J").Value = Round(sh1.Cells(num3, "E").Value * sh1.Cells(num3, "F").Value * sh1.Cells(num3, "H").Value / 1000000, 2)
                    sh3.Cells(ct5, "K").Value = sh1.Cells(num3, "I").Value
                    sh3.Cells(ct5, "L").Value = sh1.Cells(num3, "J").Value
                    sh3.Cells(ct5, "G").Value = sh1.Cells(num3, "H").Value
                    ct5 = ct5 + 1
                End If
            End If
        Next num3
    
        ct6 = ct5
        For num4 = startrow To endrow Step 1
            If sh1.Cells(num4, "M").Value = "板程序" Then
                foundrow4 = -1
                For i4 = ct5 To ct6 - 1
                    If sh3.Cells(i4, "C").Value = sh1.Cells(num4, "G").Value & "mm" & "柜体板" Then
                        foundrow4 = i4
                        Exit For
                    End If
                Next i4
                If foundrow4 >= 0 Then
                    sh3.Cells(foundrow4, "H").Value = sh3.Cells(foundrow4, "H").Value + Round(sh1.Cells(num4, "E").Value * sh1.Cells(num4, "F").Value * sh1.Cells(num4, "H").Value / 1000000, 2)
                    sh3.Cells(foundrow4, "G").Value = sh3.Cells(foundrow4, "G").Value + sh1.Cells(num4, "H").Value
                Else
                    sh3.Cells(ct6, "C").Value = sh1.Cells(num4, "G").Value & "mm" & "柜体板"
                    sh3.Cells(ct6, "H").Value = Round(sh1.Cells(num4, "E").Value * sh1.Cells(num4, "F").Value * sh1.Cells(num4, "H").Value / 1000000, 2)
                    sh3.Cells(ct6, "K").Value = sh1.Cells(num4, "I").Value
                    sh3.Cells(ct6, "L").Value = sh1.Cells(num4, "J").Value
                    sh3.Cells(ct6, "G").Value = sh1.Cells(num4, "H").Value
                    ct6 = ct6 + 1
                End If
            End If
        Next num4
    End If
    startrow = endrow
Next num0



For num0 = startrow To sh1.Range("M65536").End(xlUp).Row Step 1

    If sh1.Cells(num0, "M").Value = "成品" Then
        sh3.Cells(ct6, "B").Value = sh1.Cells(num0, "D").Value
        ct4 = ct6
        For num2 = startrow To sh1.Range("M65536").End(xlUp).Row Step 1
            If InStr(sh1.Cells(num2, "M").Value, "背板") > 0 Then
                foundrow2 = -1
                For i2 = ct6 To ct4 - 1
                    If sh3.Cells(i2, "C").Value = sh1.Cells(num2, "G").Value & "mm" & "背板" Then
                        foundrow2 = i2
                        Exit For
                    End If
                Next i2
                If foundrow2 >= 0 Then
                    sh3.Cells(foundrow2, "I").Value = sh3.Cells(foundrow2, "I").Value + Round(sh1.Cells(num2, "E").Value * sh1.Cells(num2, "F").Value / 1000000, 2)
                    sh3.Cells(foundrow2, "G").Value = sh3.Cells(foundrow2, "G").Value + sh1.Cells(num2, "H").Value
                Else
                    sh3.Cells(ct4, "C").Value = sh1.Cells(num2, "G").Value & "mm" & "背板"
                    sh3.Cells(ct4, "I").Value = Round(sh1.Cells(num2, "E").Value * sh1.Cells(num2, "F").Value / 1000000, 2)
                    sh3.Cells(ct4, "K").Value = sh1.Cells(num2, "I").Value
                    sh3.Cells(ct4, "L").Value = sh1.Cells(num2, "J").Value
                    sh3.Cells(ct4, "G").Value = sh1.Cells(num2, "H").Value
                    ct4 = ct4 + 1
                End If
            End If
        Next num2
        
        ct5 = ct4
        For num3 = startrow To sh1.Range("M65536").End(xlUp).Row Step 1
            If InStr(sh1.Cells(num3, "M").Value, "门板") > 0 Then
                foundrow3 = -1
                For i3 = ct4 To ct5 - 1
                    If sh3.Cells(i3, "C").Value = sh1.Cells(num3, "G").Value & "mm" & "门板" Then
                        foundrow3 = i3
                        Exit For
                    End If
                Next i3
                If foundrow3 >= 0 Then
                    sh3.Cells(foundrow3, "J").Value = sh3.Cells(foundrow3, "J").Value + Round(sh1.Cells(num3, "E").Value * sh1.Cells(num3, "F").Value / 1000000, 2)
                    sh3.Cells(foundrow3, "G").Value = sh3.Cells(foundrow3, "G").Value + sh1.Cells(num3, "H").Value
                Else
                    sh3.Cells(ct5, "C").Value = sh1.Cells(num3, "G").Value & "mm" & "门板"
                    sh3.Cells(ct5, "J").Value = Round(sh1.Cells(num3, "E").Value * sh1.Cells(num3, "F").Value / 1000000, 2)
                    sh3.Cells(ct5, "K").Value = sh1.Cells(num3, "I").Value
                    sh3.Cells(ct5, "L").Value = sh1.Cells(num3, "J").Value
                    sh3.Cells(ct5, "G").Value = sh1.Cells(num3, "H").Value
                    ct5 = ct5 + 1
                End If
            End If
        Next num3
    
        ct6 = ct5
        For num4 = startrow To sh1.Range("M65536").End(xlUp).Row Step 1
            If sh1.Cells(num4, "M").Value = "板程序" Then
                foundrow4 = -1
                For i4 = ct5 To ct6 - 1
                    If sh3.Cells(i4, "C").Value = sh1.Cells(num4, "G").Value & "mm" & "柜体板" Then
                        foundrow4 = i4
                        Exit For
                    End If
                Next i4
                If foundrow4 >= 0 Then
                    sh3.Cells(foundrow4, "H").Value = sh3.Cells(foundrow4, "H").Value + Round(sh1.Cells(num4, "E").Value * sh1.Cells(num4, "F").Value / 1000000, 2)
                    sh3.Cells(foundrow4, "G").Value = sh3.Cells(foundrow4, "G").Value + sh1.Cells(num4, "H").Value
                Else
                    sh3.Cells(ct6, "C").Value = sh1.Cells(num4, "G").Value & "mm" & "柜体板"
                    sh3.Cells(ct6, "H").Value = Round(sh1.Cells(num4, "E").Value * sh1.Cells(num4, "F").Value / 1000000, 2)
                    sh3.Cells(ct6, "K").Value = sh1.Cells(num4, "I").Value
                    sh3.Cells(ct6, "L").Value = sh1.Cells(num4, "J").Value
                    sh3.Cells(ct6, "G").Value = sh1.Cells(num4, "H").Value
                    ct6 = ct6 + 1
                End If
            End If
        Next num4
    End If
Next num0


ct3 = 1                                                                                 '编序号
For num0 = 7 To sh3.Range("C65536").End(xlUp).Row Step 1
    sh3.Cells(num0, "A").Value = ct3
    ct3 = ct3 + 1
Next num0




ct0 = 7                                                                     '调整柜体名称       增加柜框尺寸
For num0 = 1 To sh1.Range("M65536").End(xlUp).Row Step 1
    If sh1.Cells(num0, "M").Value = "成品" Then
        name = sh1.Cells(num0, "D").Value
        
        For num1 = ct0 To sh3.Range("B65536").End(xlUp).Row Step 1
            If sh3.Cells(num1, "B").Value <> "" Then
                sh3.Cells(num1, "B").Value = name
                sh3.Cells(num1, "D").Value = Split(sh1.Cells(num0, "C").Text, "x", 3)(0)
                sh3.Cells(num1, "E").Value = Split(sh1.Cells(num0, "C").Text, "x", 3)(1)
                sh3.Cells(num1, "F").Value = Split(sh1.Cells(num0, "C").Text, "x", 3)(2)
                
                ct0 = num1 + 1
                Exit For
            End If
        Next num1
    End If
Next num0

                                                                                                                                               
                                                                        
                                                                        
                                                                        
                                                                        
                                                                        
                                                                        
                                                                        '统计

ct0 = ct6 + 2
c1 = 0
c2 = 0
c3 = 0
c4 = 0
For num0 = 7 To sh3.Range("G65536").End(xlUp).Row Step 1
        c1 = sh3.Cells(num0, "G").Value
        c2 = sh3.Cells(num0, "H").Value
        c3 = sh3.Cells(num0, "I").Value
        c4 = sh3.Cells(num0, "J").Value
        
        sh3.Cells(ct0, "G").Value = sh3.Cells(ct0, "G").Value + c1
        sh3.Cells(ct0, "H").Value = sh3.Cells(ct0, "H").Value + c2
        sh3.Cells(ct0, "I").Value = sh3.Cells(ct0, "I").Value + c3
        sh3.Cells(ct0, "J").Value = sh3.Cells(ct0, "J").Value + c4
        
        sh3.Cells(ct0, "C").Value = "合计"
Next num0


For num0 = ct2 + 2 To sh2.Range("C65536").End(xlUp).Row Step 1
    sh3.Cells(ct0 + 1, "C").Value = sh2.Cells(num0, "C").Value
    sh3.Cells(ct0 + 1, "G").Value = sh2.Cells(num0, "G").Value
    sh3.Cells(ct0 + 1, "H").Value = sh2.Cells(num0, "H").Value
    ct0 = ct0 + 1
Next num0






                                                                                    '门板清单

 
ct1 = 7
ct2 = 7
startrow = 2
endrow = 2
For num0 = 1 To sh1.Range("M65536").End(xlUp).Row Step 1
    If sh1.Cells(num0, "M").Value = "成品" Then
        sh4.Cells(ct1, "B").Value = sh1.Cells(num0, "D").Value
        endrow = num0
    End If
    
    If endrow > 2 Then
    
        For num1 = startrow To endrow Step 1
            If InStr(sh1.Cells(num1, "M").Value, "门板") > 0 Then
                foundrow = -1
                For i = ct2 To ct1 - 1
                    
                    If sh4.Cells(i, "C").Value = sh1.Cells(num1, "D").Value Then
                        foundrow = i
                        Exit For
                    End If
                Next i
                If foundrow >= 0 Then
                    sh4.Cells(foundrow, "G").Value = sh4.Cells(foundrow, "G").Value + sh1.Cells(num1, "H").Value
                    sh4.Cells(foundrow, "H").Value = sh4.Cells(foundrow, "H").Value + Round(sh1.Cells(num1, "E").Value * sh1.Cells(num1, "F").Value * sh1.Cells(num1, "H").Value / 1000000, 2)
        
        
                Else
                    sh4.Cells(ct1, "C").Value = sh1.Cells(num1, "D").Value
                    sh4.Cells(ct1, "D").Value = sh1.Cells(num1, "E").Value
                    sh4.Cells(ct1, "E").Value = sh1.Cells(num1, "F").Value
                    sh4.Cells(ct1, "F").Value = sh1.Cells(num1, "G").Value
                    sh4.Cells(ct1, "G").Value = sh1.Cells(num1, "H").Value
                    sh4.Cells(ct1, "H").Value = Round(sh1.Cells(num1, "E").Value * sh1.Cells(num1, "F").Value * sh1.Cells(num1, "H").Value / 1000000, 2)
                    sh4.Cells(ct1, "I").Value = sh1.Cells(num1, "I").Value
                    sh4.Cells(ct1, "J").Value = sh1.Cells(num1, "J").Value
                    sh4.Cells(ct1, "K").Value = sh1.Cells(num1, "X").Value
                    ct1 = ct1 + 1
                End If
            End If
        Next num1
    End If
    startrow = endrow
    ct2 = ct1
Next num0


For num0 = startrow To sh1.Range("M65536").End(xlUp).Row Step 1
    If sh1.Cells(num0, "M").Value = "成品" Then
        sh4.Cells(ct1, "B").Value = sh1.Cells(num0, "D").Value

        For num1 = startrow To sh1.Range("M65536").End(xlUp).Row Step 1
            If InStr(sh1.Cells(num1, "M").Value, "门板") > 0 Then
                foundrow = -1
                For i = ct2 To ct1 - 1
                    
                    If sh4.Cells(i, "C").Value = sh1.Cells(num1, "D").Value Then
                        foundrow = i
                        Exit For
                    End If
                Next i
                If foundrow >= 0 Then
                    sh4.Cells(foundrow, "G").Value = sh4.Cells(foundrow, "G").Value + sh1.Cells(num1, "H").Value
                    sh4.Cells(foundrow, "H").Value = sh4.Cells(foundrow, "H").Value + Round(sh1.Cells(num1, "E").Value * sh1.Cells(num1, "F").Value * sh1.Cells(num1, "H").Value / 1000000, 2)
        
        
                Else
                    sh4.Cells(ct1, "C").Value = sh1.Cells(num1, "D").Value
                    sh4.Cells(ct1, "D").Value = sh1.Cells(num1, "E").Value
                    sh4.Cells(ct1, "E").Value = sh1.Cells(num1, "F").Value
                    sh4.Cells(ct1, "F").Value = sh1.Cells(num1, "G").Value
                    sh4.Cells(ct1, "G").Value = sh1.Cells(num1, "H").Value
                    sh4.Cells(ct1, "H").Value = Round(sh1.Cells(num1, "E").Value * sh1.Cells(num1, "F").Value * sh1.Cells(num1, "H").Value / 1000000, 2)
                    sh4.Cells(ct1, "I").Value = sh1.Cells(num1, "I").Value
                    sh4.Cells(ct1, "J").Value = sh1.Cells(num1, "J").Value
                    sh4.Cells(ct1, "K").Value = sh1.Cells(num1, "X").Value
                    ct1 = ct1 + 1
                End If
            End If
        Next num1
    End If
Next num0

ct0 = 7                                                                     '调整柜体名称
For num0 = 1 To sh1.Range("M65536").End(xlUp).Row Step 1
    If sh1.Cells(num0, "M").Value = "成品" Then
        name = sh1.Cells(num0, "D").Value
        For num1 = ct0 To sh4.Range("B65536").End(xlUp).Row Step 1
            If sh4.Cells(num1, "B").Value <> "" Then
                sh4.Cells(num1, "B").Value = name
                ct0 = num1 + 1
                Exit For
            End If
        Next num1
    End If
Next num0

ct3 = 1                                                                                 '编序号
For num0 = 7 To sh4.Range("C65536").End(xlUp).Row Step 1
    sh4.Cells(num0, "A").Value = ct3
    ct3 = ct3 + 1
Next num0



ct0 = ct1 + 2
c1 = 0
c2 = 0

For num0 = 7 To sh4.Range("G65536").End(xlUp).Row Step 1
        c1 = sh4.Cells(num0, "G").Value
        c2 = sh4.Cells(num0, "H").Value

        
        sh4.Cells(ct0, "G").Value = sh4.Cells(ct0, "G").Value + c1
        sh4.Cells(ct0, "H").Value = sh4.Cells(ct0, "H").Value + c2

        
        sh4.Cells(ct0, "C").Value = "合计"
Next num0

                                    


'___________                                                                五金清单

ct9 = 2
For num0 = 2 To sh1.Range("D65536").End(xlUp).Row Step 1
    ct9 = ct9 + 1
Next num0

For num0 = 2 To ct9 - 1 Step 1
    If sh1.Cells(num0, "V").Value = "" Then
        sh1.Cells(num0, "V").Value = sh1.Cells(num0 - 1, "V").Value
    End If
Next num0

ct9 = 7
For num0 = 2 To sh1.Range("D65536").End(xlUp).Row Step 1
    If sh1.Cells(num0, "M").Value = "五金件" Then
        foundrow1 = -1
        For i1 = 7 To ct9 - 1
            If sh5.Cells(i1, "B").Value = sh1.Cells(num0, "V").Value And sh1.Cells(num0, "V").Value = sh1.Cells(num0 - 1, "V").Value And sh5.Cells(i1, "C").Value = sh1.Cells(num0, "D").Value And sh5.Cells(i1, "D").Value = sh1.Cells(num0, "E").Value And sh5.Cells(i1, "E").Value = sh1.Cells(num0, "F").Value And sh5.Cells(i1, "F").Value = sh1.Cells(num0, "G").Value Then
                foundrow1 = i1
                Exit For
            End If
        
        
        Next i1
        
        If foundrow1 >= 0 Then
            sh5.Cells(foundrow1, "G").Value = sh5.Cells(foundrow1, "G").Value + sh1.Cells(num0, "H").Value
        Else
            sh5.Cells(ct9, "B").Value = sh1.Cells(num0, "V").Value
            sh5.Cells(ct9, "C").Value = sh1.Cells(num0, "D").Value
            sh5.Cells(ct9, "D").Value = sh1.Cells(num0, "E").Value
            sh5.Cells(ct9, "E").Value = sh1.Cells(num0, "F").Value
            sh5.Cells(ct9, "F").Value = sh1.Cells(num0, "G").Value
            sh5.Cells(ct9, "G").Value = sh1.Cells(num0, "H").Value
            sh5.Cells(ct9, "K").Value = sh1.Cells(num0, "I").Value
            ct9 = ct9 + 1
        End If
        
    
    
    End If
Next num0




'__________





ct3 = 1                                                                                '编序号
For num0 = 7 To sh5.Range("C65536").End(xlUp).Row Step 1
    sh5.Cells(num0, "A").Value = ct3
    ct3 = ct3 + 1
Next num0



ct0 = ct9 + 2
c1 = 0

For num0 = 7 To sh5.Range("G65536").End(xlUp).Row Step 1
        c1 = sh5.Cells(num0, "G").Value

        sh5.Cells(ct0, "G").Value = sh5.Cells(ct0, "G").Value + c1

        sh5.Cells(ct0, "C").Value = "合计"
Next num0




Application.DisplayAlerts = 0

For num9 = sh5.Range("B65536").End(xlUp).Row To 7 Step -1
    If sh5.Cells(num9 - 1, "B").Value = sh5.Cells(num9, "B").Value Then
       Range(sh5.Cells(num9 - 1, "B"), sh5.Cells(num9, "B")).Merge
    End If
Next num9



Call fSetConditionalFormatForBorders
Call fSetRowHeightForAllReportSheets


Application.ScreenUpdating = True
MsgBox Timer - StartTime

End Sub







