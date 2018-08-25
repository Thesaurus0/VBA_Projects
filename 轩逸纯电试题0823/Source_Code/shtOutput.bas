VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Enum Rpt
    [_first] = 1
    QType = 1
    QDesc = 2
    AOptionA = 3
    AOptionB = 4
    AOptionC = 5
    AOptionD = 6
    AOptionE = 7
    AOptionF = 8
    CorrectAnswer = 9
    QAnalysis = 10
    Point = 11
    [_last] = Point
End Enum

Public Property Get HeaderByRow()
    HeaderByRow = 1
End Property
Public Property Get DataFromRow()
    DataFromRow = HeaderByRow() + 1
End Property


