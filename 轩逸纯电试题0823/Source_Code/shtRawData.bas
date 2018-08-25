VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtRawData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Enum RawData
    QType = 3
    QDesc = 5
    AnswerOptions = 6
    CorrectAnswer = 7
End Enum

Public Property Get HeaderByRow()
    HeaderByRow = 2
End Property
Public Property Get DataFromRow()
    DataFromRow = HeaderByRow() + 1
End Property


