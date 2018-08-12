VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtBusinessDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Enum BuzDetail
    [_first] = 1
    VendorName = 2
    PLATFORM = 3
    
    Point_Qty = 8
    Point_Price = 9
    Point_CurrDayPrice = 10
    Point_DaysNum = 11
    Point_Amt = 12
    
    DownLoad_Qty = 14
    DownLoad_Price = 15
    DownLoad_Amt = 16
    
    Credit_Qty = 17
    Credit_Price = 18
    Credit_Amt = 19
    [_last] = 19
End Enum


Property Get DataStartRow()
    DataStartRow = 3
End Property
