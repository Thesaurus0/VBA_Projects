VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtPurchaseODByVendor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Enum PODByVendor
    [_first] = 1
    VendorName = 1
    ProdName = 1
    Price = 3
    Qty = 4
    Remarks = 5
    [_last] = Remarks
End Enum

Public Property Get HeaderByRow()
    HeaderByRow = 1
End Property
Public Property Get DataFromRow()
    DataFromRow = HeaderByRow() + 1
End Property


