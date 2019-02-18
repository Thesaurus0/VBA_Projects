VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "shtVendorMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Enum Vendor
    VendorName = 1
End Enum

Public Property Get HeaderByRow()
    HeaderByRow = 1
End Property
Public Property Get DataFromRow()
    DataFromRow = HeaderByRow() + 1
End Property




