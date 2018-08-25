VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Base 1

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Call fShowSheet(shtRawData)
End Sub

Private Sub Workbook_Open()
    Application.EnableEvents = False
    Call sub_WorkBookInitialization
    
    Call fHideSheet(shtSysConf)
 
    Application.EnableEvents = True
End Sub

Sub sub_WorkBookInitialization()
    'Call fHideSheet(shtRawData)
'    Call fShowSheet(shtRawData)
'    Call fHideSheet(shtCabinet)
'    Call fHideSheet(shtCabinetFrame)
'    Call fHideSheet(shtDoor)
'    Call fHideSheet(shtHardwares)
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    subRefreshRibbon
End Sub
