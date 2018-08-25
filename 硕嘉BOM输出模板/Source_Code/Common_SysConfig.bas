Attribute VB_Name = "Common_SysConfig"
Option Explicit
Option Base 1

Public Const BUSINESS_ERROR_NUMBER = 10000
'Public Const CONFIG_ERROR_NUMBER = 20000
Public Const DELIMITER = "|"

Public gErrNum As Long
Public gErrMsg As String
 
'=======================================
Public dictMiscConfig As Dictionary
'=======================================
Public arrMaster()
'=======================================
Public arrOutput()
'=======================================
Public gFSO As FileSystemObject
Public gRegExp As VBScript_RegExp_55.RegExp
Public Const PW_PROTECT_SHEET = "abcd1234"

'DEBUG.PRINT "ASDFASDFASDF"
    
