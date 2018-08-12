Attribute VB_Name = "MB_1_RibbonUI"
Option Explicit

#If VBA7 And Win64 Then  'Win64
    Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#Else
    Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

Public mRibbonObj As IRibbonUI

'=============================================================
Sub subRefreshRibbon()
    fGetRibbonReference.Invalidate
End Sub
Sub ERP_UI_Onload(ribbon As IRibbonUI)
  Set mRibbonObj = ribbon
  
  fCreateAddNameUpdateNameWhenExists "nmRibbonPointer", ObjPtr(ribbon)
  'Names("nmRibbonPointer").RefersTo = ObjPtr(ribbon)
  
  mRibbonObj.ActivateTab "ERP_2010"
  ThisWorkbook.Saved = True
End Sub
Function fGetRibbonReference() As IRibbonUI
    If Not mRibbonObj Is Nothing Then Set fGetRibbonReference = mRibbonObj: Exit Function
    
    Dim objRibbon As Object
    Dim lRibPointer As LongPtr
    
    lRibPointer = [nmRibbonPointer]
    
    CopyMemory objRibbon, lRibPointer, LenB(lRibPointer)
    
    Set fGetRibbonReference = objRibbon
    Set mRibbonObj = objRibbon
    Set objRibbon = Nothing
End Function
'---------------------------------------------------------------------
Sub Button_onAction(control As IRibbonControl)
    Call fGetControlAttributes(control, "ACTION")
End Sub
Sub Button_getImage(control As IRibbonControl, ByRef imageMso)
    Call fGetControlAttributes(control, "IMAGE", imageMso)
End Sub
Sub Button_getLabel(control As IRibbonControl, ByRef label)
    Call fGetControlAttributes(control, "LABEL", label)
End Sub
Sub Button_getSize(control As IRibbonControl, ByRef size)
    Call fGetControlAttributes(control, "SIZE", size)
End Sub

'================== toggle button common function===========================================
Sub ToggleButtonToSwitchSheet_onAction(control As IRibbonControl, pressed As Boolean)
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag(control.Tag)
    
    If Not sht Is Nothing Then
        fToggleSheetVisibleFromUIRibbonControl pressed, sht, control
    End If
    Set sht = Nothing
End Sub

Sub ToggleButtonToSwitchSheet_getPressed(control As IRibbonControl, ByRef returnedVal)
    Dim sht As Worksheet
    Set sht = fGetSheetByUIRibbonTag(control.Tag)
    
    If sht Is Nothing Then
        returnedVal = False
    Else
        returnedVal = (sht.Visible = xlSheetVisible And ActiveSheet Is sht)
    End If
End Sub
Function fGetSheetByUIRibbonTag(ByVal asButtonTag As String) As Worksheet
    Dim sht As Worksheet
    
    If fSheetExistsByCodeName(asButtonTag, sht) Then
        Set fGetSheetByUIRibbonTag = sht
    Else
        MsgBox "The button's Tag is not corresponding to any worksheet in this workbook, please check the customUI.xml you prepared," _
            & " The design thought is that the button's tag is the name of a sheet, so that the common function ToggleButtonToSwitchSheet_onAction/getPressed can get a worksheet." _
            & vbCr & vbCr & "asButtonTag: " & asButtonTag
            
    End If
    Set sht = Nothing
End Function
Function fToggleSheetVisibleFromUIRibbonControl(ByVal pressed As Boolean, sht As Worksheet, control As IRibbonControl)
    If pressed Then
        If ActiveSheet.CodeName <> sht.CodeName Then
            fActiveVisibleSwitchSheet sht
        End If
    Else
        If ActiveSheet.CodeName <> sht.CodeName Then
            fActiveVisibleSwitchSheet sht
        Else
            If fWorkbookHasMoreThanOneSheetVisible(ThisWorkbook) Then
                fVeryHideSheet sht
            End If
        End If
    End If
    
    'fGetRibbonReference.InvalidateControl (control.id)
    fGetRibbonReference.Invalidate
End Function

'---------------------------------------------------------------------

'==========================dev prod switch===================================
Sub btnSwitchDevProd_onAction(control As IRibbonControl, pressed As Boolean)
    sub_SwitchDevProdMode
End Sub

Sub btnSwitchDevProd_getPressed(control As IRibbonControl, ByRef returnedVal)
    returnedVal = fIsDev()
End Sub
Sub btnSwitchDevProd_getVisible(control As IRibbonControl, ByRef returnedVal)
    'returnedVal = fIsDev()
    returnedVal = True
End Sub
Sub grpDevFacilities_getVisible(control As IRibbonControl, ByRef returnedVal)
    returnedVal = fIsDev()
End Sub
'---------------------------------------------------------------------

'================ dev facilities ==============================================
Sub btnListAllFunctions_onAction(control As IRibbonControl)
    sub_ListAllFunctionsOfThisWorkbook
End Sub
Sub btnExportSourceCode_onAction(control As IRibbonControl)
    sub_ExportModulesSourceCodeToFolder
End Sub
Sub btnGenNumberList_onAction(control As IRibbonControl)
    sub_GenNumberList
End Sub
Sub btnGenAlphabetList_onAction(control As IRibbonControl)
    sub_GenAlpabetList
End Sub
Sub btnListAllActiveXOnCurrSheet_onAction(control As IRibbonControl)
    Sub_ListActiveXControlOnActiveSheet
End Sub
Sub btnResetOnError_onAction(control As IRibbonControl)
    sub_ResetOnError_Initialize
End Sub
'------------------------------------------------------------------------------
 
Function fGetControlAttributes(control As IRibbonControl, sType As String, Optional ByRef val)
    If Not (sType = "LABEL" Or sType = "IMAGE" Or sType = "SIZE" Or sType = "ACTION") Then
        fErr "wrong param to fGetControlAttributes: " & vbCr & "sType=" & sType & vbCr & "control=" & control.id
    End If
    
    Select Case control.id
        Case "btnCalSummaryAmount"
            Select Case sType
                Case "LABEL":   val = "计算入帐出帐结果"
                Case "IMAGE":   val = "FunctionWizard"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_CalculateBillInOut
            End Select
        Case "tbtnShowSummaryAmount"
            Select Case sType
                Case "LABEL":   val = "显示/隐藏" & vbCr & "入帐出帐汇总表"
                Case "IMAGE":   val = "ChartShowData"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                'Case "ACTION":      Call fClearRefVariables
            End Select
        Case "tbtnShowshtBillIn"
            Select Case sType
                Case "LABEL":   val = "显示/隐藏" & vbCr & "入帐表"
                Case "IMAGE":   val = "FileSaveAsExcelXlsx"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                'Case "ACTION":      Call fClearRefVariables
            End Select
        Case "tbtnShowshtBillOut"
            Select Case sType
                Case "LABEL":   val = "显示/隐藏" & vbCr & "出帐表"
                Case "IMAGE":   val = "FileSaveAsExcelXlsx"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                'Case "ACTION":      Call fClearRefVariables
            End Select
            
        Case "btnSummaryBusinessData"
            Select Case sType
                Case "LABEL":   val = "汇总明细"
                Case "IMAGE":   val = "FunctionWizard"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_SummarizeBusinssDetail
            End Select
        Case "tbtnShowshtBusinessDetails"
            Select Case sType
                Case "LABEL":   val = "显示/隐藏" & vbCr & "明细表"
                Case "IMAGE":   val = "FileSaveAsExcelXlsx"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                'Case "ACTION":      Call fClearRefVariables
            End Select
            
        Case "tbtnShowshtBusinessSumm"
            Select Case sType
                Case "LABEL":   val = "显示/隐藏" & vbCr & "汇总表"
                Case "IMAGE":   val = "FileSaveAsExcelXlsx"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                'Case "ACTION":      Call fClearRefVariables
            End Select
        Case "btnClearData"
            Select Case sType
                Case "LABEL":   val = "清除明细表中数据"
                Case "IMAGE":   val = "ReviewRejectChange"
                Case "SIZE":        val = "true"    'large=true, normal=false
                Case "SHOW_IMAGE":  val = "true"
                Case "SUPERTIP":    val = ""
                Case "SUPERTIP":    val = ""
                Case "SCREENTIP":   val = ""
                
                Case "ENABLED":     val = True
                Case "ACTION":      Call subMain_ClearBuzDetails
            End Select

    End Select
    
End Function
 
