Attribute VB_Name = "MA_0_Customize"
Option Explicit
Option Base 1

Function fSetBackToConfigSheetAndUpdategDict_UserTicket()
    Dim ckb As Object
    
    'Dim eachObj As Object
    
    'for each eachobj in shtmenu.
    Dim i As Long
    Dim sCompanyID As String
    Dim sTickValue As String
    
    For i = 0 To dictCompList.Count - 1
        sCompanyID = dictCompList.Keys(i)
         
        If Not fActiveXControlExistsInSheet(shtCurrMenu, fGetCompany_CheckBoxName(sCompanyID), ckb) Then GoTo next_company
        
        sTickValue = IIf(ckb.value, "Y", "N")
        
        Call fSetSpecifiedConfigCellValue(shtStaticData, "[Sales Company List]", "User Ticked", "Company ID=" & sCompanyID, sTickValue)
        Call fUpdateDictionaryItemValueForDelimitedElement(dictCompList, sCompanyID, Company.Selected - Company.REPORT_ID, sTickValue)
next_company:
    Next
End Function

Function fSetBackToConfigSheetAndUpdategDict_InputFiles()
    Dim i As Integer
    Dim sEachCompanyID As String
    Dim sFilePathRange As String
    Dim sEachFilePath  As String
    
    For i = 0 To dictCompList.Count - 1
        sEachCompanyID = dictCompList.Keys(i)
        'sFilePathRange = "rngSalesFilePath_" & sEachCompanyID
        
        If fGetCompany_UserTicked(sEachCompanyID) = "Y" Then
            sFilePathRange = fGetCompany_InputFileTextBoxName(sEachCompanyID)
            sEachFilePath = Trim(shtCurrMenu.Range(sFilePathRange).value)
        Else
            sEachFilePath = "User not selected."
        End If
         
        Call fSetValueBackToSysConf_InputFile_FileName(sEachCompanyID, sEachFilePath)
        Call fUpdateGDictInputFile_FileName(sEachCompanyID, sEachFilePath)
        
        'Call fSetSalesInfoFileToMainConfig(sEachCompanyID, sEachFilePath)
    Next
    
    
'    sFile = Trim(shtMenu.Range("rngSalesFilePath_GY").Value)
'
'    Call fSetValueBackToSysConf_InputFile_FileName("GY", sFile)
'    Call fUpdateGDictInputFile_FileName("GY", sFile)
    
    
End Function

Function fSetIntialValueForShtMenuInitialize()
    
End Function

Function fInitialization()
    Err.Clear
   ' gbBusinessError = False
    gErrNum = 0
    gErrMsg = ""
    
    Call fDisableExcelOptionsAll
    
    Application.ScreenUpdating = True
    Application.ScreenUpdating = True   ' for testing
    
    Call fRemoveFilterForAllSheets
End Function

Function fSetConditionFormatForFundamentalSheets()
'    Call fClearConditionFormatAndAdd(shtCompanyNameReplace, Array(1, 2), True)
'    Call fClearConditionFormatAndAdd(shtHospital, Array(1), True)
'    Call fClearConditionFormatAndAdd(shtHospitalReplace, Array(1, 2), True)
'    Call fClearConditionFormatAndAdd(shtProductMaster, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtProductNameReplace, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtProductProducerReplace, Array(1, 2), True)
'    Call fClearConditionFormatAndAdd(shtProductSeriesReplace, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtProductUnitRatio, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtProductProducerMaster, Array(1), True)
'    Call fClearConditionFormatAndAdd(shtProductNameMaster, Array(1, 2), True)
'    Call fClearConditionFormatAndAdd(shtSalesManMaster, Array(1), True)
'    Call fClearConditionFormatAndAdd(shtFirstLevelCommission, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtSecondLevelCommission, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtSelfSalesOrder, Array(1, 2, 3), True)         'to-do
'    Call fClearConditionFormatAndAdd(shtSelfSalesPreDeduct, Array(1, 2, 3, 4), True)       'to-do
'    Call fClearConditionFormatAndAdd(shtSalesManCommConfig, Array(1, 2, 3, 4, 5, 6), True)
'    Call fClearConditionFormatAndAdd(shtException, Array(1), True)
'
'    Call fClearConditionFormatAndAdd(shtSelfPurchaseOrder, Array(1, 2, 3, 4, 5), True)       'to-do
'    Call fClearConditionFormatAndAdd(shtSelfInventory, Array(1, 2, 3, 5), True)       'to-do
'
'    Call fClearConditionFormatAndAdd(shtNewRuleProducts, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtPromotionProduct, Array(2, 3, 4, 5), True)
'    Call fClearConditionFormatAndAdd(shtCZLInventory, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtSelfInventory, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtCZLPurchaseOrder, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtCZLInventory, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtCZLInvDiff, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtCZLRolloverInv, Array(1, 2, 3), True)
'    Call fClearConditionFormatAndAdd(shtSalesCompInvCalcd, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtSalesCompInvUnified, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtSalesCompRolloverInv, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtSalesCompInvDiff, Array(1, 2, 3, 4), True)
'    Call fClearConditionFormatAndAdd(shtProductTaxRate, Array(ProdTaxRate.ProductProducer, ProdTaxRate.ProductName, ProdTaxRate.ProductSeries, ProdTaxRate.TaxRate), True)
'
'    Call fClearConditionFormatAndAdd(shtRefund, Array(1, 2, 3, 4, 5), True)
End Function

Function fClearConditionFormatAndAdd(sht As Worksheet, arrKeysCols, Optional bExtendToMore10ThousRows As Boolean = True)
    Call fDeleteAllConditionFormatFromSheet(sht)
    Call fSetConditionFormatForOddEvenLine(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
    Call fSetConditionFormatForBorders(sht, , , , arrKeysCols, bExtendToMore10ThousRows)
End Function
