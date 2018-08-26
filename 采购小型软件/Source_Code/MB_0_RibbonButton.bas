Attribute VB_Name = "MB_0_RibbonButton"
'Option Explicit
'Option Base 1
'
'Private Const DELETED_FROM_NEW_VERSION = "基础版本有而新版本中没有(被删除的)"
'Private Const SAME_IN_BOTH = "两个版本都有(相同的)"
'Private Const NEWLY_ADDED_IN_NEW_VERSION = "新版本中有而基础版本中没有(新增加的)"
'Private Const BOTH_HAVE_BUT_DIFF_VALUE = "两个版本都有但其他值不同(被修改的)"
'
'Sub subMain_NewRuleProducts()
'    fActiveVisibleSwitchSheet shtNewRuleProducts, , False
'End Sub
'
'Sub subMain_ImportSalesCompanyInventory()
'    fActiveVisibleSwitchSheet shtMenuCompInvt, "A63", False
'End Sub
'Sub subMain_Ribbon_ImportSalesInfoFiles()
'    fActiveVisibleSwitchSheet shtMenu, "A74", False
'End Sub
'
'Sub subMain_Hospital()
'    fActiveVisibleSwitchSheet shtHospital, , False
'    'Call fHideAllSheetExcept(shtHospital, shtHospitalReplace)
'End Sub
'
'Sub subMain_HideHospital()
'    On Error Resume Next
'    shtHospital.Visible = xlSheetVeryHidden
'    Err.Clear
'End Sub
'Sub subMain_HospitalReplacement()
'    fActiveVisibleSwitchSheet shtHospitalReplace, , False
'    'Call fHideAllSheetExcept(shtHospital, shtHospitalReplace)
'End Sub
'
'Sub subMain_Exception()
'    fActiveVisibleSwitchSheet shtException, , False
'    'Call fHideAllSheetExcept(shtHospital, shtHospitalReplace)
'End Sub
'Sub subMain_RawSalesInfos()
'    fActiveVisibleSwitchSheet shtSalesRawDataRpt, , False
'End Sub
'
'Sub subMain_SalesInfos()
'    fActiveVisibleSwitchSheet shtSalesInfos, , False
'End Sub
'
'Sub subMain_ProductMaster()
'    fActiveVisibleSwitchSheet shtPurchaseODRaw, , False
'    'Call fHideAllSheetExcept(shtPurchaseODRaw, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
'End Sub
'
'Sub subMain_HideProductMaster()
'    On Error Resume Next
'    shtPurchaseODRaw.Visible = xlSheetVeryHidden
'    Err.Clear
'End Sub
'Sub subMain_HideProducerMaster()
'    On Error Resume Next
'    shtProductProducerMaster.Visible = xlSheetVeryHidden
'    Err.Clear
'End Sub
'
'Sub subMain_ProducerMaster()
'    fActiveVisibleSwitchSheet shtProductProducerMaster, , False
'    'Call fHideAllSheetExcept(shtPurchaseODRaw, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
'End Sub
'Sub subMain_ProductNameMaster()
'    fActiveVisibleSwitchSheet shtProductNameMaster, , False
'    'Call fHideAllSheetExcept(shtPurchaseODRaw, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
'End Sub
'Sub subMain_HideProductNameMaster()
'    On Error Resume Next
'    shtProductNameMaster.Visible = xlSheetVeryHidden
'    Err.Clear
'End Sub
'Sub subMain_ProductProducerReplace()
'    fActiveVisibleSwitchSheet shtProductProducerReplace, , False
'    'Call fHideAllSheetExcept(shtPurchaseODRaw, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
'End Sub
'Sub subMain_ProductNameReplace()
'    fActiveVisibleSwitchSheet shtProductNameReplace, , False
'    'Call fHideAllSheetExcept(shtPurchaseODRaw, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
'End Sub
'Sub subMain_ProductSeriesReplace()
'    fActiveVisibleSwitchSheet shtProductSeriesReplace, , False
'    'Call fHideAllSheetExcept(shtPurchaseODRaw, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
'End Sub
'Sub subMain_ProductUnitRatio()
'    fActiveVisibleSwitchSheet shtProductUnitRatio, , False
'    'Call fHideAllSheetExcept(shtPurchaseODRaw, shtProductProducerReplace, shtProductNameReplace, shtProductSeriesReplace, shtProductUnitRatio)
'End Sub
'
'Sub subMain_SalesMan()
'    fActiveVisibleSwitchSheet shtSalesManMaster, , False
'End Sub
'Sub subMain_SalesManCommissionConfig()
'    fActiveVisibleSwitchSheet shtSalesManCommConfig, , False
'End Sub
'
'Sub subMain_Profit()
'    fActiveVisibleSwitchSheet shtProfit, , False
'End Sub
'
'Sub subMain_SelfSalesPreDeduct()
'    fActiveVisibleSwitchSheet shtSelfSalesPreDeduct, , False
'End Sub
'
'
'Sub subMain_SelfPurchaseOrder()
'    fActiveVisibleSwitchSheet shtSelfPurchaseOrder, , False
'End Sub
'
'Sub subMain_SelfSalesOrder()
'    fActiveVisibleSwitchSheet shtSelfSalesOrder, , False
'End Sub
'
'
'Sub subMain_FirstLevelCommission()
'    fActiveVisibleSwitchSheet shtFirstLevelCommission, , False
'End Sub
'
'Sub subMain_SecondLevelCommission()
'    fActiveVisibleSwitchSheet shtSecondLevelCommission, , False
'End Sub
'
'Sub subMain_InvisibleHideAllBusinessSheets()
'    fVeryHideSheet shtCompanyNameReplace
'    fVeryHideSheet shtHospital
'    fVeryHideSheet shtHospitalReplace
'    fVeryHideSheet shtSalesRawDataRpt
'    fVeryHideSheet shtSalesInfos
'    fVeryHideSheet shtPurchaseODRaw
'    fVeryHideSheet shtProductNameReplace
'    fVeryHideSheet shtProductProducerReplace
'    fVeryHideSheet shtProductSeriesReplace
'    fVeryHideSheet shtProductUnitRatio
'    fVeryHideSheet shtProductProducerMaster
'    fVeryHideSheet shtProductNameMaster
'    fVeryHideSheet shtException
'    fVeryHideSheet shtProfit
'    fVeryHideSheet shtSelfSalesOrder
'    fVeryHideSheet shtSelfSalesPreDeduct
'    fVeryHideSheet shtSelfPurchaseOrder
'    fVeryHideSheet shtSalesManMaster
'    fVeryHideSheet shtFirstLevelCommission
'    fVeryHideSheet shtSecondLevelCommission
'    fVeryHideSheet shtSalesManCommConfig
'    fVeryHideSheet shtSelfInventory
'    fVeryHideSheet shtMenuCompInvt
'    fVeryHideSheet shtMenu
'    fVeryHideSheet shtInventoryRawDataRpt
'    fVeryHideSheet shtImportCZL2SalesCompSales
'    fVeryHideSheet shtCZLSales2CompRawData
'    fVeryHideSheet shtCZLSales2Companies
'    fVeryHideSheet shtNewRuleProducts
'    fVeryHideSheet shtPV
'    fVeryHideSheet shtPromotionProduct
'    fVeryHideSheet shtCZLInvDiff
'    fVeryHideSheet shtCZLInventory
'    fVeryHideSheet shtCZLPurchaseOrder
'    'fVeryHideSheet shtCZLInformedInvInput
'    fVeryHideSheet shtCZLRolloverInv
'    fVeryHideSheet shtSalesCompInvCalcd
'    fVeryHideSheet shtSalesCompInvUnified
'    fVeryHideSheet shtSalesCompRolloverInv
'    fVeryHideSheet shtSalesCompInvDiff
'    fVeryHideSheet shtProductTaxRate
'    fVeryHideSheet shtRefund
'    'fVeryHideSheet shtMenuRefund
'    fVeryHideSheet shtCZLSales2SCompAll
'
'    fShowSheet shtMainMenu
'    shtMainMenu.Activate
'
'    If Not mRibbonObj Is Nothing Then fGetRibbonReference.Invalidate
'End Sub
'
'Sub subMain_ShowAllBusinessSheets()
'    fShowSheet shtCompanyNameReplace
'    fShowSheet shtHospital
'    fShowSheet shtHospitalReplace
'    fShowSheet shtSalesRawDataRpt
'    fShowSheet shtSalesInfos
'    fShowSheet shtPurchaseODRaw
'    fShowSheet shtProductNameReplace
'    fShowSheet shtProductProducerReplace
'    fShowSheet shtProductSeriesReplace
'    fShowSheet shtProductUnitRatio
'    fShowSheet shtProductProducerMaster
'    fShowSheet shtProductNameMaster
'    fShowSheet shtException
'    fShowSheet shtProfit
'    fShowSheet shtSelfSalesOrder
'    fShowSheet shtSelfSalesPreDeduct
'    fShowSheet shtSelfPurchaseOrder
'    fShowSheet shtSalesManMaster
'    fShowSheet shtFirstLevelCommission
'    fShowSheet shtSecondLevelCommission
'    fShowSheet shtSalesManCommConfig
'    fShowSheet shtSelfInventory
'    fShowSheet shtMenuCompInvt
'    fShowSheet shtMenu
'    fShowSheet shtInventoryRawDataRpt
'    fShowSheet shtSalesCompInventory
'    fShowSheet shtImportCZL2SalesCompSales
'    fShowSheet shtCZLSales2CompRawData
'    fShowSheet shtCZLSales2Companies
'    fShowSheet shtPromotionProduct
'    fShowSheet shtCZLInvDiff
'    'fShowSheet shtCZLInformedInvInput
'    fShowSheet shtCZLRolloverInv
'    fShowSheet shtSalesCompInv
'    fShowSheet shtSalesCompRolloverInv
'    fShowSheet shtProductTaxRate
'
'    fShowSheet shtMainMenu
'    shtMainMenu.Activate
'End Sub
'

''Function fActiveVisibleSwitchSheet(shtToSwitch As Worksheet, Optional sRngAddrToSelect As String = "A1")
''    Dim shtCurr As Worksheet
''    Set shtCurr = ActiveSheet
''
''    On Error Resume Next
''
''    If shtToSwitch.Visible = xlSheetVisible Then
''        If ActiveSheet Is shtToSwitch Then
''            shtToSwitch.Visible = xlSheetVisible
''            shtToSwitch.Activate
''            Range(sRngAddrToSelect).Select
''        Else
''            shtToSwitch.Visible = xlSheetVeryHidden
''        End If
''    Else
''        shtToSwitch.Visible = xlSheetVisible
''        shtToSwitch.Activate
''        Range(sRngAddrToSelect).Select
''    End If
''
''    If bHidePreviousActiveSheet Then
''        If Not shtCurr Is shtToSwitch Then shtCurr.Visible = xlSheetVeryHidden
''    End If
''
''    err.Clear
''End Function
'Function fHideAllSheetExcept(ParamArray arr())
'    Dim sht 'As Worksheet
'    Dim shtConvt 'As Worksheet
'    Dim wbSht 'As Worksheet
'
'    On Error Resume Next
'
'    For Each wbSht In ThisWorkbook.Worksheets
'        For Each sht In arr
'            Set shtConvt = sht
'            If wbSht Is shtConvt Then
'                'sht.Visible = xlSheetVisible
'                GoTo next_wbsheet
'            End If
'        Next
'
'        wbSht.Visible = xlSheetVeryHidden
'next_wbsheet:
'    Next
'
'    Set shtConvt = Nothing
'    Err.Clear
'End Function
'
'Function fAppendDataToLastCellOfColumn(ByRef sht As Worksheet, alCol As Long, aValue)
'    Dim lMaxRow As Long
'    lMaxRow = sht.Cells(Rows.Count, alCol).End(xlUp).Row
'
'    If lMaxRow <= 1 Then
'        If fZero(sht.Cells(lMaxRow, alCol).Value) Then
'            sht.Cells(lMaxRow, alCol).Value = aValue
'        Else
'            sht.Cells(lMaxRow + 1, alCol).Value = aValue
'        End If
'    Else
'        sht.Cells(lMaxRow + 1, alCol).Value = aValue
'    End If
'End Function
'
'
'Function fCompareDictionaryKeysAndMultipleItems(ByRef dictBase As Dictionary, ByRef dictThis As Dictionary) As Dictionary
'    Dim dictOut As Dictionary
'    Dim i As Long
'    Dim sKey As String
'    Dim sValue As String
'
'    Set dictOut = New Dictionary
'
'    'missed from right one
'    For i = 0 To dictBase.Count - 1
'        sKey = dictBase.Keys(i)
'
'        If Not dictThis.Exists(sKey) Then
'            dictOut.Add DELETED_FROM_NEW_VERSION & DELIMITER & sKey, dictBase.Items(i) & vbLf & "新版本中没有设置"
'        Else
'            If dictBase.Items(i) <> dictThis(sKey) Then
'                dictOut.Add BOTH_HAVE_BUT_DIFF_VALUE & DELIMITER & sKey, dictBase.Items(i) & vbLf & dictThis(sKey)
'            Else
'                'dictOut.Add SAME_IN_BOTH & DELIMITER & sKey, dictBase.Items(i) & vbLf & dictThis(sKey)
'            End If
'            dictThis.Remove sKey
'        End If
'    Next
'
'    'missed from LEFT one
'    For i = 0 To dictThis.Count - 1
'        sKey = dictThis.Keys(i)
'
'        'If Not dictBase.Exists(sKey) Then
'            dictOut.Add NEWLY_ADDED_IN_NEW_VERSION & DELIMITER & sKey, dictThis.Items(i) & vbLf & "基础版本中没有设置"
'        'End If
'    Next
'
'    Set fCompareDictionaryKeysAndMultipleItems = dictOut
'    Set dictOut = Nothing
'End Function
