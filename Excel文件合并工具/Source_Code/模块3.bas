Attribute VB_Name = "模块3"
Option Explicit

Sub 宏1()
Attribute 宏1.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim a As WshShell
     
    
End Sub
Sub 宏2()
Attribute 宏2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' 宏2 宏
'

'
    Range("C22:C38").Select
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
End Sub
Sub 宏3()
Attribute 宏3.VB_ProcData.VB_Invoke_Func = " \n14"
 
    Dim arrOutput()
    Dim dictLog As Dictionary
    Dim lEachRow As Long
    Dim i As Integer
    Dim sMsg As String
    Dim sTopFolder As String
    Dim sOutputFolder As String
    Dim sProduct As String
'    Dim arrHeader()
'    Dim dictNotInProcess As Dictionary
    
    Call fInitialization
    
'    arrHeader = Array(chapter, criteria_item, FEASIBLE_TO_PROCESS, PROCESS_ON_THE_WAY, REASON_WHY_NOT, YOUR_ACTION)
    
    'On Error GoTo error_handling
    
    Set dictLog = New Dictionary
    
    sTopFolder = fSelectFolderDialog(ThisWorkbook.Path) '
    If Len(sTopFolder) <= 0 Then fErr '
    
    sOutputFolder = sTopFolder & "Output\" '& Format(Now(), "yyyymmddHHMMSS")
    
    If fFolderExists(sOutputFolder) Then
        If fPromptToConfirmToContinue("The output folder already exists, all files will be deleted first, are you sure to continue?" _
                 & vbCr & sOutputFolder) Then
            Call fDeleteAllFromFolder(sOutputFolder)
        Else
            fErr
        End If
    End If
    
    Call fDeleteRowsFromSheetLeaveHeader(shtLog)
    Call fDeleteRowsFromSheetLeaveHeader(shtReportDetails)
    
    fGetFSO
    Dim sOriginFolder As String
    Dim sFirstResponseFolder As String
    Dim sSecondResponseFolder As String
    sOriginFolder = sTopFolder & "第一次反馈\"
    sFirstResponseFolder = sTopFolder & "第一次反馈\"
    sSecondResponseFolder = sTopFolder & "第二次反馈\"
    
    If Not (gFSO.FolderExists(sOriginFolder) And gFSO.FolderExists(sFirstResponseFolder) And gFSO.FolderExists(sSecondResponseFolder)) Then
        fErr "There must be 3 folders there, but any one was not found:" & vbCr & vbCr _
             & sOriginFolder
    End If
    
    Dim arrOrigin
    Dim arrFirst
    Dim arrSecond
    arrOrigin = fGetAllExcelFileListFromSubFolders(sOriginFolder)
    
    If ArrLen(arrOrigin) <= 0 Then fErr "No file was found in folder " & vbCr & sOriginFolder
    
    arrFirst = fGetAllExcelFileListFromSubFolders(sFirstResponseFolder)
    arrSecond = fGetAllExcelFileListFromSubFolders(sSecondResponseFolder)
     
    Dim sParentFolder As String
    Dim sOrigFile As String
    Dim wbSource As Workbook
    Dim wbOut As Workbook
    Dim sht As Worksheet
    Dim sOutputFile As String
    Dim sFileNetName As String
    Dim j As Long
    Dim sSubPrtFolder As String
    Dim sBaseName As String
    
    sOutputFolder = fCheckPath(sOutputFolder)
    For i = LBound(arrOrigin) To UBound(arrOrigin)
        sOrigFile = arrOrigin(i)
        sFileNetName = fGetFileNetName(sOrigFile)
        sBaseName = fGetFileBaseName(sOrigFile)
        
        If Len(sOrigFile) - Len(sOriginFolder) - Len(sBaseName) - 1 > 0 Then
            sSubPrtFolder = Right(sOrigFile, Len(sOrigFile) - Len(sOriginFolder))
            sSubPrtFolder = Left(sSubPrtFolder, Len(sSubPrtFolder) - Len(sBaseName))
            sOutputFile = sOutputFolder & sSubPrtFolder & sFileNetName & "_tmp.xlsx"
        Else
            sOutputFile = sOutputFolder & sFileNetName & "_tmp.xlsx"
        End If
        
        Set wbSource = fOpenWorkbook(sOrigFile, , True)
        
        If wbSource.Worksheets.Count > 1 Then
            For j = 1 To wbSource.Worksheets.Count
                If j = 1 Then
                    Set wbOut = fCopySingleSheet2NewWorkbookFile(sht, sOutputFile, "原稿" & j)
                Else
                    Call fCopySingleSheet2WorkBook(sht, wbOut, "原稿" & j)
                End If
            Next
        Else
            Set wbOut = fCopySingleSheet2NewWorkbookFile(sht, sOutputFile, "原稿")
        End If
        
    Next
End Sub
Sub ddaasdfasdf()
    Dim a
    a = ActiveSheet.Range("A15").MergeArea.Address
End Sub
