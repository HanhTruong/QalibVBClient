VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExcelIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  ExcelIO.cls
 
'DESCRIPTION:  This module holds the Excel Input class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: ExcelIO.cls $
 ' 
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:29a
 ' Updated in $/QalibVBClient/Source/QalibExcel
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:09p
 ' Updated in $/QalibVBClient/Source/QalibExcel
 ' Completely revised interface for GetRuns and GetChems.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:21p
 ' Updated in $/QalibVBClient/Source/QalibExcel
 ' Added ability to get a list of chemistries from the Excel file.
 ' Added functionality to get the number of runs from the Excel file.
 ' Added named range constants.
 ' Added serial number support.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:50p
 ' Created in $/QalibVBClient/Source/QalibExcel
 ' Added to SourceSafe.

Option Explicit

' constants

Private Const CALSHEETTITLE As String = "Calibration"
Private Const BCCELL As String = "Bar_Code_String"
Private Const CHEMSCELL As String = "Chemistries_Extracted"
Private Const SAMPLEIDCELL As String = "Sample_ID"
Private Const RUNSCELL As String = "Runs_Extracted"
Private Const SPECIESCELL As String = "Run_Species"
Private Const SYSSTAT1CELL As String = "Sys_Status_1"
Private Const SERIALCELL As String = "Serial_Number"
Private Const SYSSXNCELL As String = "Sys_Sxn"
Private Const SYSTEMPCELL As String = "Sys_Temp"
Private Const DATABASERANGE As String = "Database"

Private Const MEASUREDVALTEXT As String = "_CAL_DATA_"

Private Const CHEMDELIMITER As String = ","

'***********************************************************************

'PROCEDURE:   GetRuns()

'DESCRIPTION: Reads the run data from the Excel sheet

'PARAMETERS:  inFileName - the Excel filename
'             inChemistry - the chemistry
'             inRunID - the ID of the run
'             inBarCode - the bar code of the run
'             inSerial - the serial number of the instrument for the run
'             inExpPoint - the meta data and measured values for the points
'             inSpecies - the species of the run
'             inSxn - the system sxn of the run
'             inSW - the software status of the run
'             inTemp - the temperature of the run
'             inStatus - the status of the point
              
'RETURNED:    N/A

'********************************************************************
Public Sub GetRuns(inFileName As String, inChemistry As String, inRunID() As Long, inBarCode() As String, _
    inSerial() As String, inExpPoint() As String, inSpecies() As String, inSxn() As String, inSW() As Long, _
    inTemp() As Double, inStatus() As Long)
On Error GoTo ErrTrap
    Dim runRow As Long
    Dim runStartRow As Long
    Dim runEndRow As Long
    Dim maxRunIndex As Long
    Dim runIndex As Long
    Dim expPointIndex As Long
    Dim sampleIDCol As Long
    Dim BCCol As Long
    Dim serialCol As Long
    Dim expPointCol As Long
    Dim speciesCol As Integer
    Dim sysSxnCol As Integer
    Dim swCol As Integer
    Dim sysTempCol As Integer
    Dim sampleIDCells As Variant
    Dim BCCells As Variant
    Dim serialCells As Variant
    Dim expPointCells As Variant
    Dim speciesCells As Variant
    Dim sysSxnCells As Variant
    Dim swCells As Variant
    Dim sysTempCells As Variant
    Dim curText As String
    Dim savedText As String
    Dim numSampleIDIndex As Long
    Dim expPointSetIndex As Long
    Dim maxExpPointSetIndex As Long
    Dim XlApp As Excel.Application  ' stores the reference to the Excel application
    Dim XlSht As Excel.Worksheet ' stores the reference to the Excel worksheet
    
    ' prepare the Excel
    Call OpenExcel(XlApp, inFileName, XlSht, CALSHEETTITLE)
        
    ' find named cells in the Excel sheet and assign the landmarks
    sampleIDCol = GetExcelNamedRange(XlSht, SAMPLEIDCELL).Column
    BCCol = GetExcelNamedRange(XlSht, BCCELL).Column
    serialCol = GetExcelNamedRange(XlSht, SERIALCELL).Column
    speciesCol = GetExcelNamedRange(XlSht, SPECIESCELL).Column
    sysSxnCol = GetExcelNamedRange(XlSht, SYSSXNCELL).Column
    swCol = GetExcelNamedRange(XlSht, SYSSTAT1CELL).Column
    sysTempCol = GetExcelNamedRange(XlSht, SYSTEMPCELL).Column
    runStartRow = GetExcelNamedRange(XlSht, BCCELL).Row + 1
    runEndRow = runStartRow + CLng(GetExcelNamedRange(XlSht, RUNSCELL).Value) - 1
        
    ' make sure the bounds are meaningful
    If (runEndRow < runStartRow) Then
        Call Err.Raise(COMPERR, "QalibExcel", BuildSubstString(LoadResString(SHORTFILEERRTEXT), inFileName))
    End If
        
    maxRunIndex = runEndRow - runStartRow
    
    ' the run data is communicated in a specific format
    
    ' inExpPoint is the most complicated; its format is as follows:
    ' 1.  the number of point sets
    ' 2.  the name of each point set
    ' 3.  the number of series
    ' 4.  the number of runs and the name of each series
    ' 5.  the points in each point set (all of set one, all of set two, etc.)
    
    ' inStatus has the status for each point.  It is a parallel array to inExpPoint
    ' except it does not contain the meta information from 1-4 above.
    
    ' inSerial, inBarCode, and inSpecies are parallel to inExpPoint too, except they don't
    ' have the meta information from 1-4, and since the serial number, bar code, and species are the same
    ' for each point in the run, they are not repeated for multiple point sets
    
    ' see how many sets of expPoint expPoints there are for the chemistry
    expPointIndex = 0

    On Error Resume Next
    
    Do
        curText = Trim$(CStr(GetExcelNamedRange(XlSht, inChemistry & MEASUREDVALTEXT & (expPointIndex)).Value))
        
        ' if the named range was found increment the expPoint set count and save its name
        If (Err.Number = 0) Then
            expPointIndex = expPointIndex + 1
            ReDim Preserve inExpPoint(expPointIndex) As String
            inExpPoint(0) = CStr(expPointIndex) ' increment the count of expPoint sets
            inExpPoint(expPointIndex) = curText ' save the name of the expPoint set
         End If
    Loop While (Err.Number = 0)
        
    On Error GoTo ErrTrap
    
     ' save the location of the number of sample types, which is right after the last expPoint set name
    expPointIndex = expPointIndex + 1
    ReDim Preserve inExpPoint(expPointIndex) As String
    numSampleIDIndex = expPointIndex
    
    ' prepare the arrays for the run data
    ReDim inRunID(maxRunIndex) As Long
    ReDim inBarCode(maxRunIndex) As String
    ReDim inSerial(maxRunIndex) As String
    ReDim inSpecies(maxRunIndex) As String
    ReDim inSxn(maxRunIndex) As String
    ReDim inSW(maxRunIndex) As Long
    ReDim inTemp(maxRunIndex) As Double
            
    ' prepare the status array (each point has a status)
    ReDim inStatus((maxRunIndex * CLng(inExpPoint(0))) + (CLng(inExpPoint(0)) - 1)) As Long
    
    ' make sure this does not cause a memory leak
    ' the sample types must be sorted
    Call XlSht.Range(DATABASERANGE).Sort(Key1:=SAMPLEIDCELL, Order1:=xlAscending, Header:=xlYes)
    
    ' load the cells into the variant arrays
    ' variant arrays are MUCH faster to traverse than Excel cells
    BCCells = XlSht.Range(XlSht.Cells(runStartRow, BCCol), _
        XlSht.Cells(runEndRow, BCCol))
    serialCells = XlSht.Range(XlSht.Cells(runStartRow, serialCol), _
        XlSht.Cells(runEndRow, serialCol))
    speciesCells = XlSht.Range(XlSht.Cells(runStartRow, speciesCol), _
        XlSht.Cells(runEndRow, speciesCol))
    sysSxnCells = XlSht.Range(XlSht.Cells(runStartRow, sysSxnCol), _
        XlSht.Cells(runEndRow, sysSxnCol))
    swCells = XlSht.Range(XlSht.Cells(runStartRow, swCol), _
        XlSht.Cells(runEndRow, swCol))
    sysTempCells = XlSht.Range(XlSht.Cells(runStartRow, sysTempCol), _
        XlSht.Cells(runEndRow, sysTempCol))
    sampleIDCells = XlSht.Range(XlSht.Cells(runStartRow, sampleIDCol), _
        XlSht.Cells(runEndRow, sampleIDCol))
    
    runRow = runStartRow
    savedText = ""
    inExpPoint(numSampleIDIndex) = "0" ' the number of sample types starts at zero
    
    ' cycle through the arrays to fill the run data
    For runIndex = 0 To maxRunIndex
        inRunID(runIndex) = runRow
        inBarCode(runIndex) = Trim$(CStr(BCCells(runIndex + 1, 1)))
        inSerial(runIndex) = Trim$(CStr(serialCells(runIndex + 1, 1)))
        inSpecies(runIndex) = Trim$(CStr(speciesCells(runIndex + 1, 1)))
        inSxn(runIndex) = Trim$(CStr(sysSxnCells(runIndex + 1, 1)))
        inSW(runIndex) = CLng(swCells(runIndex + 1, 1))
        inTemp(runIndex) = CDbl(sysTempCells(runIndex + 1, 1))
                        
        ' get this run's sample type
        curText = Trim$(CStr(sampleIDCells(runIndex + 1, 1)))
        
        ' see if this sample type is currently being accounted for
        If (curText <> savedText) Then
            
            ' this is a new sample type
            savedText = curText
            
            ' the number of sample types gets incremented
            inExpPoint(numSampleIDIndex) = CStr(CLng(inExpPoint(numSampleIDIndex)) + 1)
            
            ' the sample type meta information takes two elements: one for the size and one for the name
            expPointIndex = expPointIndex + 2
            ReDim Preserve inExpPoint(expPointIndex) As String
            inExpPoint(expPointIndex - 1) = "1"
            inExpPoint(expPointIndex) = savedText
        Else
            ' the sample type already exists so just increment its count
            inExpPoint(expPointIndex - 1) = CStr(CLng(inExpPoint(expPointIndex - 1)) + 1)
        End If
                
        runRow = runRow + 1
    Next runIndex
    
    ' make room for the points in each point set
    ReDim Preserve inExpPoint(expPointIndex + ((maxRunIndex + 1) * CLng(inExpPoint(0)))) As String

    ' the analyzer absorbance data gets loaded one column at a time starting with index zero
    maxExpPointSetIndex = CLng(inExpPoint(0)) - 1
    
    For expPointSetIndex = 0 To maxExpPointSetIndex
        
        ' get the location of the expPointSetIndex column
        expPointCol = GetExcelNamedRange(XlSht, inChemistry & MEASUREDVALTEXT & expPointSetIndex).Column
                            
        expPointCells = XlSht.Range(XlSht.Cells(runStartRow, expPointCol), _
            XlSht.Cells(runEndRow, expPointCol))
            
        ' load the data for the expPointSetIndex column
        For runIndex = 0 To maxRunIndex
            expPointIndex = expPointIndex + 1
            inExpPoint(expPointIndex) = Trim$(CStr(expPointCells(runIndex + 1, 1)))
            
            ' the status array does not have the meta information at the beginning so its
            ' index must be calculated
            inStatus(runIndex + ((expPointSetIndex * maxRunIndex) + expPointSetIndex)) = 0
        Next runIndex
            
    Next expPointSetIndex
    
    Call CloseExcel(XlApp)
    
    Exit Sub
ErrTrap:
    Call CloseExcel(XlApp)
    Call Err.Raise(Err.Number, Err.Source & " | ExcelIO.GetRuns", Err.Description)
End Sub

'***********************************************************************

'FUNCTION:   GetChems()

'DESCRIPTION: Reads the extracted chemistries from the Excel sheet

'PARAMETERS:  inFileName - the Excel filename
'             inChemistry - the chemistry
              
'RETURNED:    N/A

'********************************************************************
Public Sub GetChems(inFileName As String, inChemistry() As String)
On Error GoTo ErrTrap
    Dim XlApp As Excel.Application  ' stores the reference to the Excel application
    Dim XlSht As Excel.Worksheet ' stores the reference to the Excel worksheet
    Dim server As ExcelIO
    Dim startPos As Long
    Dim endPos As Long
    Dim chems As String
    Dim chemIndex As Integer
    
    Set server = New ExcelIO
    
    ' prepare the Excel
    Call OpenExcel(XlApp, inFileName, XlSht, CALSHEETTITLE)
        
    ' get the chemistries extracted cell
    chems = CStr(GetExcelNamedRange(XlSht, CHEMSCELL).Value)
    
    If (chems = "") Then
        Call Err.Raise(COMPERR, "QalibExcel", BuildSubstString(LoadResString(CHEMREADERR), inFileName))
    End If
    
    startPos = 1
    chemIndex = 0
    
    ' parse the delimited list of chemistries and load them into the return array
    Do
        
        ' find the delimiter
        endPos = InStr(startPos, chems, CHEMDELIMITER)
        
        ' increase the size of the chemistry array
        ReDim Preserve inChemistry(chemIndex) As String
        
        If (endPos = 0) Then
            inChemistry(chemIndex) = Trim$((Mid$(chems, startPos)))
        Else
            inChemistry(chemIndex) = Trim$((Mid$(chems, startPos, endPos - startPos)))
        End If
        
        If (inChemistry(chemIndex) = "") Then
            Call Err.Raise(COMPERR, "QalibExcel", BuildSubstString(LoadResString(CHEMREADERR), inFileName))
        End If
        
        ' start the next search one position in front of the delimiter
        startPos = endPos + 1
        
        chemIndex = chemIndex + 1
        
    Loop While (endPos <> 0)
    
    Call CloseExcel(XlApp)
    
    Exit Sub
ErrTrap:
    Call CloseExcel(XlApp)
    Call Err.Raise(Err.Number, Err.Source & " | ExcelIO.GetChems", Err.Description)
End Sub

