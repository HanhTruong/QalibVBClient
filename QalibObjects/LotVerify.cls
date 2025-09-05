VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LotVerify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  LotVerify.cls
 
'DESCRIPTION:  This module holds the lot verification class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: LotVerify.cls $
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:40p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Updated to 1.0.0X9.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:32p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Added chemistry property and serial number support.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe
 '
 ' *****************  Version 6  *****************
 ' User: Ballard      Date: 5/09/03    Time: 4:06p
 ' Updated in $/QalibVBClient/Source/LotVerify
 ' Changed name of file from QalibData to BeadLot.  All lot data now
 ' stored in UDT.  Added plot attribute for storing the calibration chart.
 ' Module communicates with BeadLotServer object via Lot UDT passed as
 ' function parameter.
 '
 ' *****************  Version 5  *****************
 ' User: Ballard      Date: 3/25/03    Time: 3:40p
 ' Updated in $/QalibVBClient
 ' Added mode, wavelength, rotor, species, cuvette, diluent, mold, and
 ' measurement status properties and methods to set and get them.
 '
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 3/21/03    Time: 2:34p
 ' Updated in $/QalibVBClient
 ' Renamed module "clsQalibData.cls" to "QalibData.cls."
 '
 ' This class now serves as the underlying data object for all of Qalib.
 '
 ' All of the functionality to read an Excel file and manipulate GUI
 ' components has been removed.  The class now has several private member
 ' variables to store the calibration data and public properties/methods
 ' to set/access them.  The class also has a few public methods to swap
 ' points between series and sort series.

Option Explicit

' constants

' minimum for a bead lot
Private Const MINRUNS As Long = 1

' private member variables

Private lngCalID_m As Long     ' the unique calibration ID
Private strUser_m As String ' operator
Private strChemistry_m As String ' chemistry being calibrated
'Private strSpecies_m As String ' species calibration run on
'Private strBarCode_m As String ' bar code calibration run on
Private objAnalyzerData_m As AnalyzerData ' analyzer data

' public events
Public Event UpdateStatus(percentComplete As Integer, message As String) ' tells the parent form to update the progress bar

'***********************************************************************

'PROPERTY GET:   CalID()

'DESCRIPTION: Allows other objects to get the calibration ID

'PARAMETERS:  N/A

'RETURNED:    Calibration ID

'*********************************************************************
Public Property Get CalID() As Long
On Error GoTo ErrTrap
    CalID = lngCalID_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotVerify.PropertyGet.CalID", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Chemistry()

'DESCRIPTION: Allows other objects to get the calibration ID

'PARAMETERS:  N/A

'RETURNED:    Calibration ID

'*********************************************************************
Public Property Get Chemistry() As String
On Error GoTo ErrTrap
    Chemistry = strChemistry_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotVerify.PropertyGet.Chemistry", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   AnalyzerDataSet()

'DESCRIPTION: Allows other objects to get the analyzer data

'PARAMETERS:  N/A

'RETURNED:    analyzer data set

'*********************************************************************
Public Property Get AnalyzerDataSet() As AnalyzerData
On Error GoTo ErrTrap
    Set AnalyzerDataSet = objAnalyzerData_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotVerify.PropertyGet.AnalyzerDataSet", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Class_Initialize()

'DESCRIPTION: Sets up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Initialize()
    Set objAnalyzerData_m = New AnalyzerData
End Sub

'***********************************************************************

'PROCEDURE:   Class_Terminate()

'DESCRIPTION: Cleans up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Terminate()
    Set objAnalyzerData_m = Nothing
End Sub

'***********************************************************************

'PROCEDURE:   CalibrateLot()

'DESCRIPTION: Send calibration data to Server

'PARAMETERS:  inMode - the mode of the application

'RETURNED:    N/A

'********************************************************************
Public Sub CalibrateLot(inMode As String)
On Error GoTo ErrTrap
    Dim outID As Long
    Dim curSeries As SeriesOne
    Dim curRun As Run
    Dim curExpPoint As ExpPoint
    Dim curExpPointSet As ExpPointSet
    Dim metaIndex As Long
    Dim flatExpPoint() As String
    Dim flatSerial() As String
    Dim flatBarCode() As String
    Dim flatSpecies() As String
    Dim flatStatus() As String
    Dim passExpPoint(0) As Variant
    Dim passSerial(0) As Variant
    Dim passBarCode(0) As Variant
    Dim passSpecies(0) As Variant
    Dim passStatus(0) As Variant
    Dim server As QALIBCLIENTLib.QalibClientMain
    Dim error_msg As Variant
    Dim expPointSetStartIndex As Long
    Dim maxExpPointSetIndex As Long
    Dim expPointSetIndex As Long
    Dim seriesStartIndex As Long
    Dim maxSeriesIndex As Long
    Dim seriesIndex As Long
    Dim numRuns As Long
    Dim expPointIndex As Long
    Dim expPointStartIndex As Long
    
    ' the data is communicated with the server in a specific format
    
    ' flatExpPoint is the most complicated; its format is as follows:
    ' 1.  the number of point sets
    ' 2.  the name of each point set
    ' 3.  the number of series
    ' 4.  the number of runs and the name of each series
    ' 5.  the points in each point set (all of set one, all of set two, etc.)
    
    ' flatStatus has the status for each point.  It is a parallel array to flatExpPoint
    ' except it does not contain the meta information from 1-4 above.
    
    ' flatSerial, flatBarCode, and flatSpecies are parallel to flatExpPoint too, except they don't
    ' have the meta information from 1-4, and since the serial number, bar code, and species are the same
    ' for each point in the run, they are not repeated for multiple point sets
    
     metaIndex = 0
     
    ' load the number of point sets
    ReDim flatExpPoint(metaIndex) As String
    flatExpPoint(metaIndex) = CStr(objAnalyzerData_m.ExpPointSets.Count)
    expPointSetStartIndex = metaIndex + 1
    
    ' load the names of each point set
    For Each curExpPointSet In objAnalyzerData_m.ExpPointSets
        metaIndex = metaIndex + 1
        ReDim Preserve flatExpPoint(metaIndex) As String
        flatExpPoint(metaIndex) = curExpPointSet.Name
    Next curExpPointSet
    
    maxExpPointSetIndex = metaIndex
    
    ' load the number of series
    metaIndex = metaIndex + 1
    ReDim Preserve flatExpPoint(metaIndex) As String
    flatExpPoint(metaIndex) = CStr(objAnalyzerData_m.SeriesSet.Count)
    
    seriesStartIndex = metaIndex + 2
    numRuns = 0
    
    ' load the size and name of each series
    For Each curSeries In objAnalyzerData_m.SeriesSet
                
        numRuns = numRuns + curSeries.Runs.Count
        
        metaIndex = metaIndex + 2
        ReDim Preserve flatExpPoint(metaIndex) As String
        flatExpPoint(metaIndex - 1) = CStr(curSeries.Runs.Count)
        flatExpPoint(metaIndex) = curSeries.Name
        
    Next curSeries
    
    maxSeriesIndex = metaIndex
    expPointStartIndex = metaIndex + 1
    
    ' the array needs to have enough room for the meta information and all the points
    ReDim Preserve flatExpPoint(expPointStartIndex + (numRuns * objAnalyzerData_m.ExpPointSets.Count) - 1) As String

    ' increase the array to hold each point's status
    ReDim Preserve flatStatus((numRuns * objAnalyzerData_m.ExpPointSets.Count) - 1) As String
    
    ' increase the array to hold each run's serial number
    ReDim Preserve flatSerial(numRuns - 1) As String
    
    ' increase the array to hold each run's bar code
    ReDim Preserve flatBarCode(numRuns - 1) As String

    ' increase the array to hold each run's species
    ReDim Preserve flatSpecies(numRuns - 1) As String
    
    expPointIndex = 0
    
    For seriesIndex = seriesStartIndex To maxSeriesIndex Step 2
               
        Set curSeries = objAnalyzerData_m.SeriesSet(flatExpPoint(seriesIndex))
                      
        For Each curRun In curSeries.Runs
                
            ' each run only has one of these
            flatSerial(expPointIndex) = curRun.Serial
            flatBarCode(expPointIndex) = curRun.BarCode
            flatSpecies(expPointIndex) = curRun.Species
                    
            ' each run could have more than one point
            For expPointSetIndex = expPointSetStartIndex To maxExpPointSetIndex
            
                    Set curExpPointSet = objAnalyzerData_m.ExpPointSets(flatExpPoint(expPointSetIndex))
                    Set curExpPoint = curExpPointSet.ExpPoints(curRun.ID)
                    
                    ' most runs will have just one point, so the ((expPointSetIndex - 1) * numRuns) term will be zero
                    ' for the most part.  However, for runs with more than one point this term will take care of the offset
                    ' when loading the other points
                    flatExpPoint(expPointStartIndex + (expPointIndex + ((expPointSetIndex - 1) * numRuns))) = CStr(curExpPoint.MeasuredVal)
                    flatStatus((expPointIndex + ((expPointSetIndex - 1) * numRuns))) = CStr(curExpPoint.Status)
            Next expPointSetIndex
                
            expPointIndex = expPointIndex + 1
        
        Next curRun
                
    Next seriesIndex
    
    ' This trick is necessary to pass the flat array to the COM object
    passExpPoint(0) = flatExpPoint
    passStatus(0) = flatStatus
    passSerial(0) = flatSerial
    passBarCode(0) = flatBarCode
    passSpecies(0) = flatSpecies
    
    ' send data to server
    Set server = New QALIBCLIENTLib.QalibClientMain
    error_msg = server.SendClientData(strChemistry_m, 3, inMode, strUser_m, passExpPoint(0), _
        passSerial(0), passBarCode(0), passSpecies(0), passStatus(0), outID)
               
    Call InterpretServerMsg(error_msg)
    
    ' the only thing returned by the server was the calibration ID
    lngCalID_m = outID
    
    Set curSeries = Nothing
    Set curExpPoint = Nothing
    Set curExpPointSet = Nothing
    Set curRun = Nothing
    Set server = Nothing
    
    Exit Sub
ErrTrap:
    Set curSeries = Nothing
    Set curExpPoint = Nothing
    Set curExpPointSet = Nothing
    Set curRun = Nothing
    Set server = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | LotVerify.CalibrateLot", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   LoadRuns()

'DESCRIPTION: Gets the requested runs from the database for verification

'PARAMETERS:  inCriteria - the requested runs
'             inChemistry - the chemistry to calibrate
'             inUser - the user name
'             inPassword - the password

'RETURNED:    N/A

'********************************************************************
Public Sub LoadRuns(inCriteria As Variant, inChemistry As String, inUser As String, inPassword As String)
On Error GoTo ErrTrap
    Dim passRunID() As Long
    Dim passSerial() As String
    Dim passExpPoint() As String
    Dim passStatus() As Long
    Dim passBarCode() As String
    Dim passSpecies() As String
    Dim passSxn() As String
    Dim passSW() As Long
    Dim passTemp() As Double
    Dim runIndex As Long
    Dim maxRunIndex As Long
    Dim expPointIndex As Long
    Dim numGood As Long
    Dim curSeries As SeriesOne
    Dim curRun As Run
    Dim curExpPoint As ExpPoint
    Dim curExpPointSet As ExpPointSet
    Dim objExcelIO As QalibExcel.ExcelIO
    Dim expPointSetIndex As Long
    Dim maxExpPointSetIndex As Long
    Dim curExpPointValue As String
    Dim expPointSetName() As String
    Dim seriesIndex As Long
    Dim maxSeriesIndex As Long
    Dim seriesName() As String
    Dim seriesLength() As Long
    Dim seriesRunIndex As Long
    Dim maxSeriesRunIndex As Long
    
    RaiseEvent UpdateStatus(33, "Loading Runs")
    
    strChemistry_m = inChemistry
    strUser_m = inUser
    
    Set objExcelIO = New QalibExcel.ExcelIO
    
    Call objExcelIO.GetRuns(CStr(inCriteria), inChemistry, passRunID, passBarCode, passSerial, passExpPoint, _
        passSpecies, passSxn, passSW, passTemp, passStatus)
    
    maxRunIndex = UBound(passRunID)
    
    ' make sure there's at least MINRUNS runs!
    If (maxRunIndex + 1 < MINRUNS) Then
        Call Err.Raise(COMPERR, "QalibObjects", BuildSubstString(LoadResString(MINRUNSERRTEXT), maxRunIndex + 1, MINRUNS))
    End If
                    
    Call UnpackRuns(objAnalyzerData_m, passExpPoint, passStatus, passRunID, passSerial, passBarCode, passSpecies, passSxn, passSW)

    RaiseEvent UpdateStatus(66, "Loading Assigned Values")
    
    ' get the assigned values for the runs
    Call GetAssigned(strChemistry_m, objAnalyzerData_m.SeriesSet)
        
    Set curSeries = Nothing
    Set curRun = Nothing
    Set curExpPoint = Nothing
    Set curExpPointSet = Nothing
    Set objExcelIO = Nothing
    
    Exit Sub
ErrTrap:
    Set curSeries = Nothing
    Set curRun = Nothing
    Set curExpPoint = Nothing
    Set curExpPointSet = Nothing
    Set objExcelIO = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | LotVerify.LoadRuns", Err.Description)
End Sub

