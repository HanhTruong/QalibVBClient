VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LotReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  LotReport.cls
 
'DESCRIPTION:  This module holds the lot report class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: LotReport.cls $
 ' 
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:40p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Updated to 1.0.0X9.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:34p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Added chemistry, plot, and calibration ID support.
 ' Added ability to accept calibration.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe

Option Explicit

' private constants

Private Const CALIBRATORCHECKPLOTTEXT As String = "CALIBRATOR"
Private Const CONTROLCHECKPLOTTEXT As String = "CONTROL"

' private member variables

Private lngCalID_m As Long     ' the unique calibration ID
Private strMode_m As String ' calibration mode
Private strUser_m As String ' operator
Private strCalDate_m As String ' the calibration date
Private strChemistry_m As String ' chemistry being calibrated
Private strWavelength_m As String ' wavelength being calibrated
Private strRotor_m As String ' rotor calibration run on
Private strSpecies_m As String ' species calibration run on
Private strCuvette_m As String ' cuvette calibration run on
Private strDiluent_m As String ' diluent calibration run on
Private strMold_m As String ' mold calibration run on
Private strBarCode_m As String ' bar code calibration run on
Private objAnalyzerData_m As AnalyzerData ' the calibration analyzer data
Private objResults_m As QalibCollection ' the calibration fit parameters
Private objPlots_m As QalibCollection ' the calibrtaion plots
Private strComment_m As String  ' the calibration comment
Private blnUpdatedFit_m As Boolean ' stores whether calibration has been updated

'***********************************************************************

'PROPERTY GET: CalID()

'DESCRIPTION: Allows other objects to get the calibration ID

'PARAMETERS:  N/A

'RETURNED:    calibration ID

'*********************************************************************
Public Property Get CalID() As Long
On Error GoTo ErrTrap
    CalID = lngCalID_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.CalID", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Mode()

'DESCRIPTION: Allows other objects to get the calibration mode

'PARAMETERS:  N/A

'RETURNED:    the calibration mode

'*********************************************************************
Public Property Get Mode() As String
On Error GoTo ErrTrap
     Mode = strMode_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.Mode", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   User()

'DESCRIPTION: Allows other objects to get the calibration user

'PARAMETERS:    N/A

'RETURNED:    the calibration user

'*********************************************************************
Public Property Get User() As String
On Error GoTo ErrTrap
    User = strUser_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.User", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   CalDate()

'DESCRIPTION: Allows other objects to get the calibration date

'PARAMETERS:    N/A

'RETURNED:    the calibration date

'*********************************************************************
Public Property Get CalDate() As String
On Error GoTo ErrTrap
    CalDate = strCalDate_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.CalDate", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   BarCode()

'DESCRIPTION: Allows other objects to get the bar code

'PARAMETERS:  N/A

'RETURNED:    bar code

'*********************************************************************
Public Property Get BarCode() As String
On Error GoTo ErrTrap
    BarCode = strBarCode_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.BarCode", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Chemistry()

'DESCRIPTION: Allows other objects to get the chemistry

'PARAMETERS:  N/A

'RETURNED:    chemistry

'*********************************************************************
Public Property Get Chemistry() As String
On Error GoTo ErrTrap
    Chemistry = strChemistry_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.Chemistry", Err.Description)
End Property


'***********************************************************************

'PROPERTY GET:   Wavelength()

'DESCRIPTION: Allows other objects to get the Wavelength

'PARAMETERS:  N/A

'RETURNED:    Wavelength

'*********************************************************************
Public Property Get Wavelength() As String
On Error GoTo ErrTrap
    Wavelength = strWavelength_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.Wavelength", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Rotor()

'DESCRIPTION: Allows other objects to get the Rotor

'PARAMETERS:  N/A

'RETURNED:    Rotor

'*********************************************************************
Public Property Get Rotor() As String
On Error GoTo ErrTrap
    Rotor = strRotor_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.Rotor", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Species()

'DESCRIPTION: Allows other objects to get the Species

'PARAMETERS:  N/A

'RETURNED:    Species

'*********************************************************************
Public Property Get Species() As String
On Error GoTo ErrTrap
    Species = strSpecies_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.Species", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Cuvette()

'DESCRIPTION: Allows other objects to get the Cuvette

'PARAMETERS:  N/A

'RETURNED:    Cuvette

'*********************************************************************
Public Property Get Cuvette() As String
On Error GoTo ErrTrap
    Cuvette = strCuvette_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.Cuvette", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Diluent()

'DESCRIPTION: Allows other objects to get the Diluent

'PARAMETERS:  N/A

'RETURNED:    Diluent

'*********************************************************************
Public Property Get Diluent() As String
On Error GoTo ErrTrap
    Diluent = strDiluent_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.Diluent", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Mold()

'DESCRIPTION: Allows other objects to get the Mold

'PARAMETERS:  N/A

'RETURNED:    Mold

'*********************************************************************
Public Property Get Mold() As String
On Error GoTo ErrTrap
    Mold = strMold_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.Mold", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Comment()

'DESCRIPTION: Allows other objects to get the Comment

'PARAMETERS:  N/A

'RETURNED:    Comment

'*********************************************************************
Public Property Get Comment() As String
On Error GoTo ErrTrap
    Comment = strComment_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.Comment", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   AnalyzerData()

'DESCRIPTION: Allows other objects to get the analyzer data

'PARAMETERS:  N/A

'RETURNED:    the analyzer data

'*********************************************************************
Public Property Get AnalyzerData() As AnalyzerData
On Error GoTo ErrTrap
    Set AnalyzerData = objAnalyzerData_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.AnalyzerData", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Results()

'DESCRIPTION: Allows other objects to get the fit parameters

'PARAMETERS:  N/A

'RETURNED:    fit parameters collection

'*********************************************************************
Public Property Get Results() As QalibCollection
On Error GoTo ErrTrap
    Set Results = objResults_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.Results", Err.Description)
End Property

'*********************************************************************

'PROPERTY GET:   EditableFitParams()

'DESCRIPTION: Allows other objects to get the fit parameters

'PARAMETERS:  N/A

'RETURNED:    fit parameters collection

'*********************************************************************
Public Property Get EditableFitParams() As FitParams
On Error GoTo ErrTrap
    Set EditableFitParams = New FitParams
    
    Call EditableFitParams.Load(lngCalID_m)
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.EditableFitParams", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   CalibratorCheckPlot()

'DESCRIPTION: Allows other objects to get the calibrator check plot

'PARAMETERS:  N/A

'RETURNED:    calibrator check plot

'*********************************************************************
Public Property Get CalibratorCheckPlot() As Plot
On Error GoTo ErrTrap
    Set CalibratorCheckPlot = objPlots_m(CALIBRATORCHECKPLOTTEXT)
    
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.CalibratorCheckPlot", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   ControlCheckPlot()

'DESCRIPTION: Allows other objects to get the control check plot

'PARAMETERS:  N/A

'RETURNED:    control check plot

'*********************************************************************
Public Property Get ControlCheckPlot() As Plot
On Error GoTo ErrTrap
    Set ControlCheckPlot = objPlots_m(CONTROLCHECKPLOTTEXT)
    
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.PropertyGet.ControlCheckPlot", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Class_Initialize()

'DESCRIPTION: Sets up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Initialize()
    blnUpdatedFit_m = False
End Sub

'***********************************************************************

'PROCEDURE:   Class_Terminate()

'DESCRIPTION: Cleans up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Terminate()
    Set objAnalyzerData_m = Nothing
    Set objResults_m = Nothing
    Set objPlots_m = Nothing
End Sub

'***********************************************************************

'PROCEDURE:   LoadCalibration()

'DESCRIPTION:  Loads the calibration from the database

'PARAMETERS:  inID - the calibration ID

'RETURNED:    N/A

'********************************************************************
Public Sub LoadCalibration(inID As Long)
On Error GoTo ErrTrap

    Dim server As QALIBCLIENTLib.QalibClientMain
    Dim passExpPoint(0) As Variant
    Dim passCalcExpPoint(0) As Variant
    Dim passExpPointStatus(0) As Variant
    Dim passPlots(0) As Variant
    Dim error_msg As Variant
    Dim passDescriptions(0) As Variant
    Dim passValues(0) As Variant
    Dim passAdjValues(0) As Variant
    Dim passDispositions(0) As Variant
    Dim passAdjDispositions(0) As Variant
    Dim passMinLimits(0) As Variant
    Dim passMaxLimits(0) As Variant
    Dim maxResultIndex As Long
    Dim fitParamIndex As Long
    Dim calcExpPointIndex As Long
    Dim maxCalcExpPointIndex As Long
    Dim maxPlotIndex As Long
    Dim plotIndex As Long
    Dim curResult As Result
    Dim curExpPoint As ExpPoint
    Dim curPlotPoint As PlotPoint
    Dim curPlot As Plot
    Dim curSeries As SeriesOne
        
    lngCalID_m = inID
    blnUpdatedFit_m = False
    
    Set server = New QALIBCLIENTLib.QalibClientMain
    
    ' get the calibration results for the report
    
    ' get calibration results from the server
    error_msg = server.GetCalibrationResults(lngCalID_m, strChemistry_m, strWavelength_m, _
            strSpecies_m, strDiluent_m, strMold_m, strRotor_m, strCuvette_m, strMode_m, strUser_m, _
            strCalDate_m, passDescriptions(0), passValues(0), passAdjValues(0), passMinLimits(0), _
            passMaxLimits(0), passDispositions(0), passAdjDispositions(0), passExpPoint(0), _
            passExpPointStatus(0), passCalcExpPoint(0), passPlots(0), strComment_m)
          
    Call InterpretServerMsg(error_msg)
      
    ' see how many fit parameters were returned
    maxResultIndex = UBound(passDescriptions(0))
    
    ' makes sure all fit parameter arrays are congruent
    If ((maxResultIndex <> UBound(passValues(0))) Or (maxResultIndex <> UBound(passAdjValues(0))) Or _
        (maxResultIndex <> UBound(passMinLimits(0))) Or (maxResultIndex <> UBound(passMaxLimits(0))) Or _
        (maxResultIndex <> UBound(passDispositions(0))) Or (maxResultIndex <> UBound(passAdjDispositions(0)))) Then
        Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
    End If
    
    ' load fit parameters into collection
    Set objResults_m = New QalibCollection
    
    For fitParamIndex = 0 To maxResultIndex
        Set curResult = New Result
        
        Call curResult.Load(CStr(passDescriptions(0)(fitParamIndex)), CDbl(passValues(0)(fitParamIndex)), CDbl(passAdjValues(0)(fitParamIndex)), _
            CDbl(passMinLimits(0)(fitParamIndex)), CDbl(passMaxLimits(0)(fitParamIndex)), CStr(passDispositions(0)(fitParamIndex)), CStr(passAdjDispositions(0)(fitParamIndex)))
        
        Call objResults_m.Add(curResult, curResult.Name)
        
        blnUpdatedFit_m = curResult.IsModified
    Next fitParamIndex
    
    Set objAnalyzerData_m = New AnalyzerData
    
    Call UnpackRuns(objAnalyzerData_m, passExpPoint(0), passExpPointStatus(0))
        
    ' see how many calibrators there are
    maxCalcExpPointIndex = UBound(passCalcExpPoint(0))
    
    ' the calibrators' calculated measured values come in name value pairs
    For calcExpPointIndex = 0 To maxCalcExpPointIndex Step 2
        Set curSeries = objAnalyzerData_m.SeriesSet(passCalcExpPoint(0)(calcExpPointIndex))
        
        Call curSeries.Load(inCalcMeasuredVal:=CDbl(passCalcExpPoint(0)(calcExpPointIndex + 1)), _
            inIsCalibrator:=True)
    Next calcExpPointIndex
       
    ' get the assigned values
    Call GetAssigned(strChemistry_m, objAnalyzerData_m.SeriesSet)
    
    ' load the plots
    
    Set objPlots_m = New QalibCollection
    
    maxPlotIndex = UBound(passPlots(0))
    plotIndex = 0   ' start at the beginning of the array
    
    ' load each plot
    
    ' don't go past the end of the plot array
    Do While (plotIndex <= maxPlotIndex)
        Set curPlot = New Plot
    
        Call curPlot.Load(CStr(passPlots(0)(plotIndex)), CDbl(passPlots(0)(plotIndex + 1)), _
            CDbl(passPlots(0)(plotIndex + 2)), CDbl(passPlots(0)(plotIndex + 3)), _
            CDbl(passPlots(0)(plotIndex + 4)), CDbl(passPlots(0)(plotIndex + 5)))
    
        plotIndex = plotIndex + 6
        
        ' don't go past the end of the array
        Do While (plotIndex <= maxPlotIndex)
            ' keep loading points until a new plot name is found
            If (IsNumeric(passPlots(0)(plotIndex)) = False) Then
                Exit Do
            End If
            
            Set curPlotPoint = New PlotPoint
            
            ' parameters may need to be reversed
            Call curPlotPoint.Load(CDbl(passPlots(0)(plotIndex)), CDbl(passPlots(0)(plotIndex + 1)))
        
            Call curPlot.PlotPoints.Add(curPlotPoint, plotIndex)
        
            plotIndex = plotIndex + 2
        Loop
        
        ' add the plot to the report's internal collection
        Call objPlots_m.Add(curPlot, curPlot.Name)

    Loop
     
    Set curResult = Nothing
    Set curPlotPoint = Nothing
    Set curPlot = Nothing
    Set curSeries = Nothing
    Set server = Nothing

    Exit Sub
ErrTrap:
    Set curResult = Nothing
    Set curPlotPoint = Nothing
    Set curPlot = Nothing
    Set curSeries = Nothing
    Set server = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | QalibServer.LoadCalibration", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   Comment()

'DESCRIPTION: Allows other objects to change the Comment

'PARAMETERS:  the Comment

'RETURNED:  N/A

'*********************************************************************
Public Sub UpdateComment(inComment As String)
On Error GoTo ErrTrap
    Dim server As QALIBCLIENTLib.QalibClientMain
    Dim error_msg As Variant
    
    Set server = New QALIBCLIENTLib.QalibClientMain
    
    ' send comment to the server
    error_msg = server.SendComment(lngCalID_m, inComment)
    
    Call InterpretServerMsg(error_msg)

    strComment_m = inComment
    
    Set server = Nothing
    Exit Sub
ErrTrap:
    Set server = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | LotReport.UpdateComment", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   Accept()

'DESCRIPTION:  Allows other objects to accept the calibration

'PARAMETERS:  N/A

'RETURNED:    N/A

'********************************************************************
Public Sub Accept()
On Error GoTo ErrTrap
    ' if the fit has been updated
    If ((blnUpdatedFit_m = True) And (strComment_m = "")) Then
        Call Err.Raise(COMPERR, "QalibObjects", LoadResString(UNJUSTIFIEDFITPARAMERR))
    End If
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | QalibServer.Accept", Err.Description)
End Sub

