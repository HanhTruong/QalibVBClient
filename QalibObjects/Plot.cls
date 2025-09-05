VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  Plot.cls
 
'DESCRIPTION:  This module holds the plot class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: Plot.cls $
 ' 
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:40p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Updated to 1.0.0X9.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:26p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe

Option Explicit

' private member variables
Private strName_m As String ' name
Private dblSlope_m As Double ' slope
Private dblIntercept_m As Double ' intercept
Private dblSlopeErr_m As Double ' slope error
Private dblInterceptErr_m As Double ' intercept error
Private dblCorrCoeff_m As Double ' correlation coefficient
Private objPlotPoints_m As QalibCollection  ' the points

'***********************************************************************

'PROPERTY GET:   Name()

'DESCRIPTION: Allows other objects to get the Name

'PARAMETERS:  N/A

'RETURNED:    the Name

'*********************************************************************
Public Property Get Name() As String
On Error GoTo ErrTrap
    Name = strName_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Plot.PropertyGet.Name", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Slope()

'DESCRIPTION: Allows other objects to get the Slope

'PARAMETERS:  N/A

'RETURNED:    the Slope

'*********************************************************************
Public Property Get Slope() As Double
On Error GoTo ErrTrap
    Slope = dblSlope_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Plot.PropertyGet.Slope", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Intercept()

'DESCRIPTION: Allows otherc objects to get the intercept

'PARAMETERS:  N/A

'RETURNED:    the intercept

'*********************************************************************
Public Property Get Intercept() As Double
On Error GoTo ErrTrap
    Intercept = dblIntercept_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Plot.PropertyGet.Intercept", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   SlopeErr()

'DESCRIPTION: Allows other objects to get the slope error

'PARAMETERS:  N/A

'RETURNED:    the slope error

'*********************************************************************
Public Property Get SlopeErr() As Double
On Error GoTo ErrTrap
    SlopeErr = dblSlopeErr_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Plot.PropertyGet.SlopeErr", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   InterceptErr()

'DESCRIPTION: Allows other objects to get the intercept error

'PARAMETERS:  N/A

'RETURNED:    the intercept error

'*********************************************************************
Public Property Get InterceptErr() As Double
On Error GoTo ErrTrap
    InterceptErr = dblInterceptErr_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Plot.PropertyGet.InterceptErr", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   CorrCoeff()

'DESCRIPTION: Allows other objects to get the correlation coefficient

'PARAMETERS:  N/A

'RETURNED:    the correlation coefficient

'*********************************************************************
Public Property Get CorrCoeff() As Double
On Error GoTo ErrTrap
    CorrCoeff = dblCorrCoeff_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Plot.PropertyGet.CorrCoeff", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   PlotPoints()

'DESCRIPTION: Allows other objects to get the points

'PARAMETERS:  N/A

'RETURNED:    points object

'*********************************************************************
Public Property Get PlotPoints() As QalibCollection
On Error GoTo ErrTrap
    Set PlotPoints = objPlotPoints_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Plot.PropertyGet.PlotPoints", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Load()

'DESCRIPTION: Allows friend objects to load the plot

'PARAMETERS:    inName - the name
'               inCorrCoeff - correlation coefficient
'               inSlope - the slope
'               inSlopeErr - slope error
'               inIntercept - intercept
'               inInterceptErr - intercept error

'RETURNED:    N/A

'*********************************************************************
Friend Sub Load(inName As String, inCorrCoeff As Double, inSlope As Double, inSlopeErr As Double, inIntercept As Double, _
    inInterceptErr As Double)
On Error GoTo ErrTrap
    
    strName_m = inName
    dblSlope_m = inSlope
    dblIntercept_m = inIntercept
    dblSlopeErr_m = inSlopeErr
    dblInterceptErr_m = inInterceptErr
    dblCorrCoeff_m = inCorrCoeff
    
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Plot.Load", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   Class_Initialize()

'DESCRIPTION: Sets up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Initialize()
    Set objPlotPoints_m = New QalibCollection
End Sub

'***********************************************************************

'PROCEDURE:   Class_Terminate()

'DESCRIPTION: Cleans up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Terminate()
    Set objPlotPoints_m = Nothing
End Sub








