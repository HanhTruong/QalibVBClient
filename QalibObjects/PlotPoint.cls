VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PlotPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  PlotPoint.cls
 
'DESCRIPTION:  This module holds the plot point class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: PlotPoint.cls $
 ' 
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe

Option Explicit

' private member variables
Private dblXValue_m As Double  ' the X value of the point
Private dblYValue_m As Double  ' the Y value of the point

'***********************************************************************

'PROPERTY GET:   XValue()

'DESCRIPTION: Allows other objects to get the X value

'PARAMETERS:  N/A

'RETURNED:    X value

'*********************************************************************
Public Property Get XValue() As Double
On Error GoTo ErrTrap
    XValue = Round(dblXValue_m, SAMPLEROUND)
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | PlotPoint.PropertyGet.XValue", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   YValue()

'DESCRIPTION: Allows other objects to get the Y value

'PARAMETERS:  N/A

'RETURNED:    Y value

'*********************************************************************
Public Property Get YValue() As Double
On Error GoTo ErrTrap
    YValue = Round(dblYValue_m, SAMPLEROUND)
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | PlotPoint.PropertyGet.YValue", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Load()

'DESCRIPTION: Allows friend objects to load the plot point

'PARAMETERS:    inXValue - the X value
'               inYValue - the Y value

'RETURNED:    N/A

'*********************************************************************
Friend Sub Load(inXValue As Double, inYValue As Double)
On Error GoTo ErrTrap
    
    dblXValue_m = inXValue
    dblYValue_m = inYValue
    
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | PlotPoint.Load", Err.Description)
End Sub
