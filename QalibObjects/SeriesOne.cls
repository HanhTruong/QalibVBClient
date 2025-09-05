VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SeriesOne"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  SeriesOne.cls
 
'DESCRIPTION:  This module holds the series class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: SeriesOne.cls $
 ' 
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:40p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Updated to 1.0.0X9.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe

Option Explicit

' private member variables

Private strName_m As String  ' the label for the series
Private dblAssignedVal_m As Double ' the assigned value
Private dblCalcMeasuredVal_m As Double  ' the calculated measured value
Private blnIsCalibrator_m As Boolean  ' is the series a calibrator
Private objRuns_m As QalibCollection ' the runs collection

'***********************************************************************

'PROPERTY GET:   Name()

'DESCRIPTION: Allows other objects to get the Name

'PARAMETERS:  N/A

'RETURNED:    Name

'*********************************************************************
Public Property Get Name() As String
On Error GoTo ErrTrap
    Name = strName_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | SeriesOne.PropertyGet.Name", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   AssignedVal()

'DESCRIPTION: Allows other objects to get the AssignedVal

'PARAMETERS:  N/A

'RETURNED:    AssignedVal

'*********************************************************************
Public Property Get AssignedVal() As Double
On Error GoTo ErrTrap
    AssignedVal = Round(dblAssignedVal_m, SAMPLEROUND)
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | SeriesOne.PropertyGet.AssignedVal", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   CalcMeasuredVal()

'DESCRIPTION: Allows other objects to get the calculated measured value

'PARAMETERS:  N/A

'RETURNED:    calculated measured value

'*********************************************************************
Public Property Get CalcMeasuredVal() As Double
On Error GoTo ErrTrap
    CalcMeasuredVal = Round(dblCalcMeasuredVal_m, SAMPLEROUND)
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | SeriesOne.PropertyGet.CalcMeasuredVal", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   IsCalibrator()

'DESCRIPTION: Allows other objects to see if the series is a calibrator

'PARAMETERS:  N/A

'RETURNED:    whether the series is a calibrator

'*********************************************************************
Public Property Get IsCalibrator() As Boolean
On Error GoTo ErrTrap
    IsCalibrator = blnIsCalibrator_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | SeriesOne.PropertyGet.IsCalibrator", Err.Description)
End Property


'***********************************************************************

'PROPERTY GET:   Runs()

'DESCRIPTION: Allows other objects to get the run set

'PARAMETERS:  N/A

'RETURNED:    Runs set

'*********************************************************************
Public Property Get Runs() As QalibCollection
On Error GoTo ErrTrap
    Set Runs = objRuns_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | SeriesOne.PropertyGet.Runs", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Load()

'DESCRIPTION: Allows friend objects to load the series

'PARAMETERS:  inName - the name
'             inAssignedVal - the assigned value
'             inCalcMeasuredVal - the calculated measured value
'             inIsCalibrator - is the series a calibrator

'RETURNED:    N/A

'*********************************************************************
Friend Sub Load(Optional inName As Variant, Optional inAssignedVal As Variant, _
    Optional inCalcMeasuredVal As Variant, Optional inIsCalibrator As Variant)
On Error GoTo ErrTrap
    
    ' only load the attributes that were provided
    
    If (IsMissing(inName) = False) Then
        strName_m = CStr(inName)
    End If
    
    If (IsMissing(inAssignedVal) = False) Then
        dblAssignedVal_m = CDbl(inAssignedVal)
    End If
    
    If (IsMissing(inCalcMeasuredVal) = False) Then
        dblCalcMeasuredVal_m = CDbl(inCalcMeasuredVal)
    End If
    
    If (IsMissing(inIsCalibrator) = False) Then
        blnIsCalibrator_m = CBool(inIsCalibrator)
    End If
        
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | SeriesOne.Load", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   Class_Initialize()

'DESCRIPTION: Sets up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Initialize()
    Set objRuns_m = New QalibCollection
End Sub

'***********************************************************************

'PROCEDURE:   Class_Terminate()

'DESCRIPTION: Cleans up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Terminate()
    Set objRuns_m = Nothing
End Sub

