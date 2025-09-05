VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  ExpPoint.cls
 
'DESCRIPTION:  This module holds the experiment point class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: ExpPoint.cls $
 ' 
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:40p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Updated to 1.0.0X9.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:22p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Added serial number support.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe

Option Explicit

'constants

' private member variables
Private lngID_m As Long ' the ID code for the point
Private dblMeasuredVal_m As Double  ' the measured value of the point
Private lngStatus_m As Long  ' the status of the point

'***********************************************************************

'PROPERTY GET:   Status()

'DESCRIPTION: Allows friend objects to get the status

'PARAMETERS:  N/A

'RETURNED:    status

'*********************************************************************
Friend Property Get Status() As Long
On Error GoTo ErrTrap
    Status = lngStatus_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ExpPoint.PropertyGet.Status", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   ID()

'DESCRIPTION: Allows other objects to get the ID

'PARAMETERS:  N/A

'RETURNED:    ID

'*********************************************************************
Public Property Get ID() As Long
On Error GoTo ErrTrap
    ID = lngID_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ExpPoint.PropertyGet.ID", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   MeasuredVal()

'DESCRIPTION: Allows other objects to get the measured value

'PARAMETERS:  N/A

'RETURNED:    measured value

'*********************************************************************
Public Property Get MeasuredVal() As Double
On Error GoTo ErrTrap
    MeasuredVal = dblMeasuredVal_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ExpPoint.PropertyGet.MeasuredVal", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Label()

'DESCRIPTION: Allows other objects to get the label

'PARAMETERS:  N/A

'RETURNED:    label

'*********************************************************************
Public Property Get Label() As String
On Error GoTo ErrTrap
    If ((lngStatus_m And EXCLUDESTATUS) <> 0) Then
        Label = EXCLUDETEXT
    ElseIf ((lngStatus_m And SERIESCHANGESTATUS) <> 0) Then
        Label = SERIESCHANGETEXT
    ElseIf ((lngStatus_m And OUTLIERSTATUS) <> 0) Then
        Label = OUTLIERTEXT
    Else
        Label = ""
    End If
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ExpPoint.PropertyGet.Label", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   IsExcluded()

'DESCRIPTION: Allows other objects to get the excluded status

'PARAMETERS:  N/A

'RETURNED:    excluded status

'*********************************************************************
Public Property Get IsExcluded() As Boolean
On Error GoTo ErrTrap
    If ((lngStatus_m And EXCLUDESTATUS) <> 0) Then
        IsExcluded = True
    Else
        IsExcluded = False
    End If
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ExpPoint.PropertyGet.IsExcluded", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Load()

'DESCRIPTION: Allows a friend object to load the point

'PARAMETERS:  inID - the point ID
'             inMeasuredVal - the measured value
'             inStatus - the status

'RETURNED:    N/A

'*********************************************************************
Friend Sub Load(Optional inID As Variant, Optional inMeasuredVal As Variant, Optional inStatus As Variant)
On Error GoTo ErrTrap
    
    ' only load the attributes that were provided
    If (IsMissing(inID) = False) Then
        lngID_m = CLng(inID)
    End If
    
    If (IsMissing(inMeasuredVal) = False) Then
        dblMeasuredVal_m = CDbl(inMeasuredVal)
    End If
    
    If (IsMissing(inStatus) = False) Then
        lngStatus_m = CLng(inStatus)
    End If
    
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ExpPoint.Load", Err.Description)
End Sub
