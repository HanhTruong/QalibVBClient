VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Result"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  Result.cls
 
'DESCRIPTION:  This module holds the result class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: Result.cls $
 ' 
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe

Option Explicit

' constants
Private Const RESULTROUND As Long = 3
Private Const PASSTEXT As String = "PASS"
Private Const FAILTEXT As String = "FAIL"
Private Const NOMAXLIMIT As Double = 999999999
Private Const NOMINLIMIT As Double = -NOMAXLIMIT
Private Const NATEXT As String = "N/A"

' private member variables
Private strName_m As String ' name
Private dblValue_m As Double ' original value
Private dblAdjValue_m As Double ' adjusted value
Private dblMinLimit_m As Double ' minimum limit
Private dblMaxLimit_m As Double ' maximum limit
Private strDisposition_m As String ' disposition
Private strAdjDisposition_m As String ' the adjusted disposition

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
    Call Err.Raise(Err.Number, Err.Source & " | Result.PropertyGet.Name", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Value()

'DESCRIPTION: Allows other objects to get the Value

'PARAMETERS:  N/A

'RETURNED:    the Value

'*********************************************************************
Public Property Get Value() As Double
On Error GoTo ErrTrap
    Value = dblValue_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Result.PropertyGet.Value", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   AdjValue()

'DESCRIPTION: Allows other objects to get the Adjusted Value

'PARAMETERS:  N/A

'RETURNED:    the Adjusted Value

'*********************************************************************
Public Property Get AdjValue() As Variant
On Error GoTo ErrTrap
    If (IsModified = False) Then
        AdjValue = NATEXT
    Else
        AdjValue = dblAdjValue_m
    End If
    
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Result.PropertyGet.AdjValue", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   MinLimit()

'DESCRIPTION: Allows other objects to get the Min Limit

'PARAMETERS:  N/A

'RETURNED:    the Min Limit

'*********************************************************************
Public Property Get MinLimit() As Variant
On Error GoTo ErrTrap
    If (dblMinLimit_m = NOMINLIMIT) Then
        MinLimit = NATEXT
    Else
        MinLimit = dblMinLimit_m
    End If
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Result.PropertyGet.MinLimit", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   MaxLimit()

'DESCRIPTION: Allows other objects to get the Max Limit

'PARAMETERS:  N/A

'RETURNED:    the Max Limit

'*********************************************************************
Public Property Get MaxLimit() As Variant
On Error GoTo ErrTrap
    If (dblMaxLimit_m = NOMAXLIMIT) Then
        MaxLimit = NATEXT
    Else
        MaxLimit = dblMaxLimit_m
    End If
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Result.PropertyGet.MaxLimit", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Disposition()

'DESCRIPTION: Allows other objects to get the Disposition

'PARAMETERS:  N/A

'RETURNED:    the Disposition

'*********************************************************************
Public Property Get Disposition() As String
On Error GoTo ErrTrap
    Disposition = strDisposition_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Result.PropertyGet.Disposition", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   AdjDisposition()

'DESCRIPTION: Allows other objects to get the Adjusted Disposition

'PARAMETERS:  N/A

'RETURNED:    the Adjusted Disposition

'*********************************************************************
Public Property Get AdjDisposition() As String
On Error GoTo ErrTrap
    If (IsModified = False) Then
        AdjDisposition = NATEXT
    Else
        AdjDisposition = strAdjDisposition_m
    End If
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Result.PropertyGet.AdjDisposition", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   IsModified()

'DESCRIPTION: Allows other objects to get whether the result was modified

'PARAMETERS:  N/A

'RETURNED:    whether the result was modified

'*********************************************************************
Public Property Get IsModified() As Boolean
On Error GoTo ErrTrap
    If (dblValue_m <> dblAdjValue_m) Then
        IsModified = True
    Else
        IsModified = False
    End If
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Result.PropertyGet.IsModified", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   IsValPassed()

'DESCRIPTION: Allows other objects to get whether the result passed

'PARAMETERS:  N/A

'RETURNED:    whether the result passed

'*********************************************************************
Public Property Get IsValPassed() As Boolean
On Error GoTo ErrTrap
    If (strDisposition_m = PASSTEXT) Then
        IsValPassed = True
    Else
        IsValPassed = False
    End If
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Result.PropertyGet.IsValPassed", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   IsAdjValPassed()

'DESCRIPTION: Allows other objects to get whether the adjusted result passed

'PARAMETERS:  N/A

'RETURNED:    whether the adjusted result passed

'*********************************************************************
Public Property Get IsAdjValPassed() As Boolean
On Error GoTo ErrTrap
    If (strAdjDisposition_m = PASSTEXT) Then
        IsAdjValPassed = True
    Else
        IsAdjValPassed = False
    End If
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Result.PropertyGet.IsAdjValPassed", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Load()

'DESCRIPTION: Allows friend objects to load the result

'PARAMETERS:    inName - the name
'               inValue - the value
'               inAdjValue - the adjusted value
'               inMinLimit - the minimum limit
'               inMaxLimit - the maximum limit
'               inDisposition - the disposition
'               inAdjDisposition - the adjusted disposition

'RETURNED:    N/A

'*********************************************************************
Friend Sub Load(inName As String, inValue As Double, inAdjValue As Double, inMinLimit As Double, _
    inMaxLimit As Double, inDisposition As String, inAdjDisposition As String)
On Error GoTo ErrTrap
    
    strName_m = inName
    dblValue_m = inValue
    dblAdjValue_m = inAdjValue
    dblMinLimit_m = inMinLimit
    dblMaxLimit_m = inMaxLimit
    strDisposition_m = inDisposition
    strAdjDisposition_m = inAdjDisposition

    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Result.Load", Err.Description)
End Sub
