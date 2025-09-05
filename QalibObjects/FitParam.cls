VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FitParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  FitParam.cls
 
'DESCRIPTION:  This module holds the fit parameter class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: FitParam.cls $
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
Private strName_m As String ' name
Private dblValue_m As Double ' fit parameter value
Private dblSavedValue_m As Double ' saved fit parameter value

'***********************************************************************

'PROPERTY GET:   IsModified()

'DESCRIPTION: Allows friend objects to get whether the fit parameter was modified

'PARAMETERS:  N/A

'RETURNED:    whether the fit parameter was modified

'*********************************************************************
Public Property Get IsModified() As Boolean
On Error GoTo ErrTrap
    If (dblValue_m <> dblSavedValue_m) Then
        IsModified = True
    Else
        IsModified = False
    End If
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | FitParam.PropertyGet.IsModified", Err.Description)
End Property

'***********************************************************************

'PROPERTY LET:   Value()

'DESCRIPTION: Allows other objects to change the Value

'PARAMETERS:  the Value

'RETURNED:  N/A

'*********************************************************************
Public Property Let Value(inValue As Double)
On Error GoTo ErrTrap
    dblValue_m = inValue
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | FitParam.PropertyGet.Value", Err.Description)
End Property

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
    Call Err.Raise(Err.Number, Err.Source & " | FitParam.PropertyGet.Name", Err.Description)
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
    Call Err.Raise(Err.Number, Err.Source & " | FitParam.PropertyGet.Value", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Load()

'DESCRIPTION: Allows friend objects to load the fit parameter

'PARAMETERS:  inName - the name
'             inValue - the value

'RETURNED:    N/A

'*********************************************************************
Friend Sub Load(inName As String, inValue As Double)
On Error GoTo ErrTrap
    
    strName_m = inName
    dblValue_m = inValue

    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | FitParam.Load", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   BeginEdit()

'DESCRIPTION: Initiaites editing on a fit param

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Friend Sub BeginEdit()
On Error GoTo ErrTrap
    dblSavedValue_m = dblValue_m
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | FitParam.BeginEdit", Err.Description)
End Sub

