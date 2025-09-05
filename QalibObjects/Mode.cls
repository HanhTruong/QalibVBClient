VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  Mode.cls
 
'DESCRIPTION:  This module holds the mode.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: Mode.cls $
 ' 
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:42p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe.

Option Explicit

' private member variables
Private strName_m As String ' name

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
    Call Err.Raise(Err.Number, Err.Source & " | Mode.PropertyGet.Name", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Load()

'DESCRIPTION: Allows friend objects to load the mode

'PARAMETERS:  inName - the name

'RETURNED:    N/A

'*********************************************************************
Friend Sub Load(inName As String)
On Error GoTo ErrTrap
    
    strName_m = inName

    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Mode.Load", Err.Description)
End Sub


