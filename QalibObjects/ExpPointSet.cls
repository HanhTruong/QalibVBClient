VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExpPointSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  ExpPointSet.cls
 
'DESCRIPTION:  This module holds the experiment point set class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: ExpPointSet.cls $
 ' 
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:42p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe.

Option Explicit

' private member variables

Private strName_m As String  ' the label for the point set
Private objExpPoints_m As QalibCollection ' the points collection

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
    Call Err.Raise(Err.Number, Err.Source & " | ExpPointSet.PropertyGet.Name", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   ExpPoints()

'DESCRIPTION: Allows other objects to get the point set

'PARAMETERS:  N/A

'RETURNED:    Points set

'*********************************************************************
Public Property Get ExpPoints() As QalibCollection
On Error GoTo ErrTrap
    Set ExpPoints = objExpPoints_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ExpPointSet.PropertyGet.ExpPoints", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Load()

'DESCRIPTION: Allows friend objects to load the series

'PARAMETERS:  inName - the name

'RETURNED:    N/A

'*********************************************************************
Friend Sub Load(inName As String)
On Error GoTo ErrTrap
    
    strName_m = inName
    
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ExpPointSet.Load", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   Class_Initialize()

'DESCRIPTION: Sets up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Initialize()
    Set objExpPoints_m = New QalibCollection
End Sub

'***********************************************************************

'PROCEDURE:   Class_Terminate()

'DESCRIPTION: Cleans up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Terminate()
    Set objExpPoints_m = Nothing
End Sub
