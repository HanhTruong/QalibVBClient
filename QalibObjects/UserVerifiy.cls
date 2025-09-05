VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UserVerifiy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  UserVerify.cls
 
'DESCRIPTION:  This module holds the user verification class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: UserVerifiy.cls $
 ' 
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:57a
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Added check to make sure there are modes and chemistries for the user.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:42p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe.

Option Explicit

Private objModes_m As QalibCollection
Private objChems_m As QalibCollection

'***********************************************************************

'PROPERTY GET:   Modes()

'DESCRIPTION: Allows other objects to get the modes

'PARAMETERS:  N/A

'RETURNED:    modes collection

'*********************************************************************
Public Property Get Modes() As QalibCollection
On Error GoTo ErrTrap
    Set Modes = objModes_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | UserVerify.PropertyGet.Modes", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Chems()

'DESCRIPTION: Allows other objects to get the chemistries

'PARAMETERS:  N/A

'RETURNED:    chemistries collection

'*********************************************************************
Public Property Get Chems() As QalibCollection
On Error GoTo ErrTrap
    Set Chems = objChems_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | UserVerify.PropertyGet.Chems", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Authenticate()

'DESCRIPTION: Authenticates the user and password with the system

'PARAMETERS:  inUser - the user name
'             inPassword - the password

'RETURNED:    N/A

'********************************************************************
Public Sub Authenticate(inUser As String, inPassword As String)
On Error GoTo ErrTrap
    Dim server As QALIBCLIENTLib.QalibClientMain
    Dim passMode(0) As Variant
    Dim passChem(0) As Variant
    Dim index As Long
    Dim maxIndex As Long
    Dim curMode As Mode
    Dim curChem As Chemistry
    Dim error_msg As Variant
    
    Set server = New QALIBCLIENTLib.QalibClientMain

    ' pass the username and password to the server and get the applicable modes and chemistries for the user
    error_msg = server.GetUserRights(inUser, inPassword, passMode(0), passChem(0))
   
    Call InterpretServerMsg(error_msg)
    
    ' make sure their are modes for the user
    If (IsEmpty(passMode(0)) = True) Then
        Call Err.Raise(COMPERR, "QalibObjects", BuildSubstString(LoadResString(NOUSERMODESERR), inUser))
    End If
    
    ' make sure their are chemistries for the user
    If (IsEmpty(passChem(0)) = True) Then
        Call Err.Raise(COMPERR, "QalibObjects", BuildSubstString(LoadResString(NOUSERCHEMSERR), inUser))
    End If
    
    ' load the modes
    maxIndex = UBound(passMode(0))
    
    For index = 0 To maxIndex
        Set curMode = New Mode
        
        Call curMode.Load(CStr(passMode(0)(index)))
        
        Call objModes_m.Add(curMode, curMode.Name)
    Next index
    
    ' load the chemistries
    maxIndex = UBound(passChem(0))
    
    For index = 0 To maxIndex
        Set curChem = New Chemistry
        
        Call curChem.Load(CStr(passChem(0)(index)))
        
        Call objChems_m.Add(curChem, curChem.Name)
    Next index
    
    Set curChem = Nothing
    Set curMode = Nothing
    Set server = Nothing
    
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | UserVerify.Authenticate", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   Class_Initialize()

'DESCRIPTION: Sets up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Initialize()
    Set objModes_m = New QalibCollection
    Set objChems_m = New QalibCollection
End Sub

'***********************************************************************

'PROCEDURE:   Class_Terminate()

'DESCRIPTION: Cleans up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Terminate()
    Set objModes_m = Nothing
    Set objChems_m = Nothing
End Sub
