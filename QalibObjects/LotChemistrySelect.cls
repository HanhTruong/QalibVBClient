VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LotChemistrySelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************

'File:  LotChemistrySelect.cls

'Description:  This module contains the lot chemistry select class.

'Compiler:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: LotChemistrySelect.cls $
 ' 
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:54a
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Added check to make sure user can calibrate at least one chemistry.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:40p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Updated to 1.0.0X9.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:21p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe.
 
Option Explicit

' private member variables
Private objChems_m As QalibCollection

'***********************************************************************

'PROPERTY GET:   Chems()

'DESCRIPTION: Allows other objects to get the chemistry set

'PARAMETERS:  N/A

'RETURNED:    chems set

'*********************************************************************
Public Property Get Chems() As QalibCollection
On Error GoTo ErrTrap
    Set Chems = objChems_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | LotChemistrySelect.PropertyGet.chems", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Class_Initialize()

'DESCRIPTION: Sets up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Initialize()
    Set objChems_m = New QalibCollection
End Sub

'***********************************************************************

'PROCEDURE:   Class_Terminate()

'DESCRIPTION: Cleans up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Terminate()
    Set objChems_m = Nothing
End Sub

'***********************************************************************

'PROCEDURE:    LoadChems()

'DESCRIPTION:  Requests and parses the list of chemistries from the server

'PARAMETERS:   inCriteria  - the runs to get the chemistries from
'              inUser - the user name
'              inPassword - the password

'RETURNED:    N/A

'***********************************************************************
Public Sub LoadChems(inCriteria As Variant, inUser As String, inPassword As String)
On Error GoTo ErrTrap
    Dim curChem As Chemistry
    Dim server As ExcelIO
    Dim passChem() As String
    Dim chemIndex As Integer
    Dim maxChemIndex As Integer
    Dim objUserVerify As UserVerifiy
    
    Set objUserVerify = New UserVerifiy
    Set server = New ExcelIO
    
    ' get the list of chemistries from the server
    Call server.GetChems(CStr(inCriteria), passChem)
        
    ' get a list of chemistries the user is allowed to calibrate
    Call objUserVerify.Authenticate(inUser, inPassword)
    
    maxChemIndex = UBound(passChem)
    
    ' load the chemistries into the collection
    For chemIndex = 0 To maxChemIndex
        
        ' see if the returned chemistry is in the user's list of allowed chemistries
        Set curChem = objUserVerify.Chems(passChem(chemIndex))
        
        ' if the chemistry is allowed then add it to the collection
        If (curChem Is Nothing = False) Then
        
            Set curChem = New Chemistry
        
            Call curChem.Load(passChem(chemIndex))
       
            Call objChems_m.Add(curChem, curChem.Name)
            
        End If
        
    Next chemIndex
    
    ' make sure the user has a list of chemistries to calibrate
    If (objChems_m.Count = 0) Then
        Call Err.Raise(COMPERR, "QalibObjects", BuildSubstString(LoadResString(NOAUTHCHEMSERR), inUser, inCriteria))
    End If
    
    Set curChem = Nothing
    Set server = Nothing
    
    Exit Sub
ErrTrap:
    Set curChem = Nothing
    Set server = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | LotChemistrySelect.LoadChems", Err.Description)
End Sub

