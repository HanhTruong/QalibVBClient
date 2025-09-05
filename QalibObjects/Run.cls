VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Run"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  Run.cls
 
'DESCRIPTION:  This module holds the run class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: Run.cls $
 ' 
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:42p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe.

Option Explicit

'constants

' private member variables
Private lngID_m As Long ' the ID code for the run
Private strSerial_m As String   ' the serial number of the instrument
Private strBarCode_m As String   ' the rotor bar code
Private strSpecies_m As String   ' the species from the analyzer
Private lngSavedStatus_m As Long  ' the saved status of the run while being edited
Private strOrigSeries_m As String  ' the original series of the run when it was first loaded

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
    Call Err.Raise(Err.Number, Err.Source & " | Run.PropertyGet.ID", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Serial()

'DESCRIPTION: Allows other objects to get the serial number

'PARAMETERS:  N/A

'RETURNED:    serial number

'*********************************************************************
Public Property Get Serial() As String
On Error GoTo ErrTrap
    Serial = strSerial_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Run.PropertyGet.Serial", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   BarCode()

'DESCRIPTION: Allows other objects to get the rotor bar code

'PARAMETERS:  N/A

'RETURNED:    bar code

'*********************************************************************
Public Property Get BarCode() As String
On Error GoTo ErrTrap
    BarCode = strBarCode_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Run.PropertyGet.BarCode", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Species()

'DESCRIPTION: Allows other objects to get the species

'PARAMETERS:  N/A

'RETURNED:    species

'*********************************************************************
Public Property Get Species() As String
On Error GoTo ErrTrap
    Species = strSpecies_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Run.PropertyGet.Species", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   OrigSeries()

'DESCRIPTION: Allows other objects to get the original series

'PARAMETERS:  N/A

'RETURNED:    Original series

'*********************************************************************
Public Property Get OrigSeries() As String
On Error GoTo ErrTrap
    OrigSeries = strOrigSeries_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Run.PropertyGet.OrigSeries", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Load()

'DESCRIPTION: Allows a friend object to load the run

'PARAMETERS:  inID - the run ID
'             inSerial - the serial number of the instrument
'             inBarCode - the bar code for the rotor
'             inSpecies - the species from the analyzer
'             inOrigSeries - the original series

'RETURNED:    N/A

'*********************************************************************
Friend Sub Load(inID As Long, inSerial As String, inBarCode As String, _
    inSpecies As String, inOrigSeries As String)
On Error GoTo ErrTrap
    
    lngID_m = inID
    strSerial_m = inSerial
    strBarCode_m = inBarCode
    strSpecies_m = inSpecies
    strOrigSeries_m = inOrigSeries
    
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Run.Load", Err.Description)
End Sub


