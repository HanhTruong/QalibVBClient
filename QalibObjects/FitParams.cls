VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FitParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  FitParams.cls
 
'DESCRIPTION:  This module holds the lot fit parameter edit class.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: FitParams.cls $
 ' 
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:40p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Updated to 1.0.0X9.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:26p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Now loads the fit parameters based off of a calibration ID instead of
 ' getting a pointer to the lot object.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe

Option Explicit

' private member variables
Private lngCalID_m As Long     ' the unique calibration ID
Private objFitParams_m As QalibCollection ' the fit parameters collection

'***********************************************************************

'PROPERTY GET:   FitParamsSet()

'DESCRIPTION: Allows other objects to get the fit parameters

'PARAMETERS:  N/A

'RETURNED:    fit parameters collection

'*********************************************************************
Public Property Get FitParamsSet() As QalibCollection
On Error GoTo ErrTrap
    Set FitParamsSet = objFitParams_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | FitParams.PropertyGet.FitParamsSet", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Class_Terminate()

'DESCRIPTION: Event handler for when the class terminates

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Terminate()
On Error GoTo ErrTrap
    Set objFitParams_m = Nothing
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | FitParams.Class_Terminate", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   Load()

'DESCRIPTION: Loads the editable fit parameters

'PARAMETERS:  inCalID - the calibration ID

'RETURNED:    N/A

'********************************************************************
Friend Sub Load(inCalID As Long)
On Error GoTo ErrTrap
    Dim server As QALIBCLIENTLib.QalibClientMain
    Dim passFitParams(0) As Variant
    Dim fitParamIndex As Long
    Dim maxFitParamIndex As Long
    Dim curFitParam As FitParam
    Dim error_msg As Variant
    
    lngCalID_m = inCalID
    
    Set server = New QALIBCLIENTLib.QalibClientMain
    
    ' get editable fit parameters from the server
    error_msg = server.GetEditableFitParameters(lngCalID_m, passFitParams(0))
    
    Call InterpretServerMsg(error_msg)

    ' load the editable fit parameters into the internal collection
    maxFitParamIndex = UBound(passFitParams(0))
    
    ' the array should have an even number of elements
    If (((maxFitParamIndex + 1) Mod 2) <> 0) Then
        Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
    End If
    
    Set objFitParams_m = New QalibCollection
    
    ' the fit parameters come back from the server in name, value pairs
    For fitParamIndex = 0 To maxFitParamIndex Step 2
        Set curFitParam = New FitParam
        Call curFitParam.Load(CStr(passFitParams(0)(fitParamIndex)), CDbl(passFitParams(0)(fitParamIndex + 1)))
        Call curFitParam.BeginEdit
        Call objFitParams_m.Add(curFitParam, curFitParam.Name)
    Next fitParamIndex
    
    Set server = Nothing
    Set curFitParam = Nothing
    Exit Sub
ErrTrap:
    Set server = Nothing
    Set curFitParam = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | FitParams.Load", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   CommitEdits()

'DESCRIPTION: Sends adjusted fit parameters to the server

'PARAMETERS:  N/A

'RETURNED:    N/A

'********************************************************************
Public Sub CommitEdits()
On Error GoTo ErrTrap
    Dim server As QALIBCLIENTLib.QalibClientMain
    Dim passFitParams(0) As Variant
    Dim curFitParam As FitParam
    Dim strFitParams() As String
    Dim fitParamIndex As Long
    Dim error_msg As Variant
        
    ' look for updated fit parameters and load them into the sending array
    For Each curFitParam In objFitParams_m
        If (curFitParam.IsModified = True) Then
            ReDim Preserve strFitParams(fitParamIndex + 1)
            
            strFitParams(fitParamIndex) = curFitParam.Name
            strFitParams(fitParamIndex + 1) = CStr(curFitParam.Value)
            
            ' fit parameters are loaded into array by name, value pairs
            fitParamIndex = fitParamIndex + 2
        End If
    Next curFitParam
    
    Set server = New QALIBCLIENTLib.QalibClientMain
    
    passFitParams(0) = strFitParams
        
    ' send updated fit parameters to the server
    error_msg = server.SendUpdatedFitParameters(lngCalID_m, passFitParams(0))
    
    Call InterpretServerMsg(error_msg)
    
    Set server = Nothing
    Set curFitParam = Nothing

    Exit Sub
ErrTrap:
    Set curFitParam = Nothing
    Set server = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | FitParams.CommitEdits", Err.Description)
End Sub

