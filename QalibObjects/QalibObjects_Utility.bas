Attribute VB_Name = "QalibObjects_Utility"
'**********************************************************************************
 
'FILE:  Utility.cls
 
'DESCRIPTION:  This module holds the shared constants, functions, etc. for the QalibObjects project.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: QalibObjects_Utility.bas $
 ' 
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:56a
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Added check to make sure all points for a run were valid.
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:40p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Updated to 1.0.0X9.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:28p
 ' Updated in $/QalibVBClient/Source/QalibObjects
 ' Assigned values are now chemistry specific.
 ' Added ability to interpret server error messages.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe

Option Explicit

' constants

Public Const OUTLIERTEXT As String = "Outlier"
Public Const SERIESCHANGETEXT As String = "Moved"
Public Const EXCLUDETEXT As String = "Excluded"

' released software flag
Private Const RELEASEDSW As Long = 0

' for finding measured value suppressions
Private Const MEASUREDVALERRLEN As Long = 3
Private Const MEASUREDVALERRTEXT As String = "Err"

' measured status
Public Const OUTLIERSTATUS As Long = 1
Public Const SERIESCHANGESTATUS As Long = 2
Public Const EXCLUDESTATUS As Long = 4

' component errors
Public Const COMPERR As Long = vbObjectError + &H1100

' error strings
Public Const INCONGRUENTARRAYERR As Long = 101
Public Const COLLECTIONERR As Long = 102
Public Const NOSERVERRESPONSEERR As Long = 103
Public Const UNEXPECTEDSERVERERR As Long = 104
Public Const UNJUSTIFIEDFITPARAMERR As Long = 105
Public Const MINRUNSERRTEXT As Long = 106
Public Const ALLRUNSELIMINATEDERR As Long = 107
Public Const RUNNOTFOUNDERR As Long = 108
Public Const NOUSERMODESERR As Long = 109
Public Const NOUSERCHEMSERR As Long = 110
Public Const NOAUTHCHEMSERR As Long = 111

' rounding constants
Public Const SAMPLEROUND As Long = 3

' API declares
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest _
    As Any, lpvSource As Any, ByVal cbCopy As Long)


' helper functions
'***********************************************************************

'PROCEDURE:   GetAssigned()

'DESCRIPTION: Gets the assigned values from the server for the series' names

'PARAMETERS:  inChem - the chemistry being calibrated
'             inSeriesSet - the series set object in which to put the assigned values
'

'RETURNED:    N/A

'********************************************************************
Public Sub GetAssigned(inChem As String, inSeriesSet As QalibCollection)
On Error GoTo ErrTrap
    Dim passSeries(0) As Variant
    Dim passAssigned(0) As Variant
    Dim curSeries As SeriesOne
    Dim maxSeriesIndex As Long
    Dim seriesIndex As Long
    Dim numSeries As Long
    Dim structSeries() As String
    Dim server As QALIBCLIENTLib.QalibClientMain
    Dim error_msg As Variant
    
    ' the server needs all the series labels in a string array with the first
    ' element being the number of series
    
    ' set up the flat array
    numSeries = inSeriesSet.Count
    ReDim structSeries(numSeries) As String
    
    ' the first element is the number of series.
    structSeries(0) = CStr(numSeries)
    
    ' fill in the rest of the pass array with the series names
    seriesIndex = 1
    For Each curSeries In inSeriesSet
        structSeries(seriesIndex) = curSeries.Name
        seriesIndex = seriesIndex + 1
    Next curSeries
    
    ' This trick is necessary to pass the series labels to the COM object
    passSeries(0) = structSeries

    ' Get assigned values from the database.
    Set server = New QALIBCLIENTLib.QalibClientMain
    error_msg = server.GetAssignedValues(inChem, passSeries(0), passAssigned(0))
    
    Call InterpretServerMsg(error_msg)
       
    ' load the assigned values into the series set object
    maxSeriesIndex = UBound(passAssigned(0))
    For seriesIndex = 0 To maxSeriesIndex
        Set curSeries = inSeriesSet(passSeries(0)(seriesIndex + 1))
        
        If ((curSeries Is Nothing) = True) Then
            Call Err.Raise(COMPERR, "QalibObjects", BuildSubstString(LoadResString(COLLECTIONERR), passSeries(0)(seriesIndex + 1)))
        End If
        
        Call curSeries.Load(inAssignedVal:=CDbl(passAssigned(0)(seriesIndex)))
    Next seriesIndex
    
    Set curSeries = Nothing
    Set server = Nothing
  
    Exit Sub
ErrTrap:
    Set curSeries = Nothing
    Set server = Nothing

    Call Err.Raise(Err.Number, Err.Source & " | Utility.GetAssigned", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   UnpackRuns()

'DESCRIPTION: Unpacks the run arrays and loads them into objects

'PARAMETERS:    inAnalyzerData - the analyzer data object to load
'               inExpPoint - the run data with point set and series names
'               inStatus - the run point(s) statuses
'               inRunID - the run IDs
'               inSerial - the run serial numbers
'               inBarCode - the run bar codes
'               inSpecies - the run species
'               inSxn - the run system suppressions
'               inSW - the run software status

'RETURNED:    N/A

'********************************************************************
Public Sub UnpackRuns(inAnalyzerData As AnalyzerData, inExpPoint As Variant, _
    inStatus As Variant, Optional inRunID As Variant, Optional inSerial As Variant, _
    Optional inBarCode As Variant, Optional inSpecies As Variant, Optional inSxn As Variant, _
    Optional inSW As Variant)
On Error GoTo ErrTrap
    Dim runIndex As Long
    Dim maxRunIndex As Long
    Dim expPointIndex As Long
    Dim numGood As Long
    Dim curSeries As SeriesOne
    Dim curRun As Run
    Dim curExpPoint As ExpPoint
    Dim curExpPointSet As ExpPointSet
    Dim expPointSetIndex As Long
    Dim maxExpPointSetIndex As Long
    Dim curExpPointValue As String
    Dim expPointSetName() As String
    Dim seriesIndex As Long
    Dim maxSeriesIndex As Long
    Dim seriesName() As String
    Dim seriesLength() As Long
    Dim seriesRunIndex As Long
    Dim maxSeriesRunIndex As Long
    Dim curRunID As Long
    Dim curSerial As String
    Dim curBarCode As String
    Dim curSpecies As String
    Dim curSxn As String
    Dim curSW As Long
    Dim curStatus As Long
    Dim loadedPoints As Long
                
    ' the data is communicated with the server in a specific format
    
    ' inExpPoint is the most complicated; its format is as follows:
    ' 1.  the number of point sets
    ' 2.  the name of each point set
    ' 3.  the number of series
    ' 4.  the number of runs and the name of each series
    ' 5.  the points in each point set (all of set one, all of set two, etc.)
    
    ' inStatus has the status for each point.  It is a parallel array to inExpPoint
    ' except it does not contain the meta information from 1-4 above.
    
    ' inSerial, inBarCode, and inSpecies are parallel to inExpPoint too, except they don't
    ' have the meta information from 1-4, and since the serial number, bar code, and species are the same
    ' for each point in the run, they are not repeated for multiple point sets
        
    expPointIndex = 0
    maxRunIndex = 0
    
    ' see how many points per run (point sets) there are
    maxExpPointSetIndex = CLng(inExpPoint(expPointIndex)) - 1
    ReDim expPointSetName(maxExpPointSetIndex) As String
    
    ' make sure the array has at least the point set names
    If ((maxExpPointSetIndex + 1) > UBound(inExpPoint)) Then
        Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
    End If
    
    ' get the names of the point sets
    For expPointSetIndex = 0 To maxExpPointSetIndex
        expPointIndex = expPointIndex + 1
        expPointSetName(expPointSetIndex) = inExpPoint(expPointIndex)
        
        ' set up the point set
        Set curExpPointSet = New ExpPointSet
        Call curExpPointSet.Load(expPointSetName(expPointSetIndex))
        Call inAnalyzerData.ExpPointSets.Add(curExpPointSet, curExpPointSet.Name)
    Next expPointSetIndex
        
    ' see how many series there are
    expPointIndex = expPointIndex + 1
    maxSeriesIndex = CLng(inExpPoint(expPointIndex)) - 1
    
    ReDim seriesName(maxSeriesIndex) As String
    ReDim seriesLength(maxSeriesIndex) As Long
    
    ' make sure the array has at least the series lengths and names
    If ((expPointIndex + (((maxSeriesIndex + 1) * 2))) > UBound(inExpPoint)) Then
        Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
    End If
    
    ' get the size and name of each series
    For seriesIndex = 0 To maxSeriesIndex
        expPointIndex = expPointIndex + 2
        seriesLength(seriesIndex) = CLng(inExpPoint(expPointIndex - 1))
        seriesName(seriesIndex) = inExpPoint(expPointIndex)
        
        ' keep a tally of how many runs there are
        maxRunIndex = maxRunIndex + seriesLength(seriesIndex)
    Next seriesIndex
            
    maxRunIndex = maxRunIndex - 1 ' since the runs are 0-based
                    
    ' check for congruent arrays
    If (IsMissing(inRunID) = False) Then
        If (maxRunIndex <> UBound(inRunID)) Then
            Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
        End If
    End If
    
    If (IsMissing(inBarCode) = False) Then
        If (maxRunIndex <> UBound(inBarCode)) Then
            Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
        End If
    End If
    
    If (IsMissing(inSerial) = False) Then
        If (maxRunIndex <> UBound(inSerial)) Then
            Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
        End If
    End If

    If (IsMissing(inSpecies) = False) Then
        If (maxRunIndex <> UBound(inSpecies)) Then
            Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
        End If
    End If

    If (IsMissing(inSxn) = False) Then
        If (maxRunIndex <> UBound(inSxn)) Then
            Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
        End If
    End If

    If (IsMissing(inSW) = False) Then
        If (maxRunIndex <> UBound(inSW)) Then
            Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
        End If
    End If
    
    ' there is a status for each point in the point sets
    If ((((maxRunIndex * inAnalyzerData.ExpPointSets.Count) + inAnalyzerData.ExpPointSets.Count) - 1) <> UBound(inStatus)) Then
        Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
    End If
        
    numGood = 0
    runIndex = 0
    expPointIndex = expPointIndex + 1
    
    ' the point array should be large enough for all the points in each point set
    If ((expPointIndex + ((maxRunIndex * inAnalyzerData.ExpPointSets.Count) + inAnalyzerData.ExpPointSets.Count) - 1) <> UBound(inExpPoint)) Then
        Call Err.Raise(COMPERR, "QalibObjects", LoadResString(INCONGRUENTARRAYERR))
    End If
    
    ' load all the series
    For seriesIndex = 0 To maxSeriesIndex
                        
        Set curSeries = New SeriesOne
        Call curSeries.Load(inName:=seriesName(seriesIndex))
        Call inAnalyzerData.SeriesSet.Add(curSeries, curSeries.Name)
        
        maxSeriesRunIndex = seriesLength(seriesIndex) - 1
        
        ' load the runs for a series
        For seriesRunIndex = 0 To maxSeriesRunIndex
            
            ' system suppression array is optional so only load it if it was passed in
            If (IsMissing(inSxn) = True) Then
                curSxn = ""
            Else
                curSxn = CStr(inSxn(runIndex))
            End If
            
            ' software array is optional so only load it if it was passed in
            If (IsMissing(inSW) = True) Then
                curSW = RELEASEDSW
            Else
                curSW = CLng(inSW(runIndex))
            End If
            
            ' only consider a run if it is not suppressed and run on released software
            If ((curSxn = "") And (curSW = RELEASEDSW)) Then
                
                numGood = numGood + 1
                
                Set curRun = New Run
                
                ' run ID array is optional so only load it if it was passed in
                If (IsMissing(inRunID) = True) Then
                    curRunID = runIndex
                Else
                    curRunID = CLng(inRunID(runIndex))
                End If
                
                ' serial number array is optional so only load it if it was passed in
                If (IsMissing(inSerial) = True) Then
                    curSerial = ""
                Else
                    curSerial = CStr(inSerial(runIndex))
                End If
                
                ' bar code array is optional so only load it if it was passed in
                If (IsMissing(inBarCode) = True) Then
                    curBarCode = ""
                Else
                    curBarCode = CStr(inBarCode(runIndex))
                End If
                
                ' species array is optional so only load it if it was passed in
                If (IsMissing(inSpecies) = True) Then
                    curSpecies = ""
                Else
                    curSpecies = CStr(inSpecies(runIndex))
                End If
                
                ' set up the Run
                Call curRun.Load(curRunID, curSerial, curBarCode, curSpecies, seriesName(seriesIndex))
                
                ' add the run to the series
                Call curSeries.Runs.Add(curRun, curRun.ID)
                
                ' reset the loaded points count
                loadedPoints = 0
                
                ' load the points for a run
                For expPointSetIndex = 0 To maxExpPointSetIndex
                    
                    ' most runs will have just one point, so the ((expPointSetIndex * maxRunIndex) + expPointSetIndex) term will be zero
                    ' for the most part.  However, for runs with more than one point this term will take care of the offset
                    ' when loading the other points
                    curExpPointValue = inExpPoint(expPointIndex + (expPointSetIndex * maxRunIndex) + expPointSetIndex)
                    
                    ' only consider a point if it is not suppressed
                    If (Left$(curExpPointValue, MEASUREDVALERRLEN) <> MEASUREDVALERRTEXT) Then
                        
                        ' load the status
                        curStatus = inStatus(runIndex + ((expPointSetIndex * maxRunIndex) + expPointSetIndex))
                        
                        ' load the point
                        Set curExpPoint = New ExpPoint
                        Call curExpPoint.Load(curRunID, CDbl(curExpPointValue), curStatus)
                        
                        ' get the point set
                        Set curExpPointSet = inAnalyzerData.ExpPointSets(expPointSetName(expPointSetIndex))
                        
                        ' add the point to the set
                        Call curExpPointSet.ExpPoints.Add(curExpPoint, curExpPoint.ID)
                        
                        ' the point was valid
                        loadedPoints = loadedPoints + 1
                    End If
                    
                Next expPointSetIndex
                
                ' if one of the points in a run was invalid, then remove the run
                If (loadedPoints <> inAnalyzerData.ExpPointSets.Count) Then
                    Call curSeries.Runs.Remove(curRun.ID)
                End If
                
            End If
            
            expPointIndex = expPointIndex + 1
            runIndex = runIndex + 1
        Next seriesRunIndex
            
    Next seriesIndex
    
    ' make sure all runs have not been eliminated
    If (numGood = 0) Then
        Call Err.Raise(COMPERR, "QalibObjects", LoadResString(ALLRUNSELIMINATEDERR))
    End If
                
    Set curSeries = Nothing
    Set curRun = Nothing
    Set curExpPoint = Nothing
    Set curExpPointSet = Nothing
   
    Exit Sub
ErrTrap:
    Set curSeries = Nothing
    Set curRun = Nothing
    Set curExpPoint = Nothing
    Set curExpPointSet = Nothing

    Call Err.Raise(Err.Number, Err.Source & " | Utility.UnpackRuns", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   InterpretServerMsg()

'DESCRIPTION: Interprets the response from the Qalib server and raises an error if necessary

'PARAMETERS:  inMsg - the return value from the server

'RETURNED:    N/A

'********************************************************************
Public Sub InterpretServerMsg(inMsg As Variant)
On Error GoTo ErrTrap
    ' make sure the server responded
    If (IsEmpty(inMsg) = True) Then
        Call Err.Raise(COMPERR, "QalibObjects", LoadResString(NOSERVERRESPONSEERR))
    End If
    
    Select Case (CLng(inMsg(0)))
        ' no error
        Case 0
'        Case 1
'        Case 2
'        Case 3
        Case Else
            Call Err.Raise(COMPERR, "QalibObjects", BuildSubstString(LoadResString(UNEXPECTEDSERVERERR), CStr(inMsg(1))))
    End Select
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Utility.InterpretServerMsg", Err.Description)
End Sub
