VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AnalyzerData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  AnalyzerData.cls
 
'DESCRIPTION:  This module holds the analyzer data class.  The SeriesSet holds all the series
' (indexed by series name) and the runs (indexed by run ID) for each series.  The ExpPoint set holds
' all the point sets (indexed by point set name) and the points (indexed by run ID) for each point set.
' The run ID is the link between the runs in SeriesSet and the points in ExpPointSet.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: AnalyzerData.cls $
 ' 
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:42p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe.

Option Explicit

' private member variables

Private objSeriesSet_m As QalibCollection ' the series collection
Private objExpPointSets_m As QalibCollection ' the experimental point sets collection

'***********************************************************************

'PROPERTY GET:   SeriesSet()

'DESCRIPTION: Allows other objects to get the series set

'PARAMETERS:  N/A

'RETURNED:    the series set

'*********************************************************************
Public Property Get SeriesSet() As QalibCollection
On Error GoTo ErrTrap
    Set SeriesSet = objSeriesSet_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | AnalyzerData.PropertyGet.SeriesSet", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   ExpPointSets()

'DESCRIPTION: Allows other objects to get the experimental point sets

'PARAMETERS:  N/A

'RETURNED:    the experimental point sets

'*********************************************************************
Public Property Get ExpPointSets() As QalibCollection
On Error GoTo ErrTrap
    Set ExpPointSets = objExpPointSets_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | AnalyzerData.PropertyGet.ExpPointSets", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Class_Initialize()

'DESCRIPTION: Sets up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Initialize()
    Set objSeriesSet_m = New QalibCollection
    Set objExpPointSets_m = New QalibCollection
End Sub

'***********************************************************************

'PROCEDURE:   Class_Terminate()

'DESCRIPTION: Cleans up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Terminate()
    Set objSeriesSet_m = Nothing
    Set objExpPointSets_m = Nothing
End Sub

'***********************************************************************

'PROCEDURE:   ChangeRunSeries()

'DESCRIPTION: Moves a run from series to series.

'PARAMETERS:    inRunID - the run ID
'               inNewSeries - the new series

'RETURNED:    N/A

'*********************************************************************
Public Sub ChangeRunSeries(inRunID As Long, inNewSeries As String)
On Error GoTo ErrTrap
    Dim objNewSeries As SeriesOne
    Dim objOldSeries As SeriesOne
    Dim curRun As Run
    Dim curExpPointSet As ExpPointSet
    Dim curExpPoint As ExpPoint
    Dim newStatus As Long
    
    ' find the old series for the run
    For Each objOldSeries In objSeriesSet_m
        Set curRun = objOldSeries.Runs(inRunID)
        
        If ((curRun Is Nothing) = False) Then
            Exit For
        End If
        
    Next objOldSeries
    
    ' make sure the run was found
    If ((curRun Is Nothing) = True) Then
        Call Err.Raise(COMPERR, "QalibObjects", BuildSubstString(LoadResString(RUNNOTFOUNDERR), curRun.ID))
    End If
    
    ' remove the run from the old series
    Call objOldSeries.Runs.Remove(curRun.ID)
    
    ' add the run to the new series
    Set objNewSeries = objSeriesSet_m(inNewSeries)
    Call objNewSeries.Runs.Add(curRun, curRun.ID)
        
    ' update the status for all the points in each point set corresponding to the run
    For Each curExpPointSet In objExpPointSets_m
        Set curExpPoint = curExpPointSet.ExpPoints(inRunID)
        
        ' set the new status accordingly
        If (curRun.OrigSeries <> objNewSeries.Name) Then
            newStatus = curExpPoint.Status Or SERIESCHANGESTATUS
        Else
            newStatus = curExpPoint.Status And Not SERIESCHANGESTATUS
        End If
        
        ' load the new status
        Call curExpPoint.Load(inStatus:=newStatus)
        
    Next curExpPointSet
    
    Set objNewSeries = Nothing
    Set objOldSeries = Nothing
    Set curRun = Nothing
    Set curExpPointSet = Nothing
    Set curExpPoint = Nothing
    Exit Sub
ErrTrap:
    Set objNewSeries = Nothing
    Set objOldSeries = Nothing
    Set curRun = Nothing
    Set curExpPointSet = Nothing
    Set curExpPoint = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | AnalyzerData.ChangeRunSeries", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   ExcludeRun()

'DESCRIPTION: Excludes a run

'PARAMETERS:  inRunID - the run ID
'             inExclude - whether to exclude or not

'RETURNED:    N/A

'*********************************************************************
Public Sub ExcludeRun(inRunID As Long, inExclude As Boolean)
On Error GoTo ErrTrap
    Dim curExpPointSet As ExpPointSet
    Dim curExpPoint As ExpPoint
    Dim newStatus As Long
    
    ' change the status of each point in all the point sets
    For Each curExpPointSet In objExpPointSets_m
        Set curExpPoint = curExpPointSet.ExpPoints(inRunID)
        
        ' set up the new status
        If (inExclude = True) Then
            newStatus = curExpPoint.Status Or EXCLUDESTATUS
        Else
            newStatus = curExpPoint.Status And Not EXCLUDESTATUS
        End If
        
        ' load the new status for the point
        Call curExpPoint.Load(inStatus:=newStatus)
        
    Next curExpPointSet
    
    Set curExpPointSet = Nothing
    Set curExpPoint = Nothing

    Exit Sub
ErrTrap:
    Set curExpPointSet = Nothing
    Set curExpPoint = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | AnalyzerData.ExcludeRun", Err.Description)
End Sub




