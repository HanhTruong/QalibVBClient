Attribute VB_Name = "GUIMain"
'**********************************************************************************
 
'FILE:  GUIMain.bas
 
'DESCRIPTION:  This module is the startup object for Qalib.  It holds the "Main" procedure

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: GUIMain.bas $
 '
 ' *****************  Version 14  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:36a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Removed login authentication since login form now handles that.
 ' Fixed so that plot axes are correctly labeled when user swaps them.
 '
 ' *****************  Version 13  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:39p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Updated to 1.0.0X9
 '
 ' *****************  Version 12  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:38p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Logging is now done to a local file.
 '
 ' *****************  Version 11  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:42p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added error traps.
 ' Set object references equal to nothing to free resources.
 ' Added functionality to plot calibrators since it is commonly used in
 ' the project.
 '
 ' *****************  Version 10  *****************
 ' User: Ballard      Date: 5/09/03    Time: 3:43p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Changed name from QalibMain to GUIMain.
 '
 ' *****************  Version 9  *****************
 ' User: Ballard      Date: 3/21/03    Time: 1:21p
 ' Updated in $/QalibVBClient
 ' Rename module from "modMain.bas" to "QalibMain.bas"
 '
 ' This module now shows the login, mode selection, and splash screens.
 ' It then displays the main menu modelessly and starts a calibration
 ' sequence.
 '
 ' *****************  Version 8  *****************
 ' User: Ballard      Date: 1/24/03    Time: 2:49p
 ' Updated in $/QalibVBClient
 '
 ' *****************  Version 7  *****************
 ' User: Ballard      Date: 1/17/03    Time: 4:14p
 ' Updated in $/QalibVBClient
 '
 ' *****************  Version 6  *****************
 ' User: Ballard      Date: 1/09/03    Time: 4:05p
 ' Updated in $/QalibVBClient
 ' Added file and function headers.
 ' Removed calibration functionality since it's now in the calibration
 ' object.
'
 ' *****************  Version 5  *****************
 ' User: Alves        Date: 12/06/02   Time: 10:39a
 ' Updated in $/QalibVBClient
 ' Return goodness of fit to client.
 ' Calculate rate ofr assigned values.
 ' Correlation coefficient for CALIBRATOR and CONTROLs
 '
 ' *****************  Version 4  *****************
 ' User: Alves        Date: 11/25/02   Time: 6:40p
 ' Updated in $/QalibVBClient
 ' Added get calculated assigned values from database to client
 ' functionality
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 10/04/02   Time: 4:38p
 ' Updated in $/QalibVBClient
 ' Changed subscripts in QalibResults array to reflect addition of
 ' parameter errors.
 '
 ' *****************  Version 2  *****************
 ' User: Alves        Date: 10/04/02   Time: 11:58a
 ' Updated in $/QalibVBClient
 ' Changed QResults array index from 5 to 9. To accomodate updated returns
 ' from the database.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 9/24/02    Time: 1:46p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe

Option Explicit

' private API declarations
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' constants

' application errors
Public Const APPERR As Long = vbObjectError + &H1000

' QalibObject component errors
Private Const COMPERR As Long = vbObjectError + &H1100

' log file
Private Const LOGFILE As String = "qalib.log"

' calibration image files
Public Const CALPLOTFILE As String = "CalPlot.jpg"
Public Const CALCHECKPLOTFILE As String = "CheckPlot.jpg"
Public Const CALIBRATORPREFIX As String = "Calibrator"
Public Const CONTROLPREFIX As String = "Control"


' Formatting constants for the chart
Private Const AXISLABELSIZE As Integer = 10
Private Const AXISTITLESIZE As Integer = 14
Private Const HEADERTITLESIZE As Integer = 16
Private Const LEGENDSIZE As Integer = 10
Private Const POINTSIZE As Integer = 3
Private Const MEASUREDAXISTITLE As String = "Measured"
Private Const AVAXISTITLE As String = "Assigned Value"
Private Const HEADERTITLE As String = "Data"
Private Const CALCURVETITLE As String = "Best Fit"

' Enumeration for which axis the measured values are plotted on
Public Enum enum_MeasuredAxis
    abxMeasuredX = 0
    abxMeasuredY = 1
End Enum

'resource strings
Public Const LOGFILEERR As Long = 101
Public Const REPORTFILEERR As Long = 102
Public Const COLLECTIONERR As Long = 103
Public Const MISSINGPOINTSETLABELERR As Long = 104

'resource icons
Public Const CHECKOFFICON As Long = 101
Public Const CHECKONICON As Long = 102

'registry entries
Public Const SECTIONKEY As String = "Settings"
Public Const LASTLOGINKEY As String = "LastLogin"
Public Const LASTPATHKEY As String = "LastPath"
Public Const LASTMODEKEY As String = "LastMode"


'***********************************************************************

'PROCEDURE:   Main()

'DESCRIPTION: This is the main subroutine for Qalib

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Main()
On Error GoTo ErrTrap
    ' start the log
    Call InitializeLog
    
    Dim frmSplash As Splash ' for storing the splash form
    Dim frmLogin As Login   ' for storing the login form
    Dim frmMode As ModeSelect ' for storing the mode selection form
    Dim frmMainMenu As MainMenu ' for storing the main menu form
    Dim objUserVerify As UserVerifiy  ' for storing the user verification
    
    ' get the user login
    Set frmLogin = New Login
    Set objUserVerify = New UserVerifiy
    Call frmLogin.LoadUserVerify(objUserVerify)
    Call frmLogin.Show(vbModal)
    If (frmLogin.Cancel = True) Then
        Set frmLogin = Nothing
        End
    End If
        
    ' get the mode
    Set frmMode = New ModeSelect
    Call frmMode.LoadModes(objUserVerify.Modes)
    Call frmMode.Show(vbModal)
    If (frmMode.Cancel = True) Then
        Set frmLogin = Nothing
        Set objUserVerify = Nothing
        Set frmMode = Nothing
        End
    End If
    DoEvents ' force the mode form to unload
    
    ' prepare and show the main form
    Set frmSplash = New Splash
    Call frmSplash.Show
    Call frmSplash.Refresh
    Set frmMainMenu = New MainMenu
    Call Load(frmMainMenu)
    Call Sleep(1000)
    Unload frmSplash
    Call frmMainMenu.Show
    
    ' launch the user into a calibration sequence
    Call frmMainMenu.InitializeSession(frmLogin.User, frmLogin.Password, frmMode.SelectedMode)
    
    Call Unload(frmLogin)
    Call Unload(frmMode)
    Set frmSplash = Nothing
    Set frmLogin = Nothing
    Set objUserVerify = Nothing
    Set frmMode = Nothing
    Set frmMainMenu = Nothing
    Exit Sub
ErrTrap:
    Set frmSplash = Nothing
    Set frmLogin = Nothing
    Set objUserVerify = Nothing
    Set frmMode = Nothing
    Set frmMainMenu = Nothing
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | GUIMain.Main", Err.Description)
    End ' must force end here to unload all forms
End Sub

'***********************************************************************

'PROCEDURE:   InitializeLog()

'DESCRIPTION: Prepares the log file

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub InitializeLog()
On Error GoTo ErrTrap
    ' start the logging
    Call App.StartLogging(App.Path & "\" & LOGFILE, vbLogToFile)
    
    Exit Sub
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | GUIMain.InitializeLog", BuildSubstString(LoadResString(LOGFILEERR), LOGFILE, Err.Number))
End Sub

'***********************************************************************

'PROCEDURE:   HandleError()

'DESCRIPTION: Logs errors and displays message boxes to user if necessary

'PARAMETERS:  inNumber - error number
'             inSource - source code audit trail
'             inDescription - description of error
              

'RETURNED:    N/A

'*********************************************************************
Public Sub HandleError(inNumber As Long, inSource As String, inDescription As String)
On Error GoTo ErrTrap
    Screen.MousePointer = vbDefault
    
    ' see what kind of error it is
    If (inNumber = APPERR) Then
        ' warn the user about the application error
        Call MsgBox("Number:  " & inNumber & vbCrLf & vbCrLf & "Audit Trail:  " & inSource & vbCrLf & vbCrLf & "Description:  " & inDescription, vbExclamation Or vbOKOnly, "Error")
    Else
        ' warn the user about the system error
        Call MsgBox("Number:  " & inNumber & vbCrLf & vbCrLf & "Audit Trail:  " & inSource & vbCrLf & vbCrLf & "Description:  " & inDescription, vbExclamation Or vbOKOnly, "Error")
    End If
    
    Call App.LogEvent(Now() & vbCrLf & "Number:  " & inNumber & vbCrLf & "Audit Trail:  " & inSource & vbCrLf & "Description:  " & inDescription & vbCrLf & vbCrLf)
    
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call MsgBox("Number:  " & Err.Number & vbCrLf & vbCrLf & "Audit Trail:  " & Err.Source & vbCrLf & vbCrLf & "Description:  " & Err.Description, vbExclamation Or vbOKOnly, "Error")
    Resume Next ' this entire procedure must be executed if possible
End Sub

'***********************************************************************
'
'PROCEDURE:   PlotSamples
'
'DESCRIPTION: Plots the samples' measured and assigned values on the chart
'
'PARAMETERS:  inAnalyzerData - the object with the calibrators
'             inChart - the chart to plot on
'             measuredAxis - the axis to plot the measured values
'             inTitle - the title of the chart
'
'RETURNED:    N/A
'
'********************************************************************
Public Sub PlotSamples(inAnalyzerData As AnalyzerData, inChart As TeeChart.TChart, measuredAxis As enum_MeasuredAxis, inTitle As String)
On Error GoTo ErrTrap
    Dim curExpPoint As ExpPoint
    Dim curExpPointSet As ExpPointSet
    Dim curSeries As SeriesOne
    Dim curRun As Run
    Dim pointSetIndex As Long
    Dim seriesIndex As Long
    Dim pointIndex As Long
    Dim maxSeriesIndex As Long
    Dim pointLabel As String
    Dim bottomIncrement As Double
    Dim leftIncrement As Double
    
    pointSetIndex = 0
    ' add the series to the chart control
    With inChart
        .ClearChart
        
        For Each curExpPointSet In inAnalyzerData.ExpPointSets
            
            ' a fake series gets loaded so that the point set labels will be shown in the
            ' legend.  The line for this series needs to be invisible in the legend.
            seriesIndex = .AddSeries(scLine)
            .series(seriesIndex).Title = curExpPointSet.Name
            .series(seriesIndex).asLine.LinePen.visible = False
            
            ' cycle through all the series
            For Each curSeries In inAnalyzerData.SeriesSet
                
                seriesIndex = .AddSeries(scPoint)
                
                ' cycle through all the runs
                For Each curRun In curSeries.Runs
                    
                    Set curExpPoint = curExpPointSet.ExpPoints(curRun.ID)
                    
                    ' load the point label
                    pointLabel = curExpPoint.Label
                    
                    ' the curve is loaded differently depending on what axis orientation the user selected
                    If (measuredAxis = abxMeasuredX) Then
                        pointIndex = .series(seriesIndex).AddXY(curExpPoint.MeasuredVal, _
                            curSeries.AssignedVal, pointLabel, clTeeColor)
                    Else
                        pointIndex = .series(seriesIndex).AddXY(curSeries.AssignedVal, _
                            curExpPoint.MeasuredVal, pointLabel, clTeeColor)
                    End If
                
                Next curRun
                
                ' format the series
                .series(seriesIndex).Title = curSeries.Name
                .series(seriesIndex).Marks.visible = True
                .series(seriesIndex).Marks.ArrowLength = 10
                .series(seriesIndex).asPoint.Pointer.Style = psDiamond
                .series(seriesIndex).asPoint.Pointer.HorizontalSize = POINTSIZE
                .series(seriesIndex).asPoint.Pointer.VerticalSize = POINTSIZE

            Next curSeries
            
        Next curExpPointSet
        
'        the bottom axis should show numbers not the meta information in the labels of the points
        .Axis.Bottom.Labels.Style = talValue

        ' format the chart
        .Aspect.View3D = False
        .Zoom.Enable = False
        .Scroll.Enable = pmNone

        ' format the axes
        If (measuredAxis = abxMeasuredX) Then
            .Axis.Bottom.Title.Caption = MEASUREDAXISTITLE
            .Axis.Left.Title.Caption = AVAXISTITLE
        Else
            .Axis.Bottom.Title.Caption = AVAXISTITLE
            .Axis.Left.Title.Caption = MEASUREDAXISTITLE
        End If
        .Axis.Bottom.Title.Font.Size = AXISTITLESIZE
        .Axis.Bottom.Labels.Font.Size = AXISLABELSIZE
        .Axis.Left.Title.Font.Size = AXISTITLESIZE
        .Axis.Left.Labels.Font.Size = AXISLABELSIZE

        ' format the header
        .Header.Text.Add (inTitle & " " & HEADERTITLE)
        .Header.Font.Size = HEADERTITLESIZE

        'format the legend
        .Legend.Font.Size = LEGENDSIZE
        
        ' this must be forced, otherwise the increment will be too coarse
'        bottomIncrement = .Axis.Bottom.CalcIncrement
'        leftIncrement = .Axis.Left.CalcIncrement
'        .Axis.Bottom.Increment = bottomIncrement
'        .Axis.Left.Increment = leftIncrement
        
    End With
    
    Set curExpPoint = Nothing
    Set curExpPointSet = Nothing
    Set curSeries = Nothing
    Set curRun = Nothing

    Exit Sub
ErrTrap:
    Set curExpPoint = Nothing
    Set curExpPointSet = Nothing
    Set curSeries = Nothing
    Set curRun = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | GUIMain.PlotSamples", Err.Description)
End Sub




















