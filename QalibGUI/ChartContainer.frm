VERSION 5.00
Object = "{B6C10482-FB89-11D4-93C9-006008A7EED4}#1.0#0"; "TeeChart5.ocx"
Begin VB.Form ChartContainer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Verify Data"
   ClientHeight    =   9015
   ClientLeft      =   1815
   ClientTop       =   1695
   ClientWidth     =   10530
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9015
   ScaleWidth      =   10530
   Begin TeeChart.TChart chtData 
      Height          =   8805
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Click a point to edit"
      Top             =   120
      Width           =   8805
      Base64          =   $"ChartContainer.frx":0000
   End
   Begin VB.Frame fraAxes 
      Caption         =   "Axes"
      Height          =   1215
      Left            =   9000
      TabIndex        =   2
      ToolTipText     =   "Change axis orientation"
      Top             =   1200
      Width           =   1455
      Begin VB.OptionButton optAssignedY 
         Caption         =   "Assigned on Y"
         Height          =   375
         Left            =   240
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Change axis orientation"
         Top             =   720
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optAssignedX 
         Caption         =   "Assigned on X"
         Height          =   375
         Left            =   240
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Change axis orientation"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      ToolTipText     =   "Cancel calibration"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "&Next >"
      Default         =   -1  'True
      Height          =   375
      Left            =   9240
      TabIndex        =   0
      ToolTipText     =   "Accept data and continue"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "ChartContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
'
'FILE:  ChartContainer.frm
'
'DESCRIPTION:  This module contains the form where the pool data and calibration curves
' are charted.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: ChartContainer.frm $
 ' 
 ' *****************  Version 12  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:36a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 '
 ' *****************  Version 11  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:39p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Updated to 1.0.0X9
 '
 ' *****************  Version 10  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:37p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Chemistry identification is now name driven and not part number driven
 ' on client.
 '
 ' *****************  Version 9  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:30p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added error traps.
 ' Set object references equal to nothing to free resources.
 ' Added formatting constants for chart.
 ' Underlying chart data now stored in object.
 ' Moved plot calibrator functionality to a shared module since this is
 ' not the only form to use it.
 '
 ' *****************  Version 8  *****************
 ' User: Ballard      Date: 5/09/03    Time: 3:48p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Lot object is now loaded through a PlotCalibrators function instead of
 ' being set by a property.  Chart picture is now loaded directly into lot
 ' object instead of being pasted from clipboard later.
 '
 ' *****************  Version 7  *****************
 ' User: Ballard      Date: 3/21/03    Time: 2:09p
 ' Updated in $/QalibVBClient
 ' Renamed module "frmDisplayData.frm" to "ChartContainer.frm."
 '
 ' Moved the functionality to populate, manipulate, and show the chart
 ' back to the form.
 '
 ' The form retrieves data from the underlying data object.
 '
 ' *****************  Version 6  *****************
 ' User: Ballard      Date: 1/24/03    Time: 2:53p
 ' Updated in $/QalibVBClient
 ' Query_Unload event handler added to catch if user cancels the form in
 ' any other way than clicking "Cancel."
 '
 ' *****************  Version 5  *****************
 ' User: Ballard      Date: 1/17/03    Time: 4:24p
 ' Updated in $/QalibVBClient
 ' The standard TChart control was replaced by a custom control that
 ' basically wraps the functionality of the TChart.
 ' As a result, the SwapAxes functionality was moved into the custom
 ' control.
 '
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 1/09/03    Time: 3:53p
 ' Updated in $/QalibVBClient
 ' Added file and function headings.
 ' Added ability to reverse axes.
 ' Removed calibration calculations since they're now shown on final
 ' report.
 ' Removed all chart loading, formatting, and event handling since they're
 ' now done in the calibration object.
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 12/06/02   Time: 4:59p
 ' Updated in $/QalibVBClient
 ' Change "Next" button caption.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 10/04/02   Time: 4:38p
 ' Updated in $/QalibVBClient
 ' Changed subscripts in QalibResults array to reflect addition of
 ' parameter errors.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 9/24/02    Time: 1:46p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe

Option Explicit


' Private member variables
Private blnCancel_m As Boolean ' stores whether cancel button was pressed
Private objLotVerify_m As LotVerify  ' stores a reference to the lot verification object
Private intMeasuredAxis_m As enum_MeasuredAxis ' stores which axis the measured values are on

'***********************************************************************
'
'PROPERTY GET:   Cancel
'
'DESCRIPTION: Allows other objects to see if cancel button was pressed
'
'PARAMETERS:  N/A
'
'RETURNED:  Whether the cancel button was pressed
'
'*********************************************************************
Public Property Get Cancel() As Boolean
On Error GoTo ErrTrap
    Cancel = blnCancel_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ChartContainer.PropertyGet.Cancel", Err.Description)
End Property

'***********************************************************************
'
'PROPERTY GET:   MeasuredAxis
'
'DESCRIPTION: Allows other objects to get the measured axis
'
'PARAMETERS:  N/A
'
'RETURNED:  measured axis
'
'*********************************************************************
Public Property Get measuredAxis() As enum_MeasuredAxis
On Error GoTo ErrTrap
    measuredAxis = intMeasuredAxis_m
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ChartContainer.PropertyGet.MeasuredAxis", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Form_Terminate()

'DESCRIPTION: Event handler for when the form terminates

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Form_Terminate()
On Error GoTo ErrTrap
    Set objLotVerify_m = Nothing
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ChartContainer.Form_Terminate", Err.Description)
End Sub

'***********************************************************************
'
' PROCEDURE:   Form_QueryUnload
'
' DESCRIPTION:  Executes right before form is unloaded so method of closing
' can be determined
'
' PARAMETERS:  Cancel - whether to cancel the unload
'              UnloadMode - how the form was unloaded
'
' RETURN:   N/A
'
'*********************************************************************
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo ErrTrap
    ' Only cancel the form unloading if the user closed the form by "Alt-F4",
    ' the control menu, or the "X" in the upper right hand corner.
    ' This way the unloading is controlled and the cancel property can be set.
    If (UnloadMode = vbFormControlMenu) Then
        Cancel = True
        Call cmdCancel_Click
    End If
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ChartContainer.Form_QueryUnload", Err.Description)
End Sub

'***********************************************************************
'
'PROCEDURE:   cmdCancel_Click
'
'DESCRIPTION: Hides the form and sets cancel property
'
'PARAMETERS:  N/A
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub cmdCancel_Click()
On Error GoTo ErrTrap
    blnCancel_m = True
    Me.Hide
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ChartContainer.cmdCancel_Click", Err.Description)
End Sub

'***********************************************************************
'
'PROCEDURE:   cmdNext_Click
'
'DESCRIPTION: Hides the form and sets cancel property
'
'PARAMETERS:  N/A
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub cmdNext_Click()
On Error GoTo ErrTrap
    ' perform the calibration

    blnCancel_m = False
    Me.Hide
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ChartContainer.cmdNext_Click", Err.Description)
End Sub


'***********************************************************************
'
'PROCEDURE:   optAssignedX_Click
'
'DESCRIPTION: Event handler for when user chooses assigned values on the X axis
'
'PARAMETERS:  N/A
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub optAssignedX_Click()
On Error GoTo ErrTrap
    Call SwapAxes
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ChartContainer.optAssignedX_Click", Err.Description)
End Sub

'***********************************************************************
'
'PROCEDURE:   optAssignedY_Click()
'
'DESCRIPTION: Event handler for when user chooses assigned values on the Y axis
'
'PARAMETERS:  N/A
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub optAssignedY_Click()
On Error GoTo ErrTrap
    Call SwapAxes
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ChartContainer.optAssignedY_Click", Err.Description)
End Sub

'***********************************************************************
'
'PROCEDURE:   chtData_OnClickLegend
'
'DESCRIPTION: Event handler for when user clicks on chart legend
'
'PARAMETERS:  Button - which mouse button
'             Shift - shift state
'             X - X axis mouse cursor position
'             Y - Y axis mouse cursor position
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub chtData_OnClickLegend(ByVal Button As TeeChart.EMouseButton, ByVal Shift As TeeChart.EShiftState, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrTrap
    Dim series As Long
    
    series = chtData.Legend.Clicked(X, Y) ' see what series was clicked
    
    ' if a series was clicked then process the click
    ' ignore the point set labels (their series are line style)
    If ((series > -1) And (chtData.series(series).SeriesType = scPoint)) Then
        Call ProcessChartClick(series, 0)
    End If
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ChartContainer.chtData_OnClickLegend", Err.Description)
End Sub

'***********************************************************************
'
'PROCEDURE:   chtData_OnClickSeries
'
'DESCRIPTION: Event handler for when user clicks a series
'
'PARAMETERS:  seriesIndex - which series
'             valueIndex - which point in series
'             Button - which mouse button
'             Shift - shift state
'             X - X axis mouse cursor position
'             Y - Y axis mouse cursor position
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub chtData_OnClickSeries(ByVal seriesIndex As Long, ByVal ValueIndex As Long, ByVal Button As TeeChart.EMouseButton, ByVal Shift As TeeChart.EShiftState, ByVal X As Long, ByVal Y As Long)
On Error GoTo ErrTrap
    Call ProcessChartClick(seriesIndex, ValueIndex)
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ChartContainer.chtData_OnClickSeries", Err.Description)
End Sub

'***********************************************************************
'
'PROCEDURE:   chtData_OnGetSeriesMark
'
'DESCRIPTION: Event handler for when chart formats the mark of each point
'
'PARAMETERS:  seriesIndex - which series
'             valueIndex - which point in series
'             marktext - the text of the point
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub chtData_OnGetSeriesMark(ByVal seriesIndex As Long, ByVal ValueIndex As Long, MarkText As String)
    ' the mark text of the point is contained in the point label
    MarkText = chtData.series(seriesIndex).pointLabel(ValueIndex)
End Sub

'***********************************************************************
'
'PROCEDURE:   ProcessChartClick
'
'DESCRIPTION: Displays the series edit form and updates the chart
'
'PARAMETERS:  seriesIndex - which series
'             pointIndex - which point in series
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub ProcessChartClick(seriesIndex As Long, pointIndex As Long)
On Error GoTo ErrTrap
    Dim frmSeriesEdit As SeriesEdit
    Dim curSeries As Long
    Dim pointSetLabel As String
    
    pointSetLabel = ""
    
    ' the first point set label "above" the series clicked will be the appropriate
    ' point set
    For curSeries = seriesIndex To 0 Step -1
        
        ' point set label series are line style
        If (chtData.series(curSeries).SeriesType = scLine) Then
            pointSetLabel = chtData.series(curSeries).Title
            Exit For
        End If
    Next curSeries
    
    ' make sure point set label was found
    If (pointSetLabel = "") Then
        Call Err.Raise(APPERR, "QalibGUI", LoadResString(MISSINGPOINTSETLABELERR))
    End If
                
    ' initialize and show the change series form
    Set frmSeriesEdit = New SeriesEdit
    Call frmSeriesEdit.LoadSeries(objLotVerify_m.AnalyzerDataSet, chtData.series(seriesIndex).Title, _
        pointSetLabel, pointIndex)
    Call frmSeriesEdit.Show(vbModal, Me)

'     see if user pressed cancel
    If (frmSeriesEdit.Cancel = False) Then
        Call PlotSamples(objLotVerify_m.AnalyzerDataSet, chtData, intMeasuredAxis_m, objLotVerify_m.Chemistry)     ' reload chart
    End If

    Call Unload(frmSeriesEdit)
    Set frmSeriesEdit = Nothing
    Exit Sub
ErrTrap:
    Set frmSeriesEdit = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | ChartContainer.ProcessChartClick", Err.Description)
End Sub

'***********************************************************************
'
'PROCEDURE:   SwapAxes
'
'DESCRIPTION: Switches the axes on the chart so that the X data is charted on the Y
' axis and vice versa.
'
'PARAMETERS:  N/A
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub SwapAxes()
On Error GoTo ErrTrap
    Dim tempValue As Double
    Dim seriesIndex As Integer
    Dim ValueIndex As Integer
    Dim maxSeriesIndex As Integer
    Dim maxValueIndex As Integer
    Dim tempTitle As String
    Dim tempOrder As EValueListOrder
    Dim leftIncrement As Double
    Dim bottomIncrement As Double
    
    ' save off old increments for swapping
    leftIncrement = chtData.Axis.Left.Increment
    bottomIncrement = chtData.Axis.Bottom.Increment
    
    ' get the maximum series index
    maxSeriesIndex = chtData.SeriesCount - 1
       
    ' cycle through all the chart's series
    For seriesIndex = 0 To maxSeriesIndex
        
        With chtData.series(seriesIndex)
            ' first swap the series ordering
            tempOrder = .XValues.Order
            .XValues.Order = .YValues.Order
            .YValues.Order = tempOrder
            
            ' get the maximum value index
            maxValueIndex = .XValues.Count - 1
            
            ' cycle through all the series' values
            For ValueIndex = 0 To maxValueIndex
                
                ' do the swap
                tempValue = .XValues.Value(ValueIndex)
                .XValues.Value(ValueIndex) = .YValues.Value(ValueIndex)
                .YValues.Value(ValueIndex) = tempValue
                
            Next ValueIndex
        End With
        
    Next seriesIndex
    
    ' swap the axis titles
    With chtData.Axis
        tempTitle = .Left.Title.Caption
        .Left.Title.Caption = .Bottom.Title.Caption
        .Bottom.Title.Caption = tempTitle
        
        ' this must be forced, otherwise the increment will be too coarse
        .Bottom.Increment = leftIncrement
        .Left.Increment = bottomIncrement
    End With
    
    ' update where the measured values are
    intMeasuredAxis_m = Abs(intMeasuredAxis_m - 1)
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ChartContainer.SwapAxes", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   LoadChart()

'DESCRIPTION: Loads the chart with the calibrators

'PARAMETERS:  inLotVerify - the lot verification object

'RETURNED:    N/A

'*********************************************************************
Public Sub LoadChart(inLotVerify As LotVerify)
On Error GoTo ErrTrap
    Set objLotVerify_m = inLotVerify
    
    ' default axis
    intMeasuredAxis_m = abxMeasuredX
    
    Call PlotSamples(objLotVerify_m.AnalyzerDataSet, chtData, intMeasuredAxis_m, objLotVerify_m.Chemistry)
    
    Exit Sub
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | ChartContainer.LoadChart", Err.Description)
End Sub





