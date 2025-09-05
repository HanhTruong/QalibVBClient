VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Object = "{B6C10482-FB89-11D4-93C9-006008A7EED4}#1.0#0"; "TeeChart5.ocx"
Begin VB.Form ReportContainer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Final Report"
   ClientHeight    =   11130
   ClientLeft      =   1605
   ClientTop       =   120
   ClientWidth     =   12915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11130
   ScaleWidth      =   12915
   Begin TeeChart.TChart chtCalCheck 
      Height          =   8805
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   8805
      Base64          =   $"ReportContainer.frx":0000
   End
   Begin TeeChart.TChart chtCalibration 
      Height          =   8805
      Left            =   0
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   8805
      Base64          =   $"ReportContainer.frx":0140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   11520
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdComment 
      Caption         =   "C&omment"
      Height          =   375
      Left            =   11520
      TabIndex        =   3
      ToolTipText     =   "Enter a comment"
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdAccept 
      Caption         =   "&Accept"
      Default         =   -1  'True
      Height          =   375
      Left            =   11520
      TabIndex        =   0
      ToolTipText     =   "Accept data and continue"
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "&Modify"
      Height          =   375
      Left            =   11520
      TabIndex        =   1
      ToolTipText     =   "Modify calibration"
      Top             =   600
      Width           =   1215
   End
   Begin CRVIEWER9LibCtl.CRViewer9 objCRViewer 
      Height          =   10965
      Left            =   0
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   11325
      lastProp        =   600
      _cx             =   19976
      _cy             =   19341
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "ReportContainer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
 
'FILE:  ReportContainer.frm
 
'DESCRIPTION:  This module contains the form where the final report is displayed.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: ReportContainer.frm $
 ' 
 ' *****************  Version 9  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:31a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added event handler for cancel button.
 '
 ' *****************  Version 8  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:39p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Updated to 1.0.0X9
 '
 ' *****************  Version 7  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:40p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Removes the logic that forces user to make a comment if they updated
 ' the fit parameters.  This logic is in the business objects now.
 '
 ' *****************  Version 6  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:34p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added error traps.
 ' Set object references equal to nothing to free resources.
 ' Added functionality to plot calibration.
 '
 ' *****************  Version 5  *****************
 ' User: Ballard      Date: 5/09/03    Time: 3:56p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Lot object is set through function UpdateReport instead of being set by
 ' a property.
 '
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 3/21/03    Time: 2:37p
 ' Updated in $/QalibVBClient
 ' Renamed module "frmFinalReport.frm" to "ReportContainer.frm."
 '
 ' Added ability to make a comment to the calibration.  Also added flags
 ' so that parent form would know whether to perform calibration again,
 ' proceed as is, or cancel.
 '
 ' The form retrieves data from the underlying data object.
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 1/24/03    Time: 2:52p
 ' Updated in $/QalibVBClient
 ' User can now modify the fit parameters and the report will refresh with
 ' the new goodness of calibration checks.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 1/17/03    Time: 4:22p
 ' Updated in $/QalibVBClient
 ' Added public property get so that other objects can get at the report
 ' object.  This is necessary so that other objects can prepare the
 ' report.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 1/09/03    Time: 3:45p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe.

'private constants

' Formatting constants for the chart
Private Const CALCURVETITLE As String = "Best Fit"


' private member variables
Private objReport_m As FinalReport ' stores a reference to the internal report
Private objLotReport_m As LotReport  ' stores a reference to the underlying lot report object
Private intMeasuredAxis_m As enum_MeasuredAxis ' stores which axis the measured values are on

Option Explicit

'***********************************************************************

'PROCEDURE:   Form_Terminate()

'DESCRIPTION: Event handler for when the form terminates

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Form_Terminate()
On Error GoTo ErrTrap
    Set objReport_m = Nothing
    Set objLotReport_m = Nothing
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ReportContainer.Form_Terminate", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdAccept_Click()

'DESCRIPTION: Event handler for when the user clicks "Accept"

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdAccept_Click()
On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    Call objLotReport_m.Accept
    Call Me.Hide
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ReportContainer.cmdAccept_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdCancel_Click()

'DESCRIPTION: Hides the form

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdCancel_Click()
On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    Call Me.Hide
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | Comment.cmdCancel_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdModify_Click()

'DESCRIPTION: Event handler for when the user clicks "Modify"

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdModify_Click()
On Error GoTo ErrTrap
    Dim frmFitParamsEdit As FitParamsEdit  ' stores the fit parameter edit form
    
    Screen.MousePointer = vbHourglass
    
    ' set up the fit parameters edit form
    Set frmFitParamsEdit = New FitParamsEdit
        
    ' plug the fit parameters edit object into the form
    Call frmFitParamsEdit.LoadFitParams(objLotReport_m.EditableFitParams)
    
    Screen.MousePointer = vbDefault
    
    ' give the user the ability to change the fit parameters
    Call frmFitParamsEdit.Show(vbModal, Me)
    
    Screen.MousePointer = vbHourglass
    
    ' see if user canceled changing the fit parameters
    If (frmFitParamsEdit.Cancel = False) Then
        Screen.MousePointer = vbHourglass
                 
        ' make the report object update itself
        Call objLotReport_m.LoadCalibration(objLotReport_m.CalID)
                 
        Call UpdateReport
    End If
    
    Screen.MousePointer = vbDefault
    
    Set frmFitParamsEdit = Nothing

    Exit Sub
ErrTrap:
    Set frmFitParamsEdit = Nothing
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ReportContainer.cmdModify_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdComment_Click()

'DESCRIPTION: Handles when the user clicks the "Comment" button

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdComment_Click()
On Error GoTo ErrTrap
    Screen.MousePointer = vbHourglass
    If (RequestComment = vbOK) Then
        Call UpdateReport
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ReportContainer.cmdComment_Click", Err.Description)
End Sub

'***********************************************************************
'
'PROCEDURE:   chtCalibration_OnGetSeriesMark
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
Private Sub chtCalibration_OnGetSeriesMark(ByVal seriesIndex As Long, ByVal ValueIndex As Long, MarkText As String)
    ' the mark text of the point is contained in the point label
    MarkText = chtCalibration.series(seriesIndex).pointLabel(ValueIndex)
End Sub

'***********************************************************************

'PROCEDURE:   RequestComment()

'DESCRIPTION: Handles when the user clicks the "Comment" button

'PARAMETERS:  N/A

'RETURNED:    Whether user canceled the comment

'*********************************************************************
Private Function RequestComment() As VbMsgBoxResult
On Error GoTo ErrTrap
    Dim frmComment As Comment
    
    Set frmComment = New Comment
    
    ' plug the comment object into the form
    Call frmComment.LoadComment(objLotReport_m)
    
    Screen.MousePointer = vbDefault
    
    Call frmComment.Show(vbModal, Me)
    
    Screen.MousePointer = vbHourglass
    
    ' see if the user canceled
    If (frmComment.Cancel = True) Then
        RequestComment = vbCancel
    Else
        RequestComment = vbOK
    End If
    
    Call Unload(frmComment)
    
     Set frmComment = Nothing
    Exit Function
ErrTrap:
     Set frmComment = Nothing
    Call Err.Raise(APPERR, Err.Source & " | ReportContainer.RequestComment", Err.Description)
End Function

'***********************************************************************
'
'PROCEDURE:   PlotCalibration
'
'DESCRIPTION: Loads the calibration curve stored in the lot report object
'
'PARAMETERS:  N/A
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub PlotCalibration()
On Error GoTo ErrTrap
    Dim seriesNum As Long
    Dim curSeries As SeriesOne
    
    '  add the fitted curve to the chart
    seriesNum = chtCalibration.AddSeries(scLine)
    
    With chtCalibration.series(seriesNum)
        For Each curSeries In objLotReport_m.AnalyzerData.SeriesSet
            ' only plot the calibrators
            If (curSeries.IsCalibrator = True) Then
                ' the curve is loaded differently depending on what axis orientation the user selected
                If (intMeasuredAxis_m = abxMeasuredX) Then
                    Call .AddXY(curSeries.CalcMeasuredVal, curSeries.AssignedVal, "", clTeeColor)
                Else
                    Call .AddXY(curSeries.AssignedVal, curSeries.CalcMeasuredVal, "", clTeeColor)
                End If
            End If
        Next curSeries
        
        ' title the calibration curve
        .Title = CALCURVETITLE
    End With
    
    Set curSeries = Nothing
        
    Exit Sub
ErrTrap:
    Set curSeries = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | ReportContainer.PlotCalibration", Err.Description)
End Sub

'***********************************************************************
'
'PROCEDURE:   PlotCalibrationCheck
'
'DESCRIPTION: Loads a calibration check stored in the lot report object
'
'PARAMETERS:  inCalCheck - the calibration check to plot
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub PlotCalibrationCheck(inCalCheck As String)
On Error GoTo ErrTrap
    Dim seriesNum As Long
    Dim curPlot As Plot
    Dim curPoint As PlotPoint
    
    ' get the appropriate calibration check from the lot report object
    Set curPlot = CallByName(objLotReport_m, inCalCheck & "CheckPlot", VbGet)
    
    ' make sure the calibration check is valid
    If ((curPlot Is Nothing) = True) Then
        Exit Sub
    End If
    
    With chtCalCheck
     
        ' clear out the chart
        Call .ClearChart
        
        ' add a point series
        seriesNum = .AddSeries(scPoint)
        
        ' add the points to the chart
        With .series(seriesNum)
            For Each curPoint In curPlot.PlotPoints
                Call .AddXY(curPoint.XValue, curPoint.YValue, "", clTeeColor)
            Next curPoint
        End With
        
        ' format the chart
        Call .Header.Text.Add(curPlot.Name)
        .Aspect.View3D = False
        .Zoom.Enable = False
        .Scroll.Enable = pmNone
        .Legend.visible = False
        
        ' save the calibration check plot to a file
        Call .Export.asJPEG.SaveToFile(App.Path & "\" & inCalCheck & CALCHECKPLOTFILE)
        
    End With
    
    Set curPlot = Nothing
    Set curPoint = Nothing
        
    Exit Sub
ErrTrap:
    Set curPlot = Nothing
    Set curPoint = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | ReportContainer.PlotCalibrationCheck", Err.Description)
End Sub

'***********************************************************************
'
'PROCEDURE:   UpdateReport
'
'DESCRIPTION: Forces the report to update it's records
'
'PARAMETERS:  N/A
'
'RETURNED:    N/A
'
'*********************************************************************
Private Sub UpdateReport()
On Error GoTo ErrTrap
    
    ' set up the calibration plot
    Call PlotSamples(objLotReport_m.AnalyzerData, chtCalibration, intMeasuredAxis_m, objLotReport_m.Chemistry)
    Call PlotCalibration
    Call chtCalibration.Export.asJPEG.SaveToFile(App.Path & "\" & CALPLOTFILE)
    
    ' set up the calibration check plots
    Call PlotCalibrationCheck(CALIBRATORPREFIX)
    Call PlotCalibrationCheck(CONTROLPREFIX)
    
    If ((objReport_m Is Nothing) = True) Then
        Set objReport_m = New FinalReport
        
'         connect the report viewer to the report variable
        objCRViewer.ReportSource = objReport_m
        Call objCRViewer.ViewReport
    End If
    
    Call objReport_m.UpdateReport(objLotReport_m)  ' the report needs the lot object too
    
'     may alleviate the control is busy downloading data error
    Do While (objCRViewer.IsBusy = True)
        DoEvents
    Loop
    
    Call objCRViewer.Refresh

    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | ReportContainer.UpdateReport", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   LoadReport()

'DESCRIPTION: Loads the initial report

'PARAMETERS:  inLotReport - the lot report object
'             measuredAxis - the axis the measured values are on

'RETURNED:    N/A

'*********************************************************************
Public Sub LoadReport(inLotReport As LotReport, measuredAxis As enum_MeasuredAxis)
On Error GoTo ErrTrap
    Set objLotReport_m = inLotReport
    
    intMeasuredAxis_m = measuredAxis
        
    Call UpdateReport
    Exit Sub
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | ReportContainer.LoadReport", Err.Description)
End Sub
