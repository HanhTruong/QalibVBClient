VERSION 5.00
Begin {BD4B4E61-F7B8-11D0-964D-00A0C9273C2A} FinalReport 
   ClientHeight    =   7710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11415
   OleObjectBlob   =   "FinalReport.dsx":0000
End
Attribute VB_Name = "FinalReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'FILE:  FinalReport.dsr
 
'DESCRIPTION:  This module contains the final report object for Qalib.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: FinalReport.Dsr $
 '
 ' *****************  Version 9  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:34a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 '
 ' *****************  Version 8  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:39p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Updated to 1.0.0X9
 '
 ' *****************  Version 7  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:38p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added calibrator and control plots.
 '
 ' *****************  Version 5  *****************
 ' User: Ballard      Date: 5/09/03    Time: 3:53p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Lot object is set through function UpdateReport instead of being set by
 ' a property.  Calibration chart is now loaded from lot object instead of
 ' being pasted from clipboard.  Logic to suppress and highlight rows was
 ' moved to the actual report object property pages.  This was necessary
 ' because code would not functon properly in the event handler.
 '
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 3/21/03    Time: 2:15p
 ' Updated in $/QalibVBClient
 ' Renamed module "rptQalibFinal.dsr" to "FinalReport.dsr."
 '
 ' Added functionality to suppress the adjusted value/disposition headers.
 ' Also added the ability to update the report with a function call.
 '
 ' The report retrieves data from the underlying data object.
 '
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 1/24/03    Time: 2:50p
 ' Updated in $/QalibVBClient
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 1/17/03    Time: 4:21p
 ' Updated in $/QalibVBClient
 ' Added private member variable and public property set so that another
 ' object can set the picture source for the report.  This is "safer" than
 ' getting the picture from the clipboard.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 1/09/03    Time: 3:47p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe.

Option Explicit

' constants
Private Const NAMEFLD As String = "fldName"
Private Const VALUEFLD As String = "fldValue"
Private Const DISPOSITIONFLD As String = "fldDisposition"
Private Const ADJVALUEFLD As String = "fldAdjValue"
Private Const ADJDISPOSITIONFLD As String = "fldAdjDisposition"
Private Const MINLIMITFLD As String = "fldMinLimit"
Private Const MAXLIMITFLD As String = "fldMaxLimit"
Private Const ISMODIFIEDFLD As String = "fldIsModified"
Private Const ISVALPASSEDFLD As String = "fldIsValPassed"
Private Const ISADJVALPASSEDFLD As String = "fldIsAdjValPassed"
Private Const vbGray As Long = &HC0C0C0

' private member variables
Private objLotReport_m As LotReport  ' stores a reference to the underlying lot report object

'***********************************************************************

'PROCEDURE:   Report_Terminate()

'DESCRIPTION: Event handler for when the report terminates

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Report_Terminate()
On Error GoTo ErrTrap
    Set objLotReport_m = Nothing
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | FinalReport.Report_Terminate", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   secCalAttribs_Format()

'DESCRIPTION: Event handler for when calibration attributes section is being formatted

'PARAMETERS:  pFormattingInfo - status structure

'RETURNED:    N/A

'*********************************************************************

Private Sub secCalAttribs_Format(ByVal pFormattingInfo As Object)
On Error GoTo ErrTrap

     ' add the lot attributes to the report
    Call fldDate.SetText(objLotReport_m.CalDate)
    Call fldChemistry.SetText(objLotReport_m.Chemistry)
    Call fldWavelength.SetText(objLotReport_m.Wavelength)
    Call fldSpecies.SetText(objLotReport_m.Species)
    Call fldDiluent.SetText(objLotReport_m.Diluent)
    Call fldMold.SetText(objLotReport_m.Mold)
    Call fldRotor.SetText(objLotReport_m.Rotor)
    Call fldCuvette.SetText(objLotReport_m.Cuvette)
    Call fldMode.SetText(objLotReport_m.Mode)
    Call fldUser.SetText(objLotReport_m.User)
Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | FinalReport.secCalAttribs_Format", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   secCalPlot_Format()

'DESCRIPTION: Event handler for when calibration plot section is being formatted

'PARAMETERS:  pFormattingInfo - status structure

'RETURNED:    N/A

'*********************************************************************
Private Sub secCalPlot_Format(ByVal pFormattingInfo As Object)
On Error GoTo ErrTrap

    ' paste in the calibration plot
    Set picCalibration.FormattedPicture = LoadPicture(App.Path & "\" & CALPLOTFILE)
    
    ' set the calibration ID
    Call fldRHCalID.SetText(objLotReport_m.CalID)
Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | FinalReport.secCalPlot_Format", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   secCorrPlot_Format()

'DESCRIPTION: Event handler for when correlation plot section is being formatted

'PARAMETERS:  pFormattingInfo - status structure

'RETURNED:    N/A

'*********************************************************************
Private Sub secCorrPlot_Format(ByVal pFormattingInfo As Object)
On Error GoTo ErrTrap

    ' if there is a calibrator check plot then display it and its attributes
    If ((objLotReport_m.CalibratorCheckPlot Is Nothing) = False) Then
        ' add the calibrator check attributes to the report
        Call fldCalibratorCheckCorrCoeff.SetText(objLotReport_m.CalibratorCheckPlot.CorrCoeff)
        Call fldCalibratorCheckSlope.SetText(objLotReport_m.CalibratorCheckPlot.Slope)
        Call fldCalibratorCheckSlopeErr.SetText(objLotReport_m.CalibratorCheckPlot.SlopeErr)
        Call fldCalibratorCheckIntercept.SetText(objLotReport_m.CalibratorCheckPlot.Intercept)
        Call fldCalibratorCheckInterceptErr.SetText(objLotReport_m.CalibratorCheckPlot.InterceptErr)
    
        ' paste in the calibrator check plot
        Set picCalibratorCheck.FormattedPicture = LoadPicture(App.Path & "\" & CALIBRATORPREFIX & CALCHECKPLOTFILE)
    Else
        lblCalibratorCheckCorrCoeff.Suppress = True
        lblCalibratorCheckSlope.Suppress = True
        lblCalibratorCheckSlopeErr.Suppress = True
        lblCalibratorCheckIntercept.Suppress = True
        lblCalibratorCheckInterceptErr.Suppress = True
        
        fldCalibratorCheckCorrCoeff.Suppress = True
        fldCalibratorCheckSlope.Suppress = True
        fldCalibratorCheckSlopeErr.Suppress = True
        fldCalibratorCheckIntercept.Suppress = True
        fldCalibratorCheckInterceptErr.Suppress = True
        picCalibratorCheck.Suppress = True
    End If
    
    ' if there is a control check plot then display it and its attributes
    If ((objLotReport_m.ControlCheckPlot Is Nothing) = False) Then
        ' add the control check attributes to the report
        Call fldControlCheckCorrCoeff.SetText(objLotReport_m.ControlCheckPlot.CorrCoeff)
        Call fldControlCheckSlope.SetText(objLotReport_m.ControlCheckPlot.Slope)
        Call fldControlCheckSlopeErr.SetText(objLotReport_m.ControlCheckPlot.SlopeErr)
        Call fldControlCheckIntercept.SetText(objLotReport_m.ControlCheckPlot.Intercept)
        Call fldControlCheckInterceptErr.SetText(objLotReport_m.ControlCheckPlot.InterceptErr)
        ' paste in the control check plot
        Set picControlCheck.FormattedPicture = LoadPicture(App.Path & "\" & CONTROLPREFIX & CALCHECKPLOTFILE)
    Else
        lblControlCheckCorrCoeff.Suppress = True
        lblControlCheckSlope.Suppress = True
        lblControlCheckSlopeErr.Suppress = True
        lblControlCheckIntercept.Suppress = True
        lblControlCheckInterceptErr.Suppress = True
        
        fldControlCheckCorrCoeff.Suppress = True
        fldControlCheckSlope.Suppress = True
        fldControlCheckSlopeErr.Suppress = True
        fldControlCheckIntercept.Suppress = True
        fldControlCheckInterceptErr.Suppress = True
        picControlCheck.Suppress = True
    End If
Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | FinalReport.secCorrPlot_Format", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   secDetail_Format

'DESCRIPTION: Event handler for when report detail is being formatted

'PARAMETERS:  pFormattingInfo - status structure

'RETURNED:    N/A

'*********************************************************************
Private Sub secDetail_Format(ByVal pFormattingInfo As Object)
    'see the report object for logic that highlights failed values and
    ' suppresses adjusted row as necessary
End Sub

'***********************************************************************

'PROCEDURE:   secPageFooter_Format()

'DESCRIPTION: Event handler for when page footer is being formatted

'PARAMETERS:  pFormattingInfo - status structure

'RETURNED:    N/A

'*********************************************************************
Private Sub secPageFooter_Format(ByVal pFormattingInfo As Object)
On Error GoTo ErrTrap
    Call fldPFCalID.SetText(objLotReport_m.CalID)
Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | FinalReport.secPageFooter_Format", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   secReportFooter_Format()

'DESCRIPTION: Event handler for when report footer is being formatted

'PARAMETERS:  pFormattingInfo - status structure

'RETURNED:    N/A

'*********************************************************************
Private Sub secReportFooter_Format(ByVal pFormattingInfo As Object)
On Error GoTo ErrTrap
    Call fldComment.SetText(objLotReport_m.Comment)
 Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | FinalReport.secReportFooter_Format", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   UpdateReport()

'DESCRIPTION: Update the report's records because the underlying database has changed

'PARAMETERS:  inLotReport - the lot report object

'RETURNED:    N/A

'*********************************************************************
Public Sub UpdateReport(inLotReport As LotReport)
On Error GoTo ErrTrap
    Dim objTable As CrystalComObject
    Dim curResult As Result
    
    Set objLotReport_m = inLotReport

    ' prepare the virtual database for the report
    Set objTable = New CrystalComObject
    Call objTable.AddField(NAMEFLD, vbString)
    Call objTable.AddField(VALUEFLD, vbString)
    Call objTable.AddField(MINLIMITFLD, vbString)
    Call objTable.AddField(MAXLIMITFLD, vbString)
    Call objTable.AddField(DISPOSITIONFLD, vbString)
    Call objTable.AddField(ISVALPASSEDFLD, vbBoolean)
    Call objTable.AddField(ISMODIFIEDFLD, vbBoolean)
    Call objTable.AddField(ADJVALUEFLD, vbString)
    Call objTable.AddField(ADJDISPOSITIONFLD, vbString)
    Call objTable.AddField(ISADJVALPASSEDFLD, vbBoolean)

    ' cycle through and add the results to the virtual database
    For Each curResult In objLotReport_m.Results
        With curResult
            Call objTable.AddRows(Array(.Name, CStr(.Value), CStr(.MinLimit), CStr(.MaxLimit), .Disposition, _
                .IsValPassed, .IsModified, CStr(.AdjValue), .AdjDisposition, .IsAdjValPassed))
        End With
    Next curResult

    ' link the lot object's virtual database to the report's database
    Call Database.Tables(1).SetDataSource(objTable)
    Call AutoSetUnboundFieldSource(crBMTNameAndValue)
    
    Set objTable = Nothing
    Set curResult = Nothing
    Exit Sub
ErrTrap:
    Set objTable = Nothing
    Set curResult = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | FinalReport.UpdateReport", Err.Description)
End Sub












