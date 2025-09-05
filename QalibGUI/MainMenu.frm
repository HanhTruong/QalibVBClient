VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form MainMenu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Qalib"
   ClientHeight    =   2640
   ClientLeft      =   6495
   ClientTop       =   5115
   ClientWidth     =   2460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   2460
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   960
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame fraMenu 
      Height          =   1575
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   1695
      Begin VB.CommandButton cmdQalibrate 
         Caption         =   "&Qalibrate..."
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   0
         ToolTipText     =   "Start new calibration"
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdExit 
         Cancel          =   -1  'True
         Caption         =   "E&xit"
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         TabIndex        =   1
         ToolTipText     =   "Exit Qalib"
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Main Menu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   180
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "MainMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
 
'FILE:  MainMenu.frm
 
'DESCRIPTION:  This module contains the form where the Qalib main menu is displayed.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: MainMenu.frm $
 ' 
 ' *****************  Version 14  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:34a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 '
 ' *****************  Version 13  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:39p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Updated to 1.0.0X9
 '
 ' *****************  Version 12  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:41p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Allows user to select which chemistry to calibrate now.
 '
 ' *****************  Version 11  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added error traps.
 ' Set object references equal to nothing to free resources.
 ' Added RestoreMainMenu function to perform cleanup after a calibration.
 '
 ' *****************  Version 10  *****************
 ' User: Ballard      Date: 5/09/03    Time: 3:58p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' File dialog was added (instead of using external class) to facilitate
 ' user selecting a calibration file.
 '
 ' *****************  Version 9  *****************
 ' User: Ballard      Date: 3/25/03    Time: 3:42p
 ' Updated in $/QalibVBClient
 ' Loaded calibration attributes for testing.
 '
 ' *****************  Version 8  *****************
 ' User: Ballard      Date: 3/21/03    Time: 1:29p
 ' Updated in $/QalibVBClient
 ' Changed form name from "frmMenu.frm" to "MainMenu.frm."
 '
 ' All the functionality to select, open, and close the Excel file has
 ' been removed and placed in the "QalibInput" module.  This was done
 ' because later the input will come from a database and this will make
 ' the transition easier.
 '
 ' The "Review" and "Commit" buttons were removed as this functionality
 ' will be in a separate program.
 '
 ' Completely reworked calibration flow to allow user to change fit
 ' parameters an infinite number of times and enter a comment.
 '
 ' Added event handler for child objects to show a progress bar.
 '
 ' *****************  Version 7  *****************
 ' User: Ballard      Date: 1/24/03    Time: 2:51p
 ' Updated in $/QalibVBClient
 ' Added functionality to allow user to change fit parameters.
 '
 ' *****************  Version 6  *****************
 ' User: Ballard      Date: 1/17/03    Time: 4:19p
 ' Updated in $/QalibVBClient
 ' Changed the form to diasabled by default.
 ' Removed references to database environment since local file database
 ' has been replaced by internal virtaul database (CrystalComObject).
 ' Added ToggleForm function to streamline toggling the GUI and mouse
 ' pointer before/after processing actions.
 ' Changed the common dialog box so its inital path is the same as that of
 ' the application.  Also made it so the dialog will not change the
 ' working directory based on the path the user selects.
 '
 ' *****************  Version 5  *****************
 ' User: Ballard      Date: 1/09/03    Time: 3:55p
 ' Updated in $/QalibVBClient
 ' Added file and function headers.
 ' Added common dialog control for selecting input file.
 ' Added error handling to calibration button event handler.
 ' Added functionality to open/close Excel application/workbook/worksheet.
 '
 '
 ' *****************  Version 4  *****************
 ' User: Alves        Date: 12/06/02   Time: 11:00a
 ' Updated in $/QalibVBClient
 ' Update documentation
 '
 ' *****************  Version 3  *****************
 ' User: Alves        Date: 12/06/02   Time: 10:39a
 ' Updated in $/QalibVBClient
 ' Return goodness of fit to client.
 ' Calculate rate ofr assigned values.
 ' Correlation coefficient for CALIBRATOR and CONTROLs
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 12/04/02   Time: 9:24a
 ' Updated in $/QalibVBClient
 ' Changed "Quit" procedure to unload form instead of ending whole
 ' program.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 9/24/02    Time: 1:46p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe

Option Explicit

' private member variables
Private WithEvents objLotVerify_m As LotVerify  ' object which handles the verification of lot data
Attribute objLotVerify_m.VB_VarHelpID = -1
Private frmStandby_m As Standby ' the form to show when a lot of processing is taking place
Private blnStandbyShowing_m As Boolean ' whether the standby form is showing
Private strUser_m As String ' the logged in user
Private strPassword_m As String ' the user's password
Private strMode_m As String ' the mode of the program


'***********************************************************************

'PROCEDURE:   cmdExit_Click()

'DESCRIPTION: Event handler for when the user clicks exit

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdExit_Click()
On Error GoTo ErrTrap
    ' unloading this form will terminate the application
    Call Unload(Me)
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | MainMenu.cmdExit_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   objLotVerify_m_UpdateStatus()

'DESCRIPTION: Status event handler for the lot verification object

'PARAMETERS:  percentComplete - how much to update the progress bar
'             message - text to display in the progress dialog box

'RETURNED:    N/A

'*********************************************************************
Private Sub objLotVerify_m_UpdateStatus(percent As Integer, message As String)
On Error GoTo ErrTrap
    Call UpdateStandby(percent, message) ' forward the call
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | MainMenu.objLotVerify_m_UpdateStatus", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   ToggleForm()

'DESCRIPTION: Toggles enabled state of form

'PARAMETERS:  state - whether the form is enabled or not

'RETURNED:    N/A

'*********************************************************************
Private Sub ToggleForm(state As Boolean)
On Error GoTo ErrTrap
    ' toggle buttons
    cmdQalibrate.Enabled = state
    cmdExit.Enabled = state
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | MainMenu.ToggleForm", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   ToggleStandby()

'DESCRIPTION: Handles showing and closeing the Standby form

'PARAMETERS:  visible - whether to turn the Standby form on or off

'RETURNED:    N/A

'*********************************************************************
Private Sub ToggleStandby(visible As Boolean)
On Error GoTo ErrTrap
    If (visible = True) Then
        ' setup the standby form
        Set frmStandby_m = New Standby
        Call frmStandby_m.Show(vbModeless, Me)
        blnStandbyShowing_m = True
    Else
        If (blnStandbyShowing_m = True) Then
        
            ' take down the standby form
            Call Unload(frmStandby_m)
            Set frmStandby_m = Nothing
            blnStandbyShowing_m = False
        End If
    End If
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | MainMenu.ToggleStandby", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   UpdateStandby()

'DESCRIPTION: Handles showing, updating, and closeing the Standby form

'PARAMETERS:  percentComplete - how to update the progress bar
'             message - text to display in the progress dialog box

'RETURNED:    N/A

'*********************************************************************
Private Sub UpdateStandby(percent As Integer, message As String)
On Error GoTo ErrTrap
    ' make sure the percent is reasonable
    If (percent < 0 Or percent > 100) Then
        Exit Sub
    End If
    
    ' make sure standby form is showing
    If (blnStandbyShowing_m = True) Then
        Call frmStandby_m.Update(percent, message)
    End If
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | MainMenu.UpdateStandby", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:    RestoreMainMenu()

'DESCRIPTION: Cleans up forms and variables from processing lot

'PARAMETERS:  N/A

'RETURNED:    N/A

'***********************************************************************
Private Sub RestoreMainMenu()
On Error GoTo ErrTrap

    Dim frmIndex As Form ' for iterating through forms

    dlgFile.FileName = ""
    
    ' unload open forms
    For Each frmIndex In Forms
        If (frmIndex.Caption <> Me.Caption) Then
            Call Unload(frmIndex)
        End If
    Next frmIndex
    
    ' release the memory
    Set objLotVerify_m = Nothing
    Set frmStandby_m = Nothing
    
    Call ToggleForm(True)
    Screen.MousePointer = vbDefault
    
    Set frmIndex = Nothing
    
    Exit Sub
ErrTrap:
    Set frmIndex = Nothing
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | MainMenu.RestoreMainMenu", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdQalibrate_Click()

'DESCRIPTION: Event handler for when the user clicks qalibrate; launch the
' calibration wizard

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdQalibrate_Click()
    On Error GoTo ErrHandler

    Dim frmChartContainer As ChartContainer ' for showing the chart
    Dim frmReportContainer As ReportContainer ' for showing the report
    Dim frmChemistrySelect As ChemistrySelect ' for selecting a chemistry
    Dim objLotReport As LotReport ' for storing the report data
    Dim objLotChemistrySelect As LotChemistrySelect ' for selecting a chemistry
    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    Screen.MousePointer = vbHourglass
    Call ToggleForm(False)
    
    ' set up the file dialog
    dlgFile.InitDir = GetSetting(App.Title, SECTIONKEY, LASTPATHKEY, App.Path)
    dlgFile.Flags = cdlOFNNoChangeDir

    ' let the user select a file
    Call dlgFile.ShowOpen
    
    Screen.MousePointer = vbHourglass
    Call SaveSetting(App.Title, SECTIONKEY, LASTPATHKEY, fso.GetParentFolderName(dlgFile.FileName))
    
    Set objLotChemistrySelect = New LotChemistrySelect
    Call objLotChemistrySelect.LoadChems(dlgFile.FileName, strUser_m, strPassword_m)

    ' set up the chemistry selection form
    Set frmChemistrySelect = New ChemistrySelect
    Call frmChemistrySelect.LoadChems(objLotChemistrySelect.Chems)

    Screen.MousePointer = vbDefault

    ' show the chemistry selection form and allow the user to select a chemistry
    Call frmChemistrySelect.Show(vbModal, Me)
    If (frmChemistrySelect.Cancel = True) Then
        ' run the clean up code
        Call RestoreMainMenu
        Set frmChemistrySelect = Nothing
        Set objLotChemistrySelect = Nothing
        Exit Sub
    End If
    
    Call ToggleStandby(True)
    Set objLotVerify_m = New LotVerify
    Call objLotVerify_m.LoadRuns(dlgFile.FileName, frmChemistrySelect.SelectedChemistry, strUser_m, strPassword_m)
    
    ' set up the chart
    Set frmChartContainer = New ChartContainer
    Call frmChartContainer.LoadChart(objLotVerify_m)

    Screen.MousePointer = vbDefault
    Call ToggleStandby(False) ' this will close the standby form

    ' show the chart form and allow the user to review the calibrators
    Call frmChartContainer.Show(vbModal, Me)
    Screen.MousePointer = vbHourglass

    If (frmChartContainer.Cancel = True) Then
        ' run the clean up code
        Call RestoreMainMenu
        Set frmChartContainer = Nothing
        Set frmChemistrySelect = Nothing
        Set objLotChemistrySelect = Nothing
        Exit Sub
    End If
    
    Call objLotVerify_m.CalibrateLot(strMode_m)

    ' set up the report
    Set objLotReport = New LotReport
    Call objLotReport.LoadCalibration(objLotVerify_m.CalID)
    Set frmReportContainer = New ReportContainer

    ' plug the lot object into the report
    Call frmReportContainer.LoadReport(objLotReport, frmChartContainer.measuredAxis)

    ' show the report and allow the user to update the fit parameters
    Screen.MousePointer = vbDefault
    Call frmReportContainer.Show(vbModal, Me)
    Screen.MousePointer = vbHourglass
    
    ' run the clean up code
    Call RestoreMainMenu
    
    Set fso = Nothing
    Set frmChartContainer = Nothing
    Set frmReportContainer = Nothing
    Set frmChemistrySelect = Nothing
    Set objLotReport = Nothing
    Set objLotChemistrySelect = Nothing
    Exit Sub

ErrHandler:
    If (Err.Number <> cdlCancel) Then
        ' can't raise an error here because there's nothing up the call stack to trap it
        Call HandleError(Err.Number, Err.Source & " | MainMenu.cmdQalibrate_Click", Err.Description)
    End If
    
    Screen.MousePointer = vbHourglass
    
    Call ToggleStandby(False) ' this will close the standby form

    ' run the clean up code
    Call RestoreMainMenu
    
    Set fso = Nothing
    Set frmChartContainer = Nothing
    Set frmReportContainer = Nothing
    Set frmChemistrySelect = Nothing
    Set objLotReport = Nothing
    Set objLotChemistrySelect = Nothing
End Sub

'***********************************************************************

'PROCEDURE:   InitializeSession()

'DESCRIPTION: Sets up the form global variables and starts the first file processing

'PARAMETERS:  inUser - the user logged in
'             inPassword - the user's password
'             inMode - the mode of the application

'RETURNED:    N/A

'***********************************************************************
Public Sub InitializeSession(inUser As String, inPassword As String, inMode As String)
On Error GoTo ErrTrap
    strUser_m = inUser  ' save the user for later
    strMode_m = inMode  ' save the mode for later
    strPassword_m = inPassword ' save the password for later
    Call cmdQalibrate_Click
Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | MainMenu.InitializeSession", Err.Description)
End Sub

