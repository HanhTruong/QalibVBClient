VERSION 5.00
Begin VB.Form Mode 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Mode"
   ClientHeight    =   2475
   ClientLeft      =   5610
   ClientTop       =   4455
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstModeSelect 
      Height          =   2205
      ItemData        =   "Mode.frx":0000
      Left            =   120
      List            =   "Mode.frx":0002
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Mode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
 
'FILE:  Mode.frm
 
'DESCRIPTION:  This module contains the form where the user selects the operation mode

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: Mode.frm $
 ' 
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:39p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Updated to 1.0.0X9
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:21p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added error traps.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 3/21/03    Time: 2:40p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe.

Option Explicit

'private member variables
Private blnCancel_m As Boolean  ' stores whether cancel button was pressed

'***********************************************************************

'PROPERTY GET:   Cancel()

'DESCRIPTION: Allows other objects to see if cancel button was pressed

'PARAMETERS:  N/A

'RETURNED:    Whether the cancel button was pressed

'*********************************************************************
Public Property Get Cancel() As Boolean
On Error GoTo ErrTrap
    Cancel = blnCancel_m
    Exit Property
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | Mode.PropertyGet.Cancel", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   SelectedMode()

'DESCRIPTION: Allows other objects to get the selected mode

'PARAMETERS:  N/A

'RETURNED:    the selected mode

'*********************************************************************
Public Property Get SelectedMode() As String
On Error GoTo ErrTrap
    SelectedMode = lstModeSelect.List(lstModeSelect.ListIndex)
    Exit Property
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | Mode.PropertyGet.SelectedMode", Err.Description)
End Property


'***********************************************************************

'PROCEDURE:   Form_QueryUnload()

'DESCRIPTION: Executes right before form is unloaded so method of closing
'can be determined

'PARAMETERS:  Cancel - whether to cancel the unload
'             UnloadMode - how the form was unloaded

'RETURNED:    N/A

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
    Call HandleError(Err.Number, Err.Source & " | Mode.Form_QueryUnload", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdCancel_Click()

'DESCRIPTION: Hides the form and sets cancel property

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdCancel_Click()
On Error GoTo ErrTrap
    blnCancel_m = True
    Me.Hide
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | Mode.cmdCancel_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdOK_Click()

'DESCRIPTION: Processes the mode selection

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrTrap
    blnCancel_m = False
    Me.Hide
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | Mode.cmdOK_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   LoadModes()

'DESCRIPTION: Public interface to set up the mode selection

'PARAMETERS:  inMode - the modes collection

'RETURNED:    N/A

'*********************************************************************
Public Sub LoadModes(inModes As QalibCollection)
On Error GoTo ErrTrap
    Dim curMode As Mode
    
    ' add the available modes to the list box
    For Each curMode In inModes
        Call lstModeSelect.AddItem(curMode.Name)
    Next curMode
    
    Exit Sub
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | Mode.LoadModes", Err.Description)
End Sub


