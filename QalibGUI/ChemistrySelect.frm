VERSION 5.00
Begin VB.Form ChemistrySelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Chemistry"
   ClientHeight    =   3825
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   4395
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstChemistrySelect 
      Height          =   3570
      ItemData        =   "ChemistrySelect.frx":0000
      Left            =   120
      List            =   "ChemistrySelect.frx":0002
      TabIndex        =   0
      ToolTipText     =   "Select a chemistry"
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "ChemistrySelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
 
'FILE:  ChemistrySelect.frm
 
'DESCRIPTION:  This module contains the form where the user can select a cheistry to calibrate.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: ChemistrySelect.frm $
 ' 
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:39a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Fixed tab index.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 10/10/03   Time: 1:35p
 ' Created in $/QalibVBClient/Source/QalibGUI
 ' Added to SourceSafe.
 '
Option Explicit

' private member variables
Private blnCancel_m As Boolean ' stores whether cancel button was pressed

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
    Call Err.Raise(APPERR, Err.Source & " | ChemistrySelect.PropertyGet.Cancel", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   SelectedChemistry()

'DESCRIPTION: Allows other objects to get the selected chemistry

'PARAMETERS:  N/A

'RETURNED:    the selected chemistry

'*********************************************************************
Public Property Get SelectedChemistry() As String
On Error GoTo ErrTrap
    SelectedChemistry = lstChemistrySelect.List(lstChemistrySelect.ListIndex)
    Exit Property
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | ChemistrySelect.PropertyGet.SelectedChemistry", Err.Description)
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
    Call HandleError(Err.Number, Err.Source & " | ChemistrySelect.Form_QueryUnload", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdCancel_Click()

'DESCRIPTION: Event handler for pressing "Cancel"

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdCancel_Click()
On Error GoTo ErrTrap
    blnCancel_m = True  ' let other objects know user pressed cancel
    Call Me.Hide
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ChemistrySelect.cmdCancel_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdOK_Click()

'DESCRIPTION: Event handler for pressing "OK"

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrTrap
    If (SelectedChemistry = "") Then
        Call MsgBox("Please select a chemistry.", vbOKOnly Or vbExclamation, "Selection Required")
    Else
        blnCancel_m = False  ' let other objects know user pressed ok
        Call Me.Hide
    End If
    
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | ChemistrySelect.cmdOK_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   LoadChems()

'DESCRIPTION: Public interface to set up the listbox to select the chemistry

'PARAMETERS:  inChems - the chemistries collection

'RETURNED:    N/A

'*********************************************************************
Public Sub LoadChems(inChems As QalibCollection)
On Error GoTo ErrTrap
    Dim curChem As Chemistry
    
    For Each curChem In inChems
        Call lstChemistrySelect.AddItem(curChem.Name)
    Next curChem
    
    ' set focus to first chemistry in the list
    lstChemistrySelect.ListIndex = 0
    
    Exit Sub
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | ChemistrySelect.LoadChems", Err.Description)
End Sub

