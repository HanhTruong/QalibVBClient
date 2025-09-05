VERSION 5.00
Begin VB.Form Comment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Annotate Calibration"
   ClientHeight    =   3195
   ClientLeft      =   4275
   ClientTop       =   3810
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox eboComment 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Comment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
 
'FILE:  Comment.frm
 
'DESCRIPTION:  This module contains the form where the user can enter a comment about
' the calibration

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: Comment.frm $
 ' 
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:39a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:23p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added error traps.
 ' Underlying lot data is now object-oriented.
 ' Set object references equal to nothing to free resources.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 5/09/03    Time: 3:49p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Lot object is set through function SetDataSource instead of beaing set
 ' by a property.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 3/21/03    Time: 2:40p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe.

Option Explicit

' private member variables
Private objLotReport_m As LotReport  ' stores a reference to the underlying lot report object
Private blnCancel_m As Boolean  ' stores whether the user accepts the comment

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
    Call Err.Raise(Err.Number, Err.Source & " | Comment.PropertyGet.Cancel", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Form_Terminate()

'DESCRIPTION: Event handler for when the form terminates

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Form_Terminate()
On Error GoTo ErrTrap
    Set objLotReport_m = Nothing
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | Comment.Form_Terminate", Err.Description)
End Sub

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
    Call HandleError(Err.Number, Err.Source & " | Comment.Form_QueryUnload", Err.Description)
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
    Call Me.Hide
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | Comment.cmdCancel_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdOK_Click()

'DESCRIPTION: Commits the comment and hides the form

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrTrap
    Call objLotReport_m.UpdateComment(eboComment.Text)   ' commit the comment
    blnCancel_m = False
    Call Me.Hide
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | Comment.cmdOK_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   LoadComment()

'DESCRIPTION: Allows other objects to set a reference to the underlying lot report object

'PARAMETERS:  inLotReport - reference to the underlying lot report object

'RETURNED:    N/A

'*********************************************************************
Public Sub LoadComment(inLotReport As LotReport)
On Error GoTo ErrTrap
    Set objLotReport_m = inLotReport
    eboComment.Text = inLotReport.Comment
    eboComment.SelStart = 0
    eboComment.SelLength = Len(eboComment.Text)
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Comment.LoadComment", Err.Description)
End Sub

