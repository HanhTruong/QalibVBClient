VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1545
   ClientLeft      =   5385
   ClientTop       =   5115
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox eboUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   3
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   1
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   2
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox eboPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   4
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   5
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
 
'FILE:  Login.frm
 
'DESCRIPTION:  This module contains the form where the user logs in

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: Login.frm $
 ' 
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:33a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Form now uses authentication object for login.
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

' private member variables
Private blnCancel_m As Boolean  ' stores whether cancel button was pressed
Private objUserVerify_m As UserVerifiy ' for verfiying the user

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
    Call Err.Raise(Err.Number, Err.Source & " | Login.PropertyGet.Cancel", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   User()

'DESCRIPTION: Allows other objects to get the user

'PARAMETERS:  N/A

'RETURNED:    the user

'*********************************************************************
Public Property Get User() As String
On Error GoTo ErrTrap
    User = eboUserName.Text
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Login.PropertyGet.User", Err.Description)
End Property

'***********************************************************************

'PROPERTY GET:   Password()

'DESCRIPTION: Allows other objects to get the password

'PARAMETERS:  N/A

'RETURNED:    the password

'*********************************************************************
Public Property Get Password() As String
On Error GoTo ErrTrap
    Password = eboPassword.Text
    Exit Property
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Login.PropertyGet.Password", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Form_Terminate()

'DESCRIPTION: Event handler for when the form terminates

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Form_Terminate()
On Error GoTo ErrTrap
    Set objUserVerify_m = Nothing
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | Login.Form_Terminate", Err.Description)
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
    Call HandleError(Err.Number, Err.Source & " | Login.Form_QueryUnload", Err.Description)
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
    Call HandleError(Err.Number, Err.Source & " | Login.cmdCancel_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdOK_Click()

'DESCRIPTION: Processes the login

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrTrap
    blnCancel_m = False
    
    ' authenticate the user with the user verification object
    Call objUserVerify_m.Authenticate(eboUserName.Text, eboPassword.Text)
    
    ' save the last login to the registry
    Call SaveSetting(App.Title, SECTIONKEY, LASTLOGINKEY, eboUserName.Text)

    Me.Hide
    
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | Login.cmdOK_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   LoadUserVerify()

'DESCRIPTION: Allows other objects to load the form

'PARAMETERS:  inUserVerify - the user verification object

'RETURNED:    N/A

'*********************************************************************
Public Sub LoadUserVerify(inUserVerify As UserVerifiy)
On Error GoTo ErrTrap
    Dim lastLogin As String

    Set objUserVerify_m = inUserVerify
    
    ' get the last login name from the registry
    lastLogin = GetSetting(App.Title, SECTIONKEY, LASTLOGINKEY, "")
    
    If (lastLogin = "") Then
        ' their was no last login so set the focus to the username box
        eboUserName.TabIndex = 0
    Else
        ' put the last login in the userbox and set the focus to the password box
        eboUserName.Text = lastLogin
        eboPassword.TabIndex = 0
    End If
    
    Exit Sub
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | Login.LoadUserVerify", Err.Description)
End Sub
