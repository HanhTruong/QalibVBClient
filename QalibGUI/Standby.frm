VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Standby 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Standby"
   ClientHeight    =   1335
   ClientLeft      =   5610
   ClientTop       =   3810
   ClientWidth     =   5310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   Tag             =   "Standby"
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblMessage 
      Caption         =   "Processing ..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Standby"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
 
'FILE:  Standby.frm
 
'DESCRIPTION:  This module contains the form with the progress bar to indicate to the user
' that processing is occurring

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: Standby.frm $
 ' 
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:20p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added error traps.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 5/09/03    Time: 3:56p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added message to standby form.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 3/21/03    Time: 2:40p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe.

Option Explicit

'***********************************************************************

'PROCEDURE:   Update()

'DESCRIPTION: Updates the progrss bar and message text

'PARAMETERS:  inPercent - update the progress bar with this value
'             inMessage - the text for the progress

'RETURNED:    N/A

'*********************************************************************
Public Sub Update(inPercent As Integer, inMessage As String)
On Error GoTo ErrTrap
    prgBar.Value = inPercent
    lblMessage.Caption = inMessage
    DoEvents
    Exit Sub
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | Standby.UpdateStatus", Err.Description)
End Sub
