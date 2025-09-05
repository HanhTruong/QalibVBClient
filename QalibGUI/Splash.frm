VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2685
   ClientLeft      =   5385
   ClientTop       =   3300
   ClientWidth     =   5235
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraMain 
      Height          =   2475
      Left            =   165
      TabIndex        =   0
      Top             =   120
      Width           =   4905
      Begin VB.Image imgLogo 
         Height          =   1065
         Index           =   1
         Left            =   3600
         Picture         =   "Splash.frx":000C
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1095
      End
      Begin VB.Image imgLogo 
         Height          =   1065
         Index           =   0
         Left            =   240
         Picture         =   "Splash.frx":0316
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblCopyright 
         Caption         =   "Abaxis, 2002-2004"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   1
         Top             =   1680
         Width           =   1440
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1.0.0X13"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1950
         TabIndex        =   2
         Top             =   1200
         Width           =   960
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Qalib"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   1695
         TabIndex        =   3
         Top             =   240
         Width           =   1605
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************

'FILE:  Splash.frm
 
'DESCRIPTION:  This module contains the form where the splash screen is displayed.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: Splash.frm $
 '
 ' *****************  Version 7  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:29a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 '
 ' *****************  Version 6  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:39p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Updated to 1.0.0X9
 '
 ' *****************  Version 5  *****************
 ' User: Ballard      Date: 3/21/03    Time: 1:30p
 ' Updated in $/QalibVBClient
 ' Changed name of module from "frmSplash.frm" to "Splash.frm."
 '
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 1/24/03    Time: 2:49p
 ' Updated in $/QalibVBClient
 ' Updated version number.
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 1/17/03    Time: 4:14p
 ' Updated in $/QalibVBClient
 ' Changed version number.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 1/09/03    Time: 3:49p
 ' Updated in $/QalibVBClient
 ' Added file header.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 9/24/02    Time: 1:46p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe

