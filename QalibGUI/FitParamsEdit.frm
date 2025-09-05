VERSION 5.00
Object = "{4881A3EC-DC21-11D4-8235-0010A4C42ABD}#32.32#0"; "ExtLVCTL.ocx"
Begin VB.Form FitParamsEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Fit Parameters"
   ClientHeight    =   2385
   ClientLeft      =   4500
   ClientTop       =   2505
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   6915
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraFitParams 
      Caption         =   "Fit Parameters"
      Height          =   2295
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   5295
      Begin ExtLVCTL.ExtLV lvwFitParams 
         Height          =   1935
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Click a fit parameter to edit"
         Top             =   240
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   3413
         LineEvenColor   =   0
         LineOddColor    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Verdana"
         FontSize        =   8.25
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         View            =   3
         ListIndex       =   -1
         CalendarTrailingForeColor=   -2147483631
         CalendarTitleForeColor=   -2147483630
         CalendarTitleBackColor=   -2147483633
         CalendarForeColor=   -2147483630
         CalendarBackColor=   -2147483643
         TitleHeight     =   255
         PlaySounds      =   0   'False
         DropWidth       =   0
         DropLines       =   0
         DropDelay       =   0
         TabCaptions     =   ""
         SortArrowSize   =   0
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "FitParamsEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************
 
'FILE:  FitParamsEdit.frm
 
'DESCRIPTION:  This module contains the form where the user can edit the fit parameters

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: FitParamsEdit.frm $
 ' 
 ' *****************  Version 5  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:40a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Fixed tab index.
 '
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:41p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added error traps.
 ' Set object references equal to nothing to free resources.
 ' Now use Extended ListView to display fit parameters and allow editing
 ' in place.
 ' Logic to populate Extended ListView is cleaner since the underlying fit
 ' parameters collection object exposes properties for each for parameter.
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 5/09/03    Time: 3:54p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Lot object is set through function LoadFitParams instead of being set
 ' by a property.  Extended list view (text box appears over list view)
 ' was added to make editing fit parameters easier.
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 3/21/03    Time: 2:26p
 ' Updated in $/QalibVBClient
 ' Renamed module "frmAdjustFitParams.frm" to "FitParamsEdit.frm."
 '
 ' Moved the functionality to display and manipulate the fit parameters
 ' back to the form.
 '
 ' The form retrieves data from the underlying data object.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 1/24/03    Time: 2:47p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe.

Option Explicit

' private constants
Private Const CURVALUECOL As Integer = 1 ' current value column in the list view

' private member variables
Private blnCancel_m As Boolean ' stores whether cancel button was pressed
Private objFitParams_m As FitParams  ' stores a reference to the fit parameters object

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
    Call Err.Raise(Err.Number, Err.Source & " | FitParamsEdit.PropertyGet.Cancel", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Form_Terminate()

'DESCRIPTION: Event handler for when the form terminates

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Form_Terminate()
On Error GoTo ErrTrap
    Set objFitParams_m = Nothing
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | FitParamsEdit.Form_Terminate", Err.Description)
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
    Call HandleError(Err.Number, Err.Source & " | FitParamsEdit.Form_QueryUnload", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdCancel_Click()

'DESCRIPTION: Hides the form and sets cancel property

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
    Call HandleError(Err.Number, Err.Source & " | FitParamsEdit.cmdCancel_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdOK_Click()

'DESCRIPTION: Hides the form and processes the edited fit parameters

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrTrap
    
    ' commit the changed fit parameters
    Call objFitParams_m.CommitEdits
    blnCancel_m = False
    Call Me.Hide
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | FitParamsEdit.cmdOK_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   lvwFitParams_ItemClick()

'DESCRIPTION: Event handler for when user clicks a row in the list view

'PARAMETERS:  item - which item in the list was selected
'             Button - which mouse button initiated the click
'             SecondClick - whether it's the second time the item was clicked

'RETURNED:    N/A

'*********************************************************************
Private Sub lvwFitParams_ItemClick(ByVal item As MSComctlLib.IListItem, ByVal Button As Integer, ByVal SecondClick As Boolean)
On Error GoTo ErrTrap
    
    With lvwFitParams
        ' see if user clicked new fit parameter column
        If (.SubItemClicked = CURVALUECOL) Then
            Call .ShowTextBox(item.SubItems(CURVALUECOL))  ' allow user to edit fit parameter
        End If
    End With
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | FitParamsEdit.lvwFitParams_ItemClick", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   lvwFitParams_ChangeComplete()

'DESCRIPTION: Event handler for changing the text box value

'PARAMETERS:  OrigText - previous text box value
'             NewText - current text box value

'RETURNED:    N/A

'*********************************************************************
Private Sub lvwFitParams_ChangeComplete(ByVal OrigText As String, ByVal NewText As String)
On Error GoTo ErrTrap
    Dim curFitParam As FitParam
    
    With lvwFitParams
        ' make sure user typed something
        If (Trim$(NewText) = "") Then
            Set curFitParam = Nothing
            Exit Sub
        End If
        
        ' validate the new fit parameter
        If (IsNumeric(NewText) = False) Then
            ' warn user that the entry is invalid and force them to fix it
            Call MsgBox("Please enter a valid number.", _
                vbExclamation Or vbOKOnly, "Invalid Entry")
            Call .SetFocusToTextBox
            .TextBoxControl.SelStart = 0
            .TextBoxControl.SelLength = Len(.Text)
        Else
            Set curFitParam = objFitParams_m.FitParamsSet(.SelectedItem.Text)
            ' entry is okay so update list view
            curFitParam.Value = CDbl(NewText)
            Call LoadLine(.SelectedItem, curFitParam)
            
        End If
    End With
    
    Set curFitParam = Nothing
    Exit Sub
ErrTrap:
    Set curFitParam = Nothing
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | FitParamsEdit.lvwFitParams_ChangeComplete", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   LoadLine()

'DESCRIPTION: Assigns the contents of a fit param object to the columns in the listview

'PARAMETERS:  inItem - the listview line to load
'             inFitParamt - the fit param object to load

'RETURNED:    N/A

'*********************************************************************
Private Sub LoadLine(inItem As ListItem, inFitParam As FitParam)
On Error GoTo ErrTrap
    inItem.Text = inFitParam.Name
    inItem.SubItems(CURVALUECOL) = CStr(inFitParam.Value)
Exit Sub
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | FitParamsEdit.LoadLine", Err.Description)
End Sub


'***********************************************************************

'PROCEDURE:   LoadFitParams()

'DESCRIPTION: Sets up the list view box with the fit parameters

'PARAMETERS:  inFitParams - the fit parameters object

'RETURNED:    N/A

'*********************************************************************
Public Sub LoadFitParams(inFitParams As FitParams)
On Error GoTo ErrTrap
    Dim item As ListItem
    Dim curFitParam As FitParam
    
    Set objFitParams_m = inFitParams

    ' set up the listview columns
    With lvwFitParams
        Call .ColumnHeaders.Add(, , "Description")
        Call .ColumnHeaders.Add(, , "Value")
    End With

    With lvwFitParams
        ' clear out any data in the control
        .ListItems.Clear
        .Text = ""
        
        ' cycle through all the fit parameters
        For Each curFitParam In objFitParams_m.FitParamsSet
        
            Set item = .ListItems.Add()  ' display the description for the fit parameter

            Call LoadLine(item, curFitParam)

        Next curFitParam
        
        ' size the columns automatically
        Call .SizeColumns
    End With
    
    Set curFitParam = Nothing
    Set item = Nothing
    Exit Sub
ErrTrap:
    Set curFitParam = Nothing
    Set item = Nothing
    Call Err.Raise(Err.Number, Err.Source & " | FitParamsEdit.LoadFitParams", Err.Description)
End Sub

