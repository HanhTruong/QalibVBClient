VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{4881A3EC-DC21-11D4-8235-0010A4C42ABD}#32.32#0"; "ExtLVCTL.ocx"
Begin VB.Form SeriesEdit 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3795
   ClientLeft      =   4725
   ClientTop       =   2505
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   9750
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   8400
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraListView 
      Caption         =   "Points:"
      Height          =   3615
      Left            =   120
      TabIndex        =   3
      Top             =   60
      Width           =   8055
      Begin ExtLVCTL.ExtLV lvwPoints 
         Height          =   3255
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Click a point to edit"
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   5741
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
         HideSelection   =   0   'False
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
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   8400
      TabIndex        =   2
      ToolTipText     =   "Cancel changes"
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   8400
      TabIndex        =   1
      ToolTipText     =   "Save changes"
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "SeriesEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'**********************************************************************************
 
'FILE:  SeriesEdit.frm
 
'DESCRIPTION:  This module contains the form where the user can move/delete data points
'from pools.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: SeriesEdit.frm $
 ' 
 ' *****************  Version 8  *****************
 ' User: Ballard      Date: 6/04/04    Time: 9:32a
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Fixed tab index.
 '
 ' *****************  Version 7  *****************
 ' User: Ballard      Date: 3/23/04    Time: 4:39p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Updated to 1.0.0X9
 '
 ' *****************  Version 6  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:37p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Added error traps.
 ' Set object references equal to nothing to free resources.
 ' Now use an Extended ListView control with embedded check boxes to moves
 ' points between series and/or exclude them.
 ' The logic to update the Extended ListView is cleaner now since it just
 ' has to read the point attributes from the underlying point collection
 ' to populate the control.
 '
 ' *****************  Version 5  *****************
 ' User: Ballard      Date: 5/09/03    Time: 3:55p
 ' Updated in $/QalibVBClient/Source/QalibGUI
 ' Lot object is set through function LoadSeries instead of being set by a
 ' property.  Extended list view (dropdown list appears over list view)
 ' was added to make editing fit parameters easier.
 '
 ' *****************  Version 4  *****************
 ' User: Ballard      Date: 3/21/03    Time: 2:38p
 ' Updated in $/QalibVBClient
 ' Renamed module "frmChangeData.frm" to "SeriesEdit.frm."
 '
 ' Moved the functionality to display and manipulate the series back to
 ' the form.
 '
 ' The form retrieves data from the underlying data object.
 '
 ' *****************  Version 3  *****************
 ' User: Ballard      Date: 1/24/03    Time: 2:53p
 ' Updated in $/QalibVBClient
 '
 ' *****************  Version 2  *****************
 ' User: Ballard      Date: 1/09/03    Time: 3:51p
 ' Updated in $/QalibVBClient
 ' Added file and function headers.
 ' Added custom control to handle editing series.
 ' Removed all code to process series edits since it's now done in the
 ' calibration object.
 '
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 9/24/02    Time: 1:46p
 ' Created in $/QalibVBClient
 ' Added to SourceSafe

Option Explicit

' private constants

Private Const MEASUREDVALCOL As Integer = 0 ' measured value column in the list view
Private Const CURSERIESNAMECOL As Integer = 1 ' current series name column in the list view
Private Const ASSIGNEDVALUECOL As Integer = 2  ' assigned value column in the list view
Private Const EXCLUDECOL As Integer = 3 ' ' exclude column
Private Const ORIGSERIESNAMECOL As Integer = 4 ' original series name column in the list view

Private Const CHECKOFFINDEX As Integer = 1 ' the index of the check off icon in the image list
Private Const CHECKONINDEX As Integer = 2 ' the index of the check on icon in the image list

Private Const IMAGEDIM As Integer = 16 ' the height and width of the checkbox icons

' private member variables
Private blnCancel_m As Boolean ' stores whether cancel button was pressed
Private objAnalyzerData_m As AnalyzerData   ' stores a reference to the underlying analyzer data object
Private strSeriesName_m As String   ' stores the series name being edited
Private strExpPointSetName_m As String   ' stores the point set name being edited

'***********************************************************************

'PROPERTY GET:   Cancel()

'DESCRIPTION: Allows other objects to see if cancel button was pressed

'PARAMETERS:  N/A

'RETURNED:    Whether the cancel button was pressed

'*********************************************************************
Public Property Get Cancel() As Boolean
On Error GoTo ErrTrap
    Cancel = blnCancel_m
    Exit Sub
ErrTrap:
    Call Err.Raise(APPERR, Err.Source & " | SeriesEdit.PropertyGet.Cancel", Err.Description)
End Property

'***********************************************************************

'PROCEDURE:   Form_Terminate()

'DESCRIPTION: Event handler for when the form terminates

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Form_Terminate()
On Error GoTo ErrTrap
    Set objAnalyzerData_m = Nothing
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | SeriesEdit.Form_Terminate", Err.Description)
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
    Call HandleError(Err.Number, Err.Source & " | SeriesEdit.Form_QueryUnload", Err.Description)
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
    Call HandleError(Err.Number, Err.Source & " | SeriesEdit.cmdCancel_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   cmdOK_Click()

'DESCRIPTION: Hides the form and forwards the call to process the series changes

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub cmdOK_Click()
On Error GoTo ErrTrap
    Dim curItem As MSComctlLib.ListItem
    Dim curExpPointSet As ExpPointSet
    Dim curExpPoint As ExpPoint
    
    blnCancel_m = False  ' let other objects know user pressed ok
    
    ' get the point set
    Set curExpPointSet = objAnalyzerData_m.ExpPointSets(strExpPointSetName_m)
    
    With lvwPoints
    
        For Each curItem In .ListItems
            ' get the point
            Set curExpPoint = curExpPointSet.ExpPoints(curItem.Tag)
            
             ' if the excluded status has been toggled then update the point
            If (((curExpPoint.IsExcluded = True) And (curItem.ListSubItems(EXCLUDECOL).ReportIcon = CHECKOFFINDEX)) Or _
                ((curExpPoint.IsExcluded = False) And (curItem.ListSubItems(EXCLUDECOL).ReportIcon = CHECKONINDEX))) Then
                Call objAnalyzerData_m.ExcludeRun(CStr(curItem.Tag), IIf(curItem.ListSubItems(EXCLUDECOL).ReportIcon = CHECKONINDEX, True, False))
            End If
            
            ' if the series has changed then change the series
            If (curItem.SubItems(CURSERIESNAMECOL) <> strSeriesName_m) Then
                Call objAnalyzerData_m.ChangeRunSeries(CLng(curItem.Tag), curItem.SubItems(CURSERIESNAMECOL))
            End If
                        
        Next curItem
    End With
    
    Call Me.Hide
    
    Set curItem = Nothing
    Set curExpPointSet = Nothing
    Set curExpPoint = Nothing

    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | SeriesEdit.cmdOK_Click", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   lvwPoints_DropDownChangeComplete()

'DESCRIPTION: Updates the point's series in the listview when it is changed.
' The tag for the point's row must be updated with the new series!

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub lvwPoints_DropDownChangeComplete(ByVal OrigText As String, ByVal NewText As String)
On Error GoTo ErrTrap
    Dim newSeries As SeriesOne
    
    With lvwPoints
        ' get the new series
        Set newSeries = objAnalyzerData_m.SeriesSet(NewText)
        
        .SelectedItem.SubItems(CURSERIESNAMECOL) = newSeries.Name ' set the new series name
        .SelectedItem.SubItems(ASSIGNEDVALUECOL) = CStr(newSeries.AssignedVal) ' set the new series assigned value
        .SelectedItem.ListSubItems(EXCLUDECOL).ReportIcon = CHECKOFFINDEX ' a point is automatically restored when its series is changed
                
        Call .SizeColumns
    End With
    
    Set newSeries = Nothing
    Exit Sub
ErrTrap:
    Set newSeries = Nothing
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | SeriesEdit.lvwPoints_DropDownChangeComplete", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   lvwPoints_ItemClick()

'DESCRIPTION: Allows the user to select a series form the drop-down list

'PARAMETERS:  item - which item in the list was selected
'             Button - which mouse button initiated the click
'             SecondClick - whether it's the second time the item was clicked

'RETURNED:    N/A

'*********************************************************************
Private Sub lvwPoints_ItemClick(ByVal item As MSComctlLib.IListItem, ByVal Button As Integer, ByVal SecondClick As Boolean)
On Error GoTo ErrTrap
    
    With lvwPoints
        If (.SubItemClicked = CURSERIESNAMECOL) Then
            ' show the series drop down list
            Call .ShowDropDown(item.SubItems(CURSERIESNAMECOL))
        ElseIf (.SubItemClicked = EXCLUDECOL) Then
            ' toggle the exclueded check mark
            item.ListSubItems(EXCLUDECOL).ReportIcon = IIf(item.ListSubItems(EXCLUDECOL).ReportIcon = CHECKOFFINDEX, _
                CHECKONINDEX, CHECKOFFINDEX)
        End If
    End With
    
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | SeriesEdit.lvwPoints_ItemClick", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   lvwPoints_NotInList()

'DESCRIPTION: Event handler for when user types in the drop-down list.  Do not accept
' any new entry; it must be in the list

'PARAMETERS:  OldValue - previous text
'             NewValue - new text
'             Cancel - whether to cancel the operation

'RETURNED:    N/A

'*********************************************************************
Private Sub lvwPoints_NotInList(ByVal OldValue As String, NewValue As String, Cancel As Boolean)
On Error GoTo ErrTrap
    Cancel = True ' do not allow any new text
    Exit Sub
ErrTrap:
    ' can't raise an error here because there's nothing up the call stack to trap it
    Call HandleError(Err.Number, Err.Source & " | SeriesEdit.lvwPoints_NotInList", Err.Description)
End Sub

'***********************************************************************

'PROCEDURE:   LoadSeries()

'DESCRIPTION: Public interface to set up the controls with the series being edited

'PARAMETERS:  inAnalyzerData- the analyzer data
'             inSeriesName - the series to edit
'             inExpPointSetName - the point set name
'             inExpPointIndex - the selected point

'RETURNED:    N/A

'*********************************************************************
Public Sub LoadSeries(inAnalyzerData As AnalyzerData, inSeriesName As String, inExpPointSetName As String, inExpPointIndex As Long)
On Error GoTo ErrTrap
    Dim item As ListItem
    Dim insertIndex As Long
    Dim curSeries As SeriesOne
    Dim curExpPoint As ExpPoint
    Dim curExpPointSet As ExpPointSet
    Dim curRun As Run
    
    ' set up the module variables
    Set objAnalyzerData_m = inAnalyzerData
    strSeriesName_m = inSeriesName
    strExpPointSetName_m = inExpPointSetName
        
    Me.Caption = strExpPointSetName_m & " | " & strSeriesName_m & " Edit"
        
    ' load the check on/off icons
    imgIcons.ImageHeight = IMAGEDIM
    imgIcons.ImageWidth = IMAGEDIM
    
    Call imgIcons.ListImages.Add(CHECKOFFINDEX, , LoadResPicture(CHECKOFFICON, vbResIcon))
    Call imgIcons.ListImages.Add(CHECKONINDEX, , LoadResPicture(CHECKONICON, vbResIcon))
     
    With lvwPoints
    
        ' clear out any data in the controls
        .Clear
        .ListItems.Clear
        
        ' set up the columns in the list view
        Call .ColumnHeaders.Add(, , "Measured Value")
        Call .ColumnHeaders.Add(, , "Current Series")
        Call .ColumnHeaders.Add(, , "Assigned Value")
        Call .ColumnHeaders.Add(, , "Exclude")
        Call .ColumnHeaders.Add(, , "Original Series")
        
        ' connect the icons to the listview
        Set .SmallIcons = imgIcons
        
        ' get the series
        Set curSeries = objAnalyzerData_m.SeriesSet(inSeriesName)
       
        ' cycle through all the series' runs
        ' the points in the the list view are 1-based
        For Each curRun In curSeries.Runs
            
            ' get the point set
            Set curExpPointSet = objAnalyzerData_m.ExpPointSets(strExpPointSetName_m)
            
            ' get the point out of the point set
            Set curExpPoint = curExpPointSet.ExpPoints(curRun.ID)
            
            ' default the insert index to the first position
            insertIndex = 1
            
            ' the points must be listed from lowest to highest
            For Each item In .ListItems
                If (curExpPoint.MeasuredVal > CDbl(item.Text)) Then
                    insertIndex = item.Index + 1
                Else
                    Exit For
                End If
            Next item
            
            ' add a row to the listview box
            Set item = .ListItems.Add(insertIndex)
            item.Tag = curExpPoint.ID ' store the ID of the point
            
            ' the 5 columns of the listview are the measured value, the current series title, the assigned value, the exclusion status, and original series title
            item.Text = CStr(curExpPoint.MeasuredVal)
            item.SubItems(CURSERIESNAMECOL) = strSeriesName_m
            item.SubItems(ASSIGNEDVALUECOL) = CStr(curSeries.AssignedVal)
            item.SubItems(EXCLUDECOL) = ""
            item.ListSubItems(EXCLUDECOL).ReportIcon = IIf(curExpPoint.IsExcluded = False, CHECKOFFINDEX, CHECKONINDEX)
            item.SubItems(ORIGSERIESNAMECOL) = curRun.OrigSeries
            
        Next curRun
                
        ' put the series' names in the combo box
        For Each curSeries In objAnalyzerData_m.SeriesSet
            Call .AddItem(curSeries.Name)
        Next curSeries
                
        ' set the focus to the selected point, the listview is 1-based
        .ListItems(inExpPointIndex + 1).Selected = True
        
        ' size the columns automatically
        Call .SizeColumns
        
    End With
    
    Set item = Nothing
    Set curSeries = Nothing
    Set curExpPoint = Nothing
    Set curExpPointSet = Nothing
    Set curRun = Nothing
    Exit Sub
ErrTrap:
    Set item = Nothing
    Set curSeries = Nothing
    Set curExpPoint = Nothing
    Set curExpPointSet = Nothing
    Set curRun = Nothing
    Call Err.Raise(APPERR, Err.Source & " | SeriesEdit.LoadSeries", Err.Description)
End Sub
