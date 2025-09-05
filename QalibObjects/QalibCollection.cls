VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QalibCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************************
 
'FILE:  QalibCollection.cls
 
'DESCRIPTION:  This module holds the class which contains the Qalib collection.

'COMPILER:  This module is part of a project that is designed to be edited and compiled
'in Visual Basic 6.0.  Choose "File->Make" from within the IDE to make the program.

'$History: QalibCollection.cls $
 ' 
 ' *****************  Version 1  *****************
 ' User: Ballard      Date: 7/25/03    Time: 3:48p
 ' Created in $/QalibVBClient/Source/QalibObjects
 ' Added to SourceSafe

Option Explicit

Private colQalib_m As Collection  ' the Qalib collection

'***********************************************************************

'PROCEDURE:   Class_Initialize()

'DESCRIPTION: Sets up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Initialize()
    Set colQalib_m = New Collection
End Sub

'***********************************************************************

'PROCEDURE:   Class_Terminate()

'DESCRIPTION: Cleans up the object

'PARAMETERS:  N/A

'RETURNED:    N/A

'*********************************************************************
Private Sub Class_Terminate()
    Set colQalib_m = Nothing
End Sub

'***********************************************************************

'FUNCTION:   Add()

'DESCRIPTION: Allows friend objects to add an object to the collection

'PARAMETERS:  inObject - the object to add
'             inKey - the key to use
'
'RETURNED:  N/A

'********************************************************************
Friend Function Add(inObject As Object, inKey As Variant)
On Error GoTo ErrTrap
    Call colQalib_m.Add(inObject, CStr(inKey))
    Exit Function
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | QalibCollection.Add", Err.Description)
End Function

'***********************************************************************

'PROCEDURE:   Remove()

'DESCRIPTION: Allows friend objects to remove an object from the collection

'PARAMETERS:  index - the key of the object to remove
'
'RETURNED:  N/A

'********************************************************************
Friend Sub Remove(inKey As Variant)
On Error Resume Next
    Call colQalib_m.Remove(CStr(inKey))
    Exit Sub
End Sub

'***********************************************************************

'FUNCTION:   NewEnum()

'DESCRIPTION: Allows other objects to iterate over the objects

'PARAMETERS:  N/A
'
'RETURNED:  IUnknown interface for iterating over the objects

'********************************************************************
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
On Error GoTo ErrTrap
    Set NewEnum = colQalib_m.[_NewEnum]
    Exit Function
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | QalibCollection.NewEnum", Err.Description)
End Function

'***********************************************************************

'FUNCTION:   Item()

'DESCRIPTION: Returns the object for the given key

'PARAMETERS:  inKey - the key of the object
'
'RETURNED:  the object

'********************************************************************
Public Function Item(inKey As Variant) As Object
Attribute Item.VB_UserMemId = 0
On Error Resume Next
    Set Item = colQalib_m(CStr(inKey))
        
    If (Err.Number <> 0) Then
        Set Item = Nothing
    End If
    
End Function

'***********************************************************************

'FUNCTION:   Count()

'DESCRIPTION: Returns the number of objects in the collection

'PARAMETERS:  N/A
'
'RETURNED:  Number of objects in the collection

'********************************************************************
Public Function Count() As Long
On Error GoTo ErrTrap
    Count = colQalib_m.Count
    Exit Function
ErrTrap:
    Call Err.Raise(Err.Number, Err.Source & " | QalibCollection.Count", Err.Description)
End Function
