VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataObjectSummaries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
' @remarks
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName As String = "DataObjectSummaries"

'@================================================================================
' Member variables
'@================================================================================

Private mSummaries As EnumerableCollection

Private mSpecifiers As FieldSpecifiers

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSummaries = New EnumerableCollection
Set mSpecifiers = New FieldSpecifiers
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Enumerator() As Enumerator
Const ProcName As String = "Enumerator"
Dim failpoint As String
On Error GoTo Err

Set Enumerator = mSummaries.Enumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FieldSpecifiers() As FieldSpecifiers
Set FieldSpecifiers = mSpecifiers
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add( _
                pItem As DataObjectSummary)
Const ProcName As String = "Add"
Dim failpoint As String
On Error GoTo Err

mSummaries.Add pItem

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub AddFieldDetails( _
                ByRef specifier As FieldSpecifier)
Const ProcName As String = "AddFieldDetails"
Dim failpoint As String
On Error GoTo Err

mSpecifiers.AddFieldSpecifier specifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Count() As Long
Const ProcName As String = "Count"
Dim failpoint As String
On Error GoTo Err

Count = mSummaries.Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item( _
                ByVal index As Long) As DataObjectSummary
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Set Item = mSummaries(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Remove( _
                ByVal index As Long)
Const ProcName As String = "Remove"
Dim failpoint As String
On Error GoTo Err

mSummaries.Remove (index)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
Dim failpoint As String
On Error GoTo Err

Set NewEnum = mSummaries.Enumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================












