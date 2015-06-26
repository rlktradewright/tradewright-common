VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOSummariesBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName As String = "DOSummariesBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mDataObjectSummaries As DataObjectSummaries

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mDataObjectSummaries = New DataObjectSummaries
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

Public Property Get DataObjectSummaries() As DataObjectSummaries
Set DataObjectSummaries = mDataObjectSummaries
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add( _
                pItem As DataObjectSummary)
Const ProcName As String = "Add"
Dim failpoint As String
On Error GoTo Err

mDataObjectSummaries.Add pItem

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub AddFieldDetails( _
                ByRef specifier As FieldSpecifier)
Const ProcName As String = "AddFieldDetails"
Dim failpoint As String
On Error GoTo Err

mDataObjectSummaries.AddFieldDetails specifier

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Remove(index As Long)
Const ProcName As String = "Remove"
Dim failpoint As String
On Error GoTo Err

mDataObjectSummaries.Remove (index)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================














