VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DOSummaryBuilder"
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

Private Const ModuleName                As String = "DOSummaryBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mDataObjectSummary              As DataObjectSummary

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mDataObjectSummary = New DataObjectSummary
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

Public Property Get DataObjectSummary() As DataObjectSummary
Set DataObjectSummary = mDataObjectSummary
End Property

Public Property Let FieldValue( _
                ByVal fieldName As String, _
                ByVal Value As String)
Const ProcName As String = "FieldValue"
Dim failpoint As String
On Error GoTo Err

mDataObjectSummary.FieldValue(fieldName) = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Id(ByVal Value As Long)
Const ProcName As String = "Id"
Dim failpoint As String
On Error GoTo Err

mDataObjectSummary.Id = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================









