VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataObjectSummary"
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

Private Const ModuleName As String = "DataObjectSummary"

'@================================================================================
' Member variables
'@================================================================================

Private mIdent As Long

Private mFields As Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFields = New Collection
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

Friend Property Let FieldValue( _
                ByVal fieldName As String, _
                ByVal Value As String)
Const ProcName As String = "FieldValue"
Dim failpoint As String
On Error GoTo Err

mFields.Add Value, fieldName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FieldValue( _
                ByVal fieldName As String) As String
Const ProcName As String = "FieldValue"
Dim failpoint As String
On Error GoTo Err

FieldValue = mFields(fieldName)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Id(ByVal Value As Long)
mIdent = Value
End Property

Public Property Get Id() As Long
Id = mIdent
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================










