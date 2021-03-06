VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LongStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
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

Private Const ModuleName                            As String = "LongStack"

'@================================================================================
' Member variables
'@================================================================================

Private mStack()                                    As Long
Private mStackIndex                                 As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mStack(15) As Long
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

'@================================================================================
' Methods
'@================================================================================

Public Sub Push(ByVal Value As Long)
If mStackIndex > UBound(mStack) Then ReDim Preserve mStack(2 * (UBound(mStack) + 1) - 1) As Long
mStack(mStackIndex) = Value
mStackIndex = mStackIndex + 1
End Sub

Public Function Pop() As Long
Assert mStackIndex <> 0, "Stack is empty"
mStackIndex = mStackIndex - 1
Pop = mStack(mStackIndex)
End Function

'@================================================================================
' Helper Functions
'@================================================================================




