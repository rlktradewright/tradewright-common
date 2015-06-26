VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFuture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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

Private Const ModuleName                            As String = "IFuture"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Cookie() As Variant
End Property

Public Property Get ErrorNumber() As Long
End Property

Public Property Get ErrorMessage() As String
End Property

Public Property Get ErrorSource() As String
End Property

Public Property Get IsAvailable() As Boolean
End Property

Public Property Get IsCancelled() As Boolean
End Property

Public Property Get IsFaulted() As Boolean
End Property

Public Property Get IsPending() As Boolean
End Property

Public Property Get Value() As Object
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddStateChangeListener(ByVal pListener As StateChangeListener)
End Sub

Public Sub Cancel()
End Sub

Public Sub RemoveStateChangeListener(ByVal pListener As StateChangeListener)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





