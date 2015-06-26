VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusinessObjEventSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' An object in this class is used as a source of events for a <code>BusinessDataObject</code>
' object. This is necessary because interface definitions cannot include events, so the
' <code>BusinessDataObject</code> interface includes an <code>EventSource</code>
' property that returns one of these objects.
'
' @see
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

''
' Raised when a change is made to a <code>BusinessDataObject</code> object.
'
'
'@/
Event Dirty()

''
' Raised when a changes to a <code>BusinessDataObject</code> object are committed to
' permanent storage.
'
'
'@/
Event Clean()

''
' Raised when a <code>BusinessDataObject</code> object is deleted from permanent storage.
'
'
'@/
Event Deleted()

''
' Raised when a change to a <code>BusinessDataObject</code> object puts it in a state that
' does not comply with one or more relevant business rules.
' <p>Note that this event can be fired when a <code>BusinessDataObject</code> object is
' initially created.</p>
'
'
'@/
Event Invalid()

''
' Raised when a change to a <code>BusinessDataObject</code> object puts it in a state that
' complies with all relevant business rules.
'
'
'@/
Event Valid()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName As String = "BusinessObjEventSource"

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub Dirty()
RaiseEvent Dirty
End Sub

Friend Sub Clean()
RaiseEvent Clean
End Sub

Friend Sub Deleted()
RaiseEvent Deleted
End Sub

Friend Sub Invalid()
RaiseEvent Invalid
End Sub

Friend Sub Valid()
RaiseEvent Valid
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

