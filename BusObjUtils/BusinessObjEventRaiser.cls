VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusinessObjEventRaiser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' A <code>BusinessObjEventRaiser</code> object is used to create a
' <code>BusinessObjEventSource</code> object and to cause it to fire events.
'
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

Private Const ModuleName As String = "BusinessObjEventRaiser"

'@================================================================================
' Member variables
'@================================================================================

Private mEventSource As BusinessObjEventSource

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mEventSource = New BusinessObjEventSource
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

''
' Returns the <code>BusinessObjEventSource</code> object created by this
' <code>BusinessObjEventRaiser</code> object.
'
' @return the <code>BusinessObjEventSource</code> object created by this
'           <code>BusinessObjEventRaiser</code> object.
'
'@/
Public Property Get EventSource() As BusinessObjEventSource
Set EventSource = mEventSource
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Fires the <code>Dirty</code> event in the related <code>BusinessObjEventSource</code>
' object.
'
'
'@/
Public Sub Dirty()
Const ProcName As String = "Dirty"
Dim failpoint As String
On Error GoTo Err

mEventSource.Dirty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Fires the <code>Clean</code> event in the related <code>BusinessObjEventSource</code>
' object.
'
'
'@/
Public Sub Clean()
Const ProcName As String = "Clean"
Dim failpoint As String
On Error GoTo Err

mEventSource.Clean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Fires the <code>Deleted</code> event in the related <code>BusinessObjEventSource</code>
' object.
'
'
'@/
Public Sub Deleted()
Const ProcName As String = "Deleted"
Dim failpoint As String
On Error GoTo Err

mEventSource.Deleted

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Fires the <code>Invalid</code> event in the related <code>BusinessObjEventSource</code>
' object.
'
'
'@/
Public Sub Invalid()
Const ProcName As String = "Invalid"
Dim failpoint As String
On Error GoTo Err

mEventSource.Invalid

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Fires the <code>Valid</code> event in the related <code>BusinessObjEventSource</code>
' object.
'
'
'@/
Public Sub Valid()
Const ProcName As String = "Valid"
Dim failpoint As String
On Error GoTo Err

mEventSource.Valid

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


