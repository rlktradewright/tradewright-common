VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FutureBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Event Cancelled(ev As CancelledEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "FutureBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mFuture                                     As GenericFuture
Private mFutureAsIFuture                            As IFuture

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mFuture = New GenericFuture
mFuture.Initialise Me
Set mFutureAsIFuture = mFuture
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

Public Property Let Cookie(ByVal Value As Variant)
Const ProcName As String = "Cookie"
On Error GoTo Err

mFuture.Cookie = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Cookie() As Variant
Const ProcName As String = "Cookie"
On Error GoTo Err

gSetVariant Cookie, mFuture.Cookie

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Future() As IFuture
Set Future = mFuture
End Property

Public Property Let Value(ByVal pValue As Object)
Const ProcName As String = "Value"
On Error GoTo Err

gAssert mFutureAsIFuture.IsPending, "Future is no longer pending"

mFuture.Value = pValue

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Cancel()
Const ProcName As String = "Cancel"
On Error GoTo Err

If Not mFutureAsIFuture.IsPending Then Exit Sub

mFuture.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Complete()
Const ProcName As String = "Complete"
On Error GoTo Err

gAssert mFutureAsIFuture.IsPending, "Future is no longer pending"

mFuture.Complete

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Fail(ByVal pErrorNumber As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "Fail"
On Error GoTo Err

gAssert mFutureAsIFuture.IsPending, "Future is no longer pending"

mFuture.Fail pErrorNumber, pErrorMessage, pErrorSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyCancel()
Const ProcName As String = "NotifyCancel"
On Error GoTo Err

Dim ev As CancelledEventData
Set ev.Source = Me
RaiseEvent Cancelled(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




