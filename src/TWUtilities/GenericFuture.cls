VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GenericFuture"
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

Implements IFuture

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

Private Const ModuleName                            As String = "GenericFuture"

'@================================================================================
' Member variables
'@================================================================================

Private mFutureBuilder                              As FutureBuilder

Private mIsCancelled                                As Boolean
Private mIsAvailable                                As Boolean
Private mIsFaulted                                  As Boolean
Private mErrorNumber                                As Long
Private mErrorMessage                               As String
Private mErrorSource                                As String

Private mListeners                                  As New Listeners

Private mValue                                      As Object
Private mCookie                                     As Variant

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IFuture Interface Members
'@================================================================================

Private Sub IFuture_AddStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "IFuture_AddStateChangeListener"
On Error GoTo Err

mListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IFuture_Cancel()
Const ProcName As String = "IFuture_Cancel"
On Error GoTo Err

If Not IFuture_IsPending Then Exit Sub

mFutureBuilder.NotifyCancel

mIsCancelled = True
fireStateChange FutureStateCancelled
mListeners.Clear

Set mFutureBuilder = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IFuture_Cookie() As Variant
gSetVariant IFuture_Cookie, mCookie
End Property

Private Property Get IFuture_ErrorMessage() As String
IFuture_ErrorMessage = mErrorMessage
End Property

Private Property Get IFuture_ErrorNumber() As Long
IFuture_ErrorNumber = mErrorNumber
End Property

Private Property Get IFuture_ErrorSource() As String
IFuture_ErrorSource = mErrorSource
End Property

Private Property Get IFuture_IsAvailable() As Boolean
IFuture_IsAvailable = mIsAvailable
End Property

Private Property Get IFuture_IsCancelled() As Boolean
IFuture_IsCancelled = mIsCancelled
End Property

Private Property Get IFuture_IsFaulted() As Boolean
IFuture_IsFaulted = mIsFaulted
End Property

Private Property Get IFuture_IsPending() As Boolean
IFuture_IsPending = Not (mIsAvailable Or mIsCancelled Or mIsFaulted)
End Property

Private Sub IFuture_RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "IFuture_RemoveStateChangeListener"
On Error GoTo Err

mListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get IFuture_Value() As Object
Const ProcName As String = "IFuture_Value"
On Error GoTo Err

gAssert Not mIsCancelled, "Future value creation was cancelled"
gAssert Not mIsFaulted, "Future value creation failed: error=" & mErrorNumber & "; description=" & mErrorMessage & "; source=" & mErrorSource
gAssert mIsAvailable, "Future value creation has not yet completed"

Set IFuture_Value = mValue

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Cookie(ByVal Value As Variant)
Const ProcName As String = "Cookie"
On Error GoTo Err

gSetVariant mCookie, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Cookie() As Variant
Const ProcName As String = "Cookie"
On Error GoTo Err

gSetVariant Cookie, mCookie

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Value(ByVal pValue As Object)
Set mValue = pValue
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Cancel()
Const ProcName As String = "Cancel"
On Error GoTo Err

mIsCancelled = True
fireStateChange FutureStateCancelled
mListeners.Clear

Set mFutureBuilder = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Complete()
Const ProcName As String = "Complete"
On Error GoTo Err

mIsAvailable = True
fireStateChange FutureStateCompleted
mListeners.Clear

Set mFutureBuilder = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Fail(ByVal pErrorNumber As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "Fail"
On Error GoTo Err

mIsFaulted = True
mErrorNumber = pErrorNumber
mErrorMessage = pErrorMessage
mErrorSource = pErrorSource
fireStateChange FutureStateErrored
mListeners.Clear

Set mFutureBuilder = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pFutureBuilder As FutureBuilder)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mFutureBuilder = pFutureBuilder

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireStateChange(ByVal pState As FutureStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim lListeners As Listeners
Set lListeners = mListeners

' the current listeners are effectively removed and any new ones are added to a
' new collection
Set mListeners = New Listeners

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = pState

lListeners.SetCurrentListeners
Dim i As Long
For i = 1 To lListeners.Count
    Dim lListener As IStateChangeListener
    Set lListener = lListeners.GetListener(i)
    lListener.Change ev
Next
lListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub







