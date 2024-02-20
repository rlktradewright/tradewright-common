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

Private mState                                      As FutureStates

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

mState = FutureStateCancelled
If Not mFutureBuilder Is Nothing Then mFutureBuilder.Cancel

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
IFuture_IsAvailable = (mState = FutureStateCompleted)
End Property

Private Property Get IFuture_IsCancelled() As Boolean
IFuture_IsCancelled = (mState = FutureStateCancelled)
End Property

Private Property Get IFuture_IsFaulted() As Boolean
IFuture_IsFaulted = (mState = FutureStateErrored)
End Property

Private Property Get IFuture_IsPending() As Boolean
IFuture_IsPending = (mState = FutureStatePending)
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

gAssert mState <> FutureStateCancelled, "Future Value creation was cancelled"
gAssert mState <> FutureStateErrored, "Future Value creation failed: error=" & mErrorNumber & "; description=" & mErrorMessage & "; source=" & mErrorSource
gAssert mState = FutureStateCompleted, "Future Value creation has not yet completed"

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub NotifyCancel()
Const ProcName As String = "NotifyCancel"
On Error GoTo Err

gAssert mState = FutureStateCancelled, "Invalid state"
fireStateChange FutureStateCancelled
mListeners.Clear

Set mFutureBuilder = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyComplete()
Const ProcName As String = "NotifyComplete"
On Error GoTo Err

If mState = FutureStateCancelled Then
    NotifyCancel
    Exit Sub
End If

gAssert mState = FutureStateCompleted, "Invalid state"
fireStateChange FutureStateCompleted
mListeners.Clear

Set mFutureBuilder = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub NotifyFail()
Const ProcName As String = "NotifyFail"
On Error GoTo Err

If mState = FutureStateCancelled Then
    NotifyCancel
    Exit Sub
End If

gAssert mState = FutureStateErrored, "Invalid state"

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

Friend Sub SetCancel()
mState = FutureStateCancelled
End Sub

Friend Sub SetFail(ByVal pErrorNumber As Long, ByVal pErrorMessage As String, ByVal pErrorSource As String)
Const ProcName As String = "SetFail"
On Error GoTo Err

If mState = FutureStateCancelled Then Exit Sub

mErrorNumber = pErrorNumber
mErrorMessage = pErrorMessage
mErrorSource = pErrorSource
mState = FutureStateErrored

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub SetValue(ByVal pValue As Object)
If mState = FutureStateCancelled Then Exit Sub

Set mValue = pValue
mState = FutureStateCompleted
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireStateChange(ByVal pState As FutureStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim lKey As String: lKey = gGetObjectKey(Me)
'Debug.Print "GenericFuture::fireStateChange: " & lKey

Dim lListeners As Listeners
Set lListeners = mListeners

' the current listeners are effectively removed and any new ones are added to a
' new collection
Set mListeners = New Listeners

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = pState

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not lListeners.Valid Then
    sInit = True
    sSomeListeners = lListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IStateChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        'Debug.Print "GenericFuture::fireStateChange: notify listener: " & gGetObjectKey(lListener)
        lListener.Change ev
    Next
End If
lListeners.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub







