VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimerListItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Objects of this class are created by the AddTimerListEntry method of the TimerList class."
Option Explicit

''
' A <code>TimerListItem</code> object is used to notify the application when a
' specified time has arrived. It also holds some data that will be relevant
' to the application at that time.
'
' A <code>TimerListItem</code> object is created by calling a <code>TimerList</code>
' object's <code>add</code> method.
'
' To determine when a <code>TimerListItem</code> has expired, either handle its
' <code>StateChange</code> event, or register a listener object using its
' <code>addStateChangeListener</code> method.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

''
' Fires when this <code>TimerListItem</code>'s State changes.
'
' @param ev
'   Contains details of the change of State. The <code>state</code> member of
'   the <code>StateChangeEventData</code> is a <code>TimerListItemStates</code> enum
'   value.
'@/
Event StateChange(ByRef ev As StateChangeEventData)

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                            As String = "TimerListItem"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mExpiryTime                 As Date

Private mData                       As Variant
Private mState                      As TimerListItemStates

Private mStateChangeListeners       As Listeners

Private mTimerListRef               As WeakReference

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mState = TimerListItemStatePending
Set mStateChangeListeners = New Listeners
End Sub

Private Sub Class_Terminate()
Debug.Print "TimerListItem terminated"
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
' Returns the value specified in the <code>add</code> method call on a
' <code>TimerList</code> object that created this object.
'
' @return
'   the value specified in the <code>add</code> method call on a
'   <code>TimerList</code> object that created this object.
'@/
Public Property Get Data() As Variant
gSetVariant Data, mData
End Property

''
' Returns the time at which this <code>TimerListItem</code> object will expire.
'
' @remarks
'   Note that the time returned is a UTC?????
' @return
'   The time at which this <code>TimerListItem</code> object will expire.
'@/
Friend Property Get ExpiryTime() As Date
ExpiryTime = mExpiryTime
End Property

Friend Property Let State( _
                ByVal Value As TimerListItemStates)
Const ProcName As String = "State"
On Error GoTo Err

If State = TimerListItemStateExpired Then Exit Property

mState = Value
fireStateChange mState

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the current State of this <code>TimerListItem</code> object.
'
' @return
'   The current State of this <code>TimerListItem</code> object.
'@/
Public Property Get State() As TimerListItemStates
State = mState
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Registers an object to act as a listener for changes of State.
'
' @remarks
'   The listener object must implement the <code>StateChangeListener</code>
'   interface.
' @param listener
'   The object which is to act as a listener.
'@/
Public Sub AddStateChangeListener( _
                ByVal pListener As IStateChangeListener)
Const ProcName As String = "AddStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Cancels this <code>TimerListItem</code>.
'
' @remarks
'   The <code>TimerListItem</code>'s status is set to
'   <code>TimerListItemStateCancelled</code>.
'@/
Public Sub Cancel()
Const ProcName As String = "Cancel"
On Error GoTo Err

Dim lTimerList As TimerList
Set lTimerList = mTimerListRef.Target
lTimerList.Remove Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise(ByVal pExpiryTime As Date, ByVal pData As Variant, ByVal pTimerList As TimerList)
mExpiryTime = pExpiryTime
gSetVariant mData, pData
Set mTimerListRef = gCreateWeakReference(pTimerList)
End Sub

Friend Sub Notify()
Const ProcName As String = "Notify"
On Error GoTo Err

If mState <> TimerListItemStatePending Then
    ' this can occur because notifying an earlier Item could Result in the
    ' application removing this one after it has been identified for notification
    ' but before it has actually been notified
    Exit Sub
End If
State = TimerListItemStateExpired

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Unregisters a <code>StateChangeListener</code> object.
'
' @remarks
'   No Error occurs if the object to be unregistered is not
'    actually registered.
' @param listener
'   The listener object to be unregistered.
' @see
'
'@/
Public Sub RemoveStateChangeListener(ByVal pListener As IStateChangeListener)
Const ProcName As String = "RemoveStateChangeListener"
On Error GoTo Err

mStateChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireStateChange(ByVal NewState As TimerListItemStates)
Const ProcName As String = "fireStateChange"
On Error GoTo Err

Dim ev As StateChangeEventData
Set ev.Source = Me
ev.State = NewState

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mStateChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mStateChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IStateChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If

RaiseEvent StateChange(ev)

If ev.State = TimerListItemStates.TimerListItemStateCancelled Or _
    ev.State = TimerListItemStates.TimerListItemStateExpired Then _
    mStateChangeListeners.Clear
    
Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub


