VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IntervalTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Objects of this class generate events at specified intervals, which may be regular or random."
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

''
'
' Objects of this class fire an event either once, after a specified elapsed time,
' or periodically.
'
' To create an <code>IntervalTimer</code> object, use the
' <code>CreateIntervalTimer</code> global method.
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

''
' Fires when the <code>IntervalTimer</code> object's timer interval has expired.
'
'@/
Public Event TimerExpired(ev As TimerExpiredEventData)

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                As String = "IntervalTimer"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mFirstTimerIntervalMillisecs As Long
Private mTimerIntervalMillisecs As Long
Private mRepeat As Boolean
Private mTimerNumber As Long
Private mRandomIntervals As Boolean
Private mState As TimerStates
Private mFirstInterval As Boolean
Private mData As Variant

Private mTimerExpiryListeners As Listeners

Private mSelfRef As Object

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mState = TimerStopped
Randomize
Set mTimerExpiryListeners = New Listeners
End Sub

Private Sub Class_Terminate()
'Debug.Print "IntervalTimer terminated"
If mTimerNumber <> 0 Then EndTimer mTimerNumber
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
' Returns the data passed to the <code>CreateintervalTimer</code> call.
'
' @return The data passed to the <code>CreateintervalTimer</code> call.
'
'@/
Public Property Get Data() As Variant
gSetVariant Data, mData
End Property

''
' Returns a <code>Boolean</code> value indicating whether the <code>TimerExpired</code>
' event fires repeatedly.
'
' @return A <code>Boolean</code> value indicating whether the <code>TimerExpired</code>
' event fires repeatedly.
'
'@/
Public Property Get RepeatNotifications() As Boolean
RepeatNotifications = mRepeat
End Property

''
' Indicates whether this <code>IntervalTimer</code> object is Running or Stopped.
'
' @return Returns <code>True</code> is this object is running and <code>False</code> if not..
'
'@/
Public Property Get State() As TimerStates
State = mState
End Property


'@================================================================================
' Methods
'@================================================================================

Public Sub AddTimerExpiryListener(ByVal pListener As ITimerExpiryListener)
Const ProcName As String = "AddTimerExpiryListener"
On Error GoTo Err

mTimerExpiryListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal firstExpiryIntervalMillisecs As Long, _
                ByVal repeatIntervalMillisecs As Long, _
                ByVal useRandomIntervals As Boolean, _
                ByVal pData As Variant)
mFirstTimerIntervalMillisecs = firstExpiryIntervalMillisecs
mTimerIntervalMillisecs = repeatIntervalMillisecs
If mTimerIntervalMillisecs > 0 Then
    mRepeat = True
Else
    mRepeat = False
End If
mRandomIntervals = useRandomIntervals
gSetVariant mData, pData
mState = TimerStopped
mFirstInterval = False
End Sub
                
Friend Sub Notify()
Const ProcName As String = "Notify"
On Error GoTo Err

Dim interval As Long

If Not mRepeat Then
    stopTimerEx
    mState = TimerStates.TimerStopped
ElseIf mRandomIntervals Then
    stopTimerEx
    interval = Rnd * mTimerIntervalMillisecs + 1
    mTimerNumber = BeginTimer(interval, False, Me)
ElseIf mFirstInterval Then
    stopTimerEx
    mTimerNumber = BeginTimer(mTimerIntervalMillisecs, True, Me)
End If

fireTimerExpired

If mState = TimerStopped Then Set mSelfRef = Nothing
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveTimerExpiryListener(ByVal pListener As ITimerExpiryListener)
Const ProcName As String = "RemoveTimerExpiryListener"
On Error GoTo Err

mTimerExpiryListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Tells this <code>IntervalTimer</code> object to start producing <code>TimerExpired</code> events.
'
' @Remarks Once running, the <code>IntervalTimer</code> object may be stopped by calling
' the <code>stopTimerEx</code> method.
'
'@/
Public Sub StartTimer()
Const ProcName As String = "StartTimer"
On Error GoTo Err

Dim interval As Long

Set mSelfRef = Me

If mTimerNumber <> 0 Then stopTimerEx

If mFirstTimerIntervalMillisecs <> 0 Then
    mFirstInterval = True
    interval = mFirstTimerIntervalMillisecs
    If mRandomIntervals Then
        interval = Rnd * mFirstTimerIntervalMillisecs + 1
    End If
    mTimerNumber = BeginTimer(mFirstTimerIntervalMillisecs, False, Me)
Else
    If mRandomIntervals Then
        interval = Rnd * mTimerIntervalMillisecs + 1
        mTimerNumber = BeginTimer(interval, False, Me)
    Else
        mTimerNumber = BeginTimer(mTimerIntervalMillisecs, True, Me)
    End If
End If
mState = TimerStates.TimerRunning

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Tells the <code>IntervalTimer</code> object to stop producing <code>TimerExpired</code> events.
'

'@/
Public Sub StopTimer()
Const ProcName As String = "StopTimer"

On Error GoTo Err

stopTimerEx
Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireTimerExpired()
Const ProcName As String = "fireTimerExpired"
On Error GoTo Err

Dim ev As TimerExpiredEventData

Set ev.Source = Me
gSetVariant ev.Data, mData

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mTimerExpiryListeners.Valid Then
    sInit = True
    sSomeListeners = mTimerExpiryListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ITimerExpiryListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.TimerExpired ev
    Next
End If

RaiseEvent TimerExpired(ev)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub stopTimerEx()
Const ProcName As String = "stopTimerEx"

On Error GoTo Err

mFirstInterval = False
If mTimerNumber = 0 Then Exit Sub
EndTimer mTimerNumber
mTimerNumber = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




