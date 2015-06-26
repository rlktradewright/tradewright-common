Attribute VB_Name = "GIntervalTimer"
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "GIntervalTimer"

Private Const NullIndex                     As Long = -1

Private Const MinTimerResolution            As Long = 1

Private Const TimerTableInitialSize         As Long = 16

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type TimerTableEntry
    Handle      As Long
    TimerObj    As IntervalTimer
    Periodic    As Boolean
    Fired       As Boolean
    NextFree    As Long
    Ending      As Boolean
End Type

'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mMinRes As Long

Private mTimers() As TimerTableEntry
Private mTimersIndex As Long

Private mFirstFree As Long

Private mNumRunningTimers As Long

Private mhThread As Long
Private mhTimerQueue As Long

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

Public Function BeginTimer( _
                ByVal pInterval As Long, _
                ByVal pPeriodic As Boolean, _
                ByVal pTimerObj As IntervalTimer) As Long
Const ProcName As String = "BeginTimer"

On Error GoTo Err

Dim lTimerNumber As Long
Dim i As Long

If mFirstFree <> NullIndex Then
    lTimerNumber = mFirstFree
    mFirstFree = mTimers(mFirstFree).NextFree
    'Debug.Print "Reuse timer entry: " & lTimerNumber
Else
    
    If mTimersIndex > UBound(mTimers) Then
        ReDim Preserve mTimers(1 To 2 * UBound(mTimers)) As TimerTableEntry
        Debug.Print "Timer table extended: size = " & UBound(mTimers)
        If gLogger.IsLoggable(LogLevelHighDetail) Then _
            gLogger.Log "Increased mTimers size", ProcName, ModuleName, LogLevelHighDetail, CStr(UBound(mTimers) + 1)
    End If
    lTimerNumber = mTimersIndex
    mTimersIndex = mTimersIndex + 1
    'Debug.Print "Allocate timer entry: " & lTimerNumber
End If

With mTimers(lTimerNumber)
    Set .TimerObj = pTimerObj
    .Periodic = pPeriodic
    If CreateTimerQueueTimer(VarPtr(.Handle), mhTimerQueue, AddressOf TimerProc, lTimerNumber, pInterval, IIf(.Periodic, pInterval, 0), WT_EXECUTEINTIMERTHREAD) = 0 Then gHandleWin32Error
    'If gLogger.IsLoggable(LogLevelHighDetail) Then gLogger.Log  "Started timer: handle: " & CStr(.Handle) & "; interval: " & CStr(pInterval), ProcName, ModuleName,LogLevelHighDetail
    
    mNumRunningTimers = mNumRunningTimers + 1
    
End With

BeginTimer = lTimerNumber

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub EndTimer(ByVal timerNumber As Long)

Const ProcName As String = "EndTimer"

On Error GoTo Err

With mTimers(timerNumber)
    DeleteTimerQueueTimer mhTimerQueue, .Handle, 0
    mNumRunningTimers = mNumRunningTimers - 1
    releaseEntry timerNumber
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gInit()
Const ProcName As String = "gInit"

On Error GoTo Err

Dim tc As TIMECAPS
Dim i As Long

If mMinRes <> 0 Then Exit Sub

If TimeGetDevCaps(tc, 8) <> TIMERR_NOERROR Then gHandleWin32Error

mMinRes = IIf(tc.wPeriodMin < MinTimerResolution, MinTimerResolution, tc.wPeriodMin)
If mMinRes > tc.wPeriodMax Then mMinRes = tc.wPeriodMax

TimeBeginPeriod mMinRes

If DuplicateHandle(GetCurrentProcess, GetCurrentThread, GetCurrentProcess, VarPtr(mhThread), 0, 0, DUPLICATE_SAME_ACCESS) = 0 Then gHandleWin32Error
mhTimerQueue = CreateTimerQueue
If mhTimerQueue = 0 Then gHandleWin32Error

ReDim mTimers(1 To TimerTableInitialSize) As TimerTableEntry
mTimersIndex = 1
mFirstFree = NullIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gProcessUserTimerMsg(ByVal pIndex As Long)
Const ProcName As String = "gProcessUserTimerMsg"
On Error GoTo Err

If mTimers(pIndex).Ending Then
    ' a call to end the timer was made after the TimerProc was called but
    ' before the APC executed
    mNumRunningTimers = mNumRunningTimers - 1
    releaseEntry pIndex
    Exit Sub
End If

If (Not mTimers(pIndex).Periodic) Then mNumRunningTimers = mNumRunningTimers - 1
   
If Not gInitialised Then
    Exit Sub
End If

' don't use With mTimers(pIndex) here, because if another timer is started
' in the event handler and that required the table to be ReDim'ed, that
' causes an Error (table is locked by the With)
mTimers(pIndex).Fired = True
If mTimers(pIndex).Handle <> 0 Then
    'If gLogger.IsLoggable(LogLevelHighDetail) Then gLogger.Log  "Fire timer: handle", ProcName, ModuleName, CStr(mTimers(pIndex).Handle), LogLevelHighDetail
    mTimers(pIndex).TimerObj.Notify
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gTerm()
Dim i As Long
Const ProcName As String = "gTerm"

On Error GoTo Err

For i = 1 To mTimersIndex - 1
    If Not mTimers(i).TimerObj Is Nothing Then
        EndTimer i
    End If
Next
Debug.Print "Delete timer queue"
DeleteTimerQueueEx mhTimerQueue, INVALID_HANDLE_VALUE
Debug.Print "Deleted timer queue"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub releaseEntry( _
                ByVal pIndex As Long)
Const ProcName As String = "releaseEntry"

On Error GoTo Err

With mTimers(pIndex)
    .Handle = 0
    .Fired = False
    .Periodic = False
    Set .TimerObj = Nothing
    .Ending = False
    
    .NextFree = mFirstFree
    mFirstFree = pIndex
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub TimerProc( _
                ByVal pParam As Long, _
                ByVal pTimerOrWaitFired As Long)
If Not gInitialised Then Exit Sub
' NB: trying to do anything else in this proc doesn't work because we're
' not on the VB thread
gPostUserMessage UserMessageTimer, pParam, 0
End Sub





