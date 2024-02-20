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

Private Const MinTimerResolution            As Long = 1

Private Const TimerTableInitialSize         As Long = 16

' defines the interval between ending a timer and allowing its table entry
' to be reused, to allow for the situation where a call to end the timer is made
' after the TimerProc was called but before the UserMessageTimer message arrives
' at the WindowProc - if the table entry has been reused, the new client would
' be wrongly notified
Private Const ReleaseInterval               As Double = 100# / (86400# * 1000#)

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
    ReleaseTime As Date
    Next        As Long
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

'index of first entry in free list
Private mFirstFree As Long

' index of first entry in ending queue
Private mFirstEnding As Long

' index of last entry in ending queue
Private mLastEnding As Long

Private mNumRunningTimers As Long

Private mhTimerQueue As Long

Private mPowerThrottling As PROCESS_POWER_THROTTLING_STATE

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

Static sInitialised As Boolean
If Not sInitialised Then
    allowTimerResolution
    TimeBeginPeriod 1   ' mMinRes
    sInitialised = True
End If

Dim lTimerNumber As Long
lTimerNumber = allocateEntry

With mTimers(lTimerNumber)
    Set .TimerObj = pTimerObj
    .Periodic = pPeriodic
    If CreateTimerQueueTimer(VarPtr(.Handle), mhTimerQueue, AddressOf TimerProc, lTimerNumber, pInterval, IIf(.Periodic, pInterval, 0), WT_EXECUTEINTIMERTHREAD) = 0 Then gHandleWin32Error
'    If gLogger.IsLoggable(LogLevelHighDetail) Then
'        Dim s As String: s = "CreateTimerQueueTimer " & lTimerNumber & " " & pInterval & " " & CDbl(gGetTimestamp)
'        gLogger.Log s, ProcName, ModuleName, LogLevelHighDetail
'    End If
    mNumRunningTimers = mNumRunningTimers + 1
    
End With

BeginTimer = lTimerNumber

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub EndTimer(ByVal TimerNumber As Long)
Const ProcName As String = "EndTimer"
On Error GoTo Err

With mTimers(TimerNumber)
    DeleteTimerQueueTimer mhTimerQueue, .Handle, 0
    If .Periodic Or Not .Fired Then mNumRunningTimers = mNumRunningTimers - 1
    releaseEntry TimerNumber
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gInit()
Const ProcName As String = "gInit"
On Error GoTo Err

If mMinRes <> 0 Then Exit Sub

Dim TC As TIMECAPS
If TimeGetDevCaps(TC, 8) <> TIMERR_NOERROR Then gHandleWin32Error

mMinRes = IIf(TC.wPeriodMin < MinTimerResolution, MinTimerResolution, TC.wPeriodMin)
If mMinRes > TC.wPeriodMax Then mMinRes = TC.wPeriodMax

allowTimerResolution

mhTimerQueue = CreateTimerQueue
If mhTimerQueue = 0 Then gHandleWin32Error

ReDim mTimers(1 To TimerTableInitialSize) As TimerTableEntry
mTimersIndex = 1
mFirstFree = NullIndex
mFirstEnding = NullIndex
mLastEnding = NullIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gProcessUserTimerMsg(ByVal pIndex As Long)
Const ProcName As String = "gProcessUserTimerMsg"
On Error GoTo Err

If (mTimers(pIndex).Handle <> 0 And Not mTimers(pIndex).Periodic) Then
    mTimers(pIndex).Fired = True
    mNumRunningTimers = mNumRunningTimers - 1
End If

If Not gInitialised Then
    Exit Sub
End If

' don't use With mTimers(pIndex) here, because if another timer is started
' in the event handler and that required the table to be ReDim'ed, that
' causes an error (table is locked by the With)
If mTimers(pIndex).Handle <> 0 Then
    'If gLogger.IsLoggable(LogLevelHighDetail) Then gLogger.Log  "Fire timer: handle", ProcName, ModuleName, CStr(mTimers(pIndex).Handle), LogLevelHighDetail
    mTimers(pIndex).TimerObj.Notify
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gTerm()
Const ProcName As String = "gTerm"
On Error GoTo Err

Dim i As Long
For i = 1 To mTimersIndex - 1
    If Not mTimers(i).TimerObj Is Nothing Then
        EndTimer i
    End If
Next
'Debug.Print "Delete timer queue"
DeleteTimerQueueEx mhTimerQueue, INVALID_HANDLE_VALUE
'Debug.Print "Deleted timer queue"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function allocateEntry() As Long
Const ProcName As String = "allocateEntry"

Static sMaxIndex As Long

Dim lIndex As Long
lIndex = allocateEntryIndex

With mTimers(lIndex)
    .Fired = False
    .Handle = 0
    .Next = NullIndex
    .Periodic = False
    .ReleaseTime = 0#
    Set .TimerObj = Nothing
End With

If lIndex > sMaxIndex Then
'    If gLogger.IsLoggable(LogLevelHighDetail) Then
'        gLogger.Log "Max index: " & sMaxIndex, ProcName, ModuleName, LogLevelHighDetail
'    End If
    sMaxIndex = lIndex
End If
allocateEntry = lIndex
End Function

Private Function allocateEntryIndex() As Long
Const ProcName As String = "allocateEntryIndex"

If mFirstFree <> NullIndex Then
    allocateEntryIndex = allocateFirstFree
    Exit Function
End If

If mTimersIndex <= UBound(mTimers) Then
    allocateEntryIndex = mTimersIndex
    mTimersIndex = mTimersIndex + 1
    Exit Function
End If

If mFirstEnding <> NullIndex Then
    Dim t As Date: t = gGetTimestampUtc
    Dim lCount As Long
    Do While t >= mTimers(mFirstEnding).ReleaseTime + ReleaseInterval
        Dim lNewFirstFree As Long: lNewFirstFree = mFirstEnding
        mFirstEnding = mTimers(mFirstEnding).Next
        mTimers(lNewFirstFree).Next = mFirstFree
        mFirstFree = lNewFirstFree
        lCount = lCount + 1
        If mFirstEnding = NullIndex Then
            mLastEnding = NullIndex
            Exit Do
        End If
    Loop
'    If gLogger.IsLoggable(LogLevelHighDetail) Then
'        gLogger.Log "Released timers: " & lCount, ProcName, ModuleName, LogLevelHighDetail
'    End If
End If

If mFirstFree <> NullIndex Then
    allocateEntryIndex = allocateFirstFree
    Exit Function
End If

If mTimersIndex > UBound(mTimers) Then
    ReDim Preserve mTimers(1 To 2 * UBound(mTimers)) As TimerTableEntry
    'Debug.Print "Timer table extended: size = " & UBound(mTimers)
    If gLogger.IsLoggable(LogLevelHighDetail) Then
        gLogger.Log "Increased mTimers size", ProcName, ModuleName, LogLevelHighDetail, CStr(UBound(mTimers) + 1)
    End If
End If
allocateEntryIndex = mTimersIndex
mTimersIndex = mTimersIndex + 1
End Function

Private Function allocateFirstFree() As Long
allocateFirstFree = mFirstFree
mFirstFree = mTimers(mFirstFree).Next
End Function

Private Sub allowTimerResolution()
Const ProcName As String = "allowTimerResolution"
On Error GoTo Err

Dim failPoint As String

' HighQoS
' Turn EXECUTION_SPEED throttling off.

' Don't throttle speed
failPoint = "Don't throttle speed"

mPowerThrottling.Version = PROCESS_POWER_THROTTLING_CURRENT_VERSION
mPowerThrottling.ControlMask = PROCESS_POWER_THROTTLING_EXECUTION_SPEED
mPowerThrottling.StateMask = 0

Dim lResult As Long
lResult = SetProcessInformation(GetCurrentProcess(), _
                      PROCESS_INFORMATION_CLASS.ProcessPowerThrottling, _
                      VarPtr(mPowerThrottling), _
                      Len(mPowerThrottling))
If lResult = 0 Then
    Dim lLastError As Long: lLastError = GetLastError
    failPoint = "GetLastError 1"
    'Debug.Print "Error " & lLastError & " from SetProcessInformation() for PROCESS_POWER_THROTTLING_EXECUTION_SPEED"
    If lLastError <> 0 Then MsgBox "Error " & lLastError & " from SetProcessInformation() for PROCESS_POWER_THROTTLING_EXECUTION_SPEED", vbInformation, "Error"
    If lLastError = 0 Then
    ElseIf lLastError <> 87 Then
        gHandleWin32Error
    End If
End If

' Honor Timer Resolution Requests
failPoint = "Honor Timer Resolution Requests"

mPowerThrottling.Version = PROCESS_POWER_THROTTLING_CURRENT_VERSION
mPowerThrottling.ControlMask = PROCESS_POWER_THROTTLING_IGNORE_TIMER_RESOLUTION
mPowerThrottling.StateMask = 0

lResult = SetProcessInformation(GetCurrentProcess(), _
                      PROCESS_INFORMATION_CLASS.ProcessPowerThrottling, _
                      VarPtr(mPowerThrottling), _
                      Len(mPowerThrottling))
If lResult = 0 Then
    failPoint = "GetLastError 2"
    lLastError = GetLastError
    'Debug.Print "Error " & lLastError & " from SetProcessInformation() for PROCESS_POWER_THROTTLING_IGNORE_TIMER_RESOLUTION"
    If lLastError <> 0 Then MsgBox "Error " & lLastError & " from SetProcessInformation() for PROCESS_POWER_THROTTLING_IGNORE_TIMER_RESOLUTION", vbInformation, "Error"
    If lLastError = 0 Then
    ElseIf lLastError <> 87 Then
        gHandleWin32Error
    End If
End If

Exit Sub

Err:
MsgBox Err.Description & vbCrLf & _
        Err.Source & ": " & failPoint, , "Error!"
gHandleUnexpectedError ProcName, ModuleName, failPoint
End Sub

Private Sub releaseEntry( _
                ByVal pIndex As Long)
Const ProcName As String = "releaseEntry"
On Error GoTo Err

With mTimers(pIndex)
    .Handle = 0
    .Next = NullIndex
    .ReleaseTime = gGetTimestampUtc
End With

If mFirstEnding = NullIndex Then
    mFirstEnding = pIndex
Else
    mTimers(mLastEnding).Next = pIndex
End If

mLastEnding = pIndex

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

'Dim s As String: s = "FireTimer " & pParam & " " & CDbl(gGetTimestamp)
'If gLogger.IsLoggable(LogLevelHighDetail) Then gLogger.Log s, "TimerProc", ModuleName, LogLevelHighDetail

gPostUserMessage UserMessageTimer, pParam, 0
End Sub





