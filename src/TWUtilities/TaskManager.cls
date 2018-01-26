VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
'   This class schedules the processing of <code>Task</code> objects.
'
'   There is a single <code>TaskManager</code> object per process, which is created
'   automatically.
'
'   Applications cannot create <code>TaskManager</code> objects. Their only means of
'   interacting with the global <code>TaskManager</code> object is via the
'   <code>StartTask</code> global method and the <code>RunTasksAtLowerThreadPriority</code>
'   and <code>TaskQuantumMillisecs</code> global properties.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IEnumerable
Implements IStateChangeListener

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                        As String = "TaskManager"

Private Const DefaultQuantumMillisecs                           As Long = 32
Private Const MinimumQuantumMillisecs                           As Long = 1

' per GHz of CPU
Private Const DefaultTaskConcurrency                            As Long = 4
Private Const MinimumTaskConcurrency                            As Long = 2

Private Const DefaultTaskScheduleIntervalMillisecs              As Long = 1
Private Const DefaultTimeslotMillisecs                          As Long = 16

Private Const LastIndex                                         As Long = -1

Private Const TaskSummaryInfoType                               As String = "diag.twutilities.tasks"

'@================================================================================
' Enums
'@================================================================================

Private Enum EnumeratorTypes
    Runnable = &H10000000
    Suspended = &H20000000
    Processed = &H30000000
    Pending = &H40000000
    Restartable = &H50000000
    All = &H80000000
End Enum

Private Enum TaskPrimitiveStates
    Free
    Pending
    Processed
    Running
    Suspended
    Restartable
End Enum

'@================================================================================
' Types
'@================================================================================

Private Type TaskTableEntry
    State           As TaskPrimitiveStates
    Next            As Long
    Prev            As Long
    TC              As TaskContext
    Priority        As TaskPriorities
    TimeRemaining   As Long ' number of microseconds remaining in the current
                            ' timeslot for this task
    SuspendTLI      As TimerListItem
    StartTime       As Date
    LastRunTime     As Date
    TotalCPUTime    As Long
End Type
Dim TC
'@================================================================================
' External function declarations
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTaskConcurrency                                    As Long

Private mTaskScheduleIntervalMillisecs                      As Long

Private mQuantumMillisecs                                   As Long
Private mQuantumRemainingMicrosecs                          As Long

Private mTaskSummaryLoggingIntervalSecs                     As Long
Private WithEvents mTaskInfoTimer                           As IntervalTimer
Attribute mTaskInfoTimer.VB_VarHelpID = -1
Private mSummaryLogger                                      As Logger

Private mRunTasksAtLowerThreadPriority                      As Boolean
Private mTaskTable()                                        As TaskTableEntry
Private mTaskTableNextIndex                                 As Long

' current in a doubly-linked circular queue of running entries
Private mFirstRunningIndex                                  As Long
                                    
' first in a doubly-linked circular queue of suspended entries
Private mFirstSuspendedIndex                                As Long

' first in a forward-linked list of free entries
Private mFirstFreeIndex                                     As Long

' first in a doubly-linked circular queue of pending entries
Private mFirstPendingIndex                                  As Long

' first in a doubly-linked circular queue of entries that have come out
' of suspension
Private mFirstRestartableIndex                              As Long

' first in a doubly-linked circular queue of entries that have used up their time
' allocation but have not yet completed
Private mFirstProcessedIndex                                As Long

' timing need need not be very accurate for suspended tasks
Private mTimerList                                          As TimerList
Attribute mTimerList.VB_VarHelpID = -1

Private mProcessing                                         As Boolean

Private mUpdateSequence                                     As Long

' Includes tasks on both the running and the processed list
Private mNumRunnableTasks                                   As Long

Private mScheduleInvoked                                    As Boolean

Private mInterScheduleElapsedTimer                          As New ElapsedTimer
Private mTotalInterScheduleElapsedTime                      As Double
Private mNumberOfInterSchedulePeriods                       As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mTaskConcurrency = Round(DefaultTaskConcurrency * gNumberOfProcessors * gCpuSpeedMhz / 1000)
If mTaskConcurrency < MinimumTaskConcurrency Then mTaskConcurrency = MinimumTaskConcurrency
mQuantumMillisecs = DefaultQuantumMillisecs
mTaskScheduleIntervalMillisecs = DefaultTaskScheduleIntervalMillisecs
mRunTasksAtLowerThreadPriority = True

ReDim mTaskTable(3) As TaskTableEntry

mFirstFreeIndex = LastIndex
mFirstRunningIndex = LastIndex

mFirstSuspendedIndex = LastIndex

mFirstPendingIndex = LastIndex

mFirstRestartableIndex = LastIndex

mFirstProcessedIndex = LastIndex

Set mTimerList = GTimerList.gGetTimerList
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Const ProcName As String = "IEnumerable_Enumerator"
On Error GoTo Err

Set IEnumerable_Enumerator = getEnumerator(EnumeratorTypes.Runnable)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IEnumerable_NextItem(pItem As Variant, pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_NextItem"
On Error GoTo Err

#If trace Then
    gNullTracer.EnterProcedure ProcName, ProjectName, ModuleName, ""
#End If

Const NoMoreItems As Long = &H7FFFFFFF

If CLng(pCookie) = NoMoreItems Then
    #If trace Then
        gNullTracer.ExitProcedure "IEnumerable_NextItem", ProjectName, ModuleName, ""
    #End If
    Exit Function
End If

Dim enumType As Long
enumType = CLng(pCookie) And &H70000000

Dim enumAll As Boolean
enumAll = ((CLng(pCookie) And EnumeratorTypes.All) = EnumeratorTypes.All)

Dim lIndex As Long
lIndex = CLng(pCookie) And &HFFFFFFF

If enumAll And enumType = 0 Then enumType = EnumeratorTypes.Runnable

If enumType = EnumeratorTypes.Runnable And lIndex = &HFFFFFFF Then
    If mFirstRunningIndex = LastIndex Then
        If enumAll Then
            enumType = EnumeratorTypes.Processed
        Else
            pCookie = NoMoreItems
            #If trace Then
                gNullTracer.ExitProcedure "IEnumerable_NextItem", ProjectName, ModuleName, ""
            #End If
            Exit Function
        End If
    Else
        lIndex = mFirstRunningIndex
    End If
End If

If enumType = EnumeratorTypes.Processed And lIndex = &HFFFFFFF Then
    If mFirstProcessedIndex = LastIndex Then
        If enumAll Then
            enumType = EnumeratorTypes.Restartable
        Else
            pCookie = NoMoreItems
            #If trace Then
                gNullTracer.ExitProcedure "IEnumerable_NextItem", ProjectName, ModuleName, ""
            #End If
            Exit Function
        End If
    Else
        lIndex = mFirstProcessedIndex
    End If
End If

If enumType = EnumeratorTypes.Restartable And lIndex = &HFFFFFFF Then
    If mFirstRestartableIndex = LastIndex Then
        If enumAll Then
            enumType = EnumeratorTypes.Pending
        Else
            pCookie = NoMoreItems
            #If trace Then
                gNullTracer.ExitProcedure "IEnumerable_NextItem", ProjectName, ModuleName, ""
            #End If
            Exit Function
        End If
    Else
        lIndex = mFirstRestartableIndex
    End If
End If

If enumType = EnumeratorTypes.Pending And lIndex = &HFFFFFFF Then
    If mFirstPendingIndex = LastIndex Then
        If enumAll Then
            enumType = EnumeratorTypes.Suspended
        Else
            pCookie = NoMoreItems
            #If trace Then
                gNullTracer.ExitProcedure "IEnumerable_NextItem", ProjectName, ModuleName, ""
            #End If
            Exit Function
        End If
    Else
        lIndex = mFirstPendingIndex
    End If
End If

If enumType = EnumeratorTypes.Suspended And lIndex = &HFFFFFFF Then
    If mFirstSuspendedIndex = LastIndex Then
        pCookie = NoMoreItems
        #If trace Then
            gNullTracer.ExitProcedure "IEnumerable_NextItem", ProjectName, ModuleName, ""
        #End If
        Exit Function
    Else
        lIndex = mFirstSuspendedIndex
    End If
End If

pItem = generateTaskSummary(lIndex)

If mTaskTable(lIndex).Next = LastIndex Or _
    mTaskTable(lIndex).Next = mFirstRunningIndex Or _
    mTaskTable(lIndex).Next = mFirstProcessedIndex Or _
    mTaskTable(lIndex).Next = mFirstRestartableIndex Or _
    mTaskTable(lIndex).Next = mFirstPendingIndex Or _
    mTaskTable(lIndex).Next = mFirstSuspendedIndex _
Then
    If enumAll Then
        If enumType = EnumeratorTypes.Runnable Then
            pCookie = EnumeratorTypes.Processed Or EnumeratorTypes.All
        ElseIf enumType = EnumeratorTypes.Processed Then
            pCookie = EnumeratorTypes.Restartable Or EnumeratorTypes.All
        ElseIf enumType = EnumeratorTypes.Restartable Then
            pCookie = EnumeratorTypes.Pending Or EnumeratorTypes.All
        ElseIf enumType = EnumeratorTypes.Pending Then
            pCookie = EnumeratorTypes.Suspended Or EnumeratorTypes.All
        Else
            pCookie = NoMoreItems
        End If
    Else
        pCookie = NoMoreItems
    End If
Else
    pCookie = mTaskTable(lIndex).Next Or enumType Or IIf(enumAll, EnumeratorTypes.All, 0)
End If

IEnumerable_NextItem = True

#If trace Then
    gNullTracer.ExitProcedure ProcName, ProjectName, ModuleName, ""
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IEnumerable_Remove(pCookie As Variant)
Const ProcName As String = "IEnumerable_Remove"
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Tasks cannot be removed via an enumerator"
End Sub

Private Function IEnumerable_Skip(ByVal numToSkip As Long, pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_Skip"
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Skip not suported"
End Function

Private Property Get IEnumerable_UpdateSequence() As Long
Const ProcName As String = "IEnumerable_UpdateSequence"
IEnumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "IStateChangeListener_Change", ProjectName, ModuleName, ""
#End If


If TypeOf ev.Source Is TimerListItem Then
    Dim tli As TimerListItem: Set tli = ev.Source
    Dim lIndex As Long: lIndex = tli.Data
    removeFromQueue mFirstSuspendedIndex, lIndex
    If ev.State = TimerListItemStates.TimerListItemStateExpired Then
        addToRestartableQueue lIndex
        tli.RemoveStateChangeListener Me
        Set mTaskTable(lIndex).SuspendTLI = Nothing
    ElseIf ev.State = TimerListItemStates.TimerListItemStateCancelled Then
        Set mTaskTable(lIndex).SuspendTLI = Nothing
    End If
End If
startScheduler

#If trace Then
    gNullTracer.ExitProcedure "IStateChangeListener_Change", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' mTaskInfoTimer Event Handlers
'@================================================================================

Private Sub mTaskInfoTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mTaskInfoTimer_TimerExpired"
On Error GoTo Err

#If trace Then
    gNullTracer.EnterProcedure "mTaskInfoTimer_TimerExpired", ProjectName, ModuleName, ""
#End If

If Not mSummaryLogger.IsLoggable(LogLevelDetail) Then Exit Sub

mSummaryLogger.Log LogLevelDetail, "Running Tasks ----------------------------------------------" & vbCrLf & _
                            GetRunnableTaskSummary
mSummaryLogger.Log LogLevelDetail, "Processed Tasks --------------------------------------------" & vbCrLf & _
                            GetProcessedTaskSummary
mSummaryLogger.Log LogLevelDetail, "Restartable Tasks ------------------------------------------" & vbCrLf & _
                            GetRestartableTaskSummary
mSummaryLogger.Log LogLevelDetail, "Pending Tasks ----------------------------------------------" & vbCrLf & _
                            GetPendingTaskSummary
mSummaryLogger.Log LogLevelDetail, "Suspended Tasks --------------------------------------------" & vbCrLf & _
                            GetSuspendedTaskSummary

#If trace Then
    gNullTracer.ExitProcedure "mTaskInfoTimer_TimerExpired", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get AverageInterScheduleWait() As Double
#If trace Then
    gNullTracer.EnterProcedure "AverageInterScheduleWait", ProjectName, ModuleName, ""
#End If
If mNumberOfInterSchedulePeriods <> 0 Then _
    AverageInterScheduleWait = mTotalInterScheduleElapsedTime / mNumberOfInterSchedulePeriods
mTotalInterScheduleElapsedTime = 0#
mNumberOfInterSchedulePeriods = 0
#If trace Then
    gNullTracer.ExitProcedure "AverageInterScheduleWait", ProjectName, ModuleName, ""
#End If
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the task summaries for all tasks.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get Enumerator() As Enumerator
Const ProcName As String = "Enumerator"
On Error GoTo Err

Set Enumerator = getEnumerator(EnumeratorTypes.All)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the task summaries for pending tasks.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get PendingTasksEnumerator() As Enumerator
Const ProcName As String = "PendingTasksEnumerator"
On Error GoTo Err

Set PendingTasksEnumerator = getEnumerator(EnumeratorTypes.Pending)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the task summaries for processed tasks.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get ProcessedTasksEnumerator() As Enumerator
Const ProcName As String = "ProcessedTasksEnumerator"
On Error GoTo Err

Set ProcessedTasksEnumerator = getEnumerator(EnumeratorTypes.Processed)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let QuantumMillisecs( _
                ByVal Value As Long)
Const ProcName As String = "QuantumMillisecs"
On Error GoTo Err

gAssertArgument gIsInteger(Value, MinimumQuantumMillisecs), "Value cannot be less than " & MinimumQuantumMillisecs
mQuantumMillisecs = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get QuantumMillisecs() As Long
QuantumMillisecs = mQuantumMillisecs
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the task summaries for restartable tasks.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get RestartableTasksEnumerator() As Enumerator
Const ProcName As String = "RestartableTasksEnumerator"
On Error GoTo Err

Set RestartableTasksEnumerator = getEnumerator(EnumeratorTypes.Restartable)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the task summaries for runnable tasks.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get RunnableTasksEnumerator() As Enumerator
Const ProcName As String = "RunnableTasksEnumerator"
On Error GoTo Err

Set RunnableTasksEnumerator = getEnumerator(EnumeratorTypes.Runnable)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let RunTasksAtLowerThreadPriority( _
                ByVal Value As Boolean)
mRunTasksAtLowerThreadPriority = Value
End Property

Friend Property Get RunTasksAtLowerThreadPriority() As Boolean
RunTasksAtLowerThreadPriority = mRunTasksAtLowerThreadPriority
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the task summaries for suspended tasks.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get SuspendedTasksEnumerator() As Enumerator
Const ProcName As String = "SuspendedTasksEnumerator"
On Error GoTo Err

Set SuspendedTasksEnumerator = getEnumerator(EnumeratorTypes.Suspended)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let TaskConcurrency( _
                ByVal Value As Long)
Const ProcName As String = "TaskConcurrency"
On Error GoTo Err

gAssertArgument gIsInteger(Value, MinimumTaskConcurrency), "Value must be greater than 0"

mTaskConcurrency = Value
startScheduler

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get TaskConcurrency() As Long
TaskConcurrency = mTaskConcurrency
End Property

Friend Property Let TaskSummaryLoggingIntervalSecs( _
                ByVal Value As Long)
Const ProcName As String = "TaskSummaryLoggingIntervalSecs"
On Error GoTo Err

gAssertArgument Value >= 0, "Value cannot be negative"

mTaskSummaryLoggingIntervalSecs = Value

If Not mTaskInfoTimer Is Nothing Then
    mTaskInfoTimer.StopTimer
    Set mTaskInfoTimer = Nothing
End If

If mTaskSummaryLoggingIntervalSecs > 0 Then
    Set mSummaryLogger = gLogManager.GetLogger(TaskSummaryInfoType)
    Set mTaskInfoTimer = gCreateIntervalTimer(mTaskSummaryLoggingIntervalSecs, ExpiryTimeUnitSeconds, mTaskSummaryLoggingIntervalSecs * 1000)
    mTaskInfoTimer.StartTimer
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Continue( _
                ByVal TC As TaskContext)
Const ProcName As String = "Continue"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "Continue", ProjectName, ModuleName, ""
#End If

cancelSuspendTimer TC.TaskId
removeFromQueue mFirstSuspendedIndex, TC.TaskId
addToRestartableQueue TC.TaskId
startScheduler

#If trace Then
    gNullTracer.ExitProcedure "Continue", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FinishNonRunnable( _
                ByVal TC As TaskContext)
Const ProcName As String = "FinishNonRunnable"
#If trace Then
    gNullTracer.EnterProcedure "FinishNonRunnable", ProjectName, ModuleName, ""
#End If

On Error GoTo Err

Dim id As Long: id = TC.TaskId
Select Case mTaskTable(id).State
Case TaskPrimitiveStates.Pending
    removeFromQueue mFirstPendingIndex, id
Case TaskPrimitiveStates.Suspended
    cancelSuspendTimer id
    removeFromQueue mFirstSuspendedIndex, id
Case TaskPrimitiveStates.Restartable
    removeFromQueue mFirstRestartableIndex, id
End Select
addToFreeList id

#If trace Then
    gNullTracer.ExitProcedure "FinishNonRunnable", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetPendingTaskSummary() As String
Const ProcName As String = "GetPendingTaskSummary"
On Error GoTo Err

GetPendingTaskSummary = getTaskSummaryInfoFromEnumerator(PendingTasksEnumerator)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetProcessedTaskSummary() As String
Const ProcName As String = "GetProcessedTaskSummary"
On Error GoTo Err

GetProcessedTaskSummary = getTaskSummaryInfoFromEnumerator(ProcessedTasksEnumerator)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetRunnableTaskSummary() As String
Const ProcName As String = "GetRunnableTaskSummary"
On Error GoTo Err

GetRunnableTaskSummary = getTaskSummaryInfoFromEnumerator(RunnableTasksEnumerator)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetRestartableTaskSummary() As String
Const ProcName As String = "GetRestartableTaskSummary"
On Error GoTo Err

GetRestartableTaskSummary = getTaskSummaryInfoFromEnumerator(RestartableTasksEnumerator)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetSuspendedTaskSummary() As String
Const ProcName As String = "GetSuspendedTaskSummary"
On Error GoTo Err

GetSuspendedTaskSummary = getTaskSummaryInfoFromEnumerator(SuspendedTasksEnumerator)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetTaskSummary() As String
Const ProcName As String = "GetTaskSummary"
On Error GoTo Err

GetTaskSummary = getTaskSummaryInfoFromEnumerator(Enumerator)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = getEnumerator(EnumeratorTypes.Runnable)
End Function

Friend Sub ScheduleTasks()
Const ProcName As String = "ScheduleTasks"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "ScheduleTasks", ProjectName, ModuleName, ""
#End If

Dim failPoint As String
failPoint = "100"

If Not gInitialised Then
    #If trace Then
        gNullTracer.ExitProcedure "ScheduleTasks", ProjectName, ModuleName, ""
    #End If
    Exit Sub
End If

mScheduleInvoked = False
mTotalInterScheduleElapsedTime = mTotalInterScheduleElapsedTime + mInterScheduleElapsedTimer.ElapsedTimeMicroseconds
mNumberOfInterSchedulePeriods = mNumberOfInterSchedulePeriods + 1

Dim hThread As Long
Dim currThreadPriority As Long
If mRunTasksAtLowerThreadPriority Then
    hThread = GetCurrentThread
    currThreadPriority = GetThreadPriority(hThread)
    SetThreadPriority hThread, THREAD_PRIORITY_BELOW_NORMAL
End If

failPoint = "300"

schedule

failPoint = "400"

If mRunTasksAtLowerThreadPriority Then
    SetThreadPriority hThread, currThreadPriority
End If

#If trace Then
    gNullTracer.ExitProcedure "ScheduleTasks", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, failPoint
End Sub

Friend Function StartTask( _
                ByVal pTarget As ITask, _
                ByVal pPriority As TaskPriorities, _
                Optional ByVal pName As String, _
                Optional ByVal pCookie As Variant, _
                Optional ByVal pData As Variant) As TaskController
Const ProcName As String = "StartTask"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "StartTask", ProjectName, ModuleName, ""
#End If

If pName = "" Then pName = gGenerateGUIDString
If IsMissing(pCookie) Then pCookie = Empty
If IsMissing(pData) Then pData = Empty

If gLogger.IsLoggable(LogLevelHighDetail) Then
    Dim s As String
    s = "Starting task: taskType=" & _
        TypeName(pTarget) & _
        " pCookie=" & gVariantToString(pCookie) & _
        " pData=" & gVariantToString(pData) & _
        " pName=""" & pName & """"

    gLogger.Log s, ProcName, ModuleName, , LogLevelHighDetail
End If

Dim lTaskContext As TaskContext
Set lTaskContext = New TaskContext

Dim lTaskController As TaskController
Set lTaskController = New TaskController
lTaskController.Initialise pTarget, pCookie, pName, lTaskContext

Dim id As Long
id = allocateEntry

' don't allow this task to Run until the next quantum to
' ensure it doesn't Run within this call
mTaskTable(id).Priority = pPriority
Set mTaskTable(id).TC = lTaskContext
mTaskTable(id).StartTime = gGetTimestampUtc
mTaskTable(id).LastRunTime = 0
mTaskTable(id).TotalCPUTime = 0

If mProcessing Then
    ' this task is being started from another running task. We add the new task
    ' to the restartable queue rather than the pending queue to ensure that
    ' the parent task doesn't have to wait for all the existing pending tasks
    ' to Progress before the 'subtask' gets started
    addToRestartableQueue id
Else
    addToPendingQueue id
End If

lTaskContext.Initialise pTarget, pData, id, pName, lTaskController

Set StartTask = lTaskController

startScheduler

#If trace Then
    gNullTracer.ExitProcedure "StartTask", ProjectName, ModuleName, ""
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addToFreeList( _
                ByVal pIndex As Long)
Const ProcName As String = "addToFreeList"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "addToFreeList", ProjectName, ModuleName, ""
#End If


Change
If mTaskTable(pIndex).State = TaskPrimitiveStates.Free Then
    #If trace Then
        gNullTracer.ExitProcedure "addToFreeList", ProjectName, ModuleName, ""
    #End If
    Exit Sub
End If
mTaskTable(pIndex).Next = mFirstFreeIndex
mFirstFreeIndex = pIndex
mTaskTable(pIndex).TC.State = TaskStateCompleted
Set mTaskTable(pIndex).TC = Nothing
mTaskTable(pIndex).TimeRemaining = 0
mTaskTable(pIndex).State = TaskPrimitiveStates.Free

#If trace Then
    gNullTracer.ExitProcedure "addToFreeList", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addToPendingQueue( _
                ByVal pIndex As Long)
Const ProcName As String = "addToPendingQueue"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "addToPendingQueue", ProjectName, ModuleName, ""
#End If


gAssert mTaskTable(pIndex).State <> TaskPrimitiveStates.Pending, "Task already in pending State"

addToQueue mFirstPendingIndex, pIndex
mTaskTable(pIndex).TC.State = TaskStatePending
mTaskTable(pIndex).State = TaskPrimitiveStates.Pending

#If trace Then
    gNullTracer.ExitProcedure "addToPendingQueue", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addToProcessedQueue( _
                ByVal pIndex As Long)
Const ProcName As String = "addToProcessedQueue"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "addToProcessedQueue", ProjectName, ModuleName, ""
#End If


gAssert mTaskTable(pIndex).State <> TaskPrimitiveStates.Processed, "Task already in processed State"

addToQueue mFirstProcessedIndex, pIndex
mTaskTable(pIndex).TC.State = TaskStateRunnable
mTaskTable(pIndex).State = TaskPrimitiveStates.Processed

#If trace Then
    gNullTracer.ExitProcedure "addToProcessedQueue", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addToQueue( _
                ByRef firstIndex As Long, _
                ByVal indexToAdd As Long)
Const ProcName As String = "addToQueue"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "addToQueue", ProjectName, ModuleName, ""
#End If


Change
If firstIndex = LastIndex Then
    firstIndex = indexToAdd
    mTaskTable(indexToAdd).Next = indexToAdd
    mTaskTable(indexToAdd).Prev = indexToAdd
Else
    mTaskTable(indexToAdd).Next = firstIndex
    mTaskTable(indexToAdd).Prev = mTaskTable(firstIndex).Prev
    mTaskTable(mTaskTable(firstIndex).Prev).Next = indexToAdd
    mTaskTable(firstIndex).Prev = indexToAdd
End If

#If trace Then
    gNullTracer.ExitProcedure "addToQueue", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addToRestartableQueue( _
                ByVal pIndex As Long)
Const ProcName As String = "addToRestartableQueue"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "addToRestartableQueue", ProjectName, ModuleName, ""
#End If


gAssert mTaskTable(pIndex).State <> TaskPrimitiveStates.Restartable, "Task already in restartable State"

addToQueue mFirstRestartableIndex, pIndex
mTaskTable(pIndex).TC.State = TaskStatePending
mTaskTable(pIndex).State = TaskPrimitiveStates.Restartable

#If trace Then
    gNullTracer.ExitProcedure "addToRestartableQueue", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addToRunningQueue( _
                ByVal pIndex As Long)
Const ProcName As String = "addToRunningQueue"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "addToRunningQueue", ProjectName, ModuleName, ""
#End If

gAssert mTaskTable(pIndex).State <> TaskPrimitiveStates.Running, "Task already in running State"

addToQueue mFirstRunningIndex, pIndex
mTaskTable(pIndex).TC.State = TaskStateRunnable
mTaskTable(pIndex).State = TaskPrimitiveStates.Running

#If trace Then
    gNullTracer.ExitProcedure "addToRunningQueue", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addToSuspendedQueue( _
                ByVal pIndex As Long)
Const ProcName As String = "addToSuspendedQueue"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "addToSuspendedQueue", ProjectName, ModuleName, ""
#End If

gAssert mTaskTable(pIndex).State <> TaskPrimitiveStates.Suspended, "Task already in suspended State"

addToQueue mFirstSuspendedIndex, pIndex
mTaskTable(pIndex).TC.State = TaskStateSuspended
mTaskTable(pIndex).State = TaskPrimitiveStates.Suspended

#If trace Then
    gNullTracer.ExitProcedure "addToSuspendedQueue", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub adjustRunnableTasks()
Const ProcName As String = "adjustRunnableTasks"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "adjustRunnableTasks", ProjectName, ModuleName, ""
#End If


If mNumRunnableTasks >= mTaskConcurrency Then
    #If trace Then
        gNullTracer.ExitProcedure "adjustRunnableTasks", ProjectName, ModuleName, ""
    #End If
    Exit Sub
End If

Dim lIndex As Long
Do While mNumRunnableTasks < mTaskConcurrency
    lIndex = getRestartableQueueItem
    If lIndex = LastIndex Then Exit Do
    mNumRunnableTasks = mNumRunnableTasks + 1
    If mTaskTable(lIndex).TimeRemaining > 0 Then
        addToRunningQueue lIndex
    Else
        addToProcessedQueue lIndex
    End If
Loop

If mNumRunnableTasks >= mTaskConcurrency Then
#If trace Then
    gNullTracer.ExitProcedure "adjustRunnableTasks", ProjectName, ModuleName, ""
#End If
    Exit Sub
End If

Do While mNumRunnableTasks < mTaskConcurrency
    lIndex = getPendingQueueItem
    If lIndex = LastIndex Then Exit Do
    mNumRunnableTasks = mNumRunnableTasks + 1
    addToRunningQueue lIndex
    mTaskTable(lIndex).TimeRemaining = allocateTime(mTaskTable(lIndex).Priority)
Loop

#If trace Then
    gNullTracer.ExitProcedure "adjustRunnableTasks", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function allocateEntry() As Long
Const ProcName As String = "allocateEntry"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "allocateEntry", ProjectName, ModuleName, ""
#End If


Change
If mFirstFreeIndex <> LastIndex Then
    allocateEntry = mFirstFreeIndex
    mFirstFreeIndex = mTaskTable(mFirstFreeIndex).Next
Else
    If mTaskTableNextIndex > UBound(mTaskTable) Then
        ReDim Preserve mTaskTable(2 * (UBound(mTaskTable) + 1) - 1) As TaskTableEntry
        If gLogger.IsLoggable(LogLevelHighDetail) Then _
            gLogger.Log "Increased mTaskTable size", ProcName, ModuleName, LogLevelMediumDetail, CStr(UBound(mTaskTable) + 1)
    End If
    allocateEntry = mTaskTableNextIndex
    mTaskTableNextIndex = mTaskTableNextIndex + 1
End If

#If trace Then
    gNullTracer.ExitProcedure "allocateEntry", ProjectName, ModuleName, ""
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function allocateTime( _
                ByVal Priority As TaskPriorities) As Long
Const ProcName As String = "allocateTime"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "allocateTime", ProjectName, ModuleName, ""
#End If

Dim divisor As Long
Select Case Priority
Case PriorityNormal
    divisor = 2
Case PriorityLow
    divisor = 4
Case PriorityHigh
    divisor = 1
End Select

allocateTime = 1000 * mQuantumMillisecs / divisor

#If trace Then
    gNullTracer.ExitProcedure "allocateTime", ProjectName, ModuleName, ""
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub cancelSuspendTimer( _
                ByVal id As Long)
Const ProcName As String = "cancelSuspendTimer"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "cancelSuspendTimer", ProjectName, ModuleName, ""
#End If

If Not mTaskTable(id).SuspendTLI Is Nothing Then
    mTimerList.Remove mTaskTable(id).SuspendTLI
    Set mTaskTable(id).SuspendTLI = Nothing
End If

#If trace Then
    gNullTracer.ExitProcedure "cancelSuspendTimer", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Change()
mUpdateSequence = mUpdateSequence + 1
End Sub

Private Function getEnumerator(ByVal enumType As EnumeratorTypes) As Enumerator
Const ProcName As String = "getEnumerator"
On Error GoTo Err

Dim lEnum As Enumerator: Set lEnum = New Enumerator
lEnum.Initialise Me, enumType Or &HFFFFFFF
Set getEnumerator = lEnum

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getPendingQueueItem() As Long
Const ProcName As String = "getPendingQueueItem"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "getPendingQueueItem", ProjectName, ModuleName, ""
#End If

getPendingQueueItem = getQueueItem(mFirstPendingIndex)

#If trace Then
    gNullTracer.ExitProcedure "getPendingQueueItem", ProjectName, ModuleName, ""
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getProcessedQueueItem() As Long
Const ProcName As String = "getProcessedQueueItem"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "getProcessedQueueItem", ProjectName, ModuleName, ""
#End If

getProcessedQueueItem = getQueueItem(mFirstProcessedIndex)

#If trace Then
    gNullTracer.ExitProcedure "getProcessedQueueItem", ProjectName, ModuleName, ""
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getQueueItem( _
                ByRef firstIndex) As Long
Const ProcName As String = "getQueueItem"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "getQueueItem", ProjectName, ModuleName, ""
#End If

getQueueItem = firstIndex
If firstIndex = LastIndex Then
    #If trace Then
        gNullTracer.ExitProcedure "getQueueItem", ProjectName, ModuleName, ""
    #End If
    Exit Function
End If

Change
If mTaskTable(firstIndex).Next = firstIndex Then
    ' there was only one entry in the queue
    firstIndex = LastIndex
Else
    mTaskTable(mTaskTable(firstIndex).Next).Prev = mTaskTable(firstIndex).Prev
    mTaskTable(mTaskTable(firstIndex).Prev).Next = mTaskTable(firstIndex).Next
    firstIndex = mTaskTable(firstIndex).Next
End If

#If trace Then
    gNullTracer.ExitProcedure "getQueueItem", ProjectName, ModuleName, ""
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getRestartableQueueItem() As Long
Const ProcName As String = "getRestartableQueueItem"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "getRestartableQueueItem", ProjectName, ModuleName, ""
#End If

getRestartableQueueItem = getQueueItem(mFirstRestartableIndex)

#If trace Then
    gNullTracer.ExitProcedure "getRestartableQueueItem", ProjectName, ModuleName, ""
#End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateTaskSummary( _
                ByVal pIndex As Long) As TaskSummary
Const ProcName As String = "generateTaskSummary"
On Error GoTo Err

generateTaskSummary.Name = mTaskTable(pIndex).TC.Name
generateTaskSummary.ClassName = TypeName(mTaskTable(pIndex).TC.Target)
generateTaskSummary.Priority = gTaskPrioritiesToString(mTaskTable(pIndex).Priority)
generateTaskSummary.State = TaskPrimitiveStatesToString(mTaskTable(pIndex).State)
generateTaskSummary.StartTime = mTaskTable(pIndex).StartTime
generateTaskSummary.LastRunTime = mTaskTable(pIndex).LastRunTime
generateTaskSummary.TotalCPUTime = mTaskTable(pIndex).TotalCPUTime

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getTaskSummaryInfoFromEnumerator( _
                ByVal en As Enumerator) As String
Const ProcName As String = "getTaskSummaryInfoFromEnumerator"
On Error GoTo Err

Dim s As String
Do While en.MoveNext
    Dim ts As TaskSummary
    ts = en.Current
    s = s & "Name: " & ts.Name & _
        "; Class: " & ts.ClassName & _
        "; Priority: " & ts.Priority & _
        "; State: " & ts.State & _
        "; Start time: " & gFormatTimestamp(ts.StartTime, TimestampDateAndTimeISO8601) & _
        "; Last Run time: " & gFormatTimestamp(ts.LastRunTime, TimestampDateAndTimeISO8601) & _
        "; CPU time (ms): " & Format(ts.TotalCPUTime, "0.0") & vbCrLf
        
Loop
getTaskSummaryInfoFromEnumerator = s

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub postScheduleTasksMessage()
Const ProcName As String = "postScheduleTasksMessage"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "postScheduleTasksMessage", ProjectName, ModuleName, ""
#End If

gPostUserMessage UserMessageScheduleTasks, 0, 0
mScheduleInvoked = True
mInterScheduleElapsedTimer.StartTiming

#If trace Then
    gNullTracer.ExitProcedure "postScheduleTasksMessage", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeFromQueue( _
                ByRef firstIndex As Long, _
                ByVal indexToRemove As Long)
Const ProcName As String = "removeFromQueue"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "removeFromQueue", ProjectName, ModuleName, ""
#End If

Change

If mTaskTable(indexToRemove).Next = indexToRemove Then
    ' this was the only item in the queue
    firstIndex = LastIndex
Else
    If firstIndex = indexToRemove Then
        firstIndex = mTaskTable(indexToRemove).Next
    End If
    mTaskTable(mTaskTable(indexToRemove).Prev).Next = mTaskTable(indexToRemove).Next
    mTaskTable(mTaskTable(indexToRemove).Next).Prev = mTaskTable(indexToRemove).Prev
End If

#If trace Then
    gNullTracer.ExitProcedure "removeFromQueue", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub runTask( _
                ByVal pIndex As Long, _
                ByVal Timestamp As Date)
Const ProcName As String = "runTask"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "runTask", ProjectName, ModuleName, ""
#End If

Dim failPoint As String
failPoint = 100

gAssert mTaskTable(pIndex).State = TaskPrimitiveStates.Running, "Task already in running State"

Dim timeSlotMicroseconds As Long
timeSlotMicroseconds = DefaultTimeslotMillisecs * 1000
timeSlotMicroseconds = IIf(mTaskTable(pIndex).TimeRemaining < timeSlotMicroseconds, _
                                    mTaskTable(pIndex).TimeRemaining, _
                                    timeSlotMicroseconds)
timeSlotMicroseconds = IIf(mQuantumRemainingMicrosecs < timeSlotMicroseconds, _
                                    mQuantumRemainingMicrosecs, _
                                    timeSlotMicroseconds)

mTaskTable(pIndex).LastRunTime = Timestamp


failPoint = 200

Dim trxs As TaskRunExitStatus
trxs = mTaskTable(pIndex).TC.Run(timeSlotMicroseconds)


failPoint = 300

mQuantumRemainingMicrosecs = mQuantumRemainingMicrosecs - trxs.ElapsedTimeMicrosecs
mTaskTable(pIndex).TimeRemaining = mTaskTable(pIndex).TimeRemaining - trxs.ElapsedTimeMicrosecs
mTaskTable(pIndex).TotalCPUTime = mTaskTable(pIndex).TotalCPUTime + (trxs.ElapsedTimeMicrosecs / 1000)

If trxs.Finished Then

    failPoint = 400

    mNumRunnableTasks = mNumRunnableTasks - 1
    removeFromQueue mFirstRunningIndex, pIndex
    addToFreeList pIndex
    adjustRunnableTasks
ElseIf trxs.SleepMillisecs <> 0 Then

    failPoint = 500

    mNumRunnableTasks = mNumRunnableTasks - 1
    removeFromQueue mFirstRunningIndex, pIndex
    addToSuspendedQueue pIndex
    If trxs.SleepMillisecs > 0 Then
        Set mTaskTable(pIndex).SuspendTLI = mTimerList.Add(pIndex, trxs.SleepMillisecs, ExpiryTimeUnitMilliseconds)
        mTaskTable(pIndex).SuspendTLI.AddStateChangeListener Me
    End If
    adjustRunnableTasks
ElseIf mTaskTable(pIndex).TimeRemaining <= 0 Then

    failPoint = 600

    removeFromQueue mFirstRunningIndex, pIndex
    addToProcessedQueue pIndex
Else
    mFirstRunningIndex = mTaskTable(pIndex).Next
End If

#If trace Then
    gNullTracer.ExitProcedure "runTask", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, failPoint
End Sub

Private Sub schedule()
Const ProcName As String = "schedule"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "schedule", ProjectName, ModuleName, ""
#End If

Dim failPoint As String

mProcessing = True

mQuantumRemainingMicrosecs = mQuantumMillisecs * 1000

adjustRunnableTasks

Dim Timestamp As Date: Timestamp = gGetTimestampUtc

Dim lIndex As Long
Do While mQuantumRemainingMicrosecs > 0 And _
    (mFirstRunningIndex <> LastIndex Or mFirstProcessedIndex <> LastIndex)
    
    failPoint = 100
    
    If mFirstRunningIndex = LastIndex Then
        ' all entries have used up their time allocation, so renew the allocations
        lIndex = getProcessedQueueItem
        Do While lIndex <> LastIndex
            addToRunningQueue lIndex
            mTaskTable(lIndex).TimeRemaining = mTaskTable(lIndex).TimeRemaining + allocateTime(mTaskTable(lIndex).Priority)
            lIndex = getProcessedQueueItem
        Loop
    End If
    
    failPoint = 200
    
    lIndex = mFirstRunningIndex
    
    If mTaskTable(lIndex).State <> TaskPrimitiveStates.Running Then Stop
        
    If mTaskTable(lIndex).TimeRemaining <= 0 Then
        removeFromQueue mFirstRunningIndex, lIndex
        addToProcessedQueue lIndex
    Else
        runTask lIndex, Timestamp
    End If
Loop

failPoint = 400

If mFirstRunningIndex <> LastIndex Or _
    mFirstPendingIndex <> LastIndex Or _
    mFirstProcessedIndex <> LastIndex _
Then
    postScheduleTasksMessage
End If

mProcessing = False

#If trace Then
    gNullTracer.ExitProcedure "schedule", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, failPoint
End Sub

Private Sub startScheduler()
Const ProcName As String = "startScheduler"
On Error GoTo Err
#If trace Then
    gNullTracer.EnterProcedure "startScheduler", ProjectName, ModuleName, ""
#End If

If mProcessing Then
    #If trace Then
        gNullTracer.ExitProcedure "startScheduler", ProjectName, ModuleName, ""
    #End If
    Exit Sub
End If

adjustRunnableTasks

If mScheduleInvoked Then
    #If trace Then
        gNullTracer.ExitProcedure "startScheduler", ProjectName, ModuleName, ""
    #End If
    Exit Sub
End If

If mFirstRunningIndex <> LastIndex Or _
    mFirstProcessedIndex <> LastIndex _
Then
    postScheduleTasksMessage
End If

#If trace Then
    gNullTracer.ExitProcedure "startScheduler", ProjectName, ModuleName, ""
#End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function TaskPrimitiveStatesToString( _
                ByVal pState As TaskPrimitiveStates) As String
Select Case pState
Case TaskPrimitiveStates.Free
    TaskPrimitiveStatesToString = "Free"
Case TaskPrimitiveStates.Pending
    TaskPrimitiveStatesToString = "Pending"
Case TaskPrimitiveStates.Processed
    TaskPrimitiveStatesToString = "Processed"
Case TaskPrimitiveStates.Running
    TaskPrimitiveStatesToString = "Running"
Case TaskPrimitiveStates.Suspended
    TaskPrimitiveStatesToString = "Suspended"
Case TaskPrimitiveStates.Restartable
    TaskPrimitiveStatesToString = "Restartable"
End Select
End Function



