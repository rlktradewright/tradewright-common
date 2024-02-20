VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TaskContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' An object of this class is passed to a <code>Task</code> object when the task is
' started. It provides means for the task to notify its progress and completion,
' and to Suspend its execution for a specified period.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IDeferredAction

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "TaskContext"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mTaskController                     As TaskController

Private mTaskId                             As Long
Private mName                               As String

Private mTarget                             As ITask

Private mElapsedTimer                       As ElapsedTimer

Private mSleepMillisecs                     As Long

Private mYieldCount                         As Long

Private mCancelPending                      As Boolean

Private mData                               As Variant

Private mState                              As TaskStates

Private mAppNotifiedFinish                  As Boolean

Private mResult                             As Variant

Private mErrorCode                          As Long
Private mErrorMessage                       As String
Private mErrorSource                        As String

Private mCancelled                          As Boolean

''
' This variable is needed to ensure that this object does not disappear while
' it is still on the call stack. That could happen if the application sets its
' task context reference in the same procedure as calling Finish, because the
' Finish processing causes the task manager's reference to be nothinged. If that
' happens, when control returns from the application to the Run method, an error
' occurs because this object has already been terminated.
'
'@/
Private mSelfRef                            As TaskContext

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSelfRef = Me
End Sub

Private Sub Class_Terminate()
'Debug.Print "TaskContext terminated"
End Sub

'@================================================================================
' DeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run( _
                ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"

On Error GoTo Err

If mErrorCode = 0 Then
    mTaskController.FireCompletion mResult, mCancelled, 0, "", ""
Else
    mTaskController.FireCompletion Empty, False, mErrorCode, mErrorMessage, mErrorSource
End If
Set mTaskController = Nothing
Set mTarget = Nothing
Set mSelfRef = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

''
' Returns whether cancellation has been requested.
'
' @return
'   <code>True</code> if cancellation has been requested.
'@/
Public Property Get CancelPending() As Boolean
CancelPending = mCancelPending
End Property

''
' Returns the data passed to <code>StartTask</code>.
'
' @return
'   The data passed to <code>StartTask</code>.
'@/
Public Property Get Data() As Variant
gSetVariant Data, mData
End Property

''
' Returns the task name.
'
' @return
'   The task's name.
'@/
Public Property Get Name() As String
Name = mName
End Property

Friend Property Let State( _
                ByVal Value As TaskStates)
Const ProcName As String = "State"

On Error GoTo Err

If Value = mState Then Exit Property
mState = Value
mTaskController.fireStateChange mState
If mState = TaskStateCompleted Then DeferredActionManager.DeferAction Me, Empty, 0

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the task's current State.
'
' @return
'   The task's current.
'@/
Public Property Get State() As TaskStates
State = mState
End Property

Friend Property Get Target() As ITask
Set Target = mTarget
End Property

''
' Returns the task identifier.
'
' @return
'   The task's identifier.
'@/
Public Property Get TaskId() As Long
TaskId = mTaskId
End Property

''
' Returns the number of times the task has yielded.
'
' @return
'   The number of times the task has yielded.
'@/
Public Property Get YieldCount() As Long
YieldCount = mYieldCount
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Cancel()
Const ProcName As String = "Cancel"

On Error GoTo Err

mCancelPending = True
If mState = TaskStatePending Or mState = TaskStateSuspended Then mTarget.Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Ends the suspension of this task, if it is currently suspended. If it is not
' currently suspended, or if <code>error</code> or <code>finish</code> have
' already been called, calling this method has no effect.
'
'@/
Public Sub Continue()
Const ProcName As String = "Continue"

On Error GoTo Err

If mState = TaskStateSuspended Then
    GTasks.gTaskManager.Continue Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' The task may call this method if an error occurs in its processing.
'
' @remarks
'   After calling <code>error</code>, no further calls are made to the
'   task's <code>run</code> method. The relevant <code>TaskController</code>
'   object notifies any <code>TaskCompletionListener</code>s that the task has
'   Finished with errors.
'
'   The task may call <code>error</code> either from within the <code>run</code>
'   method or from within any event handler.
'
'   The task should exit from the <code>run</code> method or the calling event handler
'   immediately after calling this method.
' @param pCode
'   Identifies the error that has occurred.
' @param pMessage
'   Describes the error.
' @param pSource
'   Indicates where the error occurred.
'@/
Public Sub Error( _
                ByVal pCode As Long, _
                ByVal pMessage As String, _
                Optional ByVal pSource As String)
Const ProcName As String = "Error"
On Error GoTo Err

If mState = TaskStateCompleted Then Exit Sub

mErrorCode = pCode
mErrorMessage = pMessage
mErrorSource = pSource

Finish Empty, False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' The task must call this method when it has completed.
'
' @remarks
'   After calling <code>finish</code>, no further calls are made to the
'   task's <code>run</code> method. The relevant <code>TaskController</code>
'   object notifies any <code>TaskCompletionListener</code>s that the task is
'   complete.
'
'   The task may call <code>finish</code> either from within the <code>run</code>
'   method or from within any event handler.
'
'   The task should exit from the <code>run</code> method or the calling event handler
'   immediately after calling this method.
' @param pResult
'   The result of processing the task (if any). It is passed in the
'   <code>TaskCompletionEventData</code> to <code>TaskCompletionListener</code>s
'   (if any are registered).
' @param pCancelled
'   Indicates whether the task was cancelled. It is passed in the
'   <code>TaskCompletionEventData</code> to <code>TaskCompletionListener</code>s
'   (if any are registered).
'@/
Public Sub Finish( _
                ByVal pResult As Variant, _
                ByVal pCancelled As Boolean)
Const ProcName As String = "Finish"
On Error GoTo Err

Dim failPoint As String

If mState = TaskStateCompleted Then Exit Sub

'Debug.Print "Finished " & mName & ": yield count=" & mYieldCount

failPoint = 100

gSetVariant mResult, pResult

mCancelled = pCancelled
mAppNotifiedFinish = True

failPoint = 200

If mState <> TaskStateRunnable Then
    GTasks.gTaskManager.FinishNonRunnable Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, failPoint
End Sub

Friend Sub Initialise( _
                ByVal pTarget As ITask, _
                ByVal pData As Variant, _
                ByVal id As Long, _
                ByVal pName As String, _
                ByVal pTaskController As TaskController)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mTaskController = pTaskController
Set mTarget = pTarget
gSetVariant mData, pData
mTaskId = id
mName = pName
Set mElapsedTimer = New ElapsedTimer
mTarget.TaskContext = Me
'Debug.Print "Starting " & mName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' The task may call this method to notify the occurrence of events of interest.
'
' @remarks
'   The relevant <code>TaskController</code> object notifies the occurrence to any
'   registered <code>TaskNotificationListener</code>s.
'
'    Details of all notifications that might occur should be included in documentation
'   for Task classes.
' @param EventCode
'   An identifier for the event.
' @param EventMessage
'   A description of the event.
'@/
Public Sub Notify( _
                ByVal EventCode As Long, _
                ByVal EventMessage As String)
Const ProcName As String = "Notify"
On Error GoTo Err

mTaskController.FireNotification EventCode, EventMessage

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' The task may call this method to notify its progress.
'
' @remarks
'   The relevant <code>TaskController</code> object notifies the progress to any
'   registered <code>TaskProgressListener</code>s.
'
'   Progress notifications are typically used to update user interface controls
'   such as progress bars. Since UI updates are relatively expensive operations,
'   it is prudent not to notify progress very frequently. Notifications at 1 percent
'   increments are more than adequate.
' @param pProgress
'   The percentage of the task's processing that has been completed.
' @param pInterimResult
'   An intermediate or partial result of the processing.
'@/
Public Sub Progress( _
                ByVal pProgress As Single, _
                ByRef pInterimResult As Variant)
Const ProcName As String = "Progress"
On Error GoTo Err

mTaskController.FireProgress pProgress, pInterimResult

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'/**
' Returns True if task is still runnable (ie not Finished or suspended)
'*/
Friend Function Run( _
                ByVal runForMicrosecs As Long) As TaskRunExitStatus
Const ProcName As String = "Run"
On Error GoTo Err

Dim failPoint As String
Dim lRunning As Boolean
Dim et As Single

mSleepMillisecs = 0

mElapsedTimer.StartTiming

Do
    lRunning = True
    On Error GoTo AppErr
    mTarget.Run
    On Error GoTo Err
    lRunning = False
    et = mElapsedTimer.ElapsedTimeMicroseconds
Loop Until mAppNotifiedFinish Or _
            mSleepMillisecs <> 0 Or _
            et > runForMicrosecs

Run.Finished = mAppNotifiedFinish
Run.ElapsedTimeMicrosecs = et
Run.SleepMillisecs = mSleepMillisecs

mYieldCount = mYieldCount + 1

Exit Function

AppErr:
Dim lErrorCode As Long: lErrorCode = Err.number
Dim lErrorMessage As String: lErrorMessage = Err.Description
Dim lErrorSource As String: lErrorSource = Err.Source

gLogger.Log "Error " & lErrorCode & " in task (" & TypeName(mTarget) & "): Name=" & Name, ProcName, ModuleName

gNotifyUnhandledError ProcName, ModuleName, , lErrorCode, lErrorMessage, lErrorSource

' if we get here, the application has handled the error in some way, so we
' cancel the task that errored and carry on

Resume CancelAfterError

CancelAfterError:
On Error GoTo Err
Finish Empty, True

Run.Finished = mAppNotifiedFinish
Run.ElapsedTimeMicrosecs = 0
Run.SleepMillisecs = 0

mYieldCount = mYieldCount + 1

Exit Function

Err:
If Err.number = VBErrorCodes.VbErrOverflow And Not lRunning Then
    ' elapsed time has overflowed - this usually happens when the task is being debugged
    ' and has sat on a breakpoint for too long
    et = 999999999
    Resume Next
End If
gHandleUnexpectedError ProcName, ModuleName, failPoint
End Function

''
' Suspends the task for the specified number of milliseconds.
'
' @remarks
'   While the task is suspended, no calls are made to its <code>run</code> method.
'
'   The task should exit from the <code>run</code> method immediately
'   after calling this method.
'
'   Calling this method from anywhere other than in the <code>run</code> method
'   has no effect.
' @param intervalMillisecs
'   The number of milliseconds that the task is to be suspended.
'
'   If this value is less than zero, the task is suspended until it calls the
'    <code>Continue</code> method. Such a value should only be specfied where the
'   task is expecting to be awoken by some external mechanism so that it can
'   call <code>Continue</code>, otherwise it will remain suspended indefinitely.
'
'   Note that after calling this method, the <code>Status</code> property does
'   not return <code>TaskStateSuspended</code> until the <code>Run</code> method has exited.
'
'   Calling <code>Finish</code> after calling <code>Suspend</code> and before the
'   <code>Run</code> method has exited cancels the effect of the call to <code>suspend</code>.
'@/
Public Sub Suspend( _
                ByVal intervalMillisecs As Long)
mSleepMillisecs = intervalMillisecs
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



