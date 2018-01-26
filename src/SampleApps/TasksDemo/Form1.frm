VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tasks Demo"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   4650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Task Manager Parameters"
      Height          =   1335
      Left            =   120
      TabIndex        =   61
      Top             =   120
      Width           =   4335
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   1050
         Left            =   120
         ScaleHeight     =   1050
         ScaleWidth      =   4095
         TabIndex        =   62
         Top             =   240
         Width           =   4095
         Begin VB.CommandButton SetParamsButton 
            Caption         =   "Set"
            Enabled         =   0   'False
            Height          =   285
            Left            =   3120
            TabIndex        =   34
            Top             =   720
            Width           =   615
         End
         Begin VB.CheckBox LowerPriorityCheck 
            Caption         =   "Run tasks at lower thread priority"
            Height          =   255
            Left            =   0
            TabIndex        =   33
            Top             =   720
            Width           =   2655
         End
         Begin VB.TextBox ConcurrencyText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            TabIndex        =   32
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox QuantumText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3120
            TabIndex        =   31
            Top             =   0
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Max concurrency"
            Height          =   255
            Left            =   0
            TabIndex        =   64
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Scheduling quantum (milliseconds)"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   63
            Top             =   0
            Width           =   2895
         End
      End
   End
   Begin VB.CommandButton SummariesButton 
      Caption         =   "Task summaries"
      Height          =   375
      Left            =   3000
      TabIndex        =   30
      Top             =   6480
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Normal priority"
      Height          =   1695
      Left            =   120
      TabIndex        =   43
      Top             =   3120
      Width           =   4335
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1365
         Left            =   120
         ScaleHeight     =   1365
         ScaleWidth      =   4095
         TabIndex        =   44
         Top             =   240
         Width           =   4095
         Begin VB.CommandButton PauseButton 
            Caption         =   "Pause"
            Enabled         =   0   'False
            Height          =   255
            Index           =   6
            Left            =   840
            TabIndex        =   16
            Top             =   1080
            Width           =   735
         End
         Begin VB.CommandButton PauseButton 
            Caption         =   "Pause"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   840
            TabIndex        =   15
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton PauseButton 
            Caption         =   "Pause"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   14
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton PauseButton 
            Caption         =   "Pause"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   840
            TabIndex        =   13
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Height          =   255
            Index           =   6
            Left            =   0
            TabIndex        =   6
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox CountText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Height          =   255
            Index           =   5
            Left            =   0
            TabIndex        =   5
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox CountText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   50
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Height          =   255
            Index           =   4
            Left            =   0
            TabIndex        =   4
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox CountText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Height          =   255
            Index           =   3
            Left            =   0
            TabIndex        =   3
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox CountText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   45
            TabStop         =   0   'False
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Index           =   3
            Left            =   1680
            TabIndex        =   23
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Index           =   4
            Left            =   1680
            TabIndex        =   24
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Index           =   5
            Left            =   1680
            TabIndex        =   25
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Index           =   6
            Left            =   1680
            TabIndex        =   26
            Top             =   1080
            Width           =   735
         End
         Begin VB.TextBox ProgressText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   3
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox ProgressText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   4
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   48
            TabStop         =   0   'False
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox ProgressText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   5
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   49
            TabStop         =   0   'False
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox ProgressText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   6
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   52
            TabStop         =   0   'False
            Top             =   1080
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Low priority"
      Height          =   1335
      Left            =   120
      TabIndex        =   53
      Top             =   5040
      Width           =   4335
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   1005
         Left            =   120
         ScaleHeight     =   1005
         ScaleWidth      =   4095
         TabIndex        =   54
         Top             =   240
         Width           =   4095
         Begin VB.CommandButton PauseButton 
            Caption         =   "Pause"
            Enabled         =   0   'False
            Height          =   255
            Index           =   9
            Left            =   840
            TabIndex        =   19
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton PauseButton 
            Caption         =   "Pause"
            Enabled         =   0   'False
            Height          =   255
            Index           =   8
            Left            =   840
            TabIndex        =   18
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton PauseButton 
            Caption         =   "Pause"
            Enabled         =   0   'False
            Height          =   255
            Index           =   7
            Left            =   840
            TabIndex        =   17
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Height          =   255
            Index           =   9
            Left            =   0
            TabIndex        =   9
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox CountText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   59
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Height          =   255
            Index           =   8
            Left            =   0
            TabIndex        =   8
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox CountText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   57
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Height          =   255
            Index           =   7
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox CountText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   55
            TabStop         =   0   'False
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Index           =   7
            Left            =   1680
            TabIndex        =   27
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Index           =   8
            Left            =   1680
            TabIndex        =   28
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Index           =   9
            Left            =   1680
            TabIndex        =   29
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox ProgressText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   7
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox ProgressText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   8
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   58
            TabStop         =   0   'False
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox ProgressText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   9
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   60
            TabStop         =   0   'False
            Top             =   720
            Width           =   615
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "High priority"
      Height          =   1335
      Left            =   120
      TabIndex        =   35
      Top             =   1560
      Width           =   4335
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1005
         Left            =   120
         ScaleHeight     =   1005
         ScaleWidth      =   4095
         TabIndex        =   36
         Top             =   240
         Width           =   4095
         Begin VB.CommandButton PauseButton 
            Caption         =   "Pause"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   840
            TabIndex        =   12
            Top             =   720
            Width           =   735
         End
         Begin VB.CommandButton PauseButton 
            Caption         =   "Pause"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   840
            TabIndex        =   11
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton PauseButton 
            Caption         =   "Pause"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   10
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   2
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox CountText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   1
            Top             =   360
            Width           =   735
         End
         Begin VB.TextBox CountText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton StartButton 
            Caption         =   "Start"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   0
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox CountText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   37
            TabStop         =   0   'False
            Top             =   0
            Width           =   855
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Index           =   0
            Left            =   1680
            TabIndex        =   20
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   21
            Top             =   360
            Width           =   735
         End
         Begin VB.CommandButton CancelButton 
            Caption         =   "Cancel"
            Enabled         =   0   'False
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   22
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox ProgressText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   0
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   0
            Width           =   615
         End
         Begin VB.TextBox ProgressText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   1
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox ProgressText 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   2
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   720
            Width           =   615
         End
      End
   End
   Begin VB.Label SchedulingIntervalLabel 
      Height          =   255
      Left            =   1800
      TabIndex        =   66
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Scheduling interval"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   65
      Top             =   6480
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements ITaskCompletionListener
Implements ITaskProgressListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mTasks(9) As CounterTask
Private mTaskControllers(9) As TaskController

Private WithEvents mSchedulingIntervalTimer As IntervalTimer
Attribute mSchedulingIntervalTimer.VB_VarHelpID = -1

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
InitialiseCommonControls
InitialiseTWUtilities
ApplicationGroupName = "TradeWright"
ApplicationName = "TasksDemo"
SetupDefaultLogging Command
#If trace = 1 Then
    EnableTracing ""
#End If
End Sub

Private Sub Form_Load()
QuantumText = TaskQuantumMillisecs
ConcurrencyText = TaskConcurrency
LowerPriorityCheck.value = IIf(RunTasksAtLowerThreadPriority, vbChecked, vbUnchecked)
Set mSchedulingIntervalTimer = CreateIntervalTimer(10, ExpiryTimeUnitSeconds, 10000)
mSchedulingIntervalTimer.StartTimer
End Sub

Private Sub Form_Unload(Cancel As Integer)
TerminateTWUtilities
End Sub

'================================================================================
' iTaskCompletionListener Interface Members
'================================================================================

Private Sub iTaskCompletionListener_TaskCompleted( _
                ByRef ev As TaskCompletionEventData)
Dim Index As Long
Index = ev.Cookie
StartButton(Index).Enabled = True
PauseButton(Index).Enabled = False
CancelButton(Index).Enabled = False

If ev.Cancelled Then
    CountText(Index).ForeColor = vbRed
End If

Set mTasks(Index) = Nothing
Set mTaskControllers(Index) = Nothing
End Sub

'================================================================================
' ITaskProgressListener Interface Members
'================================================================================

Private Sub ITaskProgressListener_Progress( _
                ByRef ev As TaskProgressEventData)
Dim Index As Long
Index = ev.Cookie
ProgressText(Index) = ev.Progress & "%"
CountText(Index).Refresh
ProgressText(Index).Refresh
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub CancelButton_Click(Index As Integer)
LogMessage "CancelButton_Click: " & Index
mTaskControllers(Index).CancelTask
End Sub

Private Sub ConcurrencyText_Change()
validateParams
End Sub

Private Sub PauseButton_Click(Index As Integer)
LogMessage "PauseButton_Click: " & Index
mTasks(Index).pause
End Sub

Private Sub QuantumText_Change()
validateParams
End Sub

Private Sub SetParamsButton_Click()
TaskQuantumMillisecs = QuantumText
TaskConcurrency = ConcurrencyText
RunTasksAtLowerThreadPriority = (LowerPriorityCheck = vbChecked)
End Sub

Private Sub StartButton_Click( _
                Index As Integer)
Dim lTask As New CounterTask
Dim priority As TaskPriorities

LogMessage "StartButton_Click: " & Index
lTask.Index = Index
Set mTasks(Index) = lTask

If Index < 3 Then
    priority = PriorityHigh
ElseIf Index < 7 Then
    priority = PriorityNormal
Else
    priority = PriorityLow
End If

Set mTaskControllers(Index) = StartTask(lTask, priority, , Index)
mTaskControllers(Index).AddTaskCompletionListener Me
mTaskControllers(Index).AddTaskProgressListener Me

CountText(Index).ForeColor = vbBlack

StartButton(Index).Enabled = False
PauseButton(Index).Enabled = True
CancelButton(Index).Enabled = True

End Sub

Private Sub SummariesButton_Click()
Dim s As String

s = "Runnable tasks------------------------------------------------------------" & vbCrLf
s = s & TaskManager.GetRunnableTaskSummary
s = s & "Processed tasks-----------------------------------------------------------" & vbCrLf
s = s & TaskManager.GetProcessedTaskSummary
s = s & "Restartable tasks---------------------------------------------------------" & vbCrLf
s = s & TaskManager.GetRestartableTaskSummary
s = s & "Pending tasks-------------------------------------------------------------" & vbCrLf
s = s & TaskManager.GetPendingTaskSummary
s = s & "Suspended tasks-----------------------------------------------------------" & vbCrLf
s = s & TaskManager.GetSuspendedTaskSummary

modelessMsgBox s, MsgBoxInformation

End Sub

'================================================================================
' mSchedulingIntervalTimer Event Handlers
'================================================================================

Private Sub mSchedulingIntervalTimer_TimerExpired(ev As TimerExpiredEventData)
SchedulingIntervalLabel.Caption = Format(TaskManager.AverageInterScheduleWait, "0.00")
SchedulingIntervalLabel.Refresh
End Sub

'================================================================================
' Properties
'================================================================================

'================================================================================
' Methods
'================================================================================

'================================================================================
' Helper Functions
'================================================================================

Private Sub modelessMsgBox( _
                ByVal prompt As String, _
                ByVal buttons As MsgBoxStyles, _
                Optional ByVal title As String)
Dim lMsgBox As New fMsgBox

lMsgBox.initialise prompt, buttons, title

lMsgBox.Show vbModeless, Me
                
End Sub

Private Sub validateParams()
SetParamsButton.Enabled = False
If Not IsInteger(QuantumText, 1) Then
ElseIf Not IsInteger(ConcurrencyText, 1) Then
Else
    SetParamsButton.Enabled = True
End If
End Sub


