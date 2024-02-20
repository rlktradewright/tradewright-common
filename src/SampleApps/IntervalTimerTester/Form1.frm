VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Timer Utilities Test Program"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton StopElapsedTimerButton 
      Caption         =   "Stop elapsed timer"
      Enabled         =   0   'False
      Height          =   495
      Left            =   8160
      TabIndex        =   19
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton StartElapsedTimerButton 
      Caption         =   "Start elapsed timer"
      Height          =   495
      Left            =   8160
      TabIndex        =   18
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Frame Interval2 
      Caption         =   "Timer 2"
      Height          =   4575
      Left            =   4320
      TabIndex        =   21
      Top             =   240
      Width           =   3735
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   240
         ScaleHeight     =   4215
         ScaleWidth      =   3375
         TabIndex        =   31
         Top             =   240
         Width           =   3375
         Begin VB.CheckBox BaseTimerCheck2 
            Caption         =   "Use base timer"
            Height          =   255
            Left            =   480
            TabIndex        =   15
            Top             =   3360
            Width           =   1695
         End
         Begin VB.CommandButton StopTimerButton2 
            Caption         =   "Stop timer"
            Enabled         =   0   'False
            Height          =   495
            Left            =   1920
            TabIndex        =   17
            Top             =   3720
            Width           =   1455
         End
         Begin VB.CommandButton StartTimerButton2 
            Caption         =   "Start timer"
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   3720
            Width           =   1455
         End
         Begin VB.TextBox FirstIntervalValue2 
            Height          =   285
            Left            =   480
            TabIndex        =   9
            Text            =   "1"
            Top             =   1200
            Width           =   2175
         End
         Begin VB.OptionButton FirstDateTimeOpt2 
            Caption         =   "Date/time"
            Height          =   195
            Left            =   600
            TabIndex        =   12
            Top             =   2040
            Width           =   1335
         End
         Begin VB.TextBox IntervalValue2 
            Height          =   285
            Left            =   480
            TabIndex        =   13
            Text            =   "0"
            Top             =   2640
            Width           =   1095
         End
         Begin VB.OptionButton FirstSecondsOpt2 
            Caption         =   "Seconds"
            Height          =   195
            Left            =   600
            TabIndex        =   10
            Top             =   1560
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton FirstMillisecondsOpt2 
            Caption         =   "Millisecs"
            Height          =   195
            Left            =   600
            TabIndex        =   11
            Top             =   1800
            Width           =   1575
         End
         Begin VB.CheckBox RandomCheck2 
            Caption         =   "Random"
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   3120
            Width           =   1575
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Average interval (millisecs)"
            Height          =   375
            Left            =   240
            TabIndex        =   42
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Event count"
            Height          =   255
            Left            =   240
            TabIndex        =   41
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Counter2 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   38
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label AvgInterval2Label 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            Height          =   255
            Left            =   1920
            TabIndex        =   37
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label3 
            Caption         =   "First interval"
            Height          =   255
            Left            =   0
            TabIndex        =   34
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Repeat interval"
            Height          =   255
            Left            =   0
            TabIndex        =   33
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label1 
            Caption         =   "(Milliseconds)"
            Height          =   255
            Left            =   1800
            TabIndex        =   32
            Top             =   2640
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Interval1 
      Caption         =   "Timer 1"
      Height          =   4575
      Left            =   240
      TabIndex        =   20
      Top             =   240
      Width           =   3735
      Begin VB.PictureBox Picture3 
         BorderStyle     =   0  'None
         Height          =   4215
         Left            =   240
         ScaleHeight     =   4215
         ScaleWidth      =   3375
         TabIndex        =   27
         Top             =   240
         Width           =   3375
         Begin VB.CheckBox BaseTimerCheck1 
            Caption         =   "Use base timer"
            Height          =   255
            Left            =   480
            TabIndex        =   6
            Top             =   3360
            Width           =   1695
         End
         Begin VB.CommandButton StartTimerButton1 
            Caption         =   "Start timer"
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   3720
            Width           =   1455
         End
         Begin VB.CommandButton StopTimerButton1 
            Caption         =   "Stop timer"
            Enabled         =   0   'False
            Height          =   495
            Left            =   1920
            TabIndex        =   8
            Top             =   3720
            Width           =   1455
         End
         Begin VB.OptionButton FirstDateTimeOpt1 
            Caption         =   "Date/time"
            Height          =   195
            Left            =   600
            TabIndex        =   3
            Top             =   2040
            Width           =   1335
         End
         Begin VB.OptionButton FirstSecondsOpt1 
            Caption         =   "Seconds"
            Height          =   195
            Left            =   600
            TabIndex        =   1
            Top             =   1560
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton FirstMillisecondsOpt1 
            Caption         =   "Millisecs"
            Height          =   195
            Left            =   600
            TabIndex        =   2
            Top             =   1800
            Width           =   1575
         End
         Begin VB.TextBox FirstIntervalValue1 
            Height          =   285
            Left            =   480
            TabIndex        =   0
            Text            =   "1"
            Top             =   1200
            Width           =   2175
         End
         Begin VB.TextBox IntervalValue1 
            Height          =   285
            Left            =   480
            TabIndex        =   4
            Text            =   "0"
            Top             =   2640
            Width           =   1095
         End
         Begin VB.CheckBox RandomCheck1 
            Caption         =   "Random"
            Height          =   255
            Left            =   480
            TabIndex        =   5
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   "Average interval (millisecs)"
            Height          =   375
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   1575
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "Event count"
            Height          =   255
            Left            =   240
            TabIndex        =   39
            Top             =   0
            Width           =   1575
         End
         Begin VB.Label Counter1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1920
            TabIndex        =   36
            Top             =   0
            Width           =   1455
         End
         Begin VB.Label AvgInterval1Label 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            Height          =   255
            Left            =   1920
            TabIndex        =   35
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Repeat interval"
            Height          =   255
            Left            =   0
            TabIndex        =   30
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label4 
            Caption         =   "First interval"
            Height          =   255
            Left            =   0
            TabIndex        =   29
            Top             =   960
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "(Milliseconds)"
            Height          =   255
            Left            =   1800
            TabIndex        =   28
            Top             =   2640
            Width           =   1215
         End
      End
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   9120
      TabIndex        =   54
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   9120
      TabIndex        =   53
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   9120
      TabIndex        =   52
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   9120
      TabIndex        =   51
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   9120
      TabIndex        =   50
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   9120
      TabIndex        =   49
      Top             =   960
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   9120
      TabIndex        =   48
      Top             =   600
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   9120
      TabIndex        =   47
      Top             =   240
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   8160
      TabIndex        =   46
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   8160
      TabIndex        =   45
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   8160
      TabIndex        =   44
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   8160
      TabIndex        =   43
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label ElapsedTimeLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   8160
      TabIndex        =   26
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   8160
      TabIndex        =   25
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   8160
      TabIndex        =   24
      Top             =   960
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   8160
      TabIndex        =   23
      Top             =   600
      Width           =   855
   End
   Begin VB.Label ValueLabel 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   8160
      TabIndex        =   22
      Top             =   240
      Width           =   855
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

Implements IStateChangeListener

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ProjectName                As String = "IntervalTimerTester"
Private Const ModuleName                As String = "Form1"

'================================================================================
' Enums
'================================================================================

Public Enum Comparison
    Greater
    Less
    Equal
End Enum

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mCounter1 As Long
Private mCounter2 As Long
Private mTotalElapsed1 As Single
Private mTotalElapsed2 As Single

Private WithEvents mTimer1 As IntervalTimer
Attribute mTimer1.VB_VarHelpID = -1
Private mTimer1Number As Long
Private WithEvents mTimer2 As IntervalTimer
Attribute mTimer2.VB_VarHelpID = -1
Private WithEvents mBaseTimer1 As IntervalTimer
Attribute mBaseTimer1.VB_VarHelpID = -1
Private WithEvents mBaseTimer2 As IntervalTimer
Attribute mBaseTimer2.VB_VarHelpID = -1

Private mFirstIntervalUnit1 As ExpiryTimeUnits
Private mFirstIntervalUnit2 As ExpiryTimeUnits

Private mTimerList As TimerList
Attribute mTimerList.VB_VarHelpID = -1
Private mTimerListItems() As TimerListItem

Private mElapsedTimer As ElapsedTimer

Private mElapsedTimer1 As ElapsedTimer
Private mElapsedTimer2 As ElapsedTimer

'================================================================================
' Form Event Handlers
'================================================================================

Private Sub Form_Initialize()
Const ProcName As String = "Form_Initialize"
On Error GoTo Err

InitialiseCommonControls
InitialiseTWUtilities
ApplicationGroupName = "TradeWright"
ApplicationName = "IntervalTimerTester"
DefaultLogLevel = LogLevelHighDetail
SetupDefaultLogging Command

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & ProcName & "." & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
MsgBox "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub Form_Load()
Dim i As Long

Dim failpoint As Long
On Error GoTo Err

Set mElapsedTimer1 = New ElapsedTimer
Set mElapsedTimer2 = New ElapsedTimer

Set mTimerList = GetGlobalTimerList

For i = 0 To ValueLabel.ubound
    ValueLabel(i).Caption = 0
Next

ReDim mTimerListItems(ValueLabel.ubound) As TimerListItem

mFirstIntervalUnit1 = ExpiryTimeUnitSeconds
mFirstIntervalUnit2 = ExpiryTimeUnitSeconds

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Form_Terminate" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
MsgBox "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim failpoint As Long
On Error GoTo Err

TerminateTWUtilities

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "Form_Terminate" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
MsgBox "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

'================================================================================
' IStateChangeListener Interface Members
'================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Dim tli As TimerListItem
If ev.State <> TimerListItemStates.TimerListItemStateExpired Then Exit Sub
Set tli = ev.Source
processTimerListItemExpiry tli
End Sub

'================================================================================
' Form Control Event Handlers
'================================================================================

Private Sub FirstDateTimeOpt1_Click()
mFirstIntervalUnit1 = ExpiryTimeUnitDateTime
If Not IsDate(FirstIntervalValue1) Then
    FirstIntervalValue1.SetFocus
    FirstIntervalValue1.SelStart = 0
    FirstIntervalValue1.SelLength = Len(FirstIntervalValue1)
End If
End Sub

Private Sub FirstDateTimeOpt2_Click()
mFirstIntervalUnit2 = ExpiryTimeUnitDateTime
If Not IsDate(FirstIntervalValue2) Then
    FirstIntervalValue2.SetFocus
    FirstIntervalValue2.SelStart = 0
    FirstIntervalValue2.SelLength = Len(FirstIntervalValue2)
End If
End Sub

Private Sub FirstIntervalValue1_Validate(Cancel As Boolean)
Select Case mFirstIntervalUnit1
Case ExpiryTimeUnitDateTime
    If Not IsDate(FirstIntervalValue1) Then Cancel = True
    If CDate(FirstIntervalValue1) <= Now Then Cancel = True
Case ExpiryTimeUnitSeconds
    If Not IsNumeric(FirstIntervalValue1) Then Cancel = True
Case ExpiryTimeUnitMilliseconds
    If Not IsNumeric(FirstIntervalValue1) Then Cancel = True
End Select
End Sub

Private Sub FirstIntervalValue2_Validate(Cancel As Boolean)
Select Case mFirstIntervalUnit2
Case ExpiryTimeUnitDateTime
    If Not IsDate(FirstIntervalValue2) Then Cancel = True
    If CDate(FirstIntervalValue2) <= Now Then Cancel = True
Case ExpiryTimeUnitSeconds
    If Not IsNumeric(FirstIntervalValue2) Then Cancel = True
Case ExpiryTimeUnitMilliseconds
    If Not IsNumeric(FirstIntervalValue2) Then Cancel = True
End Select
End Sub

Private Sub FirstMillisecondsOpt1_Click()
mFirstIntervalUnit1 = ExpiryTimeUnitMilliseconds
If Not IsNumeric(FirstIntervalValue1) Then
    FirstIntervalValue1.SetFocus
    FirstIntervalValue1.SelStart = 0
    FirstIntervalValue1.SelLength = Len(FirstIntervalValue1)
End If
End Sub

Private Sub FirstMilliSecondsOpt2_Click()
mFirstIntervalUnit2 = ExpiryTimeUnitMilliseconds
If Not IsNumeric(FirstIntervalValue2) Then
    FirstIntervalValue2.SetFocus
    FirstIntervalValue2.SelStart = 0
    FirstIntervalValue2.SelLength = Len(FirstIntervalValue2)
End If
End Sub

Private Sub FirstSecondsOpt1_Click()
mFirstIntervalUnit1 = ExpiryTimeUnitSeconds
If Not IsNumeric(FirstIntervalValue1) Then
    FirstIntervalValue1.SetFocus
    FirstIntervalValue1.SelStart = 0
    FirstIntervalValue1.SelLength = Len(FirstIntervalValue1)
End If
End Sub

Private Sub FirstSecondsOpt2_Click()
mFirstIntervalUnit2 = ExpiryTimeUnitSeconds
If Not IsNumeric(FirstIntervalValue2) Then
    FirstIntervalValue2.SetFocus
    FirstIntervalValue2.SelStart = 0
    FirstIntervalValue2.SelLength = Len(FirstIntervalValue2)
End If
End Sub

Private Sub IntervalValue1_Validate(Cancel As Boolean)
If Not IsNumeric(IntervalValue1) Then Cancel = True
End Sub

Private Sub IntervalValue2_Validate(Cancel As Boolean)
If Not IsNumeric(IntervalValue2) Then Cancel = True
End Sub

Private Sub StartElapsedTimerButton_Click()
Set mElapsedTimer = New ElapsedTimer
mElapsedTimer.StartTiming
StartElapsedTimerButton.Enabled = False
StopElapsedTimerButton.Enabled = True
End Sub

Private Sub StartTimerButton1_Click()
Const ProcName As String = "StartTimerButton1_Click"
On Error GoTo Err

If mFirstIntervalUnit1 = ExpiryTimeUnitDateTime And _
    CDate(FirstIntervalValue1) <= Now _
Then
    FirstIntervalValue1.SetFocus
    FirstIntervalValue1.SelStart = 0
    FirstIntervalValue1.SelLength = Len(FirstIntervalValue1)
    Exit Sub
End If
    
mCounter1 = 0
mTotalElapsed1 = 0
StopTimerButton1.Enabled = True
StartTimerButton1.Enabled = False

If BaseTimerCheck1.Value = vbChecked Then
    Set mBaseTimer1 = CreateIntervalTimer(FirstIntervalValue1, _
                            mFirstIntervalUnit1, _
                            IntervalValue1, _
                            IIf(RandomCheck1 = vbChecked, True, False))
    mElapsedTimer1.StartTiming
    Debug.Print "Starting timer 1"
    mBaseTimer1.StartTimer
Else
    Dim s As String: s = "Timer1 started " & CDbl(GetTimestamp)
    If gLogger.IsLoggable(LogLevelHighDetail) Then gLogger.Log s, ProcName, ModuleName, LogLevelHighDetail
    Set mTimer1 = CreateIntervalTimer(FirstIntervalValue1, _
                            mFirstIntervalUnit1, _
                            IntervalValue1, _
                            IIf(RandomCheck1 = vbChecked, True, False))
    mElapsedTimer1.StartTiming
    mTimer1.StartTimer
    mTimer1Number = mTimer1.TimerNumber
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub StartTimerButton2_Click()
If mFirstIntervalUnit2 = ExpiryTimeUnitDateTime And _
    CDate(FirstIntervalValue2) <= Now _
Then
    FirstIntervalValue2.SetFocus
    FirstIntervalValue2.SelStart = 0
    FirstIntervalValue2.SelLength = Len(FirstIntervalValue2)
    Exit Sub
End If
    
mCounter2 = 0
mTotalElapsed2 = 0
StopTimerButton2.Enabled = True
StartTimerButton2.Enabled = False

If BaseTimerCheck2.Value = vbChecked Then
    Set mBaseTimer2 = CreateIntervalTimer(FirstIntervalValue2, _
                            mFirstIntervalUnit2, _
                            IntervalValue2, _
                            IIf(RandomCheck2 = vbChecked, True, False))
    mElapsedTimer2.StartTiming
    Debug.Print "Starting timer 2"
    mBaseTimer2.StartTimer
Else
    Set mTimer2 = CreateIntervalTimer(FirstIntervalValue2, _
                            mFirstIntervalUnit2, _
                            IntervalValue2, _
                            IIf(RandomCheck2 = vbChecked, True, False))
    mElapsedTimer2.StartTiming
    mTimer2.StartTimer
End If
End Sub

Private Sub StopElapsedTimerButton_Click()
Dim elapsedTime As Double
elapsedTime = mElapsedTimer.ElapsedTimeMicroseconds
ElapsedTimeLabel.Caption = Format(elapsedTime / 1000000, "#.######") & "s"
StartElapsedTimerButton.Enabled = True
StopElapsedTimerButton.Enabled = False
End Sub

Private Sub StopTimerButton1_Click()
StopTimerButton1.Enabled = False
StartTimerButton1.Enabled = True

If BaseTimerCheck1.Value = vbChecked Then
    mBaseTimer1.StopTimer
Else
    mTimer1.StopTimer
End If
End Sub

Private Sub StopTimerButton2_Click()
StopTimerButton2.Enabled = False
StartTimerButton2.Enabled = True

If BaseTimerCheck2.Value = vbChecked Then
    mBaseTimer2.StopTimer
Else
    mTimer2.StopTimer
End If
End Sub

'================================================================================
' mBaseTimer1 Event Handlers
'================================================================================

Private Sub mBaseTimer1_TimerExpired(ev As TimerExpiredEventData)
Dim et As Single
mCounter1 = mCounter1 + 1
Counter1.Caption = CStr(mCounter1)
et = mElapsedTimer1.ElapsedTimeMicroseconds
mElapsedTimer1.StartTiming
mTotalElapsed1 = mTotalElapsed1 + et
AvgInterval1Label.Caption = Format(mTotalElapsed1 / (1000 * mCounter1), "0.000")
If Not mBaseTimer1.RepeatNotifications Then
    StartTimerButton1.Enabled = True
    StopTimerButton1.Enabled = False
End If

generateData
End Sub

'================================================================================
' mBaseTimer2 Event Handlers
'================================================================================

Private Sub mBaseTimer2_TimerExpired(ev As TimerExpiredEventData)
Dim et As Single
mCounter2 = mCounter2 + 1
Counter2.Caption = CStr(mCounter2)
et = mElapsedTimer2.ElapsedTimeMicroseconds
mElapsedTimer2.StartTiming
mTotalElapsed2 = mTotalElapsed2 + et
AvgInterval2Label.Caption = Format(mTotalElapsed2 / (1000 * mCounter2), "0.000")
If Not mBaseTimer2.RepeatNotifications Then
    StartTimerButton2.Enabled = True
    StopTimerButton2.Enabled = False
End If

generateData
End Sub

'================================================================================
' mTimer1 Event Handlers
'================================================================================

Private Sub mTimer1_TimerExpired(ev As TimerExpiredEventData)
Dim et As Single
Dim failpoint As Long
On Error GoTo Err

mCounter1 = mCounter1 + 1
Counter1.Caption = CStr(mCounter1)
et = mElapsedTimer1.ElapsedTimeMicroseconds
mElapsedTimer1.StartTiming
mTotalElapsed1 = mTotalElapsed1 + et

Dim s As String: s = "Timer1 " & mTimer1Number & " expired after " & et / 1000 & "(" & CDbl(GetTimestamp) & ")"
If gLogger.IsLoggable(LogLevelHighDetail) Then gLogger.Log s, "mTimer1_TimerExpired", ModuleName, LogLevelHighDetail

AvgInterval1Label.Caption = Format(mTotalElapsed1 / (1000 * mCounter1), "0.000")
If Not mTimer1.RepeatNotifications Then
    StartTimerButton1.Enabled = True
    StopTimerButton1.Enabled = False
End If

generateData

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "mTimer1_TimerExpired" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
MsgBox "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
End Sub

'================================================================================
' mTimer2 Event Handlers
'================================================================================

Private Sub mTimer2_TimerExpired(ev As TimerExpiredEventData)
Dim et As Single
Dim failpoint As Long
On Error GoTo Err

mCounter2 = mCounter2 + 1
Counter2.Caption = CStr(mCounter2)
et = mElapsedTimer2.ElapsedTimeMicroseconds
mElapsedTimer2.StartTiming
mTotalElapsed2 = mTotalElapsed2 + et
AvgInterval2Label.Caption = Format(mTotalElapsed2 / (1000 * mCounter2), "0.000")
If Not mTimer2.RepeatNotifications Then
    StartTimerButton2.Enabled = True
    StopTimerButton2.Enabled = False
End If

generateData

Exit Sub

Err:
Dim errNumber As Long: errNumber = Err.Number
Dim errSource As String: errSource = ProjectName & "." & ModuleName & ":" & "mTimer2_TimerExpired" & "." & failpoint & IIf(errSource <> "", vbCrLf & errSource, "")
Dim errDescription As String: errDescription = Err.Description
MsgBox "Error " & errNumber & ": " & errDescription & vbCrLf & errSource
Err.Raise errNumber, errSource, errDescription
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

Private Sub generateData()
Dim Index As Long
Dim newval As Long
Dim Data As TimerItemData

Index = Int(1000 * Rnd) Mod (ValueLabel.ubound + 1)
newval = ValueLabel(Index).Caption + Int(11 * Rnd - 5)

Set Data = New TimerItemData
Data.Index = Index
If newval > ValueLabel(Index).Caption Then
    Data.Relative = Comparison.Greater
    ValueLabel(Index).BackColor = vbBlue
    ValueLabel(Index).ForeColor = vbWhite
ElseIf newval < ValueLabel(Index).Caption Then
    Data.Relative = Comparison.Less
    ValueLabel(Index).BackColor = vbRed
    ValueLabel(Index).ForeColor = vbWhite
Else
    Data.Relative = Comparison.Equal
    ValueLabel(Index).BackColor = vbGreen
    ValueLabel(Index).ForeColor = vbWhite
End If

If Not mTimerListItems(Index) Is Nothing Then
    mTimerList.Remove mTimerListItems(Index)
    mTimerListItems(Index).RemoveStateChangeListener Me
End If

Set mTimerListItems(Index) = mTimerList.Add(Data, Int(Rnd * 300) + 200, ExpiryTimeUnitMilliseconds)
mTimerListItems(Index).AddStateChangeListener Me

ValueLabel(Index).Caption = newval
End Sub

Private Function gLogger() As FormattingLogger
Static l As FormattingLogger
If l Is Nothing Then
    Set l = CreateFormattingLogger("log", ProjectName)
End If
Set gLogger = l
End Function

Private Sub processTimerListItemExpiry( _
                ByVal Item As TimerListItem)
Dim entryData As TimerItemData
Set entryData = Item.Data
mTimerListItems(entryData.Index).RemoveStateChangeListener Me
Select Case entryData.Relative
Case Greater
    ValueLabel(entryData.Index).BackColor = vbWhite
    ValueLabel(entryData.Index).ForeColor = vbBlue
Case Less
    ValueLabel(entryData.Index).BackColor = vbWhite
    ValueLabel(entryData.Index).ForeColor = vbRed
Case Equal
    ValueLabel(entryData.Index).BackColor = vbWhite
    ValueLabel(entryData.Index).ForeColor = vbGreen
End Select
Set mTimerListItems(entryData.Index) = Nothing
End Sub


Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub



