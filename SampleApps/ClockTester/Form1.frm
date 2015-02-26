VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ClockTester"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox ElapsedText 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox CurrentTimeText 
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.ComboBox TimeZoneCombo 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton StartClockButton 
      Caption         =   "Start clock"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label TimezoneLabel 
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   4215
   End
   Begin VB.Label Label1 
      Caption         =   "Select timezone"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mClock As Clock
Attribute mClock.VB_VarHelpID = -1
Private mTimer As ElapsedTimer

Private Sub Form_Initialize()
InitialiseTWUtilities
End Sub

Private Sub Form_Load()
Dim tzNames() As String
Dim i As Long

tzNames = GetAvailableTimeZoneNames
For i = 0 To UBound(tzNames)
    TimeZoneCombo.AddItem tzNames(i)
Next

Set mTimer = New ElapsedTimer

End Sub

Private Sub Form_Terminate()
Debug.Print "Form1: Form_Terminate"
TerminateTWUtilities
End Sub

Private Sub mClock_Tick()
Dim currentTime As Date
Dim elapsed As Single

elapsed = mTimer.ElapsedTimeMicroseconds
mTimer.StartTiming

currentTime = mClock.TimeStamp
CurrentTimeText = FormatTimestamp(currentTime, TimestampCustom, "dd/mm/yy hh:mm:ss")
ElapsedText = elapsed
End Sub

Private Sub StartClockButton_Click()
Set mClock = GetClock(TimeZoneCombo.Text)
mTimer.StartTiming
TimezoneLabel.Caption = mClock.TimeZone.DisplayName
End Sub
