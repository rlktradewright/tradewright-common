VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Regular Expression Tester"
   ClientHeight    =   10320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11625
   LinkTopic       =   "Form1"
   ScaleHeight     =   10320
   ScaleWidth      =   11625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton HelpButton 
      Caption         =   "Reg Exp &help"
      Height          =   615
      Left            =   10080
      TabIndex        =   12
      Top             =   9600
      Width           =   1455
   End
   Begin VB.TextBox ExecuteResultsText 
      BackColor       =   &H00D0D0D0&
      ForeColor       =   &H00FF0000&
      Height          =   3255
      Left            =   1320
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   6960
      Width           =   8655
   End
   Begin VB.CheckBox IgnoreCaseCheck 
      Caption         =   "Ignore case"
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
   End
   Begin VB.CheckBox GlobalCheck 
      Caption         =   "Global"
      Height          =   255
      Left            =   1320
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton ExecuteButton 
      Caption         =   "&Execute"
      Height          =   615
      Left            =   10080
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton TestButton 
      Caption         =   "&Test"
      Default         =   -1  'True
      Height          =   615
      Left            =   10080
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox TextText 
      Height          =   3375
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   3480
      Width           =   8655
   End
   Begin VB.TextBox PatternText 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
   Begin VB.Label Label3 
      Caption         =   "Execute results"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   6960
      Width           =   1095
   End
   Begin VB.Label ExecuteResultLabel 
      Height          =   375
      Left            =   10080
      TabIndex        =   10
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label TestResultLabel 
      Height          =   375
      Left            =   10080
      TabIndex        =   9
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Text"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Pattern"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mRegExp As New RegExp
Private mET As New ElapsedTimer

Private Sub ExecuteButton_Click()
Dim matches As MatchCollection
Dim lMatch As Match
Dim submatch  As Variant
Dim i As Long
Dim j As Long
Dim s As String
Dim elapsed As Single

On Error GoTo Err

mRegExp.Global = (GlobalCheck = vbChecked)
mRegExp.IgnoreCase = (IgnoreCaseCheck = vbChecked)
mRegExp.Pattern = PatternText

mET.StartTiming
Set matches = mRegExp.Execute(TextText)
elapsed = mET.ElapsedTimeMicroseconds

ExecuteResultLabel.Caption = "Matches: " & matches.Count & vbCrLf & _
                            "Time: " & Int(elapsed)

For Each lMatch In matches
    s = s & i & vbTab & lMatch.FirstIndex & vbTab & lMatch.Value & vbCrLf
    j = 0
    For Each submatch In lMatch.SubMatches
        s = s & i & "." & j & vbTab & vbTab & submatch & vbCrLf
        j = j + 1
    Next
    i = i + 1
Next

ExecuteResultsText = s

Exit Sub

Err:
ExecuteResultLabel.Caption = "Failed"
End Sub

Private Sub Form_Initialize()
InitialiseTWUtilities
End Sub

Private Sub Form_Load()
Me.Move 0, 0
End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

Private Sub HelpButton_Click()
Dim f As New Form2
f.Show vbModeless, Me
f.Width = Screen.Width - Me.Width
f.Height = Screen.Height
f.Top = 0
f.Left = Screen.Width - f.Width
End Sub

Private Sub PatternText_Change()
TestResultLabel.Caption = ""
ExecuteResultLabel.Caption = ""
ExecuteResultsText = ""
End Sub

Private Sub TestButton_Click()
Dim elapsed As Single
Dim ok As Boolean

On Error GoTo Err

mRegExp.Global = (GlobalCheck = vbChecked)
mRegExp.IgnoreCase = (IgnoreCaseCheck = vbChecked)
mRegExp.Pattern = PatternText
mET.StartTiming
ok = mRegExp.Test(TextText)
elapsed = mET.ElapsedTimeMicroseconds
TestResultLabel.Caption = CStr(ok) & vbCrLf & _
                        "Time: " & Int(elapsed)

Exit Sub

Err:
TestResultLabel.Caption = "Failed"
End Sub

Private Sub TextText_Change()
TestResultLabel.Caption = ""
ExecuteResultLabel.Caption = ""
ExecuteResultsText = ""
End Sub
