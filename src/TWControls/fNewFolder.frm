VERSION 5.00
Begin VB.Form fNewFolder 
   Caption         =   "Create new folder"
   ClientHeight    =   1245
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   6030
   Begin VB.TextBox NewFolderText 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   4215
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "New folder name"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "fNewFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IThemeable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "fNewFolder"

'@================================================================================
' Member variables
'@================================================================================

Private mTheme                              As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Resize()
Dim butleft As Long
Const ProcName As String = "Form_Resize"
On Error GoTo Err

Me.Height = 1755
butleft = Me.ScaleWidth - OKButton.Width - 8 * Screen.TwipsPerPixelX
If butleft >= 1920 Then
    OKButton.Left = butleft
    CancelButton.Left = butleft
    NewFolderText.Width = butleft - NewFolderText.Left - 8 * Screen.TwipsPerPixelX
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub CancelButton_Click()
Const ProcName As String = "CancelButton_Click"
On Error GoTo Err

NewFolderText = ""
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub OKButton_Click()
Const ProcName As String = "OKButton_Click"
On Error GoTo Err

Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

Me.BackColor = mTheme.BackColor
gApplyTheme mTheme, Me.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================





