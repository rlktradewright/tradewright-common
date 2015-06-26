VERSION 5.00
Begin VB.Form fPathChooser 
   Caption         =   "Choose folder"
   ClientHeight    =   2775
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   6030
   Begin TWControls40.PathChooser PathChooser1 
      Height          =   2655
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4683
   End
   Begin VB.CommandButton NewFolderButton 
      Cancel          =   -1  'True
      Caption         =   "New folder..."
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
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
End
Attribute VB_Name = "fPathChooser"
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

Private Const ModuleName                    As String = "fPathChooser"

'@================================================================================
' Member variables
'@================================================================================

Private mCancelled                          As Boolean

Private mTheme                              As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Activate()
mCancelled = True
End Sub

Private Sub Form_Initialize()
mCancelled = True
End Sub

Private Sub Form_Resize()
Dim butleft As Long
Const ProcName As String = "Form_Resize"

On Error GoTo Err

butleft = Me.ScaleWidth - OKButton.Width - 8 * Screen.TwipsPerPixelX
If butleft >= 2160 Then
    OKButton.Left = butleft
    CancelButton.Left = butleft
    NewFolderButton.Left = butleft
    PathChooser1.Width = butleft - 8 * Screen.TwipsPerPixelX - PathChooser1.Left
End If

If Me.ScaleHeight >= 1560 Then
    PathChooser1.Height = Me.ScaleHeight - 8 * Screen.TwipsPerPixelY - PathChooser1.Top
    NewFolderButton.Top = PathChooser1.Height + PathChooser1.Top - NewFolderButton.Height
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
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

Me.Hide
mCancelled = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub NewFolderButton_Click()
Const ProcName As String = "NewFolderButton_Click"
On Error GoTo Err

PathChooser1.NewFolder

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub OKButton_Click()
Const ProcName As String = "OKButton_Click"
On Error GoTo Err

mCancelled = False
Me.Hide

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get cancelled() As Boolean
cancelled = mCancelled
End Property

Public Property Let Path(ByVal newvalue As String)
Const ProcName As String = "Path"
On Error GoTo Err

PathChooser1.Path = newvalue

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Path() As String
Const ProcName As String = "Path"
On Error GoTo Err

If Not mCancelled Then Path = PathChooser1.Path

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
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


