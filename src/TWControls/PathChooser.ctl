VERSION 5.00
Begin VB.UserControl PathChooser 
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   ScaleHeight     =   3180
   ScaleWidth      =   4560
   Begin VB.DriveListBox DriveList 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
   Begin VB.DirListBox DirList 
      Height          =   2565
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
End
Attribute VB_Name = "PathChooser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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


Private Const ModuleName                    As String = "PathChooser"

'@================================================================================
' Member variables
'@================================================================================

Private mTheme                              As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

DriveList.Width = UserControl.Width
DirList.Width = UserControl.Width
If (UserControl.Height - DirList.Top) > 66 * Screen.TwipsPerPixelY Then
    DirList.Height = UserControl.Height - DirList.Top
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

Private Sub DriveList_Change()
Const ProcName As String = "DriveList_Change"
On Error GoTo Err

DirList.Path = DriveList.Drive

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

Public Property Let Path(ByVal newvalue As String)
Const ProcName As String = "Path"
On Error GoTo Err

On Error Resume Next ' in case Path doesn't exist
If Mid$(newvalue, 2, 1) = ":" Then
    DriveList.Drive = Left$(newvalue, 2)
    DirList.Path = newvalue
ElseIf Left$(newvalue, 2) = "\\" Then
    DirList.Path = newvalue
Else
    DirList.Path = newvalue
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get Path() As String
Const ProcName As String = "Path"
On Error GoTo Err

Path = DirList.Path

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
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.BackColor
gApplyTheme mTheme, UserControl.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub NewFolder()
Dim fNew As New fNewFolder
Dim filesys As FileSystemObject
Dim folder As folder
Dim folders As folders
Dim newFolderPath As String

Const ProcName As String = "NewFolder"
On Error GoTo Err

show:

fNew.show vbModal
If fNew.NewFolderText = "" Then Unload fNew: Exit Sub

Set filesys = New FileSystemObject
Set folder = filesys.GetFolder(DirList.Path)
Set folders = folder.SubFolders
folders.Add fNew.NewFolderText

newFolderPath = DirList.Path & "\" & fNew.NewFolderText

DirList.Refresh
DirList.Path = newFolderPath
Unload fNew
Exit Sub

Err:
If Err.Number = 58 Then
    ' folder already exists
    MsgBox "Folder already exists", , "Error"
    Resume show
End If
Unload fNew

gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


