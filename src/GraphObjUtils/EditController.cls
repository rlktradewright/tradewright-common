VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EditController"
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

Private Const ModuleName                            As String = "EditController"

'@================================================================================
' Member variables
'@================================================================================

Private mEditing                                    As Boolean
Private mBoundaryToBeRecalculated                   As Boolean
Private mImageToBeCleared                           As Boolean
Private mRenderRequired                             As Boolean

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

Public Property Get BoundaryToBeRecalculated() As Boolean
BoundaryToBeRecalculated = mBoundaryToBeRecalculated
End Property

Public Property Get Editing() As Boolean
Editing = mEditing
End Property

Public Property Get ImageToBeCleared() As Boolean
ImageToBeCleared = mImageToBeCleared
End Property

Public Property Get RenderRequired() As Boolean
RenderRequired = mRenderRequired
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub BeginEdit()
Const ProcName As String = "BeginEdit"
On Error GoTo Err

If mEditing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Already editing"
End If

mEditing = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub EndEdit()
Const ProcName As String = "EndEdit"
On Error GoTo Err

If Not mEditing Then
    Err.Raise ErrorCodes.ErrIllegalStateException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Not editing"
End If
mEditing = False
mBoundaryToBeRecalculated = False
mImageToBeCleared = False
mRenderRequired = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetBoundaryToBeRecalculated()
mBoundaryToBeRecalculated = True
mRenderRequired = True
End Sub

Public Sub SetImageToBeCleared()
mImageToBeCleared = True
mRenderRequired = True
End Sub

Public Sub SetRenderRequired()
mRenderRequired = True
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




