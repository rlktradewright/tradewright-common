VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RowMover"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                    As String = "RowMover"

'@================================================================================
' Declares
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mStartRow                           As Long
Private mStartLabel                         As Label
Private mStartY                             As Long

Private mRow                                As Long
Private mLastY                              As Long

Private mTop                                As Long
Private mLeft                               As Long
Private mCellHeight                         As Long
Private mCellWidth                          As Long
Private mOffset                             As Long

Private mTransposeLine                      As line
Private mTransposeLineWidth                 As Long

Private mTranspose                          As Transpose

Private WithEvents mTimer                   As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mTimer Event Handlers
'@================================================================================

Private Sub mTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mTimer_TimerExpired"
On Error GoTo Err

Unload mTranspose
Set mTranspose = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get StartRow() As Long
Const ProcName As String = "StartRow"
On Error GoTo Err

StartRow = mStartRow

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function EndMove() As Long
Const ProcName As String = "EndMove"
On Error GoTo Err

mTransposeLine.Visible = False
mStartLabel.MousePointer = MousePointerConstants.vbDefault
'DeleteObject mPen
If mLastY <> &H80000000 Then
    EndMove = mRow
Else
    EndMove = -1
End If
If Not mTranspose Is Nothing Then Unload mTranspose

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Initialise( _
                ByVal pCell As GridCell, _
                ByVal pTop As Long, _
                ByVal pLeft As Long, _
                ByVal pStartRow As Long, _
                ByVal pStartY As Long, _
                ByVal pTransposeLine As line, _
                ByVal pTransposeLineWidth As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

mStartRow = pStartRow
mTop = pTop
mLeft = pLeft
mCellHeight = pCell.Height
mCellWidth = pCell.Width
Set mTransposeLine = pTransposeLine
mTransposeLineWidth = pTransposeLineWidth

Set mStartLabel = pCell.ValueLabel
mOffset = mStartLabel.Top - pCell.Top

mStartY = pStartY

mLastY = &H80000000

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub moveTo( _
                ByVal pRow As Long, _
                ByVal lineY As Long, _
                ByVal Y As Long)
Const ProcName As String = "MoveTo"
On Error GoTo Err

If Y = mLastY Then Exit Sub

mRow = pRow

If mTranspose Is Nothing Then
    Set mTimer = CreateIntervalTimer(500, ExpiryTimeUnitMilliseconds)
    mTimer.StartTimer
    showTransposeForm
Else
    mTimer.StopTimer
    mLastY = Y
    DrawLine lineY

    mTranspose.Top = mTop + Y - mStartY
    mTimer.StartTimer
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub
                

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub DrawLine( _
                ByVal Y As Long)
Const ProcName As String = "DrawLine"
On Error GoTo Err

mTransposeLine.Visible = False
mTransposeLine.x1 = 0
mTransposeLine.y1 = Y
mTransposeLine.x2 = mTransposeLineWidth - Screen.TwipsPerPixelX
mTransposeLine.y2 = Y
mTransposeLine.ZOrder 0
mTransposeLine.Visible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showTransposeForm()
Const ProcName As String = "showTransposeForm"
On Error GoTo Err

Set mTranspose = New Transpose
mTranspose.Left = mLeft
mTranspose.Top = mTop
mTranspose.Width = mCellWidth
mTranspose.Height = mCellHeight

mTranspose.Label1.Left = mStartLabel.Left
mTranspose.Label1.Top = mOffset
mTranspose.Label1.Width = mStartLabel.Width
mTranspose.Label1.Height = mStartLabel.Height

Set mTranspose.Label1.Font = mStartLabel.Font
mTranspose.Label1.ForeColor = mStartLabel.ForeColor
mTranspose.Label1.Alignment = mStartLabel.Alignment
mTranspose.Label1.Caption = mStartLabel.Caption

mTranspose.show vbModal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub




