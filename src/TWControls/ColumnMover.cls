VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColumnMover"
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

Private Const ModuleName                    As String = "ColumnMover"

'@================================================================================
' Member variables
'@================================================================================

Private mStartCol                           As Long
Private mStartLabel                         As Label
Private mStartX                             As Long

Private mCol                                As Long
Private mLastX                              As Long

Private mTop                                As Long
Private mLeft                               As Long
Private mCellHeight                         As Long
Private mCellWidth                          As Long

'Private mPen                                As Long

Private mTransposeLine                      As line
Private mTransposeLineHeight                As Long

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

If Not mTranspose Is Nothing Then Unload mTranspose
Set mTranspose = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get startcol() As Long
Const ProcName As String = "startcol"
On Error GoTo Err

startcol = mStartCol

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function EndMove() As Long
Const ProcName As String = "endMove"
On Error GoTo Err

If Not mTranspose Is Nothing Then mTransposeLine.Visible = False
mStartLabel.MousePointer = MousePointerConstants.vbDefault
'DeleteObject mPen
If mLastX <> &H80000000 Then
    EndMove = mCol
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
                ByVal pStartcol As Long, _
                ByVal pStartX As Long, _
                ByVal pTransposeLine As line, _
                ByVal pTransposeLineHeight As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

mStartCol = pStartcol
mTop = pTop
mLeft = pLeft
mCellHeight = pCell.Height
mCellWidth = pCell.Width
Set mTransposeLine = pTransposeLine
mTransposeLineHeight = pTransposeLineHeight

Set mStartLabel = pCell.ValueLabel

mStartX = pStartX

'mStartLabel.MousePointer = MousePointerConstants.vbCustom
mLastX = &H80000000

'mPen = CreatePen(PS_SOLID, 2, 0)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub moveTo( _
                ByVal pCol As Long, _
                ByVal lineX As Long, _
                ByVal X As Long)
Const ProcName As String = "moveTo"
On Error GoTo Err

If X = mLastX Then Exit Sub

mCol = pCol

'If mLastX <> &H80000000 Then drawLine hDC, mLastX

If mTranspose Is Nothing Then
    Set mTimer = CreateIntervalTimer(500, ExpiryTimeUnitMilliseconds)
    mTimer.StartTimer
    showTransposeForm
Else
    mTimer.StopTimer
    mLastX = X
    DrawLine lineX
    
    mTranspose.Left = mLeft + X - mStartX
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
                ByVal X As Long)
Const ProcName As String = "DrawLine"
On Error GoTo Err

mTransposeLine.Visible = False
mTransposeLine.x1 = X
mTransposeLine.y1 = 0
mTransposeLine.x2 = X
mTransposeLine.y2 = mTransposeLineHeight - Screen.TwipsPerPixelY
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

mTranspose.Label1.Left = gTextPaddingTwips
mTranspose.Label1.Top = mStartLabel.Top
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


