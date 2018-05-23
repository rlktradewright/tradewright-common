VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SpriteController"
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

Implements ITask

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Public Type SpriteControllerMetrics
    RendersPerSecond        As Long
    AverageRenderTime       As Long
    AverageMoveTime         As Long
    AverageIterationTime    As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "SpriteController"

'@================================================================================
' Member variables
'@================================================================================

Private mSprites                                    As EnumerableCollection
Private mSpritesEn                                  As Enumerator

Private mTaskContext                                As TaskContext

Private mController                                 As Controller

Private mLastRenderTime                             As Date

Private mDisableDrawingDuringUpdates                As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub ITask_Cancel()
mController.IsDrawingEnabled = True
mTaskContext.Finish 0, True
End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Static et As ElapsedTimer
Static et1 As ElapsedTimer
Static et2 As ElapsedTimer
Static et3 As ElapsedTimer
Static renderCount As Long
Static renderTime As Single
Static moveCount As Long
Static moveTime As Single
Static iterationCount As Long
Static iterationTime As Single
Dim lSprite As Sprite

If mSprites.Count = 0 Then
    mController.IsDrawingEnabled = True
    mTaskContext.Finish 0, False
    Exit Sub
End If

If et Is Nothing Then
    Set et = New ElapsedTimer
    et.StartTiming
    Set et1 = New ElapsedTimer
    et1.StartTiming
    Set et2 = New ElapsedTimer
    Set et3 = New ElapsedTimer
End If

If mTaskContext.CancelPending Then
    mController.IsDrawingEnabled = True
    mTaskContext.Finish 0, True
    Exit Sub
End If

If mDisableDrawingDuringUpdates Then mController.IsDrawingEnabled = False

et3.StartTiming
Do While mSpritesEn.MoveNext
    Set lSprite = mSpritesEn.Current
    et2.StartTiming
    lSprite.Move
    moveTime = moveTime + et2.ElapsedTimeMicroseconds
    moveCount = moveCount + 1
Loop
iterationTime = iterationTime + et3.ElapsedTimeMicroseconds
iterationCount = iterationCount + 1
Set mSpritesEn = mSprites.Enumerator

If mDisableDrawingDuringUpdates Then
    mController.IsDrawingEnabled = True
    et.StartTiming
    mController.Paint False
    mController.IsDrawingEnabled = False
    renderTime = renderTime + et.ElapsedTimeMicroseconds
    renderCount = renderCount + 1
End If

If et1.ElapsedTimeMicroseconds >= 1000000 Then
    Dim metrics As SpriteControllerMetrics
    If mDisableDrawingDuringUpdates Then
        metrics.AverageRenderTime = Int(renderTime / renderCount)
        metrics.RendersPerSecond = renderCount
    End If
    metrics.AverageMoveTime = Int(moveTime / moveCount)
    metrics.AverageIterationTime = Int(iterationTime / iterationCount / mSprites.Count)
    mTaskContext.Progress 0, metrics
    renderCount = 0
    renderTime = 0
    moveCount = 0
    moveTime = 0
    iterationTime = 0
    iterationCount = 0
    et1.StartTiming
End If

'mTaskContext.Suspend 200

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get ITask_TaskName() As String

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal pSprites As EnumerableCollection, _
                ByVal pController As Controller, _
                ByVal pDisableDrawingDuringUpdates As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSprites = pSprites
Set mSpritesEn = mSprites.Enumerator
Set mController = pController
mDisableDrawingDuringUpdates = pDisableDrawingDuringUpdates

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




