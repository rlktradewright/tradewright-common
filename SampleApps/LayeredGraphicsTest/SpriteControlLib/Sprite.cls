VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sprite"
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

Implements Task

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

Private Const ModuleName                            As String = "Sprite"

'@================================================================================
' Member variables
'@================================================================================

Private mSprite                                     As ISprite
Private mSpeed                                      As Double
Private mDirection                                  As Double
Private mPosition                                   As Point
Private mOrientation                                As Double
Private mSecondsPerRevolution                       As Double

Private mSpaceBoundary                              As TRectangle

Private mVx                                         As Double
Private mVy                                         As Double

Private WithEvents mPositionTLI                     As TimerListItem
Attribute mPositionTLI.VB_VarHelpID = -1
Private WithEvents mAppearanceTLI                   As TimerListItem
Attribute mAppearanceTLI.VB_VarHelpID = -1

Private mElapsedTimer                               As ElapsedTimer

Private mFinished                                   As Boolean

Private mTaskContext                                As TaskContext

Private mGradientFill                               As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Randomize
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()

End Sub

Private Sub Task_Run()
Dim et As Single


If Not mFinished Then
    et = mElapsedTimer.ElapsedTimeMicroseconds / 1000000
    'Debug.Print "Time since last update: " & et
    mElapsedTimer.StartTiming
    mSprite.Edit
    calcNewPosition et
    calcNewOrientation et
    mSprite.ApplyChanges
    'Debug.Print "Sprite update time: " & Format(mElapsedTimer.ElapsedTimeMicroseconds, "0")
    
    'mElapsedTimer.StartTiming
    mTaskContext.Suspend 50
End If
End Sub

Private Property Let Task_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get Task_TaskName() As String

End Property

'@================================================================================
' mAppearanceTLI Event Handlers
'@================================================================================

Private Sub mAppearanceTLI_StateChange(ev As StateChangeEventData)

Const ProcName As String = "mAppearanceTLI_StateChange"
On Error GoTo Err

Select Case ev.state
Case TimerListItemStates.TimerListItemStateExpired
    If Not mFinished Then
        changeAppearance
        Set mAppearanceTLI = GetGlobalTimerList.Add(Empty, Rnd * 9500 + 500, ExpiryTimeUnitMilliseconds)
    End If
Case Else
End Select

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mPositionTLI Event Handlers
'@================================================================================

Private Sub mPositionTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mPositionTLI_StateChange"
On Error GoTo Err

Dim et As Single

Select Case ev.state
Case TimerListItemStates.TimerListItemStateExpired
    If Not mFinished Then
        et = mElapsedTimer.ElapsedTimeMicroseconds / 1000000
        'Debug.Print "Time since last update: " & et
        mElapsedTimer.StartTiming
        mSprite.Edit
        calcNewPosition et
        calcNewOrientation et
        mSprite.ApplyChanges
        Set mPositionTLI = GetGlobalTimerList.Add(Empty, calcTimeToNextCalc, ExpiryTimeUnitMilliseconds)
        'Debug.Print "Sprite update time: " & Format(mElapsedTimer.ElapsedTimeMicroseconds, "0")
        
        'mElapsedTimer.StartTiming
    End If
Case Else
End Select

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

If Not mPositionTLI Is Nothing Then GetGlobalTimerList.Remove mPositionTLI
If Not mAppearanceTLI Is Nothing Then GetGlobalTimerList.Remove mAppearanceTLI
If Not mTaskContext Is Nothing Then mTaskContext.Finish 0, False
mFinished = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Go()
Const ProcName As String = "Go"
On Error GoTo Err

mSprite.Position = mPosition
'Set mPositionTLI = GetGlobalTimerList.Add(Empty, 20, ExpiryTimeUnitMilliseconds)
'StartTask Me, PriorityNormal
Set mAppearanceTLI = GetGlobalTimerList.Add(Empty, Rnd * 9500 + 500, ExpiryTimeUnitMilliseconds)
mElapsedTimer.StartTiming

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Initialise( _
                ByVal pSprite As ISprite, _
                ByVal pSpeed As Double, _
                ByVal pDirection As Double, _
                ByVal pPosition As Point, _
                ByVal pSecondsPerRevolution As Double, _
                ByVal pGradientFill As Boolean, _
                ByRef pSpaceBoundary As TRectangle)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mSprite = pSprite
mSpeed = pSpeed
mDirection = pDirection
Set mPosition = pPosition
mSecondsPerRevolution = pSecondsPerRevolution
mGradientFill = pGradientFill
mSpaceBoundary = pSpaceBoundary

Set mElapsedTimer = New ElapsedTimer

calcVelocityComponents

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Move()
Dim et As Single

Const ProcName As String = "Move"
On Error GoTo Err

If Not mFinished Then
    et = mElapsedTimer.ElapsedTimeMicroseconds / 1000000
    mElapsedTimer.StartTiming
    mSprite.Edit
    calcNewPosition et
    calcNewOrientation et
    mSprite.ApplyChanges
End If

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcNewOrientation(ByVal pElapsedtime As Single)
Const ProcName As String = "calcNewOrientation"
On Error GoTo Err

If mSecondsPerRevolution = 0 Then Exit Sub
mOrientation = mOrientation + pElapsedtime * 2 * Pi / mSecondsPerRevolution
mSprite.Orientation = mOrientation

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub calcNewPosition(ByVal pElapsedtime As Single)
Dim dx As Double
Dim dy As Double
Dim lBoundary As TRectangle
Dim lNewBoundary As TRectangle

Const ProcName As String = "calcNewPosition"
On Error GoTo Err

dx = pElapsedtime * mVx
dy = pElapsedtime * mVy

lBoundary = mSprite.BoundingRectangle
lNewBoundary = lBoundary

Do While dx <> 0
    If dx > 0 Then
        If lNewBoundary.Right + dx >= mSpaceBoundary.Right Then
            dx = -(dx - (mSpaceBoundary.Right - lNewBoundary.Right))
            TRectangleTranslate lNewBoundary, mSpaceBoundary.Right - lNewBoundary.Right, 0
            reverseX
        Else
            TRectangleTranslate lNewBoundary, dx, 0
            dx = 0
        End If
    Else
        If lNewBoundary.Left + dx <= mSpaceBoundary.Left Then
            dx = -(dx - (mSpaceBoundary.Left - lNewBoundary.Left))
            TRectangleTranslate lNewBoundary, mSpaceBoundary.Left - lNewBoundary.Left, 0
            reverseX
        Else
            TRectangleTranslate lNewBoundary, dx, 0
            dx = 0
        End If
    End If
Loop

Do While dy <> 0
    If dy > 0 Then
        If lNewBoundary.Top + dy >= mSpaceBoundary.Top Then
            dy = -(dy - (mSpaceBoundary.Top - lNewBoundary.Top))
            TRectangleTranslate lNewBoundary, 0, mSpaceBoundary.Top - lNewBoundary.Top
            reverseY
        Else
            TRectangleTranslate lNewBoundary, 0, dy
            dy = 0
        End If
    Else
        If lNewBoundary.Bottom + dy <= mSpaceBoundary.Bottom Then
            dy = -(dy - (mSpaceBoundary.Bottom - lNewBoundary.Bottom))
            TRectangleTranslate lNewBoundary, 0, mSpaceBoundary.Bottom - lNewBoundary.Bottom
            reverseY
        Else
            TRectangleTranslate lNewBoundary, 0, dy
            dy = 0
        End If
    End If
Loop

Dim lPosn As Point
Set lPosn = mSprite.Position
mSprite.Position = NewPoint(lPosn.X + lNewBoundary.Left - lBoundary.Left, _
                            lPosn.Y + lNewBoundary.Top - lBoundary.Top)


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function calcTimeToNextCalc() As Long

'calcTimeToNextCalc = 1000 / (mSpeed / 0.25)
'If calcTimeToNextCalc > 10 Then calcTimeToNextCalc = 10
'If calcTimeToNextCalc < 5 Then calcTimeToNextCalc = 5
Const ProcName As String = "calcTimeToNextCalc"
On Error GoTo Err

calcTimeToNextCalc = 50

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub calcVelocityComponents()
mVx = mSpeed * Cos(mDirection)
mVy = mSpeed * Sin(mDirection)
End Sub

Private Sub changeAppearance()
Const ProcName As String = "changeAppearance"
On Error GoTo Err

ReDim lcolors(1) As Long
lcolors(0) = CLng(Rnd * &HFFFFFF)
lcolors(1) = CLng(Rnd * &HFFFFFF)
If mGradientFill Then
    mSprite.Brush = CreateRadialGradientBrush(TPoint(0.75, 0.75), TPoint(0.5, 0.5), 0.5, 0.5, True, lcolors)
Else
    mSprite.Brush = CreateBrush(lcolors(0))
End If
mSprite.Pen = CreatePixelPen(lcolors(1), Int(Rnd * 5) + 1)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub reverseX()
Const ProcName As String = "reverseX"
On Error GoTo Err

mDirection = Pi - mDirection
calcVelocityComponents

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub reverseY()
Const ProcName As String = "reverseY"
On Error GoTo Err

mDirection = -mDirection
calcVelocityComponents

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub
