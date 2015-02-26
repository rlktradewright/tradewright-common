VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H8000000C&
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13230
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   13230
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   0
      ScaleHeight     =   6015
      ScaleWidth      =   13230
      TabIndex        =   6
      Top             =   0
      Width           =   13230
      Begin VB.CommandButton BigButton 
         Height          =   4575
         Left            =   3600
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   5775
      End
   End
   Begin VB.PictureBox ControlsPicture 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   0
      ScaleHeight     =   1575
      ScaleWidth      =   13230
      TabIndex        =   10
      Top             =   6780
      Width           =   13230
      Begin VB.TextBox DeferIntervalText 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5880
         TabIndex        =   20
         Text            =   "20"
         Top             =   1080
         Width           =   615
      End
      Begin VB.CheckBox DeferredPaintingCheck 
         Caption         =   "Use deferred painting"
         Height          =   255
         Left            =   3960
         TabIndex        =   19
         Top             =   840
         Width           =   2415
      End
      Begin VB.CheckBox ShowSunCheck 
         Caption         =   "Show sun"
         Height          =   255
         Left            =   3960
         TabIndex        =   18
         Top             =   600
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox GradientFillSpritesCheck 
         Caption         =   "Use gradient fill for sprites"
         Height          =   255
         Left            =   3960
         TabIndex        =   17
         Top             =   360
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CheckBox GradientBackgroundCheck 
         Caption         =   "Use gradient fill background"
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Top             =   120
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.CommandButton GoStopButton 
         Caption         =   "Go"
         Default         =   -1  'True
         Height          =   495
         Left            =   12600
         TabIndex        =   2
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton LeftButton 
         Caption         =   "<-"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
      Begin VB.CommandButton RightButton 
         Caption         =   "->"
         Enabled         =   0   'False
         Height          =   495
         Left            =   720
         TabIndex        =   4
         Top             =   120
         Width           =   495
      End
      Begin VB.TextBox NumSpritesText 
         Height          =   285
         Left            =   11880
         TabIndex        =   0
         Text            =   "10"
         Top             =   120
         Width           =   615
      End
      Begin VB.CheckBox DrawOnButtonCheck 
         Alignment       =   1  'Right Justify
         Caption         =   "Draw on button"
         Height          =   255
         Left            =   10440
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
      Begin VB.CommandButton RotateLeftButton 
         Caption         =   "Rot left"
         Enabled         =   0   'False
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton RotateRightButton 
         Caption         =   "Rot right"
         Enabled         =   0   'False
         Height          =   495
         Left            =   720
         TabIndex        =   8
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton ChangeSizeButton 
         Caption         =   "Change size"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Defer interval (millisecs)"
         Height          =   255
         Left            =   4200
         TabIndex        =   21
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label NumSpritesLabel 
         Caption         =   "Number of sprites"
         Height          =   255
         Left            =   10440
         TabIndex        =   15
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label LogLabel 
         Height          =   255
         Left            =   1320
         TabIndex        =   14
         Top             =   120
         Width           =   2415
      End
      Begin VB.Label RenderTimeLabel 
         Height          =   255
         Left            =   1320
         TabIndex        =   13
         Top             =   630
         Width           =   2535
      End
      Begin VB.Label MoveTimeLabel 
         Height          =   255
         Left            =   1320
         TabIndex        =   12
         Top             =   375
         Width           =   2415
      End
      Begin VB.Label IterationTimeLabel 
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   885
         Width           =   2535
      End
   End
End
Attribute VB_Name = "Form1"
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

Implements IExtendedEventListener

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

Private Const ModuleName                            As String = "Form1"

'@================================================================================
' Member variables
'@================================================================================

Private mGraphics                                   As Graphics

Private mModel                                      As LayeredGraphicsModel
Private mController                                 As Controller

Private mRectangleSeries1                           As RectangleSeries
Private mRectangleSeries2                           As RectangleSeries

Private mEllipseSeries1                             As EllipseSeries

Private WithEvents mUnhandledErrorHandler           As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1
Private mIsInDev                                    As Boolean

Private mRect1                                      As Rectangle
Private mRect2                                      As Rectangle

Private mEllipse1                                   As Ellipse

Private mSprites                                    As New EnumerableCollection

Private WithEvents mSpriteControllerTC              As TaskController
Attribute mSpriteControllerTC.VB_VarHelpID = -1

Private WithEvents mSun                             As Sun
Attribute mSun.VB_VarHelpID = -1
Private mSky                                        As Rectangle

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
Debug.Print "Running in development environment: " & CStr(inDev)
InitialiseTWUtilities
TaskQuantumMillisecs = 16
RunTasksAtLowerThreadPriority = True

Set mUnhandledErrorHandler = UnhandledErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "LayeredGraphicsTest1"
SetupDefaultLogging Command
Randomize
End Sub

Private Sub Form_Resize()
GoStopButton.Left = Me.ScaleWidth - GoStopButton.Width - 120
NumSpritesText.Left = GoStopButton.Left - NumSpritesText.Width - 120
NumSpritesLabel.Left = NumSpritesText.Left - NumSpritesLabel.Width - 120
DrawOnButtonCheck.Left = NumSpritesLabel.Left
Picture1.Height = Me.ScaleHeight - ControlsPicture.Height
End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

'@================================================================================
' IExtendedEventListener Interface Members
'@================================================================================

Private Sub IExtendedEventListener_Notify(ev As ExtendedEventData)
Const ProcName As String = "IExtendedEventListener_Notify"
On Error GoTo Err

Dim lRect As Rectangle
If ev.ExtendedEvent.Name = "MouseEnter" Then
    For Each lRect In mRectangleSeries1
        lRect.Brush = GetBrush(vbMagenta)
    Next
ElseIf ev.ExtendedEvent.Name = "MouseLeave" Then
    For Each lRect In mRectangleSeries1
        lRect.Brush = GetBrush(vbGreen)
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub BigButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "BigButton_MouseDown"
On Error GoTo Err

If Not mGraphics Is Nothing Then mGraphics.Refresh

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub BigButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "BigButton_MouseUp"
On Error GoTo Err

If Not mGraphics Is Nothing Then mGraphics.Refresh

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub ChangeSizeButton_Click()
mRectangleSeries2.Size = NewSize(Rnd * 8 + 2, Rnd * 8 + 2, ScaleUnitLogical, ScaleUnitLogical)
mEllipseSeries1.Size = NewSize(Rnd * 8 + 2, Rnd * 8 + 2, ScaleUnitLogical, ScaleUnitLogical)
End Sub

Private Sub DeferredPaintingCheck_Click()
DeferIntervalText.Enabled = (DeferredPaintingCheck.Value = vbChecked)
End Sub

Private Sub GoStopButton_Click()
Const ProcName As String = "GoStopButton_Click"
On Error GoTo Err

If GoStopButton.Caption = "Go" Then
    If IsInteger(NumSpritesText.Text, 0, 500) And _
        IsInteger(DeferIntervalText.Text, 5, 200) _
    Then
        LeftButton.Enabled = True
        RightButton.Enabled = True
        RotateLeftButton.Enabled = True
        RotateRightButton.Enabled = True
        ChangeSizeButton.Enabled = True
        GoStopButton.Caption = "Stop"
        getGraphics
        startSprites CLng(NumSpritesText.Text)
        createSunAndSky
    End If
Else
    LeftButton.Enabled = False
    RightButton.Enabled = False
    RotateLeftButton.Enabled = False
    RotateRightButton.Enabled = False
    ChangeSizeButton.Enabled = False
    stopSprites
    If Not mSun Is Nothing Then mSun.Finish
    mGraphics.PaintBackground
    mRectangleSeries1.MouseEnterEvent.RemoveExtendedEventListener mRectangleSeries1, Me
    mRectangleSeries1.MouseLeaveEvent.RemoveExtendedEventListener mRectangleSeries1, Me
    Set mRectangleSeries1 = Nothing
    Set mRectangleSeries2 = Nothing
    Set mEllipseSeries1 = Nothing
    Set mRect1 = Nothing
    Set mRect2 = Nothing
    Set mEllipse1 = Nothing
    Set mGraphics = Nothing
    mController.Clear
    mController.Finish
    Set mController = Nothing
    Set mModel = Nothing
    Picture1.Cls
    GoStopButton.Caption = "Go"
End If

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub LeftButton_Click()
Const ProcName As String = "LeftButton_Click"
On Error GoTo Err

mEllipse1.Position = NewPoint(mEllipse1.Position.X - 2, mEllipse1.Position.Y)

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub RightButton_Click()
Const ProcName As String = "RightButton_Click"
On Error GoTo Err

mEllipse1.Position = NewPoint(mEllipse1.Position.X + 2, mEllipse1.Position.Y)

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub RotateLeftButton_Click()
Const ProcName As String = "RotateLeftButton_Click"
On Error GoTo Err

mEllipse1.Orientation = mEllipse1.Orientation + Pi / 8

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub RotateRightButton_Click()
Const ProcName As String = "RotateRightButton_Click"
On Error GoTo Err

mEllipse1.Orientation = mEllipse1.Orientation - Pi / 8

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mSpriteControllerTC Event Handlers
'@================================================================================

Private Sub mSpriteControllerTC_Progress(ev As TaskProgressEventData)
Const ProcName As String = "mSpriteControllerTC_Progress"
On Error GoTo Err

Dim metrics As SpriteControllerMetrics
metrics = ev.InterimResult
LogLabel.Caption = "Renders/sec: " & metrics.RendersPerSecond
RenderTimeLabel.Caption = "Avg render time (microsecs): " & metrics.AverageRenderTime
MoveTimeLabel.Caption = "Avg move time (microsecs): " & metrics.AverageMoveTime
IterationTimeLabel.Caption = "Avg iteration time (microsecs): " & metrics.AverageIterationTime

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mSun Event Handlers
'@================================================================================

Private Sub mSun_PositionChanged()
Const ProcName As String = "mSun_PositionChanged"
On Error GoTo Err

'setSkyBrush mSun.Position, mSun.OrbitSize
If Not mSky Is Nothing Then mSky.Position = mSun.Position

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As ErrorEventData)

handleFatalError

' Tell TWUtilities that we've now handled this unhandled error. Not actually
' needed here because HandleFatalError never returns anyway
UnhandledErrorHandler.Handled = True
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addEllipse( _
                ByVal pSeries As EllipseSeries, _
                ByVal pFillColor As Long, _
                ByVal pLineColor As Long, _
                ByVal pLineThickness As Long, _
                ByVal pPosition As Point, _
                ByVal pLayer As LayerNumbers) As Ellipse
Const ProcName As String = "addEllipse"
On Error GoTo Err

Dim lEllipse As Ellipse
Set lEllipse = pSeries.Add

lEllipse.Position = pPosition

ReDim lcolors(1) As Long
lcolors(0) = pFillColor
lcolors(1) = pLineColor
If GradientFillSpritesCheck.Value = vbChecked Then
    lEllipse.Brush = CreateRadialGradientBrush(TPoint(0.75, 0.75), TPoint(0.5, 0.5), 0.5, 0.5, True, lcolors)
Else
    lEllipse.Brush = CreateBrush(lcolors(1))
End If

If pLineThickness <> 0 Then
    lEllipse.Pen = CreatePixelPen(pLineColor, pLineThickness)
Else
    lEllipse.ClearPen
End If

lEllipse.Layer = pLayer

Set addEllipse = lEllipse

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function addRect( _
                ByVal pSeries As RectangleSeries, _
                ByVal pFillColor As Long, _
                ByVal pLineColor As Long, _
                ByVal pLineThickness As Long, _
                ByVal pPosition As Point, _
                ByVal pLayer As LayerNumbers) As Rectangle
Const ProcName As String = "addRect"
On Error GoTo Err

Dim lRect As Rectangle
Set lRect = pSeries.Add

lRect.Position = pPosition

ReDim lcolors(1) As Long
lcolors(0) = pFillColor
lcolors(1) = pLineColor
If GradientFillSpritesCheck.Value = vbChecked Then
    lRect.Brush = CreateRadialGradientBrush(TPoint(0.75, 0.75), TPoint(0.5, 0.5), 0.5, 0.5, True, lcolors)
Else
    lRect.Brush = CreateBrush(lcolors(0))
End If

If pLineThickness <> 0 Then
    lRect.Pen = CreatePixelPen(pLineColor, pLineThickness)
Else
    lRect.ClearPen
End If
lRect.Layer = pLayer

Set addRect = lRect

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub createSunAndSky()
Const ProcName As String = "createSunAndSky"
On Error GoTo Err

If ShowSunCheck.Value = vbUnchecked Then Exit Sub

Set mSun = New Sun
mSun.Initialise mEllipseSeries1.Add, NewPoint(80, -30), NewSize(130, 220), 5# * Pi / 4#, mGraphics

Set mSky = mRectangleSeries1.Add
mSky.Layer = LayerBackground + 1
mSky.Size = NewSize(300, 300)
mSky.Position = mSun.Position

setSkyBrush mSun.Position, mSun.OrbitSize

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub getGraphics()
Const ProcName As String = "getGraphics"
On Error GoTo Err

If Not mGraphics Is Nothing Then Stop

ReDim lcolors(1) As Long
lcolors(0) = &HF8F0F0
lcolors(1) = &H7F0000
BigButton.Visible = (DrawOnButtonCheck.Value = vbChecked)
Set mGraphics = CreateGraphics(IIf(DrawOnButtonCheck.Value = vbChecked, BigButton.hWnd, Picture1.hWnd), _
                                0, _
                                0, _
                                100, _
                                100, _
                                IIf(GradientBackgroundCheck.Value = vbChecked, CreateRadialGradientBrush(TPoint(0.25, -0.05), TPoint(0.25, -0.05), 1.5, 1.5, True, lcolors), CreateBrush(lcolors(1))))

If mGraphics Is Nothing Then Stop

mGraphics.PaintBackground

Set mController = CreateLayeredGraphicsEngine(mGraphics, 1, 0, (DeferredPaintingCheck.Value = vbChecked), CLng(DeferIntervalText.Text))
Set mModel = mController.Model

Set mRectangleSeries1 = mModel.AddGraphicObjectSeries(New RectangleSeries, LayerNumbers.LayerLowestUser)
mRectangleSeries1.MouseEnterEvent.AddExtendedEventListener mRectangleSeries1, Me
mRectangleSeries1.MouseLeaveEvent.AddExtendedEventListener mRectangleSeries1, Me
mRectangleSeries1.Size = NewSize(5, 5, ScaleUnitLogical, ScaleUnitLogical)
Set mRect1 = addRect(mRectangleSeries1, vbRed, vbBlue, 0, NewPoint(50, 50), LayerLowestUser + (LayerHighestUser - LayerLowestUser) / 2)
Set mRect2 = addRect(mRectangleSeries1, vbYellow, vbWhite, 5, NewPoint(25, 25), LayerLowestUser + (LayerHighestUser - LayerLowestUser) / 2)

Set mEllipseSeries1 = mModel.AddGraphicObjectSeries(New EllipseSeries, LayerNumbers.LayerLowestUser)
mEllipseSeries1.Size = NewSize(5, 5, ScaleUnitLogical, ScaleUnitLogical)
Set mEllipse1 = addEllipse(mEllipseSeries1, vbYellow, &H7F00, 3, NewPoint(55, 47), LayerLowestUser + (LayerHighestUser - LayerLowestUser) * 2 / 3)

Set mRectangleSeries2 = mModel.AddGraphicObjectSeries(New RectangleSeries, LayerNumbers.LayerLowestUser)
mRectangleSeries2.Size = NewSize(3, 3, ScaleUnitLogical, ScaleUnitLogical)
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub handleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly.
'
' Note that normally one would use the End statement to terminate a VB6 program abruptly. But
' the TWUtilities component interferes with the End statement's processing and may prevent
' proper shutdown, so we use the TWUtilities component's EndProcess method instead.
'
' However if we are running in the development environment, then we call End because the
' EndProcess method kills the entire development environment as well which can have undesirable
' side effects if other components are also loaded.

If mIsInDev Then
    End
Else
    EndProcess
End If

End Sub

Private Function inDev() As Boolean
mIsInDev = True
inDev = True
End Function

Private Sub setSkyBrush( _
                ByVal pSunPosition As Point, _
                ByVal pOrbitSize As Size)
Const ProcName As String = "setSkyBrush"
On Error GoTo Err

ReDim lcolors(1) As Long
lcolors(0) = &HFFFCFC
lcolors(1) = &H400000

Dim et As New ElapsedTimer
et.StartTiming

Dim lBrush As IBrush
Set lBrush = CreateRadialGradientBrush( _
                TPoint(pSunPosition.XLogical(mGraphics) / 100#, pSunPosition.YLogical(mGraphics) / 100#), _
                TPoint(pSunPosition.XLogical(mGraphics) / 100#, pSunPosition.YLogical(mGraphics) / 100#), _
                pOrbitSize.WidthLogical(mGraphics) / 100#, _
                pOrbitSize.HeightLogical(mGraphics) / 100#, _
                True, _
                lcolors)
Debug.Print "CreateBrush: " & et.ElapsedTimeMicroseconds

et.StartTiming
mSky.Brush = lBrush
Debug.Print "SetBrush: " & et.ElapsedTimeMicroseconds

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub startSprite()
Const ProcName As String = "startSprite"
On Error GoTo Err

Dim lGraphObj As ISprite
If Rnd < 0.5 Then
    Set lGraphObj = addRect(mRectangleSeries2, _
            CLng(Rnd * &HFFFFFF), _
            CLng(Rnd * &HFFFFFF), _
            CLng(Rnd * 5) + 1, _
            NewPoint(Rnd * 50 + 25, Rnd * 50 + 25), _
            LayerNumbers.LayerMin + Int(Rnd * LayerNumbers.LayerMax - LayerNumbers.LayerMin))
Else
    Set lGraphObj = addEllipse(mEllipseSeries1, _
            CLng(Rnd * &HFFFFFF), _
            CLng(Rnd * &HFFFFFF), _
            CLng(Rnd * 5) + 1, _
            NewPoint(Rnd * 50 + 25, Rnd * 50 + 25), _
            LayerNumbers.LayerMin + Int(Rnd * LayerNumbers.LayerMax - LayerNumbers.LayerMin))
End If

Dim lSprite As Sprite
Set lSprite = New Sprite

lSprite.Initialise lGraphObj, _
                     Rnd * Rnd * 100, _
                    Rnd * 2 * Pi, _
                    NewPoint(Rnd * 50 + 25, Rnd * 50 + 25), _
                    IIf(Rnd < 0.2, 0, -100 + Rnd * 200), _
                    (GradientFillSpritesCheck.Value = vbChecked), _
                    mGraphics.Boundary
mSprites.Add lSprite, CStr(ObjPtr(lSprite))
lSprite.Go

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub startSprites(ByVal pNumber As Long)
Const ProcName As String = "startSprites"
Dim failpoint As String
On Error GoTo Err

failpoint = "Create sprites"
Dim i As Long
For i = 1 To pNumber
    startSprite
Next

failpoint = "Create SpriteController"
Dim lSpriteController As New SpriteController

failpoint = "Initialise SpriteController"
lSpriteController.Initialise mSprites, mController, (DeferredPaintingCheck.Value = vbUnchecked)

failpoint = "Start SpriteController"
Set mSpriteControllerTC = StartTask(lSpriteController, PriorityNormal)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, pFailpoint:=failpoint
End Sub

Private Sub stopSprites()
Const ProcName As String = "stopSprites"
On Error GoTo Err

If Not mSpriteControllerTC Is Nothing Then mSpriteControllerTC.CancelTask

'mRectangleSeries2.Clear

Dim i As Long
For i = mSprites.Count To 1 Step -1
    Dim lSprite As Sprite
    Set lSprite = mSprites(i)
    lSprite.Finish
    mSprites.Remove i
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

