VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12930
   LinkTopic       =   "Form1"
   ScaleHeight     =   9720
   ScaleWidth      =   12930
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox ControlsPicture 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12930
      TabIndex        =   1
      Top             =   8865
      Width           =   12930
      Begin VB.CommandButton DoItButton 
         Caption         =   "Do it"
         Default         =   -1  'True
         Height          =   495
         Left            =   5400
         TabIndex        =   2
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox DrawingPicture 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   0
      ScaleHeight     =   2655
      ScaleWidth      =   12930
      TabIndex        =   0
      Top             =   0
      Width           =   12930
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

Private mBarSeries                                  As OHLCBarSeries
Private mDataPointSeries                            As DataPointSeries

Private WithEvents mUnhandledErrorHandler           As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1
Private mIsInDev                                    As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
InitialiseTWUtilities

Set mUnhandledErrorHandler = UnhandledErrorHandler
ApplicationGroupName = "TradeWright"
ApplicationName = "TestGraphObj"
SetupDefaultLogging Command
Randomize
End Sub

Private Sub Form_Resize()
DrawingPicture.Height = ControlsPicture.Top
DoItButton.Left = Me.ScaleWidth - DoItButton.Width - 120
End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

If Not mGraphics Is Nothing Then mGraphics.Finish

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

Private Sub IExtendedEventListener_Notify(ev As ExtendedEventData)

End Sub

'@================================================================================
' Form Control Handlers
'@================================================================================

Private Sub DoItButton_Click()
Const ProcName As String = "DoItButton_Click"
On Error GoTo Err

LogMessage "Setup graphics"
setupGraphics

LogMessage "Disable drawing"
mController.IsDrawingEnabled = False

LogMessage "Generate bars"
showBars

LogMessage "Generate datapoints"
showDataPoints

LogMessage "Enable drawing"
mController.IsDrawingEnabled = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Private Function addBar(ByVal pX As Long, ByVal pPrevClose As Double) As OHLCBar
Const ProcName As String = "addBar"
On Error GoTo Err

Dim lBar As OHLCBar

Set lBar = mBarSeries.Add

lBar.X = pX
lBar.OpenValue = pPrevClose + (11 * Rnd - 5) * 0.5
lBar.HighValue = lBar.OpenValue + (40 * Rnd) * 0.5
lBar.LowValue = lBar.OpenValue - (40 * Rnd) * 0.5
lBar.CloseValue = 0.5 * Int(((Int(Rnd * 11) / 10) * (lBar.HighValue - lBar.LowValue)) / 0.5) + lBar.LowValue
lBar.Width = Rnd * 2 + 0.3
lBar.Orientation = Rnd * 0.3 - 0.15

Debug.Print "Bar: " & pX & "," & lBar.OpenValue & "," & lBar.HighValue & "," & lBar.LowValue & "," & lBar.CloseValue
Set addBar = lBar
Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub addBars()
Const ProcName As String = "addBars"
On Error GoTo Err

Dim i As Long
Dim lPrevClose As Double

lPrevClose = 5250
For i = 1 To 300
    lPrevClose = addBar(i, lPrevClose).CloseValue
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub addDataPoints()
Const ProcName As String = "addDataPoints"
On Error GoTo Err

Dim lBar As OHLCBar
Dim lDp As DataPoint

For Each lBar In mBarSeries
    Set lDp = mDataPointSeries.Add
    lDp.X = lBar.X
    lDp.Value = (lBar.OpenValue + lBar.HighValue + lBar.LowValue + lBar.CloseValue) / 4
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub createBarSeries()
Const ProcName As String = "createBarSeries"
On Error GoTo Err

Set mBarSeries = mModel.AddGraphicObjectSeries(New OHLCBarSeries, LayerNumbers.LayerLowestUser)
mBarSeries.MouseEnterEvent.AddExtendedEventListener mBarSeries, Me
mBarSeries.MouseLeaveEvent.AddExtendedEventListener mBarSeries, Me

'mBarSeries.UpBrush = CreateBrush(&H7F00&)
mBarSeries.UpPen = CreatePixelPen(&H7F00&, 3, LineInsideSolid)
'mBarSeries.DownBrush = CreateBrush(&H7F&)
mBarSeries.DownPen = CreatePixelPen(&H7F&, 3, LineInsideSolid)
'mBarSeries.Pen = CreatePixelPen(vbWhite, , LineInsideSolid)

ReDim lcolors(1) As Long
lcolors(0) = &HFF00&
lcolors(1) = &H7F00&
mBarSeries.UpBrush = CreateRadialGradientBrush(TPoint(0.75, 0.75), TPoint(0.5, 0.5), 0.75, 0.75, True, lcolors)

ReDim lcolors(1) As Long
lcolors(0) = &HFF&
lcolors(1) = &H7F&
mBarSeries.DownBrush = CreateRadialGradientBrush(TPoint(0.75, 0.25), TPoint(0.5, 0.5), 0.75, 0.75, True, lcolors)

mBarSeries.IncludeInAutoscale = True
mBarSeries.DisplayMode = OHLCBarDisplayModeCandlestick
'mBarSeries.Orientation = -Pi / 36
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub createDataPointSeries()
Const ProcName As String = "createDataPointSeries"
On Error GoTo Err

Set mDataPointSeries = mModel.AddGraphicObjectSeries(New DataPointSeries, LayerNumbers.LayerLowestUser + 5)
mDataPointSeries.MouseEnterEvent.AddExtendedEventListener mDataPointSeries, Me
mDataPointSeries.MouseLeaveEvent.AddExtendedEventListener mDataPointSeries, Me

'mDataPointSeries.UpBrush = CreateBrush(&H7F00&)
mDataPointSeries.UpPen = CreatePixelPen(&HFF7F7F, , LineInsideSolid)
'mDataPointSeries.DownBrush = CreateBrush(&H7F&)
mDataPointSeries.DownPen = CreatePixelPen(&H7F7FFF, , LineInsideSolid)
'mDataPointSeries.Pen = CreatePixelPen(vbWhite, , LineInsideSolid)

ReDim lcolors(1) As Long
lcolors(0) = &HFF0000
lcolors(1) = &HFF7F7F
mDataPointSeries.UpBrush = CreateRadialGradientBrush(TPoint(0.5, 0.5), TPoint(0.5, 0.5), 0.75, 0.75, True, lcolors)

ReDim lcolors(1) As Long
lcolors(0) = &HFF&
lcolors(1) = &H7F7FFF
mDataPointSeries.DownBrush = CreateRadialGradientBrush(TPoint(0.5, 0.5), TPoint(0.5, 0.5), 0.75, 0.75, True, lcolors)

mDataPointSeries.IncludeInAutoscale = True
mDataPointSeries.DisplayMode = DataPointDisplayModeEllipse
mDataPointSeries.NumberOfSides = 5
mDataPointSeries.Orientation = Pi / 4#
mDataPointSeries.LineMode = DataPointLineModeNone
mDataPointSeries.LinePen = CreatePixelPen(vbCyan, 1)
mDataPointSeries.Size = NewSize(20, 10, ScaleUnitPixels, ScaleUnitPixels)
Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Private Sub setupGraphics()
Const ProcName As String = "setupGraphics"
On Error GoTo Err

If Not mGraphics Is Nothing Then
    mBarSeries.MouseEnterEvent.RemoveExtendedEventListener mBarSeries, Me
    mBarSeries.MouseLeaveEvent.RemoveExtendedEventListener mBarSeries, Me
    mGraphics.Finish
    Set mGraphics = Nothing
    mController.Clear
    mController.Finish
    Set mController = Nothing
    Set mModel = Nothing
End If

Set mGraphics = CreateGraphics(DrawingPicture.hWnd, 150, 5200, 270, 5300, CreateBrush(&H440000))
If mGraphics Is Nothing Then Stop
mGraphics.PaintBackground

Set mController = CreateLayeredGraphicsEngine(mGraphics, 1, 0, True, 10)
Set mModel = mController.Model
mController.Autoscaling = True

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub showBars()
Const ProcName As String = "showBars"
On Error GoTo Err

createBarSeries

addBars

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub showDataPoints()
Const ProcName As String = "showDataPoints"
On Error GoTo Err

createDataPointSeries

addDataPoints

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub


