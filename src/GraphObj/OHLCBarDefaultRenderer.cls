VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OHLCBarDefaultRenderer"
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

Implements IRenderer

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

Private Const ModuleName                            As String = "OHLCBarDefaultRenderer"

'@================================================================================
' Member variables
'@================================================================================

Private mGraphics                                   As Graphics

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IRenderer Interface Members
'@================================================================================

Private Function IRenderer_GetBoundingRectangle(ByVal pObject As IGraphicObject) As TRectangle
Const ProcName As String = "IRenderer_GetBoundingRectangle"
On Error GoTo Err

Dim lBar As OHLCBar

Set lBar = pObject

IRenderer_GetBoundingRectangle = determineBoundingRect(lBar)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IRenderer_Render(ByVal pObject As IGraphicObject)
Const ProcName As String = "IRenderer_Render"
On Error GoTo Err

Dim lBar As OHLCBar

Set lBar = pObject

' uncomment the following line during development to make the bounding rectangle
' visible in the display
'mGraphics.FillBarFromTRectangle GetBrush(&HF0F0F0), pObject.BoundingRectangle

paint lBar, chooseBrush(lBar), choosePen(lBar)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Graphics(ByVal Value As Graphics)
Set mGraphics = Value
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcYAdjustment( _
                ByVal pValue As Double, _
                ByVal pBar As OHLCBar, _
                ByVal pMaxAdjustment As Double) As Double
If (pValue - pMaxAdjustment) < pBar.LowValue Then
    calcYAdjustment = pMaxAdjustment - (pValue - pBar.LowValue)
ElseIf (pValue + pMaxAdjustment) > pBar.HighValue Then
    calcYAdjustment = pBar.HighValue - pValue - pMaxAdjustment
Else
    calcYAdjustment = 0#
End If
End Function

Private Function chooseBrush(ByVal pBar As OHLCBar) As IBrush
Set chooseBrush = pBar.Brush
If chooseBrush Is Nothing Then
    If upBar(pBar) Then
        Set chooseBrush = pBar.UpBrush
    Else
        Set chooseBrush = pBar.DownBrush
    End If
End If
End Function

Private Function choosePen(ByVal pBar As OHLCBar) As Pen
Set choosePen = pBar.Pen
If choosePen Is Nothing Then
    If upBar(pBar) Then
        Set choosePen = pBar.UpPen
    Else
        Set choosePen = pBar.DownPen
    End If
End If
End Function

Private Function determineBoundingRect(ByVal pBar As OHLCBar) As TRectangle
Const ProcName As String = "determineBoundingRect"
On Error GoTo Err

mGraphics.BeginBoundingRectangleDetermination
paint pBar, chooseBrush(pBar), choosePen(pBar)
determineBoundingRect = mGraphics.EndBoundingRectangleDetermination

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub paint( _
                ByVal pBar As OHLCBar, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paint"

On Error GoTo Err

setOrientation pBar

Select Case pBar.DisplayMode
Case OHLCBarDisplayModeCandlestick
    paintAsCandlestick pBar, pBrush, pPen
Case OHLCBarDisplayModeLine
    paintAsLine pBar, pPen
Case Else
    paintAsBar pBar, pPen
End Select

resetOrientation pBar

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub paintAsBar( _
                ByVal pBar As OHLCBar, _
                ByVal pPen As Pen)
Const ProcName As String = "paintBar"

On Error GoTo Err

Dim lIsPixelPen As Boolean
Dim lLogicalSemiThicknessX As Double
Dim lLogicalSemiThicknessY As Double
Dim lWidth As Long

lIsPixelPen = pPen.IsPixelPen
lWidth = pPen.Width

If lIsPixelPen And lWidth = 1& Then
    paintTheBar pBar, pPen, 0#, 0#, 0#
Else
    If lIsPixelPen Then
        lLogicalSemiThicknessX = mGraphics.ConvertPixelsToLogicalX(lWidth) / 2#
        lLogicalSemiThicknessY = mGraphics.ConvertPixelsToLogicalY(lWidth) / 2#
    Else
        lLogicalSemiThicknessX = lWidth / 2#
        lLogicalSemiThicknessY = mGraphics.ConvertLogicalXtoLogicalY(lWidth) / 2#
    End If
    
    paintTheBar pBar, _
                pPen, _
                calcYAdjustment(pBar.OpenValue, pBar, lLogicalSemiThicknessY), _
                calcYAdjustment(pBar.CloseValue, pBar, lLogicalSemiThicknessY), _
                lLogicalSemiThicknessX

End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub paintAsCandlestick( _
                ByVal pBar As OHLCBar, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paintCandlestick"
On Error GoTo Err

Dim lX As Double
Dim lSemiWidth As Double

lX = pBar.X
lSemiWidth = pBar.Width / 2#

paintTail pBar, pPen
paintBody pBrush, pPen, lX - lSemiWidth, pBar.OpenValue, lX + lSemiWidth, pBar.CloseValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub paintAsLine( _
                ByVal pBar As OHLCBar, _
                ByVal pPen As Pen)
Dim prevBar As OHLCBar
Const ProcName As String = "paintAsLine"
On Error GoTo Err

Dim lX As Double
Dim lClose As Double

lX = pBar.X
lClose = pBar.CloseValue

Set prevBar = pBar.PreviousBar
If Not prevBar Is Nothing Then
    mGraphics.DrawLine pPen, NewPoint(prevBar.X, prevBar.CloseValue), NewPoint(lX, lClose)
Else
    mGraphics.DrawLine pPen, NewPoint(lX, lClose), NewPoint(lX, lClose)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub paintBody( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByVal pX1 As Double, _
                ByVal pY1 As Double, _
                ByVal pX2 As Double, _
                ByVal pY2 As Double)
Const ProcName As String = "paintBody"

On Error GoTo Err

mGraphics.PaintRectangleLogical pBrush, pPen, pX1, pY1, pX2, pY2

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub paintTail( _
                ByVal pBar As OHLCBar, _
                ByVal pPen As Pen)
Const ProcName As String = "paintTail"

On Error GoTo Err

Dim lX As Double
lX = pBar.X
mGraphics.DrawLine pPen, NewPoint(lX, pBar.LowValue), NewPoint(lX, pBar.HighValue)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub paintTheBar( _
                ByVal pBar As OHLCBar, _
                ByVal pPen As Pen, _
                ByVal pOpenAdjustment As Double, _
                ByVal pCloseAdjustment As Double, _
                ByVal pMidAdjustment As Double)
                
Dim lX As Double
Dim lOpen As Double
Dim lClose As Double
Dim lSemiWidth As Double

lX = pBar.X
lOpen = pBar.OpenValue
lClose = pBar.CloseValue
lSemiWidth = pBar.Width / 2#

mGraphics.DrawLine pPen, _
                    NewPoint(lX - lSemiWidth, lOpen + pOpenAdjustment), _
                    NewPoint(lX + pMidAdjustment, lOpen + pOpenAdjustment)
mGraphics.DrawLine pPen, _
                    NewPoint(lX, pBar.LowValue), _
                    NewPoint(lX, pBar.HighValue)
mGraphics.DrawLine pPen, _
                    NewPoint(lX - pMidAdjustment, lClose + pCloseAdjustment), _
                    NewPoint(lX + lSemiWidth + mGraphics.ConvertPixelsToLogicalX(1), lClose + pCloseAdjustment)
End Sub

Private Sub resetOrientation(ByVal pBar As OHLCBar)
Const ProcName As String = "resetOrientation"
On Error GoTo Err

If pBar.Orientation <> 0 Then mGraphics.Reset

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrientation(ByVal pBar As OHLCBar)
Const ProcName As String = "setOrientation"
On Error GoTo Err

If pBar.Orientation = 0 Then Exit Sub

mGraphics.RotateAboutPoint pBar.Orientation, TPoint(pBar.X, (pBar.LowValue + pBar.HighValue) / 2#)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function upBar(ByVal pBar As OHLCBar) As Boolean
upBar = (pBar.CloseValue >= pBar.OpenValue)
End Function








