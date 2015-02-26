VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPointDefaultRenderer"
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

' TODO: replace DataPoint with the related graphic object class name
Private Const ModuleName                            As String = "DataPointDefaultRenderer"

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

Dim lObj As DataPoint

Set lObj = pObject

IRenderer_GetBoundingRectangle = determineBoundingRect(lObj)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub IRenderer_Render(ByVal pObject As IGraphicObject)
Const ProcName As String = "IRenderer_Render"
On Error GoTo Err

Dim lObj As DataPoint

Set lObj = pObject

' uncomment the following line during development to make the bounding rectangle
' visible in the display
'mGraphics.FillObjFromTRectangle GetBrush(&HF0F0F0), pObject.BoundingRectangle

paint lObj, chooseBrush(lObj), choosePen(lObj)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Private Function chooseBrush(ByVal pObj As DataPoint) As IBrush
If upPoint(pObj) Then
    Set chooseBrush = pObj.UpBrush
Else
    Set chooseBrush = pObj.DownBrush
End If
If chooseBrush Is Nothing Then Set chooseBrush = pObj.Brush
End Function

Private Function chooseLinePen(ByVal pObj As DataPoint) As Pen
If upPoint(pObj) Then
    Set chooseLinePen = pObj.UpLinePen
Else
    Set chooseLinePen = pObj.DownLinePen
End If
If chooseLinePen Is Nothing Then Set chooseLinePen = pObj.LinePen
End Function

Private Function choosePen(ByVal pObj As DataPoint) As Pen
If upPoint(pObj) Then
    Set choosePen = pObj.UpPen
Else
    Set choosePen = pObj.DownPen
End If
If choosePen Is Nothing Then Set choosePen = pObj.Pen
End Function

Private Function determineBoundingRect(ByVal pObj As DataPoint) As TRectangle
Const ProcName As String = "determineBoundingRect"
On Error GoTo Err

mGraphics.BeginBoundingRectangleDetermination
paint pObj, chooseBrush(pObj), choosePen(pObj)
determineBoundingRect = mGraphics.EndBoundingRectangleDetermination

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub paint( _
                ByVal pObj As DataPoint, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paint"
On Error GoTo Err

paintLine pObj, chooseLinePen(pObj)
paintPrevBody pObj, pBrush, pPen

paintBody pObj, pBrush, pPen

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintAsCross( _
                ByVal pObj As DataPoint, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paintAsCross"
On Error GoTo Err

Dim lSize As Size
Dim lX As Double
Dim lY As Double
Dim lHorizOffset As Double
Dim lVertOffset As Double
lX = pObj.X
lY = pObj.Value
Set lSize = pObj.Size
lHorizOffset = InverseSqrtOf2 * lSize.WidthLogical(mGraphics) / 2#
lVertOffset = InverseSqrtOf2 * lSize.HeightLogical(mGraphics) / 2#
mGraphics.DrawLine pPen, NewPoint(lX - lHorizOffset, lY - lVertOffset), NewPoint(lX + lHorizOffset, lY + lVertOffset)
mGraphics.DrawLine pPen, NewPoint(lX - lHorizOffset, lY + lVertOffset), NewPoint(lX + lHorizOffset, lY - lVertOffset)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintAsDash( _
                ByVal pObj As DataPoint, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paintAsDash"
On Error GoTo Err

Dim lSemiWidth As Double
Dim lX As Double
Dim lY As Double
lX = pObj.X
lY = pObj.Value
lSemiWidth = pObj.Size.WidthLogical(mGraphics) / 2#
mGraphics.DrawLine pPen, NewPoint(lX - lSemiWidth, lY), NewPoint(lX + lSemiWidth, lY)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintAsEllipse( _
                ByVal pObj As DataPoint, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paintAsEllipse"
On Error GoTo Err

mGraphics.PaintEllipseFromPositionAndSize pBrush, pPen, NewPoint(pObj.X, pObj.Value), pObj.Size

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintAsHistogram( _
                ByVal pObj As DataPoint, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paintAsHistogram"
On Error GoTo Err

Dim lSemiWidth As Double
Dim lX
Dim lHistBase As Double
lX = pObj.X
lSemiWidth = pObj.HistogramBarWidth / 2#
lHistBase = pObj.HistogramBaseValue
mGraphics.paintRectangle pBrush, _
                        pPen, _
                        NewPoint(lX - lSemiWidth, lHistBase), _
                        NewPoint(lX + lSemiWidth, pObj.Value - lHistBase)

Dim lBaselineX As Double
If Not pObj.PreviousDataPoint Is Nothing Then
    lBaselineX = pObj.PreviousDataPoint.X + pObj.PreviousDataPoint.HistogramBarWidth / 2#
Else
    lBaselineX = lX - lSemiWidth
End If
mGraphics.DrawLine pObj.HistogramBaselinePen, NewPoint(lBaselineX, lHistBase), NewPoint(lX + lSemiWidth, lHistBase)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintAsPoint( _
                ByVal pObj As DataPoint, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paintAsPoint"
On Error GoTo Err

mGraphics.DrawPoint pPen, NewPoint(pObj.X, pObj.Value)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintAsPolygon( _
                ByVal pObj As DataPoint, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paintAsPolygon"
On Error GoTo Err

mGraphics.PaintPolygonFromPositionAndSize pBrush, pPen, pObj.NumberOfSides, NewPoint(pObj.X, pObj.Value), pObj.Size

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintBody( _
                ByVal pObj As DataPoint, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paintBody"
On Error GoTo Err

setOrientation pObj

Select Case pObj.DisplayMode
Case DataPointDisplayModeNone
Case DataPointDisplayModePoint
    paintAsPoint pObj, pBrush, pPen
Case DataPointDisplayModeDash
    paintAsDash pObj, pBrush, pPen
Case DataPointDisplayModeHistogram
    paintAsHistogram pObj, pBrush, pPen
Case DataPointDisplayModePolygon
    paintAsPolygon pObj, pBrush, pPen
Case DataPointDisplayModeEllipse
    paintAsEllipse pObj, pBrush, pPen
Case DataPointDisplayModeCross
    paintAsCross pObj, pBrush, pPen
End Select

resetOrientation pObj

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintLine( _
                ByVal pObj As DataPoint, _
                ByVal pPen As Pen)
Const ProcName As String = "paintLine"
On Error GoTo Err

Dim lXFrom As Double
Dim lXTo As Double
Dim lYFrom As Double
Dim lYTo As Double
Dim lXMid As Double

If pObj.LineMode = DataPointLineModeNone Then Exit Sub

If pObj.PreviousDataPoint Is Nothing Then Exit Sub

lXFrom = pObj.PreviousDataPoint.X
lXTo = pObj.X
lYFrom = pObj.PreviousDataPoint.Value
lYTo = pObj.Value
lXMid = (lXFrom + lXTo) / 2#

Select Case pObj.LineMode
Case DataPointLineModeNone
Case DataPointLineModeStraight
    mGraphics.DrawLine pPen, NewPoint(lXFrom, lYFrom), NewPoint(lXTo, lYTo)
Case DataPointLineModeStepped
    mGraphics.DrawLine pPen, NewPoint(lXFrom, lYFrom), NewPoint(lXMid, lYFrom)
    mGraphics.DrawLine pPen, NewPoint(lXMid, lYFrom), NewPoint(lXMid, lYTo)
    mGraphics.DrawLine pPen, NewPoint(lXMid, lYTo), NewPoint(lXTo, lYTo)
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub paintPrevBody( _
                ByVal pObj As DataPoint, _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen)
Const ProcName As String = "paintPrevBody"
On Error GoTo Err

Dim lPrevObj As DataPoint
Set lPrevObj = pObj.PreviousDataPoint
If lPrevObj Is Nothing Then Exit Sub

paintBody lPrevObj, chooseBrush(lPrevObj), choosePen(lPrevObj)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub resetOrientation(ByVal pObj As DataPoint)
Const ProcName As String = "resetOrientation"
On Error GoTo Err

If pObj.Orientation <> 0 Then mGraphics.Reset

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub setOrientation(ByVal pObj As DataPoint)
Const ProcName As String = "setOrientation"
On Error GoTo Err

If pObj.Orientation = 0 Then Exit Sub

Select Case pObj.DisplayMode
Case DataPointDisplayModeDash
Case DataPointDisplayModePolygon
Case DataPointDisplayModeEllipse
Case DataPointDisplayModeCross
Case Else
    Exit Sub
End Select

mGraphics.RotateAboutPoint pObj.Orientation, TPoint(pObj.X, pObj.Value)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function upPoint(ByVal pObj As DataPoint) As Boolean
If pObj.PreviousDataPoint Is Nothing Then
    upPoint = False
Else
    upPoint = pObj.Value >= pObj.PreviousDataPoint.Value
End If
End Function
