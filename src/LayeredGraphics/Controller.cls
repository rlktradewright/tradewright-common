VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Controller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "Controller"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mGraphics                        As Graphics
Attribute mGraphics.VB_VarHelpID = -1

Private mModel                                      As LayeredGraphicsModel

Private mPaintingInProgress                         As Boolean

Private mIsDrawingEnabled                           As Boolean

Private mDeferredPainter                            As DeferredPainter

Private mIsCleared                                  As Boolean

Private mAutoscaling                                As Boolean

Private mUseDeferredPainting                        As Boolean

Private mSelectionManager                           As SelectionManager
Private mMouseMovementManager                       As MouseMovementManager

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mModel = New LayeredGraphicsModel
mIsDrawingEnabled = True

Set mSelectionManager = New SelectionManager
mSelectionManager.Initialise mModel

Set mMouseMovementManager = New MouseMovementManager
mMouseMovementManager.Initialise mModel

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mGraphics Event Handlers
'@================================================================================

Private Sub mGraphics_Click()
Const ProcName As String = "mGraphics_Click"
On Error GoTo Err

mSelectionManager.ClickSelectedObjectOrSeries

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mGraphics_DblClick()
Const ProcName As String = "mGraphics_DblClick"
On Error GoTo Err

mSelectionManager.DblClickSelectedObjectOrSeries

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mGraphics_MouseDown( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Const ProcName As String = "mGraphics_MouseDown"

On Error GoTo Err

mSelectionManager.SetSelectedObjectOrSeries X, Y

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mGraphics_MouseMove(pButton As Integer, pShift As Integer, pX As Single, pY As Single)
Const ProcName As String = "mGraphics_MouseMove"
On Error GoTo Err

mMouseMovementManager.MouseMove pButton, pShift, pX, pY

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mGraphics_Resize()
Const ProcName As String = "mGraphics_Resize"
On Error GoTo Err

Dim lBoundary As TRectangle
lBoundary = GetScopeRect

mGraphics.SetScales lBoundary.Left, lBoundary.Bottom, lBoundary.Right, lBoundary.Top
If mUseDeferredPainting Then
    mDeferredPainter.ChangeGaugeAndRepaint
Else
    ChangeGaugeAndRepaint
End If

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Autoscaling() As Boolean
Const ProcName As String = "Autoscaling"

On Error GoTo Err

Autoscaling = mAutoscaling

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Autoscaling(ByVal Value As Boolean)
Const ProcName As String = "Autoscaling"

On Error GoTo Err

mAutoscaling = Value
setAutoscaling

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Graphics() As Graphics
Set Graphics = mGraphics
End Property

Friend Property Get IsCleared() As Boolean
IsCleared = mIsCleared
End Property

Public Property Let IsDrawingEnabled(ByVal Value As Boolean)
mIsDrawingEnabled = Value
End Property

Public Property Get IsDrawingEnabled() As Boolean
IsDrawingEnabled = mIsDrawingEnabled
End Property

Public Property Let IsSelectionEnabled(ByVal Value As Boolean)
mSelectionManager.IsSelectionEnabled = Value
End Property

Public Property Get IsSelectionEnabled() As Boolean
IsSelectionEnabled = mSelectionManager.IsSelectionEnabled
End Property

Public Property Get Model() As LayeredGraphicsModel
Set Model = mModel
End Property

Public Property Get SelectedObject() As IGraphicObject
Set SelectedObject = mSelectionManager.SelectedObject
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub CalculateAndRepaint()
Const ProcName As String = "CalculateAndRepaint"

On Error GoTo Err

mModel.DetermineInScopeObjects mGraphics.Boundary
calculateVerticalBounds
mModel.DetermineVisibleObjects mGraphics.Boundary
PaintAll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ChangeGaugeAndRepaint()
Const ProcName As String = "changeGaugeAndRepaint"

On Error GoTo Err

mModel.CalculateInscopeGaugeDependentObjects mGraphics.Boundary
PaintAll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"

On Error GoTo Err

If mIsCleared Then Exit Sub

mIsCleared = True

mModel.Clear
Set mModel = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearSelection()
Const ProcName As String = "ClearSelection"

On Error GoTo Err

mSelectionManager.ClearSelection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetScopeRect() As TRectangle
Const ProcName As String = "GetScopeRect"
On Error GoTo Err

GetScopeRect = mGraphics.Boundary

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Sub Finish()
mGraphics.Finish
Set mGraphics = Nothing
End Sub

Friend Sub Initialise( _
                ByVal pGraphics As Graphics, _
                ByVal pUseDeferredPainting As Boolean, _
                ByVal pDeferredPaintDelay As Long, _
                ByVal pStripWidth As Double, _
                ByVal pStripOrigin As Double)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mGraphics = pGraphics
mModel.Initialise Me, pStripWidth, pStripOrigin

If pUseDeferredPainting Then
    mUseDeferredPainting = pUseDeferredPainting
    Set mDeferredPainter = New DeferredPainter
    mDeferredPainter.Initialise Me, pDeferredPaintDelay
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub OnObjectChanged( _
                ByVal pGraphObj As IGraphicObject)
Const ProcName As String = "OnObjectChanged"

On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub
If mPaintingInProgress Then Exit Sub

If Not pGraphObj.BoundingRectangle.isValid Then Exit Sub
If pGraphObj.NoDraw Then Exit Sub

If needToRescale(pGraphObj) Then
    doScaleAndRepaint
Else
    rePaintChangedObject pGraphObj
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Paint( _
                ByVal recalculate As Boolean)
Const ProcName As String = "Paint"

On Error GoTo Err

If recalculate Then
    doCalculateAndRepaint
Else
    doRepaint
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PaintAll()
Const ProcName As String = "PaintAll"

On Error GoTo Err

mPaintingInProgress = True

mGraphics.PaintBackground

mGraphics.SaveClippingRegion
mGraphics.SetClippingRegion mGraphics.Boundary


paintObjects mModel.GetAllVisibleObjects

mGraphics.RestoreClippingRegion

mPaintingInProgress = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PaintObject( _
                ByVal pGraphObj As IGraphicObject, _
                ByRef pArea As TRectangle)
Const ProcName As String = "PaintObject"
On Error GoTo Err

mGraphics.SaveClippingRegion
mGraphics.IntersectClippingRegion pArea
mGraphics.IntersectClippingRegion pGraphObj.BoundingRectangle

pGraphObj.Draw pArea
mGraphics.RestoreClippingRegion

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function RepaintAreas( _
                ByRef pAreas() As TRectangle, _
                ByVal pCount As Long) As Long
                
Const ProcName As String = "RepaintAreas"
On Error GoTo Err

Dim i As Long

If mGraphics Is Nothing Then Exit Function

mGraphics.SaveClippingRegion
mGraphics.SetClippingRegion pAreas(0)
For i = 1 To pCount - 1
    mGraphics.AddToClippingRegion pAreas(i)
Next

mGraphics.PaintBackground

RepaintAreas = paintObjects(mModel.GetObjectsInAreas(pAreas, pCount, LayerNumbers.LayerMin))

mGraphics.RestoreClippingRegion

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ScaleAndRepaint()
Const ProcName As String = "ScaleAndRepaint"

On Error GoTo Err

calculateVerticalBounds
mModel.DetermineVisibleObjects mGraphics.Boundary
PaintAll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScaleUp(ByVal factor As Single)
Const ProcName As String = "ScaleUp"

On Error GoTo Err

With mGraphics
    SetVerticalScale .Bottom + factor * .Height / 2, _
                                .Top - factor * .Height / 2
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetVerticalScale(ByVal pLowValue As Double, _
                            ByVal pHighValue As Double)
Const ProcName As String = "SetVerticalScale"

On Error GoTo Err

Autoscaling = False

mGraphics.SetScales mGraphics.Left, pLowValue, mGraphics.Right, pHighValue

doCalculateAndRepaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub UndrawObject( _
                ByVal pGraphObj As IGraphicObject)
Const ProcName As String = "UndrawObject"

On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub
If mPaintingInProgress Then Exit Sub
If mGraphics Is Nothing Then Exit Sub

If pGraphObj.Handle < 0 Then Exit Sub ' check if object has been removed from the model

    If mUseDeferredPainting Then
        mDeferredPainter.AddAreaToRedraw pGraphObj.BoundingRectangleInArea(mGraphics.Boundary)
    Else
        undrawArea pGraphObj.BoundingRectangleInArea(mGraphics.Boundary)
    End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calculateVerticalBounds()
Dim scaleMax As Double
Dim scaleMin As Double

Const ProcName As String = "calculateVerticalBounds"

On Error GoTo Err

If Not Autoscaling Then Exit Sub

scaleMax = MinusInfinityDouble
scaleMin = PlusInfinityDouble

If mModel.CalcVerticalBoundsForNonGaugeDependentObjects( _
                                mGraphics.Boundary, _
                                scaleMin, _
                                scaleMax) _
Then
    ' there are gauge-dependent objects that may affect the max and min scales.
    ' We need to recalculate their bounds with the new gauge determined from
    ' the other objects to get a better estimate of their sizes.
    mGraphics.SetScales mGraphics.Left, scaleMin, mGraphics.Right, scaleMax
    mModel.CalcVerticalBoundsForGaugeDependentObjects _
                                mGraphics.Boundary, _
                                scaleMin, _
                                scaleMax
End If

If scaleMin = PlusInfinityDouble Or scaleMax = MinusInfinityDouble Then Exit Sub

mGraphics.SetScales mGraphics.Left, _
                    scaleMin - (scaleMax - scaleMin) * 0.025, _
                    mGraphics.Right, _
                    scaleMax + (scaleMax - scaleMin) * 0.025

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doCalculateAndRepaint()
Const ProcName As String = "doCalculateAndRepaint"

On Error GoTo Err

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        mDeferredPainter.CalculateAndRepaint
    Else
        CalculateAndRepaint
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doRepaint()
Const ProcName As String = "doRepaint"

On Error GoTo Err

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        mDeferredPainter.RepaintAll
    Else
        PaintAll
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doScaleAndRepaint()
Const ProcName As String = "doScaleAndRepaint"

On Error GoTo Err

If mIsDrawingEnabled Then
    If mUseDeferredPainting Then
        mDeferredPainter.ScaleAndRepaint
    Else
        ScaleAndRepaint
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function needToRescale( _
                ByVal pGraphObj As IGraphicObject) As Boolean
Const ProcName As String = "needToRescale"
On Error GoTo Err

If Autoscaling And pGraphObj.IncludeInAutoscale Then
    Dim verticalRangeInViewport As TInterval
    verticalRangeInViewport = pGraphObj.VerticalRangeInXinterval(TRectangleGetXInterval(mGraphics.Boundary))
    
    If verticalRangeInViewport.isValid And _
        (verticalRangeInViewport.endValue > mGraphics.Boundary.Top Or _
            verticalRangeInViewport.startValue < (mGraphics.Boundary.Bottom)) _
    Then
        needToRescale = True
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function paintObjects( _
                ByVal pEnumble As Enumerable) As Long
Const ProcName As String = "paintObjects"
On Error GoTo Err

Dim En As Enumerator
Dim lObj As IGraphicObject
Dim lCount As Long

Set En = pEnumble.Enumerator
Do While En.MoveNext
    lCount = lCount + 1
    Set lObj = mModel.GetGraphicObject(CLng(En.Current))
    PaintObject lObj, lObj.BoundingRectangleInArea(mGraphics.Boundary)
Loop

paintObjects = lCount
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub rePaintChangedObject( _
                ByVal pGraphObj As IGraphicObject)
Const ProcName As String = "rePaintChangedObject"

On Error GoTo Err

If mUseDeferredPainting Then
    mDeferredPainter.AddAreaToRedraw pGraphObj.BoundingRectangleInArea(mGraphics.Boundary)
Else
    rePaintObject pGraphObj
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub rePaintObject( _
                ByVal pGraphObj As IGraphicObject)
Const ProcName As String = "rePaintObject"
On Error GoTo Err

Dim lRect As TRectangle
lRect = pGraphObj.BoundingRectangleInArea(mGraphics.Boundary)

mGraphics.SaveClippingRegion
mGraphics.IntersectClippingRegion lRect
mGraphics.IntersectClippingRegion pGraphObj.BoundingRectangle

pGraphObj.Draw lRect

Dim lAreas(0) As TRectangle
lAreas(0) = lRect

paintObjects mModel.GetObjectsInAreas(lAreas, 1, pGraphObj.Layer + 1)

mGraphics.RestoreClippingRegion
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setAutoscaling()
Const ProcName As String = "setAutoscaling"

On Error GoTo Err

If Autoscaling Then doScaleAndRepaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub undrawArea( _
                ByRef pAreaToUndraw As TRectangle)
Const ProcName As String = "undrawArea"
On Error GoTo Err

If Not mIsDrawingEnabled Then Exit Sub
If Not pAreaToUndraw.isValid Then Exit Sub

mGraphics.SaveClippingRegion
mGraphics.IntersectClippingRegion pAreaToUndraw
mGraphics.PaintBackground

' redraw objects in the undrawn area, preserving layering
    
Dim lAreas(0) As TRectangle
lAreas(0) = pAreaToUndraw

paintObjects mModel.GetObjectsInAreas(lAreas, 1, LayerNumbers.LayerMin)

mGraphics.RestoreClippingRegion

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




