VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPoint"
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

Implements ChangeListener
Implements IExtendedEventsSource
Implements IGraphicObject

'@================================================================================
' Events
'@================================================================================

Event Click()

Event DblCLick()

Event SelectionStateChanged()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

' TODO: set the ModuleName constant to the class name
Private Const ModuleName                            As String = "DataPoint"

'@================================================================================
' Member variables
'@================================================================================

Private mBase                                       As GraphicObjectBase
Attribute mBase.VB_VarHelpID = -1

Private mEPhost                                     As ExtendedPropertyHost

Private mHandle                                     As Long

' TODO: correct the model interface type
Private mModel                                      As IDataPointModel

' TODO: correct the style type
Private mStyle                                      As DataPointStyle

Private mPrevDataPointRef                           As WeakReference

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mEPhost = New ExtendedPropertyHost
Set mBase = New GraphicObjectBase

' TODO: correct the related module name
GDataPoint.gRegisterProperties
GDataPoint.gRegisterExtendedEvents

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEventData)
' TODO: correct the change types enum name
Dim lChangeType As DataPointChangeTypes
lChangeType = ev.changeType
Select Case lChangeType
Case DataPointChangeTypes.DataPointValueChanged
    recordUpdateFromModel GDataPoint.gValueProperty, mModel.Value
Case DataPointChangeTypes.DataPointOrientationChanged
    recordUpdateFromModel GDataPoint.gOrientationProperty, mModel.Orientation
Case DataPointChangeTypes.DataPointOrientationCleared
    recordClearFromModel GDataPoint.gOrientationProperty
Case DataPointChangeTypes.DataPointSizeChanged
    recordUpdateFromModel GDataPoint.gSizeProperty, mModel.Size
Case DataPointChangeTypes.DataPointSizeCleared
    recordClearFromModel GDataPoint.gSizeProperty
End Select
End Sub

'@================================================================================
' IExtendedEventsSource Interface Members
'@================================================================================

Private Property Get IExtendedEventsSource_Parent() As Object
Set IExtendedEventsSource_Parent = mBase.Series
End Property

'@================================================================================
' IGraphicObject Interface Members
'@================================================================================

Private Property Get IGraphicObject_BoundingRectangle() As TRectangle
IGraphicObject_BoundingRectangle = mBase.BoundingRectangle
End Property

Private Property Get IGraphicObject_BoundingRectangleInArea(pArea As TRectangle) As TRectangle
' TODO: consider the validity of the following statement and correct if necessary. For example,
' the bounding rectangle for a line has the line's start point at one corner and its end point
' at the diagonally opposite corner. However, the bounding rectangle within a specific area
' is based on the points of intersection of the line with the area's boundary, which is not
' the same thing.
IGraphicObject_BoundingRectangleInArea = TRectangleIntersection(IGraphicObject_BoundingRectangle, pArea)
End Property

Private Sub IGraphicObject_Click()
RaiseEvent Click
End Sub

Private Sub IGraphicObject_DblCLick()
RaiseEvent DblCLick
End Sub

Private Sub IGraphicObject_Draw(pArea As TRectangle)
mBase.Draw pArea
End Sub

Private Sub IGraphicObject_Finish()
Finish
End Sub

Private Property Get IGraphicObject_GaugeDependent() As Boolean
' TODO: consider whether the object is gauge-dependent
End Property

Private Property Let IGraphicObject_Handle(ByVal Value As Long)
mHandle = Value
End Property

Private Property Get IGraphicObject_Handle() As Long
IGraphicObject_Handle = mHandle
End Property

Private Function IGraphicObject_HitTest(ByVal X As Double, ByVal Y As Double) As Boolean
' TODO: consider the validity of the following statement and correct if necessary
IGraphicObject_HitTest = TRectangleContainsPoint(mBase.BoundingRectangle, X, Y)
End Function

Private Property Get IGraphicObject_IncludeInAutoscale() As Boolean
IGraphicObject_IncludeInAutoscale = IncludeInAutoscale
End Property

Private Property Get IGraphicObject_IsSelectable() As Boolean
IGraphicObject_IsSelectable = IsSelectable
End Property

Private Property Let IGraphicObject_IsSelected(ByVal Value As Boolean)
If Value = mEPhost.GetLocalValue(GDataPoint.gIsSelectedProperty) Then
Else
    mEPhost.SetValue GDataPoint.gIsSelectedProperty, Value
    RaiseEvent SelectionStateChanged
End If
End Property

Private Property Get IGraphicObject_IsSelected() As Boolean
IGraphicObject_IsSelected = IsSelected
End Property

Private Property Get IGraphicObject_Layer() As LayerNumbers
IGraphicObject_Layer = Layer
End Property

Private Sub IGraphicObject_MouseEnter(ByVal pButton As Integer, ByVal pShift As Integer, ByVal pX As Single, ByVal pY As Single)
' TODO: consider if any different action is needed
GDataPoint.gMouseEnterEvent.fire Me, Empty
End Sub

Private Sub IGraphicObject_MouseLeave(ByVal pButton As Integer, ByVal pShift As Integer, ByVal pX As Single, ByVal pY As Single)
' TODO: consider if any different action is needed
GDataPoint.gMouseLeaveEvent.fire Me, Empty
End Sub

Private Property Get IGraphicObject_NoDraw() As Boolean
IGraphicObject_NoDraw = (Not canBeDrawn)
End Property

Private Property Get IGraphicObject_SeriesID() As Long
IGraphicObject_SeriesID = mBase.Series.Id
End Property

Private Property Get IGraphicObject_VerticalRangeInXinterval(pXInterval As TInterval) As TInterval
' TODO: consider the validity of the following statement and correct if necessary
If TIntervalOverlaps(pXInterval, TRectangleGetXInterval(mBase.BoundingRectangle)) Then
    IGraphicObject_VerticalRangeInXinterval = TRectangleGetYInterval(mBase.BoundingRectangle)
End If
End Property

Private Property Let IGraphicObject_Visible(ByVal Value As Boolean)
mEPhost.SetValue GDataPoint.gIsVisibleProperty, Value
End Property

'@================================================================================
' Properties
'@================================================================================

Public Property Let Brush(ByVal Value As IBrush)
Const ProcName As String = "Brush"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gBrushProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Brush() As IBrush
Const ProcName As String = "Brush"
On Error GoTo Err

Set Brush = mEPhost.GetValue(GDataPoint.gBrushProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DisplayMode(ByRef Value As DataPointDisplayModes)
Const ProcName As String = "DisplayMode"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gDisplayModeProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
Const ProcName As String = "DisplayMode"
On Error GoTo Err

DisplayMode = mEPhost.GetValue(GDataPoint.gDisplayModeProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DownBrush(ByVal Value As IBrush)
Const ProcName As String = "DownBrush"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gDownBrushProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DownBrush() As IBrush
Const ProcName As String = "DownBrush"
On Error GoTo Err

Set DownBrush = mEPhost.GetValue(GDataPoint.gDownBrushProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DownLinePen(ByVal Value As Pen)
Const ProcName As String = "DownLinePen"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gDownLinePenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DownLinePen() As Pen
Const ProcName As String = "DownLinePen"
On Error GoTo Err

Set DownLinePen = mEPhost.GetValue(GDataPoint.gDownLinePenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DownPen(ByVal Value As Pen)
Const ProcName As String = "DownPen"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gDownPenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DownPen() As Pen
Const ProcName As String = "DownPen"
On Error GoTo Err

Set DownPen = mEPhost.GetValue(GDataPoint.gDownPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HistogramBarWidth(ByVal Value As Double)
Const ProcName As String = "HistogramBarWidth"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gHistogramBarWidthProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HistogramBarWidth() As Double
Const ProcName As String = "HistogramBarWidth"
On Error GoTo Err

HistogramBarWidth = mEPhost.GetValue(GDataPoint.gHistogramBarWidthProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HistogramBaselinePen(ByVal Value As Pen)
Const ProcName As String = "HistogramBaselinePen"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gHistogramBaselinePenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HistogramBaselinePen() As Pen
Const ProcName As String = "HistogramBaselinePen"
On Error GoTo Err

Set HistogramBaselinePen = mEPhost.GetValue(GDataPoint.gHistogramBaselinePenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HistogramBaseValue() As Double
Const ProcName As String = "HistogramBaseValue"
On Error GoTo Err

HistogramBaseValue = mEPhost.GetValue(GDataPoint.gHistogramBaseValueProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gIncludeInAutoscaleProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

IncludeInAutoscale = mEPhost.GetValue(GDataPoint.gIncludeInAutoscaleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
Const ProcName As String = "IsSelectable"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gIsSelectableProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsSelectable() As Boolean
Const ProcName As String = "IsSelectable"
On Error GoTo Err

IsSelectable = mEPhost.GetValue(GDataPoint.gIsSelectableProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsSelected() As Boolean
Const ProcName As String = "IsSelected"
On Error GoTo Err

IsSelected = mEPhost.GetValue(GDataPoint.gIsSelectedProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsVisible() As Boolean
Const ProcName As String = "IsVisible"
On Error GoTo Err

IsVisible = mEPhost.GetValue(GDataPoint.gIsVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gLayerProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mEPhost.GetValue(GDataPoint.gLayerProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LineMode() As DataPointLineModes
Const ProcName As String = "LineMode"
On Error GoTo Err

LineMode = mEPhost.GetValue(GDataPoint.gLineModeProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let LinePen(ByVal Value As Pen)
Const ProcName As String = "LinePen"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gLinePenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LinePen() As Pen
Const ProcName As String = "LinePen"
On Error GoTo Err

Set LinePen = mEPhost.GetValue(GDataPoint.gLinePenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MouseEnterEvent() As ExtendedEvent
Set MouseEnterEvent = GDataPoint.gMouseEnterEvent
End Property

Public Property Get MouseLeaveEvent() As ExtendedEvent
Set MouseLeaveEvent = GDataPoint.gMouseLeaveEvent
End Property

Public Property Let NumberOfSides(ByVal Value As Long)
Const ProcName As String = "NumberOfSides"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gNumberOfSidesProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get NumberOfSides() As Long
Const ProcName As String = "NumberOfSides"
On Error GoTo Err

NumberOfSides = mEPhost.GetValue(GDataPoint.gNumberOfSidesProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Orientation(ByVal Value As Double)
Const ProcName As String = "Orientation"
On Error GoTo Err

mModel.Orientation = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Orientation() As Double
Const ProcName As String = "Layer"
On Error GoTo Err

Orientation = mEPhost.GetValue(GDataPoint.gOrientationProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Pen(ByVal Value As Pen)
Const ProcName As String = "Pen"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gPenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Pen() As Pen
Const ProcName As String = "Pen"
On Error GoTo Err

Set Pen = mEPhost.GetValue(GDataPoint.gPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PreviousDataPoint() As DataPoint
Const ProcName As String = "PreviousDataPoint"

On Error GoTo Err

If mPrevDataPointRef Is Nothing Then Exit Property
Set PreviousDataPoint = prevDataPointObj

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let PreviousDataPoint(ByVal Value As DataPoint)
Const ProcName As String = "PreviousDataPoint"

On Error GoTo Err

Set mPrevDataPointRef = CreateWeakReference(Value)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Size(ByVal Value As Size)
Const ProcName As String = "Size"
On Error GoTo Err

mModel.Size = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Size() As Size
Const ProcName As String = "Size"
On Error GoTo Err

Set Size = mEPhost.GetValue(GDataPoint.gSizeProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

' TODO: correct the style type name
Public Property Let Style(ByVal Value As DataPointStyle)
Const ProcName As String = "Style"
On Error GoTo Err

Set mStyle = Value
mEPhost.Style = mStyle.ExtendedPropertyHost

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

' TODO: correct the style type name
Public Property Get Style() As DataPointStyle
Set Style = mStyle
End Property

Public Property Let UpBrush(ByVal Value As IBrush)
Const ProcName As String = "UpBrush"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gUpBrushProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get UpBrush() As IBrush
Const ProcName As String = "UpBrush"
On Error GoTo Err

Set UpBrush = mEPhost.GetValue(GDataPoint.gUpBrushProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let UpLinePen(ByVal Value As Pen)
Const ProcName As String = "UpLinePen"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gUpLinePenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get UpLinePen() As Pen
Const ProcName As String = "UpLinePen"
On Error GoTo Err

Set UpLinePen = mEPhost.GetValue(GDataPoint.gUpLinePenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let UpPen(ByVal Value As Pen)
Const ProcName As String = "UpPen"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gUpPenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get UpPen() As Pen
Const ProcName As String = "UpPen"
On Error GoTo Err

Set UpPen = mEPhost.GetValue(GDataPoint.gUpPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Value(ByVal Value As Double)
Const ProcName As String = "Value"
On Error GoTo Err

mModel.Value = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Value() As Double
Const ProcName As String = "Value"
On Error GoTo Err

Value = mEPhost.GetValue(GDataPoint.gValueProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let X(ByVal Value As Double)
Const ProcName As String = "X"
On Error GoTo Err

mEPhost.SetValue GDataPoint.gXProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get X() As Double
Const ProcName As String = "X"
On Error GoTo Err

X = mEPhost.GetValue(GDataPoint.gXProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ApplyChanges()
mBase.ApplyChanges
End Sub

Public Sub ClearBrush()
Const ProcName As String = "ClearBrush"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDisplayMode()
Const ProcName As String = "ClearDisplayMode"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gDisplayModeProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDownBrush()
Const ProcName As String = "ClearDownBrush"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gDownBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDownLinePen()
Const ProcName As String = "ClearDownLinePen"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gDownLinePenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDownPen()
Const ProcName As String = "ClearDownPen"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gDownPenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearHistogramBarWidth()
Const ProcName As String = "ClearHistogramBarWidth"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gHistogramBarWidthProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearHistogramBaselinePen()
Const ProcName As String = "ClearHistogramBaselinePen"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gHistogramBaselinePenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearHistogramBaseValue()
Const ProcName As String = "ClearHistogramBaseValue"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gHistogramBaseValueProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearIncludeInAutoscale()
Const ProcName As String = "ClearIncludeInAutoscale"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gIncludeInAutoscaleProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearIsSelectable()
Const ProcName As String = "ClearIsSelectable"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gIsSelectableProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearLayer()
Const ProcName As String = "ClearLayer"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gLayerProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearLineMode()
Const ProcName As String = "ClearLineMode"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gLineModeProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearLinePen()
Const ProcName As String = "ClearLinePen"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gLinePenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearNumberOfSides()
Const ProcName As String = "ClearNumberOfSides"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gNumberOfSidesProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearOrientation()
Const ProcName As String = "ClearOrientation"
On Error GoTo Err

mModel.ClearOrientation

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearPen()
Const ProcName As String = "ClearPen"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gPenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearPosition()
Const ProcName As String = "ClearPosition"
On Error GoTo Err

mModel.ClearPosition

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearSize()
Const ProcName As String = "ClearSize"
On Error GoTo Err

mModel.ClearSize

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearUpBrush()
Const ProcName As String = "ClearUpBrush"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gUpBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearUpLinePen()
Const ProcName As String = "ClearUpLinePen"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gUpLinePenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearUpPen()
Const ProcName As String = "ClearUpPen"
On Error GoTo Err

mEPhost.ClearValue GDataPoint.gUpPenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Edit()
mBase.Edit
End Sub

Friend Sub Finish()
mModel.RemoveChangeListener Me
End Sub

' TODO: correct the type names in the method signature
Friend Sub Initialise( _
                ByVal pSeries As DataPointSeries, _
                ByVal pModel As IDataPointModel)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

mBase.Initialise Me, pSeries, mEPhost

Set mModel = pModel
mModel.AddChangeListener Me

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function canBeDrawn() As Boolean
canBeDrawn = (Value <> MaxDouble)
End Function

Private Function prevDataPointObj() As DataPoint
Const ProcName As String = "prevDataPointObj"

On Error GoTo Err

Set prevDataPointObj = mPrevDataPointRef.Target

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub recordClearFromModel(ByVal pExtProp As ExtendedProperty)
Const ProcName As String = "recordClearFromModel"
On Error GoTo Err

mEPhost.ClearValue pExtProp

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub recordUpdateFromModel(ByVal pExtProp As ExtendedProperty, ByVal pValue As Variant)
Const ProcName As String = "recordUpdateFromModel"
On Error GoTo Err

mEPhost.SetValue pExtProp, pValue

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub



