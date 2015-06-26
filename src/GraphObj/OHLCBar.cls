VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OHLCBar"
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
Private Const ModuleName                            As String = "OHLCBar"

'@================================================================================
' Member variables
'@================================================================================

Private mBase                                       As GraphicObjectBase
Attribute mBase.VB_VarHelpID = -1

Private mEPhost                                     As ExtendedPropertyHost

Private mHandle                                     As Long

Private mModel                                      As IOHLCBarModel

Private mStyle                                      As OHLCBarStyle

Private mPrevBarRef                                 As WeakReference

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mEPhost = New ExtendedPropertyHost
Set mBase = New GraphicObjectBase

' TODO: correct the related module name
GOHLCBar.gRegisterProperties
GOHLCBar.gRegisterExtendedEvents

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEventData)
' TODO: correct the change types enum name
Dim lChangeType As OHLCBarChangeTypes
lChangeType = ev.changeType
Select Case lChangeType
Case OHLCBarChangeTypes.OHLCBarCloseValueChanged
    recordUpdateFromModel GOHLCBar.gCloseValueProperty, mModel.CloseValue
Case OHLCBarChangeTypes.OHLCBarLowValueChanged
    recordUpdateFromModel GOHLCBar.gLowValueProperty, mModel.LowValue
Case OHLCBarChangeTypes.OHLCBarHighValueChanged
    recordUpdateFromModel GOHLCBar.gHighValueProperty, mModel.HighValue
Case OHLCBarChangeTypes.OHLCBarOpenValueChanged
    recordUpdateFromModel GOHLCBar.gOpenValueProperty, mModel.OpenValue
Case OHLCBarChangeTypes.OHLCBarXChanged
    recordUpdateFromModel GOHLCBar.gXProperty, mModel.X
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
If Value = mEPhost.GetLocalValue(GOHLCBar.gIsSelectedProperty) Then
Else
    mEPhost.SetValue GOHLCBar.gIsSelectedProperty, Value
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
GOHLCBar.gMouseEnterEvent.fire Me, Empty
End Sub

Private Sub IGraphicObject_MouseLeave(ByVal pButton As Integer, ByVal pShift As Integer, ByVal pX As Single, ByVal pY As Single)
' TODO: consider if any different action is needed
GOHLCBar.gMouseLeaveEvent.fire Me, Empty
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
mEPhost.SetValue GOHLCBar.gIsVisibleProperty, Value
End Property

'@================================================================================
' Properties
'@================================================================================

Public Property Let Brush(ByVal Value As IBrush)
Const ProcName As String = "Brush"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gBrushProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Brush() As IBrush
Const ProcName As String = "Brush"
On Error GoTo Err

Set Brush = mEPhost.GetValue(GOHLCBar.gBrushProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let CloseValue(ByVal Value As Double)
Const ProcName As String = "CloseValue"
On Error GoTo Err

mModel.CloseValue = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get CloseValue() As Double
Const ProcName As String = "CloseValue"
On Error GoTo Err

CloseValue = mEPhost.GetValue(GOHLCBar.gCloseValueProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DisplayMode(ByVal Value As OHLCBarDisplayModes)
Const ProcName As String = "DisplayMode"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gDisplayModeProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DisplayMode() As OHLCBarDisplayModes
Const ProcName As String = "DisplayMode"
On Error GoTo Err

DisplayMode = mEPhost.GetValue(GOHLCBar.gDisplayModeProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DownBrush(ByVal Value As IBrush)
Const ProcName As String = "DownBrush"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gDownBrushProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DownBrush() As IBrush
Const ProcName As String = "DownBrush"
On Error GoTo Err

Set DownBrush = mEPhost.GetValue(GOHLCBar.gDownBrushProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DownPen(ByVal Value As Pen)
Const ProcName As String = "DownPen"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gDownPenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DownPen() As Pen
Const ProcName As String = "DownPen"
On Error GoTo Err

Set DownPen = mEPhost.GetValue(GOHLCBar.gDownPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HighValue(ByVal Value As Double)
Const ProcName As String = "HighValue"
On Error GoTo Err

mModel.HighValue = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HighValue() As Double
Const ProcName As String = "HighValue"
On Error GoTo Err

HighValue = mEPhost.GetValue(GOHLCBar.gHighValueProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gIncludeInAutoscaleProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

IncludeInAutoscale = mEPhost.GetValue(GOHLCBar.gIncludeInAutoscaleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
Const ProcName As String = "IsSelectable"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gIsSelectableProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsSelectable() As Boolean
Const ProcName As String = "IsSelectable"
On Error GoTo Err

IsSelectable = mEPhost.GetValue(GOHLCBar.gIsSelectableProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsSelected() As Boolean
Const ProcName As String = "IsSelected"
On Error GoTo Err

IsSelected = mEPhost.GetValue(GOHLCBar.gIsSelectedProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsVisible() As Boolean
Const ProcName As String = "IsVisible"
On Error GoTo Err

IsVisible = mEPhost.GetValue(GOHLCBar.gIsVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gLayerProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mEPhost.GetValue(GOHLCBar.gLayerProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let LowValue(ByVal Value As Double)
Const ProcName As String = "LowValue"
On Error GoTo Err

mModel.LowValue = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LowValue() As Double
Const ProcName As String = "LowValue"
On Error GoTo Err

LowValue = mEPhost.GetValue(GOHLCBar.gLowValueProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MouseEnterEvent() As ExtendedEvent
Set MouseEnterEvent = GOHLCBar.gMouseEnterEvent
End Property

Public Property Get MouseLeaveEvent() As ExtendedEvent
Set MouseLeaveEvent = GOHLCBar.gMouseLeaveEvent
End Property

Public Property Let OpenValue(ByVal Value As Double)
Const ProcName As String = "OpenValue"
On Error GoTo Err

mModel.OpenValue = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get OpenValue() As Double
Const ProcName As String = "OpenValue"
On Error GoTo Err

OpenValue = mEPhost.GetValue(GOHLCBar.gOpenValueProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Orientation(ByVal Value As Double)
Const ProcName As String = "Orientation"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gOrientationProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Orientation() As Double
Const ProcName As String = "Layer"
On Error GoTo Err

Orientation = mEPhost.GetValue(GOHLCBar.gOrientationProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Pen(ByVal Value As Pen)
Const ProcName As String = "Pen"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gPenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Pen() As Pen
Const ProcName As String = "Pen"
On Error GoTo Err

Set Pen = mEPhost.GetValue(GOHLCBar.gPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get PreviousBar() As OHLCBar
Const ProcName As String = "PreviousBar"

On Error GoTo Err

If mPrevBarRef Is Nothing Then Exit Property
Set PreviousBar = prevBarObj

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Let PreviousBar(ByVal Value As OHLCBar)
Const ProcName As String = "PreviousBar"

On Error GoTo Err

Set mPrevBarRef = CreateWeakReference(Value)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

' TODO: correct the style type name
Public Property Let Style(ByVal Value As OHLCBarStyle)
Const ProcName As String = "Style"
On Error GoTo Err

Set mStyle = Value
mEPhost.Style = mStyle.ExtendedPropertyHost

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

' TODO: correct the style type name
Public Property Get Style() As OHLCBarStyle
Set Style = mStyle
End Property

Public Property Let UpBrush(ByVal Value As IBrush)
Const ProcName As String = "UpBrush"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gUpBrushProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get UpBrush() As IBrush
Const ProcName As String = "UpBrush"
On Error GoTo Err

Set UpBrush = mEPhost.GetValue(GOHLCBar.gUpBrushProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let UpPen(ByVal Value As Pen)
Const ProcName As String = "UpPen"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gUpPenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get UpPen() As Pen
Const ProcName As String = "UpPen"
On Error GoTo Err

Set UpPen = mEPhost.GetValue(GOHLCBar.gUpPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Width(ByVal Value As Double)
Const ProcName As String = "Width"
On Error GoTo Err

mEPhost.SetValue GOHLCBar.gWidthProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Width() As Double
Const ProcName As String = "Width"
On Error GoTo Err

Width = mEPhost.GetValue(GOHLCBar.gWidthProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let X(ByVal Value As Double)
Const ProcName As String = "X"
On Error GoTo Err

mModel.X = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get X() As Double
Const ProcName As String = "X"
On Error GoTo Err

X = mEPhost.GetValue(GOHLCBar.gXProperty)

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

mEPhost.ClearValue GOHLCBar.gBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDisplayMode()
Const ProcName As String = "ClearDisplayMode"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gDisplayModeProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDownBrush()
Const ProcName As String = "ClearDownBrush"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gDownBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDownPen()
Const ProcName As String = "ClearDownPen"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gDownPenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearIncludeInAutoscale()
Const ProcName As String = "ClearIncludeInAutoscale"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gIncludeInAutoscaleProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearIsSelectable()
Const ProcName As String = "ClearIsSelectable"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gIsSelectableProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearLayer()
Const ProcName As String = "ClearLayer"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gLayerProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearOrientation()
Const ProcName As String = "ClearOrientation"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gOrientationProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearPen()
Const ProcName As String = "ClearPen"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gPenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearUpBrush()
Const ProcName As String = "ClearUpBrush"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gUpBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearUpPen()
Const ProcName As String = "ClearUpPen"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gUpPenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearWidth()
Const ProcName As String = "ClearWidth"
On Error GoTo Err

mEPhost.ClearValue GOHLCBar.gWidthProperty

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
                ByVal pSeries As OHLCBarSeries, _
                ByVal pModel As IOHLCBarModel)
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
If OpenValue = MaxDouble Then
ElseIf HighValue = MaxDouble Then
ElseIf LowValue = MaxDouble Then
ElseIf CloseValue = MaxDouble Then
Else
    canBeDrawn = True
End If
End Function

Private Function prevBarObj() As OHLCBar
Const ProcName As String = "prevBarObj"

On Error GoTo Err

Set prevBarObj = mPrevBarRef.Target

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Private Sub recordClearFromModel(ByVal pExtProp As ExtendedProperty)
Const ProcName As String = "recordUpdateFromModel"
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



