VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Rectangle"
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
Implements ISprite

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

Private Const ModuleName                            As String = "Rectangle"

'@================================================================================
' Member variables
'@================================================================================

Private mBase                                       As GraphicObjectBase
Attribute mBase.VB_VarHelpID = -1

Private mEPhost                                     As ExtendedPropertyHost

Private mHandle                                     As Long

Private mModel                                      As IRectangleModel

Private mStyle                                      As RectangleStyle

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mEPhost = New ExtendedPropertyHost
Set mBase = New GraphicObjectBase

GRectangle.gRegisterProperties

GRectangle.gRegisterExtendedEvents

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub Class_Terminate()
Debug.Print "Rectangle terminated"
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub ChangeListener_Change(ev As ChangeEventData)
Dim lChangeType As RectangleChangeTypes
lChangeType = ev.ChangeType
Select Case lChangeType
Case RectangleChangeTypes.RectangleOrientationChanged
    recordUpdateFromModel GRectangle.gOrientationProperty, mModel.Orientation
Case RectangleChangeTypes.RectangleOrientationCleared
    recordClearFromModel GRectangle.gOrientationProperty
Case RectangleChangeTypes.RectanglePositionChanged
    recordUpdateFromModel GRectangle.gPositionProperty, mModel.Position
Case RectangleChangeTypes.RectanglePositionCleared
    recordClearFromModel GRectangle.gPositionProperty
Case RectangleChangeTypes.RectangleSizeChanged
    recordUpdateFromModel GRectangle.gSizeProperty, mModel.Size
Case RectangleChangeTypes.RectangleSizeCleared
    recordClearFromModel GRectangle.gSizeProperty
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

Private Property Get IGraphicObject_ExtendedObject() As Boolean

End Property

Private Sub IGraphicObject_Finish()
Finish
End Sub

Private Property Get IGraphicObject_GaugeDependent() As Boolean

End Property

Private Property Let IGraphicObject_Handle(ByVal Value As Long)
mHandle = Value
End Property

Private Property Get IGraphicObject_Handle() As Long
IGraphicObject_Handle = mHandle
End Property

Private Function IGraphicObject_HitTest(ByVal X As Double, ByVal Y As Double) As Boolean
IGraphicObject_HitTest = TRectangleContainsPoint(mBase.BoundingRectangle, X, Y)
End Function

Private Property Get IGraphicObject_IncludeInAutoscale() As Boolean
IGraphicObject_IncludeInAutoscale = IncludeInAutoscale
End Property

Private Property Get IGraphicObject_IsSelectable() As Boolean
IGraphicObject_IsSelectable = IsSelectable
End Property

Private Property Let IGraphicObject_IsSelected(ByVal Value As Boolean)
If Value = mEPhost.GetLocalValue(GRectangle.gIsSelectedProperty) Then
Else
    mEPhost.SetValue GRectangle.gIsSelectedProperty, Value
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
GRectangle.gMouseEnterEvent.fire Me, Empty
End Sub

Private Sub IGraphicObject_MouseLeave(ByVal pButton As Integer, ByVal pShift As Integer, ByVal pX As Single, ByVal pY As Single)
GRectangle.gMouseLeaveEvent.fire Me, Empty
End Sub

Private Property Get IGraphicObject_NoDraw() As Boolean
IGraphicObject_NoDraw = (Not canBeDrawn)
End Property

Private Property Get IGraphicObject_SeriesID() As Long
IGraphicObject_SeriesID = mBase.Series.Id
End Property

Private Property Get IGraphicObject_VerticalRangeInXinterval(pXInterval As TInterval) As TInterval
If TIntervalOverlaps(pXInterval, TRectangleGetXInterval(mBase.BoundingRectangle)) Then
    IGraphicObject_VerticalRangeInXinterval = TRectangleGetYInterval(mBase.BoundingRectangle)
End If
End Property

Private Property Let IGraphicObject_Visible(ByVal Value As Boolean)
mEPhost.SetValue GRectangle.gIsVisibleProperty, Value
End Property

'@================================================================================
' ISprite Interface Members
'@================================================================================

Private Sub ISprite_ApplyChanges()
ApplyChanges
End Sub

Private Property Get ISprite_BoundingRectangle() As TRectangle
ISprite_BoundingRectangle = mBase.BoundingRectangle
End Property

Private Property Let ISprite_Brush(ByVal RHS As IBrush)
Brush = RHS
End Property

Private Property Get ISprite_Brush() As IBrush
Set ISprite_Brush = Brush
End Property

Private Sub ISprite_Edit()
Edit
End Sub

Private Property Let ISprite_IncludeInAutoscale(ByVal RHS As Boolean)
IncludeInAutoscale = RHS
End Property

Private Property Get ISprite_IncludeInAutoscale() As Boolean
ISprite_IncludeInAutoscale = IncludeInAutoscale
End Property

Private Property Let ISprite_IsSelectable(ByVal RHS As Boolean)
IsSelectable = RHS
End Property

Private Property Get ISprite_IsSelectable() As Boolean
ISprite_IsSelectable = IsSelectable
End Property

Private Property Get ISprite_IsSelected() As Boolean
ISprite_IsSelected = IsSelected
End Property

Private Property Get ISprite_IsVisible() As Boolean
ISprite_IsVisible = IsVisible
End Property

Private Property Let ISprite_Layer(ByVal RHS As LayerNumbers)
Layer = RHS
End Property

Private Property Get ISprite_Layer() As LayerNumbers
ISprite_Layer = Layer
End Property

Private Property Let ISprite_Orientation(ByVal RHS As Double)
Orientation = RHS
End Property

Private Property Get ISprite_Orientation() As Double
ISprite_Orientation = Orientation
End Property

Private Property Let ISprite_Pen(ByVal RHS As Pen)
Pen = RHS
End Property

Private Property Get ISprite_Pen() As Pen
Set ISprite_Pen = Pen
End Property

Private Property Let ISprite_Position(ByVal RHS As Point)
Position = RHS
End Property

Private Property Get ISprite_Position() As Point
Set ISprite_Position = Position
End Property

Private Property Let ISprite_Size(ByVal RHS As Size)
Size = RHS
End Property

Private Property Get ISprite_Size() As Size
Set ISprite_Size = Size
End Property

'@================================================================================
' Properties
'@================================================================================

Public Property Let Brush(ByVal Value As IBrush)
Const ProcName As String = "Brush"
On Error GoTo Err

mEPhost.SetValue GRectangle.gBrushProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Brush() As IBrush
Const ProcName As String = "Brush"
On Error GoTo Err

Set Brush = mEPhost.GetValue(GRectangle.gBrushProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

mEPhost.SetValue GRectangle.gIncludeInAutoscaleProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

IncludeInAutoscale = mEPhost.GetValue(GRectangle.gIncludeInAutoscaleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
Const ProcName As String = "IsSelectable"
On Error GoTo Err

mEPhost.SetValue GRectangle.gIsSelectableProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsSelectable() As Boolean
Const ProcName As String = "IsSelectable"
On Error GoTo Err

IsSelectable = mEPhost.GetValue(GRectangle.gIsSelectableProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsSelected() As Boolean
Const ProcName As String = "IsSelected"
On Error GoTo Err

IsSelected = mEPhost.GetValue(GRectangle.gIsSelectedProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsVisible() As Boolean
Const ProcName As String = "IsVisible"
On Error GoTo Err

IsVisible = mEPhost.GetValue(GRectangle.gIsVisibleProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
On Error GoTo Err

mEPhost.SetValue GRectangle.gLayerProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mEPhost.GetValue(GRectangle.gLayerProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get MouseEnterEvent() As ExtendedEvent
Set MouseEnterEvent = GRectangle.gMouseEnterEvent
End Property

Public Property Get MouseLeaveEvent() As ExtendedEvent
Set MouseLeaveEvent = GRectangle.gMouseLeaveEvent
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

Orientation = mEPhost.GetValue(GRectangle.gOrientationProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Pen(ByVal Value As Pen)
Const ProcName As String = "Pen"
On Error GoTo Err

mEPhost.SetValue GRectangle.gPenProperty, Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Pen() As Pen
Const ProcName As String = "Pen"
On Error GoTo Err

Set Pen = mEPhost.GetValue(GRectangle.gPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Position(ByVal Value As Point)
Const ProcName As String = "Position"
On Error GoTo Err

mModel.Position = Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Position() As Point
Const ProcName As String = "Position"
On Error GoTo Err

Set Position = mEPhost.GetValue(GRectangle.gPositionProperty)

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

Set Size = mEPhost.GetValue(GRectangle.gSizeProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Style(ByVal Value As RectangleStyle)
Const ProcName As String = "Style"
On Error GoTo Err

Set mStyle = Value
mEPhost.Style = mStyle.ExtendedPropertyHost

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Style() As RectangleStyle
Set Style = mStyle
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

mEPhost.ClearValue GRectangle.gBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearIncludeInAutoscale()
Const ProcName As String = "ClearIncludeInAutoscale"
On Error GoTo Err

mEPhost.ClearValue GRectangle.gIncludeInAutoscaleProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearIsSelectable()
Const ProcName As String = "ClearIsSelectable"
On Error GoTo Err

mEPhost.ClearValue GRectangle.gIsSelectableProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearLayer()
Const ProcName As String = "ClearLayer"
On Error GoTo Err

mEPhost.ClearValue GRectangle.gLayerProperty

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

mEPhost.ClearValue GRectangle.gPenProperty

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

Public Sub Edit()
mBase.Edit
End Sub

Friend Sub Finish()
mModel.RemoveChangeListener Me
End Sub

Friend Sub Initialise( _
                ByVal pSeries As RectangleSeries, _
                ByVal pModel As IRectangleModel)
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
canBeDrawn = Not (Position Is Nothing Or Size Is Nothing)
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

