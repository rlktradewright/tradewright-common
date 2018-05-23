VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPointSeries"
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

Implements IExtendedEventsSource
Implements IGraphicObjectSeries

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

' TODO: replace DataPoint with the name of the related graphic object class
Private Const ModuleName                            As String = "DataPointSeries"

'@================================================================================
' Member variables
'@================================================================================

Private mController                                 As Controller
Private mModel                                      As LayeredGraphicsModel

Private WithEvents mGraphics                        As Graphics
Attribute mGraphics.VB_VarHelpID = -1
Private mGraphicsUpdateNumber                       As Long
Private mLastGraphicsChange                         As GraphicsChangeTypes

Private WithEvents mBase                            As GraphicObjectSeriesBase
Attribute mBase.VB_VarHelpID = -1

Private mName                                       As String
Private mSeriesID                                   As Long
Private mSeriesUpdateNumber                         As Long

Private mFinished                                   As Boolean

Private mRenderer                                   As IRenderer

Private mGraphicObjects                             As Collection

Private mStyle                                      As DataPointStyle

Private mPrevDataPointRef                           As WeakReference

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mBase = New GraphicObjectSeriesBase

' TODO: correct the module name
GDataPoint.gRegisterProperties
GDataPoint.gRegisterExtendedEvents

Set mGraphicObjects = New Collection

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' IExtendedEventsSource Interface Members
'@================================================================================

Private Property Get IExtendedEventsSource_Parent() As Object
Set IExtendedEventsSource_Parent = Nothing
End Property

'@================================================================================
' IGraphicObjectSeries Interface Members
'@================================================================================

Private Sub IGraphicObjectSeries_Click()
RaiseEvent Click
End Sub

Private Property Get IGraphicObjectSeries_Controller() As Controller
Set IGraphicObjectSeries_Controller = mController
End Property

Private Property Get IGraphicObjectSeries_Count() As Long
IGraphicObjectSeries_Count = mGraphicObjects.Count
End Property

Private Sub IGraphicObjectSeries_DblCLick()
RaiseEvent DblCLick
End Sub

Private Property Get IGraphicObjectSeries_ExtendedPropertyHost() As ExtendedPropertyHost
Set IGraphicObjectSeries_ExtendedPropertyHost = mBase.ExtendedPropertyHost
End Property

Private Sub IGraphicObjectSeries_Finish()
mFinished = True
End Sub

Private Property Get IGraphicObjectSeries_GraphicsUpdateNumber() As Long
IGraphicObjectSeries_GraphicsUpdateNumber = mGraphicsUpdateNumber
End Property

Private Property Get IGraphicObjectSeries_Id() As Long
IGraphicObjectSeries_Id = Id
End Property

Private Sub IGraphicObjectSeries_Initialise( _
                ByVal pName As String, _
                ByVal pLayer As LayerNumbers, _
                ByVal pSeriesID As Long, _
                ByVal pController As Controller)
mName = pName
Layer = pLayer
mSeriesID = pSeriesID
Set mController = pController
Set mModel = mController.Model
Set mGraphics = mController.Graphics
End Sub

Private Property Get IGraphicObjectSeries_IsSelectable() As Boolean
IGraphicObjectSeries_IsSelectable = IsSelectable
End Property

Private Property Let IGraphicObjectSeries_IsSelected(ByVal Value As Boolean)
If Value = mBase.ExtendedPropertyHost.GetLocalValue(GDataPoint.gIsSelectedProperty) Then
Else
    mBase.ExtendedPropertyHost.SetValue GDataPoint.gIsSelectedProperty, Value
    RaiseEvent SelectionStateChanged
End If
End Property

Private Property Get IGraphicObjectSeries_IsSelected() As Boolean
IGraphicObjectSeries_IsSelected = IsSelected
End Property

Private Function IGraphicObjectSeries_Item(ByVal pIndex As Long) As IGraphicObject
Const ProcName As String = "IGraphicObjectSeries_item"
Dim failpoint As String
On Error GoTo Err

Set IGraphicObjectSeries_Item = mGraphicObjects.Item(pIndex)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IGraphicObjectSeries_LastGraphicsChange() As GraphicsChangeTypes
IGraphicObjectSeries_LastGraphicsChange = mLastGraphicsChange
End Property

Private Property Get IGraphicObjectSeries_LayeredGraphicsModel() As LayeredGraphicsModel
Set IGraphicObjectSeries_LayeredGraphicsModel = mModel
End Property

Private Property Get IGraphicObjectSeries_Name() As String
IGraphicObjectSeries_Name = mName
End Property

Private Sub IGraphicObjectSeries_Remove(ByVal pIndex As Long)
mGraphicObjects.Remove pIndex
End Sub

Private Sub IGraphicObjectSeries_RemoveObject(ByVal Value As IGraphicObject)
mGraphicObjects.Remove Hex$(ObjPtr(Value))
End Sub

Private Property Let IGraphicObjectSeries_Renderer(ByVal Value As IRenderer)
Renderer = Value
End Property

Private Property Get IGraphicObjectSeries_Renderer() As IRenderer
Set IGraphicObjectSeries_Renderer = Renderer
End Property

Private Property Get IGraphicObjectSeries_SeriesID() As Long
IGraphicObjectSeries_SeriesID = mSeriesID
End Property

Private Property Get IGraphicObjectSeries_SeriesUpdateNumber() As Long
IGraphicObjectSeries_SeriesUpdateNumber = mSeriesUpdateNumber
End Property

'@================================================================================
' mBase Event Handlers
'@================================================================================

Private Sub mBase_RepaintRegion(ByVal pRecalcBoundaries As Boolean)
Const ProcName As String = "mBase_RepaintRegion"
On Error GoTo Err

RepaintRegion pRecalcBoundaries

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mGraphics Event Handlers
'@================================================================================

Private Sub mGraphics_Change(ev As ChangeEventData)
mGraphicsUpdateNumber = mGraphicsUpdateNumber + 1
mLastGraphicsChange = ev.ChangeType
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Brush(ByVal Value As IBrush)
Const ProcName As String = "Brush"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gBrushProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Brush() As IBrush
Const ProcName As String = "Brush"
On Error GoTo Err

Set Brush = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gBrushProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Count() As Long
Const ProcName As String = "Count"
Dim failpoint As String
On Error GoTo Err

Count = mGraphicObjects.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DisplayMode(ByRef Value As DataPointDisplayModes)
Const ProcName As String = "DisplayMode"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gDisplayModeProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
Const ProcName As String = "DisplayMode"
On Error GoTo Err

DisplayMode = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gDisplayModeProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DownBrush(ByVal Value As IBrush)
Const ProcName As String = "DownBrush"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gDownBrushProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DownBrush() As IBrush
Const ProcName As String = "DownBrush"
On Error GoTo Err

Set DownBrush = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gDownBrushProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DownLinePen(ByVal Value As Pen)
Const ProcName As String = "DownLinePen"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gDownLinePenProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DownLinePen() As Pen
Const ProcName As String = "DownLinePen"
On Error GoTo Err

Set DownLinePen = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gDownLinePenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DownPen(ByVal Value As Pen)
Const ProcName As String = "DownPen"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gDownPenProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DownPen() As Pen
Const ProcName As String = "DownPen"
On Error GoTo Err

Set DownPen = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gDownPenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HistogramBarWidth(ByVal Value As Double)
Const ProcName As String = "HistogramBarWidth"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gHistogramBarWidthProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HistogramBarWidth() As Double
Const ProcName As String = "HistogramBarWidth"
On Error GoTo Err

HistogramBarWidth = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gHistogramBarWidthProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HistogramBaselinePen(ByVal Value As Pen)
Const ProcName As String = "HistogramBaselinePen"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gHistogramBaselinePenProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HistogramBaselinePen() As Pen
Const ProcName As String = "HistogramBaselinePen"
On Error GoTo Err

Set HistogramBaselinePen = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gHistogramBaselinePenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let HistogramBaseValue(ByVal Value As Double)
Const ProcName As String = "HistogramBaseValue"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gHistogramBaseValueProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HistogramBaseValue() As Double
Const ProcName As String = "HistogramBaseValue"
On Error GoTo Err

HistogramBaseValue = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gHistogramBaseValueProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Id() As Long
Id = mSeriesID
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gIncludeInAutoscaleProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IncludeInAutoscale() As Boolean
Const ProcName As String = "IncludeInAutoscale"
On Error GoTo Err

IncludeInAutoscale = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gIncludeInAutoscaleProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
Const ProcName As String = "IsSelectable"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gIsSelectableProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsSelectable() As Boolean
Const ProcName As String = "IsSelectable"
On Error GoTo Err

IsSelectable = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gIsSelectableProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsSelected() As Boolean
Const ProcName As String = "IsSelected"
On Error GoTo Err

IsSelected = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gIsSelectedProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gLayerProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gLayerProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LineMode(ByVal Value As DataPointLineModes)
Const ProcName As String = "LineMode"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gLineModeProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LineMode() As DataPointLineModes
Const ProcName As String = "LineMode"
On Error GoTo Err

LineMode = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gLineModeProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let LinePen(ByVal Value As Pen)
Const ProcName As String = "LinePen"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gLinePenProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LinePen() As Pen
Const ProcName As String = "LinePen"
On Error GoTo Err

Set LinePen = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gLinePenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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

mBase.ExtendedPropertyHost.SetValue GDataPoint.gNumberOfSidesProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get NumberOfSides() As Long
Const ProcName As String = "NumberOfSides"
On Error GoTo Err

NumberOfSides = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gNumberOfSidesProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Orientation(ByVal Value As Double)
Const ProcName As String = "Orientation"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gOrientationProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Orientation() As Double
Const ProcName As String = "Layer"
On Error GoTo Err

Orientation = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gOrientationProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Pen(ByVal Value As Pen)
Const ProcName As String = "Pen"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gPenProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Pen() As Pen
Const ProcName As String = "Pen"
On Error GoTo Err

Set Pen = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gPenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Renderer() As IRenderer
If mRenderer Is Nothing Then
    ' TODO: correct the default renderer class name
    Dim lRenderer  As New DataPointDefaultRenderer
    lRenderer.Graphics = mController.Graphics
    Set mRenderer = lRenderer
End If
Set Renderer = mRenderer
End Property

Public Property Let Renderer(ByVal Value As IRenderer)
Set mRenderer = Value
End Property

Public Property Let Size(ByVal Value As Size)
Const ProcName As String = "Size"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gSizeProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Size() As Size
Const ProcName As String = "Size"
On Error GoTo Err

Set Size = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gSizeProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' TODO: correct the style class name
Public Property Let Style(ByVal Value As DataPointStyle)
Const ProcName As String = "Style"
On Error GoTo Err

Set mStyle = Value
mBase.ExtendedPropertyHost.Style = mStyle.ExtendedPropertyHost

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' TODO: correct the style class name
Public Property Get Style() As DataPointStyle
Set Style = mStyle
End Property

Public Property Let UpBrush(ByVal Value As IBrush)
Const ProcName As String = "UpBrush"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gUpBrushProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get UpBrush() As IBrush
Const ProcName As String = "UpBrush"
On Error GoTo Err

Set UpBrush = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gUpBrushProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let UpLinePen(ByVal Value As Pen)
Const ProcName As String = "UpLinePen"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gUpLinePenProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get UpLinePen() As Pen
Const ProcName As String = "UpLinePen"
On Error GoTo Err

Set UpLinePen = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gUpLinePenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let UpPen(ByVal Value As Pen)
Const ProcName As String = "UpPen"
On Error GoTo Err

mBase.ExtendedPropertyHost.SetValue GDataPoint.gUpPenProperty, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get UpPen() As Pen
Const ProcName As String = "UpPen"
On Error GoTo Err

Set UpPen = mBase.ExtendedPropertyHost.GetValue(GDataPoint.gUpPenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add(Optional ByVal pModel As IDataPointModel) As DataPoint
Const ProcName As String = "Add"
On Error GoTo Err

If mFinished Then Err.Raise ErrorCodes.ErrIllegalStateException, , ModuleName & " is finished"

Set Add = New DataPoint

If pModel Is Nothing Then Set pModel = New DataPointDefaultModel
Add.Initialise Me, pModel
mGraphicObjects.Add Add, Hex$(ObjPtr(Add))

If Not mPrevDataPointRef Is Nothing Then Add.PreviousDataPoint = prevDataPointObj
Set mPrevDataPointRef = CreateWeakReference(Add)

mModel.AddGraphicObject Add

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Sub ApplyChanges()
mBase.ApplyChanges
End Sub

Public Function Clear() As TaskController
Const ProcName As String = "Clear"
On Error GoTo Err

Set Clear = mModel.ClearGraphicObjectSeries(Me)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Sub ClearBrush()
Const ProcName As String = "ClearBrush"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gBrushProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearDisplayMode()
Const ProcName As String = "ClearDisplayMode"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gDisplayModeProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearDownBrush()
Const ProcName As String = "ClearDownBrush"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gDownBrushProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearDownLinePen()
Const ProcName As String = "ClearDownLinePen"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gDownLinePenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearDownPen()
Const ProcName As String = "ClearDownPen"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gDownPenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearHistogramBarWidth()
Const ProcName As String = "ClearHistogramBarWidth"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gHistogramBarWidthProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearHistogramBaselinePen()
Const ProcName As String = "ClearHistogramBaselinePen"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gHistogramBaselinePenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearHistogramBaseValue()
Const ProcName As String = "ClearHistogramBaseValue"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gHistogramBaseValueProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearIncludeInAutoscale()
Const ProcName As String = "ClearIncludeInAutoscale"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gIncludeInAutoscaleProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearIsSelectable()
Const ProcName As String = "ClearIsSelectable"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gIsSelectableProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearLayer()
Const ProcName As String = "ClearLayer"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gLayerProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearLineMode()
Const ProcName As String = "ClearLineMode"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gLineModeProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearLinePen()
Const ProcName As String = "ClearLinePen"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gLinePenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearNumberOfSides()
Const ProcName As String = "ClearNumberOfSides"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gNumberOfSidesProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearOrientation()
Const ProcName As String = "ClearOrientation"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gOrientationProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearPen()
Const ProcName As String = "ClearPen"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gPenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearSize()
Const ProcName As String = "ClearSize"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gSizeProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearUpBrush()
Const ProcName As String = "ClearUpBrush"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gUpBrushProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearUpLinePen()
Const ProcName As String = "ClearUpLinePen"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gUpLinePenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearUpPen()
Const ProcName As String = "ClearUpPen"
On Error GoTo Err

mBase.ExtendedPropertyHost.ClearValue GDataPoint.gUpPenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Edit()
mBase.Edit
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
Dim failpoint As String
On Error GoTo Err

Set NewEnum = mGraphicObjects.[_NewEnum]

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function prevDataPointObj() As DataPoint
Const ProcName As String = "prevDataPointObj"

On Error GoTo Err

Set prevDataPointObj = mPrevDataPointRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub RepaintRegion( _
                ByVal pRecalcBoundaries As Boolean)
Const ProcName As String = "RepaintRegion"
Dim failpoint As String
On Error GoTo Err

If mGraphicObjects.Count <> 0 Then
    If pRecalcBoundaries Then
        mSeriesUpdateNumber = mSeriesUpdateNumber + 1
        mModel.RecalcBoundaryForObjectsInSeries mSeriesID
        mController.paint True
    Else
        mController.paint False
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub








