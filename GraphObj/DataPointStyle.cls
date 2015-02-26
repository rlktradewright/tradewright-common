VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPointStyle"
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

Implements IGraphicObjectStyle

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

' TODO: replace DataPoint by the related graphic object class name
Private Const ModuleName                            As String = "DataPointStyle"

' TODO: replace DataPoint by the related graphic object class name
Private Const RelatedClassName                      As String = "DataPoint"

'@================================================================================
' Member variables
'@================================================================================

Private mName                                       As String
Private mStyleType                                  As String
Private mConfig                                     As ConfigurationSection

Private mEPhost                                     As ExtendedPropertyHost

Private mBasedOn                                    As DataPointStyle

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mEPhost = New ExtendedPropertyHost

' TODO: replace DataPoint by the related graphic object class name
GDataPoint.gRegisterProperties
GDataPoint.gRegisterExtendedEvents

mName = GenerateGUIDString
mStyleType = RelatedClassName & " Styles"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' IGraphicObjectStyle Interface Members
'@================================================================================

Private Property Let IGraphicObjectStyle_ConfigurationSection(ByVal Value As ConfigurationSection)
Me.ConfigurationSection = Value
End Property

Private Sub IGraphicObjectStyle_LoadFromConfig(ByVal pConfig As ConfigurationSection)
Me.LoadFromConfig pConfig
End Sub

Private Property Get IGraphicObjectStyle_Name() As String
IGraphicObjectStyle_Name = mName
End Property

Private Property Get IGraphicObjectStyle_ProgId() As String
IGraphicObjectStyle_ProgId = ProjectName & "." & ModuleName
End Property

Private Property Get IGraphicObjectStyle_StyleType() As String
IGraphicObjectStyle_StyleType = mStyleType
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

' TODO: replace DataPoint by the related graphic object class name
Public Property Let BasedOn(ByVal Value As DataPointStyle)
Set mBasedOn = Value
mEPhost.Style = mBasedOn.ExtendedPropertyHost
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBasedOn, Value.Name
End Property

' TODO: replace DataPoint by the related graphic object class name
Public Property Get BasedOn() As DataPointStyle
Set BasedOn = mBasedOn
End Property

Public Property Let Brush(ByVal Value As IBrush)
Const ProcName As String = "Brush"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As IBrush
If setProperty(GDataPoint.gBrushProperty, Value, prevValue) Then
        If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gBrushProperty))
                If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
        End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Brush() As IBrush
Const ProcName As String = "Brush"
On Error GoTo Err

Set Brush = mEPhost.GetLocalValue(GDataPoint.gBrushProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
Dim failpoint As String
On Error GoTo Err

If Value Is Nothing Then
    RemoveFromConfig
    Set mConfig = Nothing
    Exit Property
End If

If Value Is mConfig Then Exit Property
Set mConfig = Value

' TODO: add similar statements for custom properties
mConfig.SetSetting ConfigSettingName, mName
mConfig.SetSetting ConfigSettingStyleType, mStyleType

If Not mBasedOn Is Nothing Then mConfig.SetSetting ConfigSettingBasedOn, mBasedOn.Name

' TODO: add similar statements for custom properties
If isLocalValueSet(GDataPoint.gDownLinePenProperty) Then mEPhost.GetLocalValue(GDataPoint.gDownLinePenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gDownLinePenProperty))
If isLocalValueSet(GDataPoint.gUpLinePenProperty) Then mEPhost.GetLocalValue(GDataPoint.gUpLinePenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gUpLinePenProperty))
If isLocalValueSet(GDataPoint.gLinePenProperty) Then mEPhost.GetLocalValue(GDataPoint.gLinePenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gLinePenProperty))
If isLocalValueSet(GDataPoint.gHistogramBaselinePenProperty) Then mEPhost.GetLocalValue(GDataPoint.gHistogramBaselinePenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gHistogramBaselinePenProperty))
If isLocalValueSet(GDataPoint.gHistogramBaseValueProperty) Then mConfig.SetSetting gGetConfigName(GDataPoint.gHistogramBaseValueProperty), mEPhost.GetLocalValue(GDataPoint.gHistogramBaseValueProperty)
If isLocalValueSet(GDataPoint.gLineModeProperty) Then mConfig.SetSetting gGetConfigName(GDataPoint.gLineModeProperty), mEPhost.GetLocalValue(GDataPoint.gLineModeProperty)
If isLocalValueSet(GDataPoint.gHistogramBarWidthProperty) Then mConfig.SetSetting gGetConfigName(GDataPoint.gHistogramBarWidthProperty), mEPhost.GetLocalValue(GDataPoint.gHistogramBarWidthProperty)
If isLocalValueSet(GDataPoint.gDownPenProperty) Then mEPhost.GetLocalValue(GDataPoint.gDownPenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gDownPenProperty))
If isLocalValueSet(GDataPoint.gUpPenProperty) Then mEPhost.GetLocalValue(GDataPoint.gUpPenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gUpPenProperty))
If isLocalValueSet(GDataPoint.gDownBrushProperty) Then mEPhost.GetLocalValue(GDataPoint.gDownBrushProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gDownBrushProperty))
If isLocalValueSet(GDataPoint.gUpBrushProperty) Then mEPhost.GetLocalValue(GDataPoint.gUpBrushProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gUpBrushProperty))
If isLocalValueSet(GDataPoint.gNumberOfSidesProperty) Then mConfig.SetSetting gGetConfigName(GDataPoint.gNumberOfSidesProperty), mEPhost.GetLocalValue(GDataPoint.gNumberOfSidesProperty)
If isLocalValueSet(GDataPoint.gDisplayModeProperty) Then mConfig.SetSetting gGetConfigName(GDataPoint.gDisplayModeProperty), mEPhost.GetLocalValue(GDataPoint.gDisplayModeProperty)
If isLocalValueSet(GDataPoint.gBrushProperty) Then mEPhost.GetLocalValue(GDataPoint.gBrushProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gBrushProperty))
If isLocalValueSet(GDataPoint.gIsSelectableProperty) Then mConfig.SetSetting gGetConfigName(GDataPoint.gIsSelectableProperty), mEPhost.GetLocalValue(GDataPoint.gIsSelectableProperty)
If isLocalValueSet(GDataPoint.gLayerProperty) Then mConfig.SetSetting gGetConfigName(GDataPoint.gLayerProperty), mEPhost.GetLocalValue(GDataPoint.gLayerProperty)
If isLocalValueSet(GDataPoint.gOrientationProperty) Then mConfig.SetSetting gGetConfigName(GDataPoint.gOrientationProperty), mEPhost.GetLocalValue(GDataPoint.gOrientationProperty)
If isLocalValueSet(GDataPoint.gPenProperty) Then mEPhost.GetLocalValue(GDataPoint.gPenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gPenProperty))
If isLocalValueSet(GDataPoint.gSizeProperty) Then mEPhost.GetLocalValue(GDataPoint.gSizeProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gSizeProperty))

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DisplayMode(ByRef Value As DataPointDisplayModes)
Const ProcName As String = "DisplayMode"
On Error GoTo Err

setProperty GDataPoint.gDisplayModeProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GDataPoint.gDisplayModeProperty), Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DisplayMode() As DataPointDisplayModes
Const ProcName As String = "DisplayMode"
On Error GoTo Err

DisplayMode = mEPhost.GetLocalValue(GDataPoint.gDisplayModeProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DownBrush(ByVal Value As IBrush)
Const ProcName As String = "DownBrush"
On Error GoTo Err

Dim prevValue As IBrush
If setProperty(GDataPoint.gDownBrushProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gDownBrushProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DownBrush() As IBrush
Const ProcName As String = "DownBrush"
On Error GoTo Err

Set DownBrush = mEPhost.GetLocalValue(GDataPoint.gDownBrushProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DownLinePen(ByVal Value As Pen)
Const ProcName As String = "DownLinePen"
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GDataPoint.gDownLinePenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gDownLinePenProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DownLinePen() As Pen
Const ProcName As String = "DownLinePen"
On Error GoTo Err

Set DownLinePen = mEPhost.GetLocalValue(GDataPoint.gDownLinePenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let DownPen(ByVal Value As Pen)
Const ProcName As String = "DownPen"
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GDataPoint.gDownPenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gDownPenProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get DownPen() As Pen
Const ProcName As String = "DownPen"
On Error GoTo Err

Set DownPen = mEPhost.GetLocalValue(GDataPoint.gDownPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HistogramBarWidth(ByVal Value As Double)
Const ProcName As String = "HistogramBarWidth"
On Error GoTo Err

setProperty GDataPoint.gHistogramBarWidthProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GDataPoint.gHistogramBarWidthProperty), Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HistogramBarWidth() As Double
Const ProcName As String = "HistogramBarWidth"
On Error GoTo Err

HistogramBarWidth = mEPhost.GetLocalValue(GDataPoint.gHistogramBarWidthProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let HistogramBaselinePen(ByVal Value As Pen)
Const ProcName As String = "HistogramBaselinePen"
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GDataPoint.gHistogramBaselinePenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gHistogramBaselinePenProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get HistogramBaselinePen() As Pen
Const ProcName As String = "HistogramBaselinePen"
On Error GoTo Err

Set HistogramBaselinePen = mEPhost.GetLocalValue(GDataPoint.gHistogramBaselinePenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Friend Property Get ExtendedPropertyHost() As ExtendedPropertyHost
Set ExtendedPropertyHost = mEPhost
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
Const ProcName As String = "IsSelectable"
Dim failpoint As String
On Error GoTo Err

setProperty GDataPoint.gIsSelectableProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GDataPoint.gIsSelectableProperty), Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsSelectable() As Boolean
Const ProcName As String = "IsSelectable"
On Error GoTo Err

IsSelectable = mEPhost.GetLocalValue(GDataPoint.gIsSelectableProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

setProperty GDataPoint.gLayerProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GDataPoint.gLayerProperty), Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mEPhost.GetLocalValue(GDataPoint.gLayerProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let LineMode(ByVal Value As DataPointLineModes)
Const ProcName As String = "LineMode"
On Error GoTo Err

setProperty GDataPoint.gLineModeProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GDataPoint.gLineModeProperty), Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LineMode() As DataPointLineModes
Const ProcName As String = "LineMode"
On Error GoTo Err

LineMode = mEPhost.GetLocalValue(GDataPoint.gLineModeProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let LinePen(ByVal Value As Pen)
Const ProcName As String = "LinePen"
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GDataPoint.gLinePenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gLinePenProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get LinePen() As Pen
Const ProcName As String = "LinePen"
On Error GoTo Err

Set LinePen = mEPhost.GetLocalValue(GDataPoint.gLinePenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let NumberOfSides(ByVal Value As Long)
Const ProcName As String = "NumberOfSides"
On Error GoTo Err

setProperty GDataPoint.gNumberOfSidesProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GDataPoint.gNumberOfSidesProperty), Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get NumberOfSides() As Long
Const ProcName As String = "NumberOfSides"
On Error GoTo Err

NumberOfSides = mEPhost.GetLocalValue(GDataPoint.gNumberOfSidesProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Orientation(ByVal Value As Double)
Const ProcName As String = "Orientation"
Dim failpoint As String
On Error GoTo Err

setProperty GDataPoint.gOrientationProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GDataPoint.gOrientationProperty), Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Orientation() As Double
Const ProcName As String = "Orientation"
On Error GoTo Err

Orientation = mEPhost.GetLocalValue(GDataPoint.gOrientationProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Pen(ByVal Value As Pen)
Const ProcName As String = "Pen"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GDataPoint.gPenProperty, Value, prevValue) Then
        If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gPenProperty))
                If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
        End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Pen() As Pen
Const ProcName As String = "Pen"
On Error GoTo Err

Set Pen = mEPhost.GetLocalValue(GDataPoint.gPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Size(ByVal Value As Size)
Const ProcName As String = "Size"
On Error GoTo Err

Dim prevValue As Size
If setProperty(GDataPoint.gSizeProperty, Value, prevValue) Then
        If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gSizeProperty))
                If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
        End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Size() As Size
Const ProcName As String = "Size"
On Error GoTo Err

Set Size = mEPhost.GetLocalValue(GDataPoint.gSizeProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let UpBrush(ByVal Value As IBrush)
Const ProcName As String = "UpBrush"
On Error GoTo Err

Dim prevValue As IBrush
If setProperty(GDataPoint.gUpBrushProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gUpBrushProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get UpBrush() As IBrush
Const ProcName As String = "UpBrush"
On Error GoTo Err

Set UpBrush = mEPhost.GetLocalValue(GDataPoint.gUpBrushProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let UpLinePen(ByVal Value As Pen)
Const ProcName As String = "UpLinePen"
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GDataPoint.gUpLinePenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gUpLinePenProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get UpLinePen() As Pen
Const ProcName As String = "UpLinePen"
On Error GoTo Err

Set UpLinePen = mEPhost.GetLocalValue(GDataPoint.gUpLinePenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let UpPen(ByVal Value As Pen)
Const ProcName As String = "UpPen"
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GDataPoint.gUpPenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GDataPoint.gUpPenProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get UpPen() As Pen
Const ProcName As String = "UpPen"
On Error GoTo Err

Set UpPen = mEPhost.GetLocalValue(GDataPoint.gUpPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ClearBrush()
Const ProcName As String = "ClearBrush"
Dim failpoint As String
On Error GoTo Err

clearProperty GDataPoint.gBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDisplayMode()
Const ProcName As String = "ClearDisplayMode"
On Error GoTo Err

clearProperty GDataPoint.gDisplayModeProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDownBrush()
Const ProcName As String = "ClearDownBrush"
On Error GoTo Err

clearProperty GDataPoint.gDownBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDownLinePen()
Const ProcName As String = "ClearDownLinePen"
On Error GoTo Err

clearProperty GDataPoint.gDownLinePenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearDownPen()
Const ProcName As String = "ClearDownPen"
On Error GoTo Err

clearProperty GDataPoint.gDownPenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearHistogramBarWidth()
Const ProcName As String = "ClearHistogramBarWidth"
On Error GoTo Err

clearProperty GDataPoint.gHistogramBarWidthProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearHistogramBaselinePen()
Const ProcName As String = "ClearHistogramBaselinePen"
On Error GoTo Err

clearProperty GDataPoint.gHistogramBaselinePenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearHistogramBaseValue()
Const ProcName As String = "ClearHistogramBaseValue"
On Error GoTo Err

clearProperty GDataPoint.gHistogramBaseValueProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearLayer()
Const ProcName As String = "ClearLayer"
Dim failpoint As String
On Error GoTo Err

clearProperty GDataPoint.gLayerProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearLineMode()
Const ProcName As String = "ClearLineMode"
On Error GoTo Err

clearProperty GDataPoint.gLineModeProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearLinePen()
Const ProcName As String = "ClearLinePen"
On Error GoTo Err

clearProperty GDataPoint.gLinePenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearNumberOfSides()
Const ProcName As String = "ClearNumberOfSides"
On Error GoTo Err

clearProperty GDataPoint.gNumberOfSidesProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearOrientation()
Const ProcName As String = "ClearOrientation"
On Error GoTo Err

clearProperty GDataPoint.gOrientationProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearPen()
Const ProcName As String = "ClearPen"
Dim failpoint As String
On Error GoTo Err

clearProperty GDataPoint.gPenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearSize()
Const ProcName As String = "ClearSize"
Dim failpoint As String
On Error GoTo Err

clearProperty GDataPoint.gSizeProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearUpBrush()
Const ProcName As String = "ClearUpBrush"
On Error GoTo Err

clearProperty GDataPoint.gUpBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearUpLinePen()
Const ProcName As String = "ClearUpLinePen"
On Error GoTo Err

clearProperty GDataPoint.gUpLinePenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearUpPen()
Const ProcName As String = "ClearUpPen"
On Error GoTo Err

clearProperty GDataPoint.gUpPenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

' TODO: replace DataPoint by the related graphic object class name
Public Function Clone() As DataPointStyle
Const ProcName As String = "Clone"
Dim failpoint As String
On Error GoTo Err

' TODO: replace DataPoint by the related graphic object class name
Set Clone = New DataPointStyle

' TODO: add similar statements for custom properties
Clone.BasedOn = mBasedOn
If isLocalValueSet(GDataPoint.gDownLinePenProperty) Then Clone.DownLinePen = mEPhost.GetLocalValue(GDataPoint.gDownLinePenProperty)
If isLocalValueSet(GDataPoint.gUpLinePenProperty) Then Clone.UpLinePen = mEPhost.GetLocalValue(GDataPoint.gUpLinePenProperty)
If isLocalValueSet(GDataPoint.gLinePenProperty) Then Clone.LinePen = mEPhost.GetLocalValue(GDataPoint.gLinePenProperty)
If isLocalValueSet(GDataPoint.gHistogramBaselinePenProperty) Then Clone.HistogramBaselinePen = mEPhost.GetLocalValue(GDataPoint.gHistogramBaselinePenProperty)
If isLocalValueSet(GDataPoint.gLineModeProperty) Then Clone.LineMode = mEPhost.GetLocalValue(GDataPoint.gLineModeProperty)
If isLocalValueSet(GDataPoint.gHistogramBarWidthProperty) Then Clone.HistogramBarWidth = mEPhost.GetLocalValue(GDataPoint.gHistogramBarWidthProperty)
If isLocalValueSet(GDataPoint.gDownPenProperty) Then Clone.DownPen = mEPhost.GetLocalValue(GDataPoint.gDownPenProperty)
If isLocalValueSet(GDataPoint.gUpPenProperty) Then Clone.UpPen = mEPhost.GetLocalValue(GDataPoint.gUpPenProperty)
If isLocalValueSet(GDataPoint.gDownBrushProperty) Then Clone.DownBrush = mEPhost.GetLocalValue(GDataPoint.gDownBrushProperty)
If isLocalValueSet(GDataPoint.gUpBrushProperty) Then Clone.UpBrush = mEPhost.GetLocalValue(GDataPoint.gUpBrushProperty)
If isLocalValueSet(GDataPoint.gNumberOfSidesProperty) Then Clone.NumberOfSides = mEPhost.GetLocalValue(GDataPoint.gNumberOfSidesProperty)
If isLocalValueSet(GDataPoint.gDisplayModeProperty) Then Clone.DisplayMode = mEPhost.GetLocalValue(GDataPoint.gDisplayModeProperty)
If isLocalValueSet(GDataPoint.gBrushProperty) Then Clone.Brush = mEPhost.GetLocalValue(GDataPoint.gBrushProperty)
If isLocalValueSet(GDataPoint.gIsSelectableProperty) Then Clone.IsSelectable = mEPhost.GetLocalValue(GDataPoint.gIsSelectableProperty)
If isLocalValueSet(GDataPoint.gLayerProperty) Then Clone.Layer = mEPhost.GetLocalValue(GDataPoint.gLayerProperty)
If isLocalValueSet(GDataPoint.gOrientationProperty) Then Clone.Orientation = mEPhost.GetLocalValue(GDataPoint.gOrientationProperty)
If isLocalValueSet(GDataPoint.gPenProperty) Then Clone.Pen = mEPhost.GetLocalValue(GDataPoint.gPenProperty)
If isLocalValueSet(GDataPoint.gSizeProperty) Then Clone.Size = mEPhost.GetLocalValue(GDataPoint.gSizeProperty)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
On Error GoTo Err

Set mConfig = pConfig
mName = mConfig.GetSetting(ConfigSettingName)
mStyleType = mConfig.GetSetting(ConfigSettingStyleType)

' TODO: add similar statements for custom properties
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gUpBrushProperty)) Is Nothing Then UpBrush = LoadIBrushFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gUpBrushProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gDownBrushProperty)) Is Nothing Then DownBrush = LoadIBrushFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gDownBrushProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gUpPenProperty)) Is Nothing Then UpPen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gUpPenProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gDownPenProperty)) Is Nothing Then DownPen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gDownPenProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gHistogramBaselinePenProperty)) Is Nothing Then HistogramBaselinePen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gHistogramBaselinePenProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gLinePenProperty)) Is Nothing Then LinePen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gLinePenProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gUpLinePenProperty)) Is Nothing Then UpLinePen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gUpLinePenProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gDownLinePenProperty)) Is Nothing Then DownLinePen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gDownLinePenProperty)))
If mConfig.GetSetting(gGetConfigName(GDataPoint.gLineModeProperty)) <> "" Then LineMode = mConfig.GetSetting(gGetConfigName(GDataPoint.gLineModeProperty))
If mConfig.GetSetting(gGetConfigName(GDataPoint.gHistogramBarWidthProperty)) <> "" Then HistogramBarWidth = mConfig.GetSetting(gGetConfigName(GDataPoint.gHistogramBarWidthProperty))
If mConfig.GetSetting(gGetConfigName(GDataPoint.gNumberOfSidesProperty)) <> "" Then NumberOfSides = mConfig.GetSetting(gGetConfigName(GDataPoint.gNumberOfSidesProperty))
If mConfig.GetSetting(gGetConfigName(GDataPoint.gDisplayModeProperty)) <> "" Then DisplayMode = mConfig.GetSetting(gGetConfigName(GDataPoint.gDisplayModeProperty))
If mConfig.GetSetting(gGetConfigName(GDataPoint.gLayerProperty)) <> "" Then Layer = mConfig.GetSetting(gGetConfigName(GDataPoint.gLayerProperty))
If mConfig.GetSetting(gGetConfigName(GDataPoint.gIsSelectableProperty)) <> "" Then IsSelectable = mConfig.GetSetting(gGetConfigName(GDataPoint.gIsSelectableProperty))
If mConfig.GetSetting(gGetConfigName(GDataPoint.gOrientationProperty)) <> "" Then Orientation = mConfig.GetSetting(gGetConfigName(GDataPoint.gOrientationProperty))
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gBrushProperty)) Is Nothing Then Brush = LoadIBrushFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gBrushProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gPenProperty)) Is Nothing Then Pen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gPenProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gSizeProperty)) Is Nothing Then Size = LoadSizeFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GDataPoint.gSizeProperty)))

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
                
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As String
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clearProperty(ByVal pExtProp As ExtendedProperty)
Const ProcName As String = "clearProperty"
On Error GoTo Err

mEPhost.ClearValue pExtProp
If Not mConfig Is Nothing Then mConfig.RemoveSetting gGetConfigName(pExtProp)

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function isLocalValueSet(ByVal pExtProp As ExtendedProperty) As Boolean
Const ProcName As String = "isLocalValueSet"
On Error GoTo Err

isLocalValueSet = Not IsEmpty(mEPhost.GetLocalValue(pExtProp))

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function


Private Function setProperty( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewValue As Variant, _
                Optional ByRef pPrevValue As Variant) As Boolean
Const ProcName As String = "setProperty"
On Error GoTo Err

setProperty = gSetProperty(mEPhost, pExtProp, pNewValue, pPrevValue)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function






