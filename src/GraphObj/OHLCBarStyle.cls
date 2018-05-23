VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OHLCBarStyle"
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

' TODO: replace OHLCBar by the related graphic object class name
Private Const ModuleName                            As String = "OHLCBarStyle"

' TODO: replace OHLCBar by the related graphic object class name
Private Const RelatedClassName                      As String = "OHLCBar"

'@================================================================================
' Member variables
'@================================================================================

Private mName                                       As String
Private mStyleType                                  As String
Private mConfig                                     As ConfigurationSection

Private mEPhost                                     As ExtendedPropertyHost

Private mBasedOn                                    As OHLCBarStyle

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mEPhost = New ExtendedPropertyHost

GOHLCBar.gRegisterProperties
GOHLCBar.gRegisterExtendedEvents

mName = GenerateGUIDString
mStyleType = RelatedClassName & " Styles"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
' OHLCBar Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

' TODO: replace OHLCBar by the related graphic object class name
Public Property Let BasedOn(ByVal Value As OHLCBarStyle)
Set mBasedOn = Value
mEPhost.Style = mBasedOn.ExtendedPropertyHost
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBasedOn, Value.Name
End Property

' TODO: replace OHLCBar by the related graphic object class name
Public Property Get BasedOn() As OHLCBarStyle
Set BasedOn = mBasedOn
End Property

Public Property Let Brush(ByVal Value As IBrush)
Const ProcName As String = "Brush"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As IBrush
If setProperty(GOHLCBar.gBrushProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gBrushProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Brush() As IBrush
Const ProcName As String = "Brush"
On Error GoTo Err

Set Brush = mEPhost.GetLocalValue(GOHLCBar.gBrushProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
If isLocalValueSet(GOHLCBar.gWidthProperty) Then mConfig.SetSetting gGetConfigName(GOHLCBar.gWidthProperty), mEPhost.GetLocalValue(GOHLCBar.gWidthProperty)
If isLocalValueSet(GOHLCBar.gUpPenProperty) Then mEPhost.GetLocalValue(GOHLCBar.gUpPenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gUpPenProperty))
If isLocalValueSet(GOHLCBar.gDownPenProperty) Then mEPhost.GetLocalValue(GOHLCBar.gDownPenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gDownPenProperty))
If isLocalValueSet(GOHLCBar.gUpBrushProperty) Then mEPhost.GetLocalValue(GOHLCBar.gUpBrushProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gUpBrushProperty))
If isLocalValueSet(GOHLCBar.gDownBrushProperty) Then mEPhost.GetLocalValue(GOHLCBar.gDownBrushProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gDownBrushProperty))
If isLocalValueSet(GOHLCBar.gDisplayModeProperty) Then mConfig.SetSetting gGetConfigName(GOHLCBar.gDisplayModeProperty), mEPhost.GetLocalValue(GOHLCBar.gDisplayModeProperty)
If isLocalValueSet(GOHLCBar.gBrushProperty) Then mEPhost.GetLocalValue(GOHLCBar.gBrushProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gBrushProperty))
If isLocalValueSet(GOHLCBar.gIsSelectableProperty) Then mConfig.SetSetting gGetConfigName(GOHLCBar.gIsSelectableProperty), mEPhost.GetLocalValue(GOHLCBar.gIsSelectableProperty)
If isLocalValueSet(GOHLCBar.gLayerProperty) Then mConfig.SetSetting gGetConfigName(GOHLCBar.gLayerProperty), mEPhost.GetLocalValue(GOHLCBar.gLayerProperty)
If isLocalValueSet(GOHLCBar.gOrientationProperty) Then mConfig.SetSetting gGetConfigName(GOHLCBar.gOrientationProperty), mEPhost.GetLocalValue(GOHLCBar.gOrientationProperty)
If isLocalValueSet(GOHLCBar.gPenProperty) Then mEPhost.GetLocalValue(GOHLCBar.gPenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gPenProperty))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DisplayMode(ByVal Value As OHLCBarDisplayModes)
Const ProcName As String = "DisplayMode"
On Error GoTo Err

setProperty GOHLCBar.gDisplayModeProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GOHLCBar.gDisplayModeProperty), Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DisplayMode() As OHLCBarDisplayModes
Const ProcName As String = "DisplayMode"
On Error GoTo Err

DisplayMode = mEPhost.GetLocalValue(GOHLCBar.gDisplayModeProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DownBrush(ByVal Value As IBrush)
Const ProcName As String = "DownBrush"
On Error GoTo Err

Dim prevValue As IBrush
If setProperty(GOHLCBar.gDownBrushProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gDownBrushProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DownBrush() As IBrush
Const ProcName As String = "DownBrush"
On Error GoTo Err

Set DownBrush = mEPhost.GetLocalValue(GOHLCBar.gDownBrushProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let DownPen(ByVal Value As Pen)
Const ProcName As String = "DownPen"
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GOHLCBar.gDownPenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gDownPenProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DownPen() As Pen
Const ProcName As String = "DownPen"
On Error GoTo Err

Set DownPen = mEPhost.GetLocalValue(GOHLCBar.gDownPenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ExtendedPropertyHost() As ExtendedPropertyHost
Set ExtendedPropertyHost = mEPhost
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
Const ProcName As String = "IsSelectable"
Dim failpoint As String
On Error GoTo Err

setProperty GOHLCBar.gIsSelectableProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GOHLCBar.gIsSelectableProperty), Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsSelectable() As Boolean
Const ProcName As String = "IsSelectable"
On Error GoTo Err

IsSelectable = mEPhost.GetLocalValue(GOHLCBar.gIsSelectableProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

setProperty GOHLCBar.gLayerProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GOHLCBar.gLayerProperty), Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mEPhost.GetLocalValue(GOHLCBar.gLayerProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Orientation(ByVal Value As Double)
Const ProcName As String = "Orientation"
Dim failpoint As String
On Error GoTo Err

setProperty GOHLCBar.gOrientationProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GOHLCBar.gOrientationProperty), Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Orientation() As Double
Const ProcName As String = "Orientation"
On Error GoTo Err

Orientation = mEPhost.GetLocalValue(GOHLCBar.gOrientationProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Pen(ByVal Value As Pen)
Const ProcName As String = "Pen"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GOHLCBar.gPenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gPenProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Pen() As Pen
Const ProcName As String = "Pen"
On Error GoTo Err

Set Pen = mEPhost.GetLocalValue(GOHLCBar.gPenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let UpBrush(ByVal Value As IBrush)
Const ProcName As String = "UpBrush"
On Error GoTo Err

Dim prevValue As IBrush
If setProperty(GOHLCBar.gUpBrushProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gUpBrushProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get UpBrush() As IBrush
Const ProcName As String = "UpBrush"
On Error GoTo Err

Set UpBrush = mEPhost.GetLocalValue(GOHLCBar.gUpBrushProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let UpPen(ByVal Value As Pen)
Const ProcName As String = "UpPen"
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GOHLCBar.gUpPenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GOHLCBar.gUpPenProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get UpPen() As Pen
Const ProcName As String = "UpPen"
On Error GoTo Err

Set UpPen = mEPhost.GetLocalValue(GOHLCBar.gUpPenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Width(ByVal Value As Double)
Const ProcName As String = "Width"
On Error GoTo Err

setProperty GOHLCBar.gWidthProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GOHLCBar.gWidthProperty), Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Width() As Double
Const ProcName As String = "Width"
On Error GoTo Err

Width = mEPhost.GetLocalValue(GOHLCBar.gWidthProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ClearBrush()
Const ProcName As String = "ClearBrush"
Dim failpoint As String
On Error GoTo Err

clearProperty GOHLCBar.gBrushProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearDisplayMode()
Const ProcName As String = "ClearDisplayMode"
On Error GoTo Err

clearProperty GOHLCBar.gDisplayModeProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearDownBrush()
Const ProcName As String = "ClearDownBrush"
On Error GoTo Err

clearProperty GOHLCBar.gDownBrushProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearDownPen()
Const ProcName As String = "ClearDownPen"
On Error GoTo Err

clearProperty GOHLCBar.gDownPenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearLayer()
Const ProcName As String = "ClearLayer"
Dim failpoint As String
On Error GoTo Err

clearProperty GOHLCBar.gLayerProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearOrientation()
Const ProcName As String = "ClearOrientation"
On Error GoTo Err

clearProperty GOHLCBar.gOrientationProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearPen()
Const ProcName As String = "ClearPen"
Dim failpoint As String
On Error GoTo Err

clearProperty GOHLCBar.gPenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearUpBrush()
Const ProcName As String = "ClearUpBrush"
On Error GoTo Err

clearProperty GOHLCBar.gUpBrushProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearUpPen()
Const ProcName As String = "ClearUpPen"
On Error GoTo Err

clearProperty GOHLCBar.gUpPenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearWidth()
Const ProcName As String = "ClearWidth"
On Error GoTo Err

clearProperty GOHLCBar.gWidthProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

' TODO: replace OHLCBar by the related graphic object class name
Public Function Clone() As OHLCBarStyle
Const ProcName As String = "Clone"
Dim failpoint As String
On Error GoTo Err

' TODO: replace OHLCBar by the related graphic object class name
Set Clone = New OHLCBarStyle

' TODO: add similar statements for custom properties
Clone.BasedOn = mBasedOn
If isLocalValueSet(GOHLCBar.gWidthProperty) Then Clone.Width = mEPhost.GetLocalValue(GOHLCBar.gWidthProperty)
If isLocalValueSet(GOHLCBar.gUpPenProperty) Then Clone.UpPen = mEPhost.GetLocalValue(GOHLCBar.gUpPenProperty)
If isLocalValueSet(GOHLCBar.gDownPenProperty) Then Clone.DownPen = mEPhost.GetLocalValue(GOHLCBar.gDownPenProperty)
If isLocalValueSet(GOHLCBar.gUpBrushProperty) Then Clone.UpBrush = mEPhost.GetLocalValue(GOHLCBar.gUpBrushProperty)
If isLocalValueSet(GOHLCBar.gDownBrushProperty) Then Clone.DownBrush = mEPhost.GetLocalValue(GOHLCBar.gDownBrushProperty)
If isLocalValueSet(GOHLCBar.gDisplayModeProperty) Then Clone.DisplayMode = mEPhost.GetLocalValue(GOHLCBar.gDisplayModeProperty)
If isLocalValueSet(GOHLCBar.gBrushProperty) Then Clone.Brush = mEPhost.GetLocalValue(GOHLCBar.gBrushProperty)
If isLocalValueSet(GOHLCBar.gIsSelectableProperty) Then Clone.IsSelectable = mEPhost.GetLocalValue(GOHLCBar.gIsSelectableProperty)
If isLocalValueSet(GOHLCBar.gLayerProperty) Then Clone.Layer = mEPhost.GetLocalValue(GOHLCBar.gLayerProperty)
If isLocalValueSet(GOHLCBar.gOrientationProperty) Then Clone.Orientation = mEPhost.GetLocalValue(GOHLCBar.gOrientationProperty)
If isLocalValueSet(GOHLCBar.gPenProperty) Then Clone.Pen = mEPhost.GetLocalValue(GOHLCBar.gPenProperty)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
If Not mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gDownBrushProperty)) Is Nothing Then DownBrush = LoadIBrushFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gDownBrushProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gUpBrushProperty)) Is Nothing Then UpBrush = LoadIBrushFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gUpBrushProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gDownPenProperty)) Is Nothing Then DownPen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gDownPenProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gUpPenProperty)) Is Nothing Then UpPen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gUpPenProperty)))
If mConfig.GetSetting(gGetConfigName(GOHLCBar.gWidthProperty)) <> "" Then Width = mConfig.GetSetting(gGetConfigName(GOHLCBar.gWidthProperty))
If mConfig.GetSetting(gGetConfigName(GOHLCBar.gDisplayModeProperty)) <> "" Then DisplayMode = mConfig.GetSetting(gGetConfigName(GOHLCBar.gDisplayModeProperty))
If mConfig.GetSetting(gGetConfigName(GOHLCBar.gLayerProperty)) <> "" Then Layer = mConfig.GetSetting(gGetConfigName(GOHLCBar.gLayerProperty))
If mConfig.GetSetting(gGetConfigName(GOHLCBar.gIsSelectableProperty)) <> "" Then IsSelectable = mConfig.GetSetting(gGetConfigName(GOHLCBar.gIsSelectableProperty))
If mConfig.GetSetting(gGetConfigName(GOHLCBar.gOrientationProperty)) <> "" Then Orientation = mConfig.GetSetting(gGetConfigName(GOHLCBar.gOrientationProperty))
If Not mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gBrushProperty)) Is Nothing Then Brush = LoadIBrushFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gBrushProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gPenProperty)) Is Nothing Then Pen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GOHLCBar.gPenProperty)))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
                
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As String
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function isLocalValueSet(ByVal pExtProp As ExtendedProperty) As Boolean
Const ProcName As String = "isLocalValueSet"
On Error GoTo Err

isLocalValueSet = Not IsEmpty(mEPhost.GetLocalValue(pExtProp))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
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
gHandleUnexpectedError ProcName, ModuleName
End Function






