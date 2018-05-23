VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RectangleStyle"
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

Private Const ModuleName                            As String = "RectangleStyle"

Private Const RelatedClassName                      As String = "Rectangle"

'@================================================================================
' Member variables
'@================================================================================

Private mName                                       As String
Private mStyleType                                  As String
Private mConfig                                     As ConfigurationSection

Private mEPhost                                     As ExtendedPropertyHost

Private mBasedOn                                    As RectangleStyle

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mEPhost = New ExtendedPropertyHost
GRectangle.gRegisterProperties
GRectangle.gRegisterExtendedEvents
mName = GenerateGUIDString
mStyleType = RelatedClassName & " Styles"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' IGraphicObjectStyle Interface Members
'@================================================================================

Private Property Let IGraphicObjectStyle_ConfigurationSection(ByVal RHS As ConfigurationSection)
Me.ConfigurationSection = RHS
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

Public Property Let BasedOn(ByVal Value As RectangleStyle)
Set mBasedOn = Value
mEPhost.Style = mBasedOn.ExtendedPropertyHost
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBasedOn, Value.Name
End Property

Public Property Get BasedOn() As RectangleStyle
Set BasedOn = mBasedOn
End Property

Public Property Let Brush(ByVal Value As IBrush)
Const ProcName As String = "Brush"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As IBrush
If setProperty(GRectangle.gBrushProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GRectangle.gBrushProperty))
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

Set Brush = mEPhost.GetLocalValue(GRectangle.gBrushProperty)

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

mConfig.SetSetting ConfigSettingName, mName
mConfig.SetSetting ConfigSettingStyleType, mStyleType

If Not mBasedOn Is Nothing Then mConfig.SetSetting ConfigSettingBasedOn, mBasedOn.Name

If isLocalValueSet(GRectangle.gBrushProperty) Then mEPhost.GetLocalValue(GRectangle.gBrushProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GRectangle.gBrushProperty))
If isLocalValueSet(GRectangle.gIsSelectableProperty) Then mConfig.SetSetting gGetConfigName(GRectangle.gIsSelectableProperty), mEPhost.GetLocalValue(GRectangle.gIsSelectableProperty)
If isLocalValueSet(GRectangle.gLayerProperty) Then mConfig.SetSetting gGetConfigName(GRectangle.gLayerProperty), mEPhost.GetLocalValue(GRectangle.gLayerProperty)
If isLocalValueSet(GRectangle.gOrientationProperty) Then mConfig.SetSetting gGetConfigName(GRectangle.gOrientationProperty), mEPhost.GetLocalValue(GRectangle.gOrientationProperty)
If isLocalValueSet(GRectangle.gPenProperty) Then mEPhost.GetLocalValue(GRectangle.gPenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GRectangle.gPenProperty))
If isLocalValueSet(GRectangle.gPositionProperty) Then mEPhost.GetLocalValue(GRectangle.gPositionProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GRectangle.gPositionProperty))
If isLocalValueSet(GRectangle.gSizeProperty) Then mEPhost.GetLocalValue(GRectangle.gSizeProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GRectangle.gSizeProperty))

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

setProperty GRectangle.gIsSelectableProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GRectangle.gIsSelectableProperty), Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsSelectable() As Boolean
Const ProcName As String = "IsSelectable"
On Error GoTo Err

IsSelectable = mEPhost.GetLocalValue(GRectangle.gIsSelectableProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

setProperty GRectangle.gLayerProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GRectangle.gLayerProperty), Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mEPhost.GetLocalValue(GRectangle.gLayerProperty)

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

setProperty GRectangle.gOrientationProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GRectangle.gOrientationProperty), Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Orientation() As Double
Const ProcName As String = "Orientation"
On Error GoTo Err

Orientation = mEPhost.GetLocalValue(GRectangle.gOrientationProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Pen(ByVal Value As Pen)
Const ProcName As String = "Pen"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GRectangle.gPenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GRectangle.gPenProperty))
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

Set Pen = mEPhost.GetLocalValue(GRectangle.gPenProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Position(ByVal Value As Point)
Const ProcName As String = "Position"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As Point
If setProperty(GRectangle.gPositionProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GRectangle.gPositionProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Position() As Point
Const ProcName As String = "Position"
On Error GoTo Err

Set Position = mEPhost.GetLocalValue(GRectangle.gPositionProperty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Size(ByVal Value As Size)
Const ProcName As String = "Size"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As Size
If setProperty(GRectangle.gSizeProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GRectangle.gSizeProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Size() As Size
Const ProcName As String = "Size"
On Error GoTo Err

Set Size = mEPhost.GetLocalValue(GRectangle.gSizeProperty)

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

clearProperty GRectangle.gBrushProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearLayer()
Const ProcName As String = "ClearLayer"
Dim failpoint As String
On Error GoTo Err

clearProperty GRectangle.gLayerProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearOrientation()
Const ProcName As String = "ClearOrientation"
On Error GoTo Err

clearProperty GRectangle.gOrientationProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearPen()
Const ProcName As String = "ClearPen"
Dim failpoint As String
On Error GoTo Err

clearProperty GRectangle.gPenProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearPosition()
Const ProcName As String = "ClearPosition"
Dim failpoint As String
On Error GoTo Err

clearProperty GRectangle.gPositionProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearSize()
Const ProcName As String = "ClearSize"
Dim failpoint As String
On Error GoTo Err

clearProperty GRectangle.gSizeProperty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Clone() As RectangleStyle
Const ProcName As String = "Clone"
Dim failpoint As String
On Error GoTo Err

Set Clone = New RectangleStyle
Clone.BasedOn = mBasedOn
If isLocalValueSet(GRectangle.gBrushProperty) Then Clone.Brush = mEPhost.GetLocalValue(GRectangle.gBrushProperty)
If isLocalValueSet(GRectangle.gIsSelectableProperty) Then Clone.IsSelectable = mEPhost.GetLocalValue(GRectangle.gIsSelectableProperty)
If isLocalValueSet(GRectangle.gLayerProperty) Then Clone.Layer = mEPhost.GetLocalValue(GRectangle.gLayerProperty)
If isLocalValueSet(GRectangle.gOrientationProperty) Then Clone.Orientation = mEPhost.GetLocalValue(GRectangle.gOrientationProperty)
If isLocalValueSet(GRectangle.gPenProperty) Then Clone.Pen = mEPhost.GetLocalValue(GRectangle.gPenProperty)
If isLocalValueSet(GRectangle.gPositionProperty) Then Clone.Position = mEPhost.GetLocalValue(GRectangle.gPositionProperty)
If isLocalValueSet(GRectangle.gSizeProperty) Then Clone.Size = mEPhost.GetLocalValue(GRectangle.gSizeProperty)

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

If mConfig.GetSetting(gGetConfigName(GRectangle.gLayerProperty)) <> "" Then Layer = mConfig.GetSetting(gGetConfigName(GRectangle.gLayerProperty))
If mConfig.GetSetting(gGetConfigName(GRectangle.gIsSelectableProperty)) <> "" Then IsSelectable = mConfig.GetSetting(gGetConfigName(GRectangle.gIsSelectableProperty))
If mConfig.GetSetting(gGetConfigName(GRectangle.gOrientationProperty)) <> "" Then Orientation = mConfig.GetSetting(gGetConfigName(GRectangle.gOrientationProperty))
If Not mConfig.GetConfigurationSection(gGetConfigName(GRectangle.gBrushProperty)) Is Nothing Then Brush = LoadIBrushFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GRectangle.gBrushProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GRectangle.gPenProperty)) Is Nothing Then Pen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GRectangle.gPenProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GRectangle.gPositionProperty)) Is Nothing Then Position = LoadPointFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GRectangle.gPositionProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GRectangle.gSizeProperty)) Is Nothing Then Size = LoadSizeFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GRectangle.gSizeProperty)))

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

isLocalValueSet = mEPhost.IsPropertySet(pExtProp)

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




