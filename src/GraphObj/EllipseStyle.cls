VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EllipseStyle"
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

Private Const ModuleName                            As String = "EllipseStyle"

Private Const RelatedClassName                      As String = "Ellipse"

'@================================================================================
' Member variables
'@================================================================================

Private mName                                       As String
Private mStyleType                                  As String
Private mConfig                                     As ConfigurationSection

Private mEPhost                                     As ExtendedPropertyHost

Private mBasedOn                                    As EllipseStyle

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mEPhost = New ExtendedPropertyHost

GEllipse.gRegisterProperties
GEllipse.gRegisterExtendedEvents

mName = GenerateGUIDString
mStyleType = RelatedClassName & " Styles"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
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

Public Property Let BasedOn(ByVal Value As EllipseStyle)
Set mBasedOn = Value
mEPhost.Style = mBasedOn.ExtendedPropertyHost
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBasedOn, Value.Name
End Property

Public Property Get BasedOn() As EllipseStyle
Set BasedOn = mBasedOn
End Property

Public Property Let Brush(ByVal Value As IBrush)
Const ProcName As String = "Brush"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As IBrush
If setProperty(GEllipse.gBrushProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GEllipse.gBrushProperty))
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

Set Brush = mEPhost.GetLocalValue(GEllipse.gBrushProperty)

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

mConfig.SetSetting ConfigSettingName, mName
mConfig.SetSetting ConfigSettingStyleType, mStyleType

If Not mBasedOn Is Nothing Then mConfig.SetSetting ConfigSettingBasedOn, mBasedOn.Name

If isLocalValueSet(GEllipse.gBrushProperty) Then mEPhost.GetLocalValue(GEllipse.gBrushProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GEllipse.gBrushProperty))
If isLocalValueSet(GEllipse.gIsSelectableProperty) Then mConfig.SetSetting gGetConfigName(GEllipse.gIsSelectableProperty), mEPhost.GetLocalValue(GEllipse.gIsSelectableProperty)
If isLocalValueSet(GEllipse.gLayerProperty) Then mConfig.SetSetting gGetConfigName(GEllipse.gLayerProperty), mEPhost.GetLocalValue(GEllipse.gLayerProperty)
If isLocalValueSet(GEllipse.gOrientationProperty) Then mConfig.SetSetting gGetConfigName(GEllipse.gOrientationProperty), mEPhost.GetLocalValue(GEllipse.gOrientationProperty)
If isLocalValueSet(GEllipse.gPenProperty) Then mEPhost.GetLocalValue(GEllipse.gPenProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GEllipse.gPenProperty))
If isLocalValueSet(GEllipse.gPositionProperty) Then mEPhost.GetLocalValue(GEllipse.gPositionProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GEllipse.gPositionProperty))
If isLocalValueSet(GEllipse.gSizeProperty) Then mEPhost.GetLocalValue(GEllipse.gSizeProperty).ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GEllipse.gSizeProperty))

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

setProperty GEllipse.gIsSelectableProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GEllipse.gIsSelectableProperty), Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get IsSelectable() As Boolean
Const ProcName As String = "IsSelectable"
On Error GoTo Err

IsSelectable = mEPhost.GetLocalValue(GEllipse.gIsSelectableProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
Const ProcName As String = "Layer"
Dim failpoint As String
On Error GoTo Err

setProperty GEllipse.gLayerProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GEllipse.gLayerProperty), Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Layer() As LayerNumbers
Const ProcName As String = "Layer"
On Error GoTo Err

Layer = mEPhost.GetLocalValue(GEllipse.gLayerProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Name() As String
Name = mName
End Property

Public Property Let Orientation(ByVal Value As Double)
Const ProcName As String = "Orientation"
Dim failpoint As String
On Error GoTo Err

setProperty GEllipse.gOrientationProperty, Value
If Not mConfig Is Nothing Then mConfig.SetSetting gGetConfigName(GEllipse.gOrientationProperty), Value

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Orientation() As Double
Const ProcName As String = "Orientation"
On Error GoTo Err

Orientation = mEPhost.GetLocalValue(GEllipse.gOrientationProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Pen(ByVal Value As Pen)
Const ProcName As String = "Pen"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As Pen
If setProperty(GEllipse.gPenProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GEllipse.gPenProperty))
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

Set Pen = mEPhost.GetLocalValue(GEllipse.gPenProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Position(ByVal Value As Point)
Const ProcName As String = "Position"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As Point
If setProperty(GEllipse.gPositionProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GEllipse.gPositionProperty))
        If Not prevValue Is Nothing Then prevValue.RemoveFromConfig
    End If
End If

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Position() As Point
Const ProcName As String = "Position"
On Error GoTo Err

Set Position = mEPhost.GetLocalValue(GEllipse.gPositionProperty)

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Let Size(ByVal Value As Size)
Const ProcName As String = "Size"
Dim failpoint As String
On Error GoTo Err

Dim prevValue As Size
If setProperty(GEllipse.gSizeProperty, Value, prevValue) Then
    If Not mConfig Is Nothing Then
        Value.ConfigurationSection = mConfig.AddConfigurationSection(gGetConfigName(GEllipse.gSizeProperty))
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

Set Size = mEPhost.GetLocalValue(GEllipse.gSizeProperty)

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

clearProperty GEllipse.gBrushProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearLayer()
Const ProcName As String = "ClearLayer"
Dim failpoint As String
On Error GoTo Err

clearProperty GEllipse.gLayerProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearOrientation()
Const ProcName As String = "ClearOrientation"
On Error GoTo Err

clearProperty GEllipse.gOrientationProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearPen()
Const ProcName As String = "ClearPen"
Dim failpoint As String
On Error GoTo Err

clearProperty GEllipse.gPenProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearPosition()
Const ProcName As String = "ClearPosition"
Dim failpoint As String
On Error GoTo Err

clearProperty GEllipse.gPositionProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub ClearSize()
Const ProcName As String = "ClearSize"
Dim failpoint As String
On Error GoTo Err

clearProperty GEllipse.gSizeProperty

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Function Clone() As EllipseStyle
Const ProcName As String = "Clone"
Dim failpoint As String
On Error GoTo Err

Set Clone = New EllipseStyle
Clone.BasedOn = mBasedOn
If isLocalValueSet(GEllipse.gBrushProperty) Then Clone.Brush = mEPhost.GetLocalValue(GEllipse.gBrushProperty)
If isLocalValueSet(GEllipse.gIsSelectableProperty) Then Clone.IsSelectable = mEPhost.GetLocalValue(GEllipse.gIsSelectableProperty)
If isLocalValueSet(GEllipse.gLayerProperty) Then Clone.Layer = mEPhost.GetLocalValue(GEllipse.gLayerProperty)
If isLocalValueSet(GEllipse.gOrientationProperty) Then Clone.Orientation = mEPhost.GetLocalValue(GEllipse.gOrientationProperty)
If isLocalValueSet(GEllipse.gPenProperty) Then Clone.Pen = mEPhost.GetLocalValue(GEllipse.gPenProperty)
If isLocalValueSet(GEllipse.gPositionProperty) Then Clone.Position = mEPhost.GetLocalValue(GEllipse.gPositionProperty)
If isLocalValueSet(GEllipse.gSizeProperty) Then Clone.Size = mEPhost.GetLocalValue(GEllipse.gSizeProperty)

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

If mConfig.GetSetting(gGetConfigName(GEllipse.gLayerProperty)) <> "" Then Layer = mConfig.GetSetting(gGetConfigName(GEllipse.gLayerProperty))
If mConfig.GetSetting(gGetConfigName(GEllipse.gIsSelectableProperty)) <> "" Then IsSelectable = mConfig.GetSetting(gGetConfigName(GEllipse.gIsSelectableProperty))
If mConfig.GetSetting(gGetConfigName(GEllipse.gOrientationProperty)) <> "" Then Orientation = mConfig.GetSetting(gGetConfigName(GEllipse.gOrientationProperty))
If Not mConfig.GetConfigurationSection(gGetConfigName(GEllipse.gBrushProperty)) Is Nothing Then Brush = LoadIBrushFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GEllipse.gBrushProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GEllipse.gPenProperty)) Is Nothing Then Pen = LoadPenFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GEllipse.gPenProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GEllipse.gPositionProperty)) Is Nothing Then Position = LoadPointFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GEllipse.gPositionProperty)))
If Not mConfig.GetConfigurationSection(gGetConfigName(GEllipse.gSizeProperty)) Is Nothing Then Size = LoadSizeFromConfig(mConfig.GetConfigurationSection(gGetConfigName(GEllipse.gSizeProperty)))

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

isLocalValueSet = mEPhost.IsPropertySet(pExtProp)

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






