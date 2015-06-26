Attribute VB_Name = "GOHLCBar"
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

Public Enum OHLCBarChangeTypes
    OHLCBarCloseValueChanged
    OHLCBarLowValueChanged
    OHLCBarHighValueChanged
    OHLCBarOpenValueChanged
    OHLCBarXChanged
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

' TODO: replace OHLCBar with the related class name
Private Const ModuleName                            As String = "GOHLCBar"

'@================================================================================
' Member variables
'@================================================================================

Public gMouseEnterEvent                             As ExtendedEvent
Public gMouseLeaveEvent                             As ExtendedEvent

Public gBrushProperty                               As ExtendedProperty
Public gCloseValueProperty                          As ExtendedProperty
Public gDisplayModeProperty                         As ExtendedProperty
Public gDownBrushProperty                           As ExtendedProperty
Public gDownPenProperty                             As ExtendedProperty
Public gHighValueProperty                           As ExtendedProperty
Public gIncludeInAutoscaleProperty                  As ExtendedProperty
Public gIsSelectableProperty                        As ExtendedProperty
Public gIsSelectedProperty                          As ExtendedProperty
Public gIsVisibleProperty                           As ExtendedProperty
Public gLayerProperty                               As ExtendedProperty
Public gLowValueProperty                            As ExtendedProperty
Public gOpenValueProperty                           As ExtendedProperty
Public gOrientationProperty                         As ExtendedProperty
Public gPenProperty                                 As ExtendedProperty
Public gUpBrushProperty                             As ExtendedProperty
Public gUpPenProperty                               As ExtendedProperty
Public gWidthProperty                               As ExtendedProperty
Public gXProperty                                   As ExtendedProperty

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' OHLCBar Interface Members
'@================================================================================

'@================================================================================
' OHLCBar Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub gRegisterExtendedEvents()
Set gMouseEnterEvent = RegisterExtendedEvent("MouseEnter", ExtendedEventModeBubble, "Rectangle")
Set gMouseLeaveEvent = RegisterExtendedEvent("MouseLeave", ExtendedEventModeBubble, "Rectangle")
End Sub

' TODO: amend as required
Public Sub gRegisterProperties()
Static sRegistered As Boolean
If sRegistered Then Exit Sub
sRegistered = True

RegisterGraphicObjectExtProperty gWidthProperty, _
                pName:="Width", _
                pType:=vbDouble, _
                pTypename:="", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=False, _
                pDefaultValue:=0.6, _
                pValidatorPointer:=AddressOf gValidateWidth

RegisterGraphicObjectExtProperty gUpPenProperty, _
                pName:="UpPen", _
                pType:=vbObject, _
                pTypename:="Pen", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=CreatePixelPen(vbBlack, , LineInsideSolid), _
                pValidatorPointer:=AddressOf gValidatePen

RegisterGraphicObjectExtProperty gDownPenProperty, _
                pName:="DownPen", _
                pType:=vbObject, _
                pTypename:="Pen", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=CreatePixelPen(vbBlack, , LineInsideSolid), _
                pValidatorPointer:=AddressOf gValidatePen

RegisterGraphicObjectExtProperty gUpBrushProperty, _
                pName:="UpBrush", _
                pType:=vbObject, _
                pTypename:="IBrush", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=CreateBrush(vbBlack), _
                pValidatorPointer:=AddressOf gValidateBrush

RegisterGraphicObjectExtProperty gDownBrushProperty, _
                pName:="DownBrush", _
                pType:=vbObject, _
                pTypename:="IBrush", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=CreateBrush(vbBlack), _
                pValidatorPointer:=AddressOf gValidateBrush

RegisterGraphicObjectExtProperty gDisplayModeProperty, _
                pName:="DisplayMode", _
                pType:=vbLong, _
                pTypename:="OHLCBarDisplayModes", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=False, _
                pDefaultValue:=OHLCBarDisplayModes.OHLCBarDisplayModeBar, _
                pValidatorPointer:=AddressOf gValidateDisplayMode

RegisterGraphicObjectExtProperty gCloseValueProperty, _
                pName:="CloseValue", _
                pType:=vbDouble, _
                pTypename:="", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=False, _
                pDefaultValue:=MaxDouble

RegisterGraphicObjectExtProperty gLowValueProperty, _
                pName:="LowValue", _
                pType:=vbDouble, _
                pTypename:="", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=False, _
                pDefaultValue:=MaxDouble

RegisterGraphicObjectExtProperty gHighValueProperty, _
                pName:="HighValue", _
                pType:=vbDouble, _
                pTypename:="", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=False, _
                pDefaultValue:=MaxDouble

RegisterGraphicObjectExtProperty gOpenValueProperty, _
                pName:="OpenValue", _
                pType:=vbDouble, _
                pTypename:="", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=False, _
                pDefaultValue:=MaxDouble

RegisterGraphicObjectExtProperty gXProperty, _
                pName:="X", _
                pType:=vbDouble, _
                pTypename:="", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=True, _
                pAffectsSize:=False, _
                pAffectsRender:=False, _
                pDefaultValue:=0#

RegisterGraphicObjectExtProperty gBrushProperty, _
                pName:="Brush", _
                pType:=vbObject, _
                pTypename:="IBrush", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=Nothing, _
                pValidatorPointer:=AddressOf gValidateBrush

RegisterGraphicObjectExtProperty gIncludeInAutoscaleProperty, _
                pName:="IncludeInAutoscale", _
                pType:=vbBoolean, _
                pTypename:=TypeName(True), _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=True

RegisterGraphicObjectExtProperty gIsSelectableProperty, _
                pName:="IsSelectable", _
                pType:=vbBoolean, _
                pTypename:=TypeName(True), _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=False, _
                pDefaultValue:=False
                
RegisterGraphicObjectExtProperty gIsSelectedProperty, _
                pName:="IsSelected", _
                pType:=vbBoolean, _
                pTypename:=TypeName(True), _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=False, _
                pDefaultValue:=False

RegisterGraphicObjectExtProperty gPenProperty, _
                pName:="Pen", _
                pType:=vbObject, _
                pTypename:="Pen", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=Nothing, _
                pValidatorPointer:=AddressOf gValidatePen

RegisterGraphicObjectExtProperty gIsVisibleProperty, _
                pName:="IsVisible", _
                pType:=vbBoolean, _
                pTypename:=TypeName(True), _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=False

RegisterGraphicObjectExtProperty gLayerProperty, _
                pName:="Layer", _
                pType:=vbLong, _
                pTypename:=TypeName(1&), _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=LayerNumbers.LayerLowestUser, _
                pValidatorPointer:=AddressOf gValidateLayer
                
RegisterGraphicObjectExtProperty gOrientationProperty, _
                pName:="Orientation", _
                pType:=vbDouble, _
                pTypename:=TypeName(1#), _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=True, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=0#
                
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





Public Sub gValidateDisplayMode(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidateDisplayMode"
On Error GoTo Err

Dim lValue As OHLCBarDisplayModes

lValue = pValue

Select Case lValue
Case OHLCBarDisplayModeBar
Case OHLCBarDisplayModeCandlestick
Case OHLCBarDisplayModeLine
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a member of the OHLCBarDisplayModes enum"
End Select

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub gValidateWidth(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidateWidth"
On Error GoTo Err

If pValue <= 0# Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value must be positive"

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

