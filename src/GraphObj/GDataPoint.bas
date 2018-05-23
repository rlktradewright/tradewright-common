Attribute VB_Name = "GDataPoint"
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

' TODO: correct the enum and value names and add/remove any others needed
Public Enum DataPointChangeTypes
    DataPointValueChanged
    DataPointPositionChanged
    DataPointPositionCleared
    DataPointSizeChanged
    DataPointSizeCleared
    DataPointOrientationChanged
    DataPointOrientationCleared
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

' TODO: replace DataPoint with the related class name
Private Const ModuleName                            As String = "GDataPoint"

'@================================================================================
' Member variables
'@================================================================================

' TODO: add/remove as required
Public gMouseEnterEvent                             As ExtendedEvent
Public gMouseLeaveEvent                             As ExtendedEvent

Public gBrushProperty                               As ExtendedProperty
Public gDisplayModeProperty                         As ExtendedProperty
Public gDownBrushProperty                           As ExtendedProperty
Public gDownLinePenProperty                         As ExtendedProperty
Public gDownPenProperty                             As ExtendedProperty
Public gHistogramBarWidthProperty                   As ExtendedProperty
Public gHistogramBaselinePenProperty                As ExtendedProperty
Public gHistogramBaseValueProperty                  As ExtendedProperty
Public gIncludeInAutoscaleProperty                  As ExtendedProperty
Public gIsSelectableProperty                        As ExtendedProperty
Public gIsSelectedProperty                          As ExtendedProperty
Public gIsVisibleProperty                           As ExtendedProperty
Public gLayerProperty                               As ExtendedProperty
Public gLineModeProperty                            As ExtendedProperty
Public gLinePenProperty                             As ExtendedProperty
Public gNumberOfSidesProperty                       As ExtendedProperty
Public gOrientationProperty                         As ExtendedProperty
Public gPenProperty                                 As ExtendedProperty
Public gSizeProperty                                As ExtendedProperty
Public gUpBrushProperty                             As ExtendedProperty
Public gUpLinePenProperty                           As ExtendedProperty
Public gUpPenProperty                               As ExtendedProperty
Public gValueProperty                               As ExtendedProperty
Public gXProperty                                   As ExtendedProperty

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Sub gRegisterExtendedEvents()
Static sRegistered As Boolean
If sRegistered Then Exit Sub
sRegistered = True

Set gMouseEnterEvent = RegisterExtendedEvent("MouseEnter", ExtendedEventModeBubble, "Rectangle")
Set gMouseLeaveEvent = RegisterExtendedEvent("MouseLeave", ExtendedEventModeBubble, "Rectangle")
End Sub

' TODO: amend as required
Public Sub gRegisterProperties()
Static sRegistered As Boolean
If sRegistered Then Exit Sub
sRegistered = True

RegisterGraphicObjectExtProperty gDownLinePenProperty, _
                pName:="DownLinePen", _
                pType:=vbObject, _
                pTypename:="Pen", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=Nothing, _
                pValidatorPointer:=AddressOf gValidatePen

RegisterGraphicObjectExtProperty gUpLinePenProperty, _
                pName:="UpLinePen", _
                pType:=vbObject, _
                pTypename:="Pen", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=Nothing, _
                pValidatorPointer:=AddressOf gValidatePen

RegisterGraphicObjectExtProperty gLinePenProperty, _
                pName:="LinePen", _
                pType:=vbObject, _
                pTypename:="Pen", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=CreatePixelPen(vbBlack, , LineInsideSolid), _
                pValidatorPointer:=AddressOf gValidatePen

RegisterGraphicObjectExtProperty gHistogramBaselinePenProperty, _
                pName:="HistogramBaselinePen", _
                pType:=vbObject, _
                pTypename:="Pen", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=Nothing, _
                pValidatorPointer:=AddressOf gValidatePen

RegisterGraphicObjectExtProperty gHistogramBaseValueProperty, _
                pName:="HistogramBaseValue", _
                pType:=vbDouble, _
                pTypename:="", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=0#

RegisterGraphicObjectExtProperty gLineModeProperty, _
                pName:="LineMode", _
                pType:=vbLong, _
                pTypename:="DataPointLineModes", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=DataPointLineModes.DataPointLineModeStraight, _
                pValidatorPointer:=AddressOf gValidateLineMode

RegisterGraphicObjectExtProperty gHistogramBarWidthProperty, _
                pName:="HistogramBarWidth", _
                pType:=vbDouble, _
                pTypename:="", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=0.6, _
                pValidatorPointer:=AddressOf gValidateHistogramBarWidth

RegisterGraphicObjectExtProperty gDownPenProperty, _
                pName:="DownPen", _
                pType:=vbObject, _
                pTypename:="Pen", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=Nothing, _
                pValidatorPointer:=AddressOf gValidatePen

RegisterGraphicObjectExtProperty gUpPenProperty, _
                pName:="UpPen", _
                pType:=vbObject, _
                pTypename:="Pen", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=Nothing, _
                pValidatorPointer:=AddressOf gValidatePen

RegisterGraphicObjectExtProperty gDownBrushProperty, _
                pName:="DownBrush", _
                pType:=vbObject, _
                pTypename:="IBrush", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=Nothing, _
                pValidatorPointer:=AddressOf gValidateBrush

RegisterGraphicObjectExtProperty gUpBrushProperty, _
                pName:="UpBrush", _
                pType:=vbObject, _
                pTypename:="IBrush", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=Nothing, _
                pValidatorPointer:=AddressOf gValidateBrush

RegisterGraphicObjectExtProperty gNumberOfSidesProperty, _
                pName:="NumberOfSides", _
                pType:=vbLong, _
                pTypename:="", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=4, _
                pValidatorPointer:=AddressOf gValidatePolygonNumberOfSides

RegisterGraphicObjectExtProperty gDisplayModeProperty, _
                pName:="DisplayMode", _
                pType:=vbLong, _
                pTypename:="DataPointDisplayModes", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=DataPointDisplayModes.DataPointDisplayModeNone, _
                pValidatorPointer:=AddressOf gValidateDisplayMode

RegisterGraphicObjectExtProperty gValueProperty, _
                pName:="Value", _
                pType:=vbDouble, _
                pTypename:="", _
                pAffectsPaintingRegion:=True, _
                pAffectsPosition:=True, _
                pAffectsSize:=False, _
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
                pDefaultValue:=GetBrush, _
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
                pDefaultValue:=CreatePixelPen(&HC0C0C0), _
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
                
RegisterGraphicObjectExtProperty gSizeProperty, _
                pName:="Size", _
                pType:=vbObject, _
                pTypename:="Size", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=NewSize(0.2, 0.2, ScaleUnitCm, ScaleUnitCm), _
                pValidatorPointer:=AddressOf gValidateSize
                
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Public Sub gValidateDisplayMode(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidateDisplayMode"
On Error GoTo Err

Dim lValue As DataPointDisplayModes

lValue = pValue

Select Case lValue
Case DataPointDisplayModeNone
Case DataPointDisplayModePoint
Case DataPointDisplayModeDash
Case DataPointDisplayModeHistogram
Case DataPointDisplayModePolygon
Case DataPointDisplayModeEllipse
Case DataPointDisplayModeCross
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a member of the DataPointDisplayModes enum"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gValidateHistogramBarWidth(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidateHistogramBarWidth"
On Error GoTo Err

If pValue <= 0 Or pValue > 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "HistogramBarWidth must be greater than zero but not greater than 1"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gValidateLineMode(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidateLineMode"
On Error GoTo Err

Dim lValue As DataPointLineModes

lValue = pValue

Select Case lValue
Case DataPointLineModeNone
Case DataPointLineModeStraight
Case DataPointLineModeStepped
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a member of the DataPointLineModes enum"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

