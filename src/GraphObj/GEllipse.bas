Attribute VB_Name = "GEllipse"

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

Public Enum EllipseChangeTypes
    EllipsePositionChanged
    EllipsePositionCleared
    EllipseSizeChanged
    EllipseSizeCleared
    EllipseOrientationChanged
    EllipseOrientationCleared
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GEllipse"

'@================================================================================
' Member variables
'@================================================================================

Public gMouseEnterEvent                             As ExtendedEvent
Public gMouseLeaveEvent                             As ExtendedEvent

Public gBrushProperty                               As ExtendedProperty
Public gIncludeInAutoscaleProperty                  As ExtendedProperty
Public gIsSelectableProperty                        As ExtendedProperty
Public gIsSelectedProperty                          As ExtendedProperty
Public gIsVisibleProperty                           As ExtendedProperty
Public gLayerProperty                               As ExtendedProperty
Public gOrientationProperty                         As ExtendedProperty
Public gPenProperty                                 As ExtendedProperty
Public gPositionProperty                            As ExtendedProperty
Public gSizeProperty                                As ExtendedProperty

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
Set gMouseEnterEvent = RegisterExtendedEvent("MouseEnter", ExtendedEventModeBubble, "Rectangle")
Set gMouseLeaveEvent = RegisterExtendedEvent("MouseLeave", ExtendedEventModeBubble, "Rectangle")
End Sub

' TODO: amend as required
Public Sub gRegisterProperties()
Static sRegistered As Boolean
If sRegistered Then Exit Sub
sRegistered = True

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
                
RegisterGraphicObjectExtProperty gPositionProperty, _
                pName:="Position", _
                pType:=vbObject, _
                pTypename:="Point", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=True, _
                pAffectsSize:=False, _
                pAffectsRender:=True, _
                pDefaultValue:=NewPoint(0, 0), _
                pValidatorPointer:=AddressOf gValidatePosition
                
RegisterGraphicObjectExtProperty gSizeProperty, _
                pName:="Size", _
                pType:=vbObject, _
                pTypename:="Size", _
                pAffectsPaintingRegion:=False, _
                pAffectsPosition:=False, _
                pAffectsSize:=True, _
                pAffectsRender:=True, _
                pDefaultValue:=NewSize(1, 1), _
                pValidatorPointer:=AddressOf gValidateSize
                
End Sub

'@================================================================================
' Helper Functions
'@================================================================================





