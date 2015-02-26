VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphObjUtils"
Attribute VB_GlobalNameSpace = True
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

Private Const ModuleName                            As String = "GraphObjUtils"

'@================================================================================
' Member variables
'@================================================================================

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

Public Sub RegisterGraphicObjectExtProperty( _
                ByRef pExtProp As ExtendedProperty, _
                ByVal pName As String, _
                ByVal pType As VbVarType, _
                ByVal pTypename As String, _
                ByVal pAffectsPaintingRegion As Boolean, _
                ByVal pAffectsPosition As Boolean, _
                ByVal pAffectsSize As Boolean, _
                ByVal pAffectsRender As Boolean, _
                Optional ByVal pDefaultValue As Variant, _
                Optional ByVal pValidatorPointer As Long, _
                Optional ByVal pConfigName As String)
If pExtProp Is Nothing Then
    Dim lMetadata As New GraphicExtPropMetadata
    Set pExtProp = RegisterExtendedProperty(pName, _
                                        pType, _
                                        pTypename, _
                                        pDefaultValue, _
                                        lMetadata, _
                                        pValidatorPointer)
lMetadata.SetAffectsPaintingRegion(pAffectsPaintingRegion) _
                    .SetAffectsPosition(pAffectsPosition) _
                    .SetAffectsSize(pAffectsSize) _
                    .SetAffectsRender(pAffectsRender) _
                    .SetConfigName (IIf(pConfigName <> "", pConfigName, IIf(pExtProp.ValueIsObject, "", "&") & pExtProp.PropertyName))
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




