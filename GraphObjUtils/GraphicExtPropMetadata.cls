VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicExtPropMetadata"
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

Private Const ModuleName                            As String = "GraphicExtPropMetadata"

'@================================================================================
' Member variables
'@================================================================================

Private mAffectsSize                                As Boolean
Private mAffectsPosition                            As Boolean
Private mAffectsPaintingRegion                      As Boolean
Private mAffectsRender                              As Boolean
Private mConfigName                                 As String

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

Public Property Get AffectsPaintingRegion() As Boolean
AffectsPaintingRegion = mAffectsPaintingRegion
End Property

Public Property Get AffectsPosition() As Boolean
AffectsPosition = mAffectsPosition
End Property

Public Property Get AffectsSize() As Boolean
AffectsSize = mAffectsSize
End Property

Public Property Get AffectsRender() As Boolean
AffectsRender = mAffectsRender
End Property

Public Property Get ConfigName() As String
ConfigName = mConfigName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function SetAffectsPaintingRegion(ByVal Value As Boolean) As GraphicExtPropMetadata
mAffectsPaintingRegion = Value
Set SetAffectsPaintingRegion = Me
End Function

Public Function SetAffectsPosition(ByVal Value As Boolean) As GraphicExtPropMetadata
mAffectsPosition = Value
Set SetAffectsPosition = Me
End Function

Public Function SetAffectsRender(ByVal Value As Boolean) As GraphicExtPropMetadata
mAffectsRender = Value
Set SetAffectsRender = Me
End Function

Public Function SetAffectsSize(ByVal Value As Boolean) As GraphicExtPropMetadata
mAffectsSize = Value
Set SetAffectsSize = Me
End Function

Public Function SetConfigName(ByVal Value As String) As GraphicExtPropMetadata
mConfigName = Value
Set SetConfigName = Me
End Function

'@================================================================================
' Helper Functions
'@================================================================================




