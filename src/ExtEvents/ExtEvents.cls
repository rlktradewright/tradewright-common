VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtEvents"
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

Public Enum ExtendedEventModes
    ExtendedEventModeNone
    ExtendedEventModeTunnel
    ExtendedEventModeBubble
    ExtendedEventModeDirect
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type ExtendedEventData
    Source          As Object
    OriginalSource  As Object
    Handled         As Boolean
    ExtendedEvent   As ExtendedEvent
    Data            As Variant
End Type

Public Type ExtendedEventListenerEntry
    Listener            As IExtendedEventListener
    NotifyHandledEvents As Boolean
End Type

Public Type ExtendedEventListenerCollectionEntry
    Source          As Object
    Collection      As Collection
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ExtEvents"

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

Public Function RegisterExtendedEvent( _
                ByVal pName As String, _
                ByVal pMode As ExtendedEventModes, _
                ByVal pOwnerClassName As String) As ExtendedEvent
                
Set RegisterExtendedEvent = gRegisterExtendedEvent(pName, pMode, pOwnerClassName)
End Function

'@================================================================================
' Helper Functions
'@================================================================================



