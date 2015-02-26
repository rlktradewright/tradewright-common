VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtendedEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "ExtendedEvent"

'@================================================================================
' Member variables
'@================================================================================

Private mName                                       As String
Private mMode                                       As ExtendedEventModes

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

Public Property Get Mode() As ExtendedEventModes
Mode = mMode
End Property

Public Property Get Name() As String
Name = mName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddExtendedEventListener( _
                ByVal pSource As IExtendedEventsSource, _
                ByVal pListener As IExtendedEventListener, _
                Optional ByVal pNotifyHandledEvents As Boolean = False)
gAddListener pSource, pListener, Me, pNotifyHandledEvents
End Sub

Public Sub Fire( _
                ByVal pSource As IExtendedEventsSource, _
                ByVal pValue As Variant)
gFire pSource, pValue, Me

End Sub

Friend Sub Initialise( _
                ByVal pName As String, _
                ByVal pMode As ExtendedEventModes)
mName = pName
mMode = pMode
End Sub

Public Sub RemoveExtendedEventListener( _
                ByVal pSource As IExtendedEventsSource, _
                ByVal pListener As IExtendedEventListener)
gRemoveListener pSource, pListener, Me
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




