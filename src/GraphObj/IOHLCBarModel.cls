VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOHLCBarModel"
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

Private Const ModuleName                            As String = "IOHLCBarModel"

Public Property Let CloseValue(ByVal Value As Double)

End Property

Public Property Get CloseValue() As Double

End Property

Public Property Let HighValue(ByVal Value As Double)

End Property

Public Property Get HighValue() As Double

End Property

Public Property Let LowValue(ByVal Value As Double)

End Property

Public Property Get LowValue() As Double

End Property

Public Property Let OpenValue(ByVal Value As Double)

End Property

Public Property Get OpenValue() As Double

End Property

Public Property Let X(ByVal Value As Double)

End Property

Public Property Get X() As Double

End Property

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

Public Sub AddChangeListener(ByVal pListener As IChangeListener)

End Sub

Public Sub RemoveChangeListener(ByVal pListener As IChangeListener)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================





