VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDataPointModel"
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

' TODO: replace XXX with the related graphic object class name
Private Const ModuleName                            As String = "IDataPointModel"

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

' TODO: add custom properties

Public Property Let Orientation(ByVal Value As Double)

End Property

Public Property Get Orientation() As Double

End Property

Public Property Let Position(ByVal Value As Point)

End Property

Public Property Get Position() As Point

End Property

Public Property Let Size(ByVal Value As Size)

End Property

Public Property Get Size() As Size

End Property

Public Property Let Value(ByVal Value As Double)

End Property

Public Property Get Value() As Double

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddChangeListener(ByVal pListener As IChangeListener)

End Sub

Public Sub ClearOrientation()

End Sub

Public Sub ClearPosition()

End Sub

Public Sub ClearSize()

End Sub

Public Sub RemoveChangeListener(ByVal pListener As IChangeListener)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================





