VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISprite"
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

Private Const ModuleName                            As String = "ISprite"

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

Public Property Get BoundingRectangle() As TRectangle
End Property

Public Property Let Brush(ByVal Value As IBrush)
End Property

Public Property Get Brush() As IBrush
End Property

Public Property Let IncludeInAutoscale(ByVal Value As Boolean)
End Property

Public Property Get IncludeInAutoscale() As Boolean
End Property

Public Property Let IsSelectable(ByVal Value As Boolean)
End Property

Public Property Get IsSelectable() As Boolean
End Property

Public Property Get IsSelected() As Boolean
End Property

Public Property Get IsVisible() As Boolean
End Property

Public Property Let Layer(ByVal Value As LayerNumbers)
End Property

Public Property Get Layer() As LayerNumbers
End Property

Public Property Let Orientation(ByVal Value As Double)
End Property

Public Property Get Orientation() As Double
End Property

Public Property Let Pen(ByVal Value As Pen)
End Property

Public Property Get Pen() As Pen
End Property

Public Property Let Position(ByVal Value As Point)
End Property

Public Property Get Position() As Point
End Property

Public Property Let Size(ByVal Value As Size)
End Property

Public Property Get Size() As Size
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ApplyChanges()
End Sub

Public Sub Edit()
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




