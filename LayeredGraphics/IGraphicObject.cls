VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGraphicObject"
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

Private Const ModuleName                    As String = "IGraphicObject"

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

Public Property Get BoundingRectangleInArea(ByRef pArea As TRectangle) As TRectangle

End Property

Public Property Get GaugeDependent() As Boolean

End Property

Public Property Let Handle(ByVal Value As Long)

End Property

Public Property Get Handle() As Long

End Property

Public Property Get IncludeInAutoscale() As Boolean

End Property

Public Property Get IsSelectable() As Boolean

End Property

Public Property Let IsSelected(ByVal Value As Boolean)

End Property

Public Property Get IsSelected() As Boolean

End Property

Public Property Get Layer() As LayerNumbers

End Property

Public Property Get NoDraw() As Boolean

End Property

Public Property Get SeriesID() As Long

End Property

Public Property Get VerticalRangeInXinterval(ByRef pXInterval As TInterval) As TInterval

End Property

Public Property Let Visible(ByVal Value As Boolean)

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Click()

End Sub

Public Sub DblCLick()

End Sub

Public Sub Draw(ByRef pArea As TRectangle)

End Sub

Public Sub Finish()

End Sub

Public Function HitTest( _
                ByVal pX As Double, _
                ByVal pY As Double) As Boolean
                
End Function

Public Sub MouseEnter(ByVal pButton As Integer, ByVal pShift As Integer, ByVal pX As Single, ByVal pY As Single)

End Sub

Public Sub MouseLeave(ByVal pButton As Integer, ByVal pShift As Integer, ByVal pX As Single, ByVal pY As Single)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================




