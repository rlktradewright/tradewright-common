VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RectangleDefaultRenderer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IRenderer

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

Private Const ModuleName                            As String = "RectangleDefaultRenderer"

'@================================================================================
' Member variables
'@================================================================================

Private mGraphics                                   As Graphics

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IRenderer Interface Members
'@================================================================================

Private Function IRenderer_GetBoundingRectangle(ByVal pObject As IGraphicObject) As TRectangle
Const ProcName As String = "IRenderer_GetBoundingRectangle"
On Error GoTo Err

Dim lRectangle As Rectangle

Set lRectangle = pObject

IRenderer_GetBoundingRectangle = determineBoundingRect(lRectangle)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IRenderer_Render(ByVal pObject As IGraphicObject)
Const ProcName As String = "IRenderer_Render"
On Error GoTo Err

Dim lRectangle As Rectangle

Set lRectangle = pObject

' uncomment the following line during development to make the bounding rectangle
' visible in the display
'mGraphics.FillRectangleFromTRectangle GetBrush(&HF0F0F0), pObject.BoundingRectangle

paintRectangle lRectangle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Graphics(ByVal Value As Graphics)
Set mGraphics = Value
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Function determineBoundingRect(ByVal pRectangle As Rectangle) As TRectangle
Const ProcName As String = "determineBoundingRect"
On Error GoTo Err

mGraphics.BeginBoundingRectangleDetermination
paintRectangle pRectangle
determineBoundingRect = mGraphics.EndBoundingRectangleDetermination

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub paintRectangle(ByVal pRectangle As Rectangle)
Const ProcName As String = "paintRectangle"
On Error GoTo Err

setOrientation pRectangle

mGraphics.PaintRectangleFromPositionAndSize pRectangle.Brush, pRectangle.Pen, pRectangle.Position, pRectangle.Size

resetOrientation pRectangle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resetOrientation(ByVal pRectangle As Rectangle)
Const ProcName As String = "resetOrientation"
On Error GoTo Err

If pRectangle.Orientation <> 0 Then mGraphics.Reset

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOrientation(ByVal pRectangle As Rectangle)
Const ProcName As String = "setOrientation"
On Error GoTo Err

Dim lPosition As Point

If pRectangle.Orientation = 0 Then Exit Sub

Set lPosition = pRectangle.Position
mGraphics.RotateAboutPoint pRectangle.Orientation, TPoint(lPosition.XLogical(mGraphics), lPosition.YLogical(mGraphics))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


