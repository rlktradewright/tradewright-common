Attribute VB_Name = "GGdi"
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

Private Const ModuleName                            As String = "GGdi"

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

Public Function GdiPointToString( _
                ByRef pPoint As GDI_POINT) As String
GdiPointToString = "X=" & pPoint.X & "; Y=" & pPoint.Y
End Function

Public Function GdiPolygonToRectangle(ByRef pPoints() As GDI_POINT) As GDI_RECT
Dim i As Long
With GdiPolygonToRectangle
    .Bottom = MinLong
    .Left = MaxLong
    .Right = MinLong
    .Top = MaxLong
    For i = 0 To UBound(pPoints)
        If pPoints(i).X < .Left Then .Left = pPoints(i).X
        If pPoints(i).X > .Right Then .Right = pPoints(i).X
        If pPoints(i).Y > .Bottom Then .Bottom = pPoints(i).Y
        If pPoints(i).Y < .Top Then .Top = pPoints(i).Y
    Next
End With
End Function

Public Function GdiRectangle( _
                ByVal x1 As Long, _
                ByVal y1 As Long, _
                ByVal x2 As Long, _
                ByVal y2 As Long) As GDI_RECT
With GdiRectangle
    .Left = IIf(x1 <= x2, x1, x2)
    .Top = IIf(y1 <= y2, y1, y2)
    .Bottom = IIf(y1 <= y2, y2, y1)
    .Right = IIf(x1 <= x2, x2, x1)
End With
End Function

Public Function GdiRectangleCompare( _
                ByRef pRect1 As GDI_RECT, _
                ByRef pRect2 As GDI_RECT) As Boolean
If pRect1.Bottom <> pRect2.Bottom Then Exit Function
If pRect1.Left <> pRect2.Left Then Exit Function
If pRect1.Right <> pRect2.Right Then Exit Function
If pRect1.Top <> pRect2.Top Then Exit Function
GdiRectangleCompare = True
End Function

Public Sub GdiRectangleSetFields( _
                ByRef pRect As GDI_RECT, _
                ByVal x1 As Long, _
                ByVal y1 As Long, _
                ByVal x2 As Long, _
                ByVal y2 As Long)
With pRect
    .Left = IIf(x1 <= x2, x1, x2)
    .Top = IIf(y1 <= y2, y1, y2)
    .Bottom = IIf(y1 <= y2, y2, y1)
    .Right = IIf(x1 <= x2, x2, x1)
End With
End Sub

Public Sub GdiRectangleSetFieldsByPositionAndSize( _
                ByRef pRect As GDI_RECT, _
                ByVal X As Long, _
                ByVal Y As Long, _
                ByVal pWidth As Long, _
                ByVal pHeight As Long)
With pRect
    .Left = X
    .Top = Y
    .Bottom = .Top - pHeight
    .Right = .Left + pWidth
End With
End Sub

Public Function GdiRectangleToString( _
                ByRef pRect As GDI_RECT) As String
GdiRectangleToString = "Bottom=" & pRect.Bottom & "; Left=" & pRect.Left & "; Top=" & pRect.Top & "; Right=" & pRect.Right
End Function

'@================================================================================
' Helper Functions
'@================================================================================




