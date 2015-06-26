Attribute VB_Name = "GRectangle"
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


Private Const ModuleName                    As String = "GRectangle"

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

'================================================================================
' Rectangle functions
'
' NB: these are implemented as functions rather than class methods for
' efficiency reasons, due to the very large numbers of rectangles made
' use of
'================================================================================

Public Function TIntervalContains( _
                ByRef pInterval As TInterval, _
                ByVal X As Double) As Boolean
If Not pInterval.isValid Then Exit Function
TIntervalContains = (X >= pInterval.startValue) And (X <= pInterval.endValue)
End Function

Public Function TIntervalIntersection( _
                ByRef int1 As TInterval, _
                ByRef int2 As TInterval) As TInterval
With TIntervalIntersection
    .startValue = gMaxDouble(int1.startValue, int2.startValue)
    .endValue = gMinDouble(int1.endValue, int2.endValue)
    .isValid = .startValue <= .endValue
End With

'Dim startValue1 As Double
'Dim endValue1 As Double
'Dim startValue2 As Double
'Dim endValue2 As Double
'
'If Not int1.isValid Or Not int2.isValid Then Exit Function
'
'startValue1 = int1.startValue
'endValue1 = int1.endValue
'startValue2 = int2.startValue
'endValue2 = int2.endValue
'
'With TIntervalIntersection
'    If startValue1 >= startValue2 And startValue1 <= endValue2 Then
'        .startValue = startValue1
'        If endValue1 >= startValue2 And endValue1 <= endValue2 Then
'            .endValue = endValue1
'        Else
'            .endValue = endValue2
'        End If
'        .isValid = True
'        Exit Function
'    End If
'    If endValue1 >= startValue2 And endValue1 <= endValue2 Then
'        .endValue = endValue1
'        .startValue = startValue2
'        .isValid = True
'        Exit Function
'    End If
'    If startValue1 < startValue2 And endValue1 > endValue2 Then
'        .startValue = startValue2
'        .endValue = endValue2
'        .isValid = True
'        Exit Function
'    End If
'End With
End Function

Public Function TIntervalOverlaps( _
                ByRef int1 As TInterval, _
                ByRef int2 As TInterval) As Boolean
TIntervalOverlaps = TIntervalIntersection(int1, int2).isValid

'TIntervalOverlaps = True
'
'If int1.startValue >= int2.startValue And int1.startValue <= int2.endValue Then
'    Exit Function
'End If
'If int1.endValue >= int2.startValue And int1.endValue <= int2.endValue Then
'    Exit Function
'End If
'If int1.startValue < int2.startValue And int1.endValue > int2.endValue Then
'    Exit Function
'End If
'TIntervalOverlaps = False
End Function

Public Function TPoint( _
                ByVal pX As Double, _
                ByVal pY As Double) As TPoint
TPoint.X = pX
TPoint.Y = pY
End Function

Public Function TPointAdd( _
                ByRef pPoint1 As TPoint, _
                ByRef pPoint2 As TPoint) As TPoint
TPointAdd.X = pPoint1.X + pPoint2.X
TPointAdd.Y = pPoint1.Y + pPoint2.Y
End Function

Public Function TPointFromShortString( _
                ByRef pInput As String) As TPoint
Dim ar As Variant
Const ProcName As String = "TPointFromShortString"
On Error GoTo Err

ar = Split(pInput, ",")
TPointFromShortString.X = ar(0)
TPointFromShortString.Y = ar(1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function TPointSubtract( _
                ByRef pPoint1 As TPoint, _
                ByRef pPoint2 As TPoint) As TPoint
TPointSubtract.X = pPoint1.X - pPoint2.X
TPointSubtract.Y = pPoint1.Y - pPoint2.Y
End Function

Public Function TPointToShortString( _
                ByRef pPoint As TPoint) As String
TPointToShortString = pPoint.X & "," & pPoint.Y
End Function

Public Function TPointToString( _
                ByRef pPoint As TPoint) As String
TPointToString = "X=" & pPoint.X & "; Y=" & pPoint.Y
End Function

Public Function TRectangle( _
                ByVal x1 As Double, _
                ByVal y1 As Double, _
                ByVal x2 As Double, _
                ByVal y2 As Double, _
                Optional allowZeroDimensions As Boolean = False) As TRectangle
With TRectangle
    .Left = IIf(x1 <= x2, x1, x2)
    .Top = IIf(y1 <= y2, y2, y1)
    .Bottom = IIf(y1 <= y2, y1, y2)
    .Right = IIf(x1 <= x2, x2, x1)
    
    If allowZeroDimensions Then
        .isValid = True
    Else
        TRectangleValidate TRectangle, False
    End If
End With
End Function

Public Function TRectangleBottomCentre( _
                ByRef pRect As TRectangle) As TPoint
TRectangleBottomCentre.X = (pRect.Right + pRect.Left) / 2
TRectangleBottomCentre.Y = pRect.Bottom
End Function

Public Function TRectangleBottomLeft( _
                ByRef pRect As TRectangle) As TPoint
TRectangleBottomLeft.X = pRect.Left
TRectangleBottomLeft.Y = pRect.Bottom
End Function

Public Function TRectangleBottomRight( _
                ByRef pRect As TRectangle) As TPoint
TRectangleBottomRight.X = pRect.Right
TRectangleBottomRight.Y = pRect.Bottom
End Function

Public Function TRectangleCentreCentre( _
                ByRef pRect As TRectangle) As TPoint
TRectangleCentreCentre.X = (pRect.Right + pRect.Left) / 2
TRectangleCentreCentre.Y = (pRect.Top + pRect.Bottom) / 2
End Function

Public Function TRectangleCentreLeft( _
                ByRef pRect As TRectangle) As TPoint
TRectangleCentreLeft.X = pRect.Left
TRectangleCentreLeft.Y = (pRect.Top + pRect.Bottom) / 2
End Function

Public Function TRectangleCentreRight( _
                ByRef pRect As TRectangle) As TPoint
TRectangleCentreRight.X = pRect.Right
TRectangleCentreRight.Y = (pRect.Top + pRect.Bottom) / 2
End Function

Public Function TRectangleContainsPoint( _
                ByRef pRect As TRectangle, _
                ByVal X As Double, _
                ByVal Y As Double) As Boolean
If Not pRect.isValid Then Exit Function
If X < pRect.Left Then Exit Function
If X > pRect.Right Then Exit Function
If Y < pRect.Bottom Then Exit Function
If Y > pRect.Top Then Exit Function
TRectangleContainsPoint = True
End Function

Public Function TRectangleContainsRect( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
If Not rect1.isValid Then Exit Function
If Not rect2.isValid Then Exit Function
If rect2.Left < rect1.Left Then Exit Function
If rect2.Right > rect1.Right Then Exit Function
If rect2.Bottom < rect1.Bottom Then Exit Function
If rect2.Top > rect1.Top Then Exit Function
TRectangleContainsRect = True
End Function

Public Function TRectangleEquals( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
With rect1
    If Not .isValid Or Not rect2.isValid Then Exit Function
    If .Bottom <> rect2.Bottom Then Exit Function
    If .Left <> rect2.Left Then Exit Function
    If .Top <> rect2.Top Then Exit Function
    If .Right <> rect2.Right Then Exit Function
End With
TRectangleEquals = True
End Function

Public Sub TRectangleExpand( _
                ByRef pRect As TRectangle, _
                ByVal xIncrement As Double, _
                ByVal yIncrement As Double)
With pRect
    If Not .isValid Then Exit Sub
    .Left = .Left - xIncrement
    .Right = .Right + xIncrement
    .Top = .Top + yIncrement
    .Bottom = .Bottom - yIncrement
End With
End Sub

Public Sub TRectangleExpandByRotation( _
                ByRef pRect As TRectangle, _
                ByVal pAngle As Double, _
                ByVal pGraphics As Graphics)
Dim centreX As Double
Dim centreY As Double
Dim halfWidth As Double
Dim halfHeight As Double
Dim sinOfAngle As Double
Dim cosOfAngle As Double
With pRect
    If Not .isValid Then Exit Sub
    centreX = (.Left + .Right) / 2
    centreY = (.Bottom + .Top) / 2
    halfWidth = (.Right - .Left) / 2
    halfHeight = (.Top - .Bottom) / 2
    sinOfAngle = Abs(Sin(pAngle))
    cosOfAngle = Abs(Cos(pAngle))
    
    .Left = centreX - (halfWidth * cosOfAngle + halfHeight * sinOfAngle / pGraphics.AspectRatio)
    .Right = centreX + halfWidth * cosOfAngle + halfHeight * sinOfAngle / pGraphics.AspectRatio
    .Top = centreY + pGraphics.AspectRatio * halfWidth * sinOfAngle + halfHeight * cosOfAngle
    .Bottom = centreY - (pGraphics.AspectRatio * halfWidth * sinOfAngle + halfHeight * cosOfAngle)
End With
End Sub

Public Sub TRectangleExpandBySize( _
                ByRef pRect As TRectangle, _
                ByVal pSize As size, _
                ByVal pGraphics As Graphics)
TRectangleExpand pRect, pSize.WidthLogical(pGraphics), pSize.HeightLogical(pGraphics)
End Sub

Public Function TRectangleFromShortString( _
                ByRef pInput As String, _
                Optional allowZeroDimensions As Boolean = False) As TRectangle
Dim ar As Variant
Const ProcName As String = "TRectangleFromShortString"
On Error GoTo Err

ar = Split(pInput, ",")
TRectangleFromShortString = TRectangle(ar(0), ar(1), ar(2), ar(3), allowZeroDimensions)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function TRectangleGetXInterval( _
                ByRef pRect As TRectangle) As TInterval
With TRectangleGetXInterval
    .startValue = pRect.Left
    .endValue = pRect.Right
    .isValid = pRect.isValid
End With
End Function

Public Function TRectangleGetYInterval( _
                ByRef pRect As TRectangle) As TInterval
With TRectangleGetYInterval
    .startValue = pRect.Bottom
    .endValue = pRect.Top
    .isValid = pRect.isValid
End With
End Function

Public Sub TRectangleInitialise( _
                ByRef pRect As TRectangle)
With pRect
    .isValid = False
    .Left = MaxDouble
    .Right = MinDouble
    .Bottom = MaxDouble
    .Top = MinDouble
End With
End Sub

Public Function TRectangleIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TRectangle
With TRectangleIntersection
    .Left = gMaxDouble(rect1.Left, rect2.Left)
    
    
    .Right = gMinDouble(rect1.Right, rect2.Right)
    
    .Bottom = gMaxDouble(rect1.Bottom, rect2.Bottom)
    
    .Top = gMinDouble(rect1.Top, rect2.Top)
End With
TRectangleValidate TRectangleIntersection, False
        
               
'Dim xInt As TInterval
'Dim yint As TInterval
'xInt = TIntervalIntersection(TRectangleGetXInterval(rect1), TRectangleGetXInterval(rect2))
'yint = TIntervalIntersection(TRectangleGetYInterval(rect1), TRectangleGetYInterval(rect2))
'With TRectangleIntersection
'    .Left = xInt.startValue
'    .Right = xInt.endValue
'    .Bottom = yint.startValue
'    .Top = yint.endValue
'    If xInt.isValid And yint.isValid Then .isValid = True
'End With
End Function


Public Function TRectangleOverlaps( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
TRectangleOverlaps = TRectangleIntersection(rect1, rect2).isValid
'TRectangleOverlaps = TIntervalOverlaps(TRectangleGetXInterval(rect1), TRectangleGetXInterval(rect2)) And _
'            TIntervalOverlaps(TRectangleGetYInterval(rect1), TRectangleGetYInterval(rect2))
End Function

Public Sub TRectangleSetFields( _
                ByRef pRect As TRectangle, _
                ByVal x1 As Double, _
                ByVal y1 As Double, _
                ByVal x2 As Double, _
                ByVal y2 As Double, _
                Optional allowZeroDimensions As Boolean = False)
With pRect
    .Left = IIf(x1 <= x2, x1, x2)
    .Top = IIf(y1 <= y2, y2, y1)
    .Bottom = IIf(y1 <= y2, y1, y2)
    .Right = IIf(x1 <= x2, x2, x1)
    
    If allowZeroDimensions Then
        .isValid = True
    Else
        TRectangleValidate pRect, False
    End If
End With
End Sub

Public Sub TRectangleSetXInterval( _
                ByRef pRect As TRectangle, _
                ByRef pInterval As TInterval)
With pRect
'    If pInterval.startValue <= pInterval.endValue Then
        .Left = pInterval.startValue
        .Right = pInterval.endValue
'    Else
'        .Left = pInterval.endValue
'        .Right = pInterval.startValue
'    End If
    .isValid = .isValid And pInterval.isValid
End With
End Sub

Public Sub TRectangleSetYInterval( _
                ByRef pRect As TRectangle, _
                ByRef pInterval As TInterval)
With pRect
'    If pInterval.startValue <= pInterval.endValue Then
        .Bottom = pInterval.startValue
        .Top = pInterval.endValue
'    Else
'        .Bottom = pInterval.endValue
'        .Top = pInterval.startValue
'    End If
    .isValid = pInterval.isValid
End With
End Sub

Public Function TRectangleTopCentre( _
                ByRef pRect As TRectangle) As TPoint
TRectangleTopCentre.X = (pRect.Right + pRect.Left) / 2
TRectangleTopCentre.Y = pRect.Top
End Function

Public Function TRectangleTopLeft( _
                ByRef pRect As TRectangle) As TPoint
TRectangleTopLeft.X = pRect.Left
TRectangleTopLeft.Y = pRect.Top
End Function

Public Function TRectangleTopRight( _
                ByRef pRect As TRectangle) As TPoint
TRectangleTopRight.X = pRect.Right
TRectangleTopRight.Y = pRect.Top
End Function

Public Function TRectangleToString( _
                ByRef pRect As TRectangle) As String
TRectangleToString = IIf(pRect.isValid, "Valid: ", "Invalid: ") & "Bottom=" & pRect.Bottom & "; Left=" & pRect.Left & "; Top=" & pRect.Top & "; Right=" & pRect.Right
End Function

Public Function TRectangleToShortString( _
                ByRef pRect As TRectangle) As String
TRectangleToShortString = IIf(pRect.isValid, "True,", "False,") & pRect.Bottom & "," & pRect.Left & "," & pRect.Top & "," & pRect.Right
End Function

Public Sub TRectangleTranslate( _
                ByRef pRect As TRectangle, _
                ByVal pdX As Double, _
                ByVal pdY As Double)

If Not pRect.isValid Then Exit Sub

With pRect
    .Left = .Left + pdX
    .Right = .Right + pdX
    .Bottom = .Bottom + pdY
    .Top = .Top + pdY
End With
End Sub

Public Sub TRectangleTranslateBySize( _
                ByRef pRect As TRectangle, _
                ByRef pOffset As size, _
                ByVal pGraphics As Graphics)
TRectangleTranslate pRect, pOffset.WidthLogical(pGraphics), pOffset.HeightLogical(pGraphics)
End Sub

Public Sub TRectangleTranslateByPoint( _
                ByRef pRect As TRectangle, _
                ByRef pOffset As TPoint)
TRectangleTranslate pRect, pOffset.X, pOffset.Y
End Sub

Public Function TRectangleUnion( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TRectangle
If Not (rect1.isValid And rect2.isValid) Then
    If rect1.isValid Then
        TRectangleUnion = rect1
    ElseIf rect2.isValid Then
        TRectangleUnion = rect2
    End If
    Exit Function
End If

With TRectangleUnion
    .Left = gMinDouble(rect1.Left, rect2.Left)
    .Bottom = gMinDouble(rect1.Bottom, rect2.Bottom)
    .Top = gMaxDouble(rect1.Top, rect2.Top)
    .Right = gMaxDouble(rect1.Right, rect2.Right)
    .isValid = True
End With
End Function

Public Function TRectangleValidate( _
                ByRef pRect As TRectangle, _
                Optional allowZeroDimensions As Boolean = False) As Boolean
With pRect
    .isValid = False
    If allowZeroDimensions Then
        If .Left <= .Right And .Bottom <= .Top Then .isValid = True
    Else
        If .Left < .Right And .Bottom < .Top Then .isValid = True
    End If
    TRectangleValidate = .isValid
End With
End Function

Public Function TRectangleXIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TInterval
With TRectangleXIntersection
    .startValue = gMaxDouble(rect1.Left, rect2.Left)
    .endValue = gMinDouble(rect1.Right, rect2.Right)
    .isValid = .startValue <= .endValue
End With
'TRectangleXIntersection = TIntervalIntersection(TRectangleGetXInterval(rect1), TRectangleGetXInterval(rect2))
End Function

Public Function TRectangleYIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TInterval
With TRectangleYIntersection
    .startValue = gMaxDouble(rect1.Bottom, rect2.Bottom)
    .endValue = gMinDouble(rect1.Top, rect2.Top)
    .isValid = .startValue <= .endValue
End With
'TRectangleYIntersection = TIntervalIntersection(TRectangleGetYInterval(rect1), TRectangleGetYInterval(rect2))
End Function

Public Sub TPointMultiply( _
                ByRef pPoint As TPoint, _
                ByVal pFactor As Double)
pPoint.X = pPoint.X * pFactor
pPoint.Y = pPoint.Y * pFactor
End Sub

'@================================================================================
' Helper Functions
'@================================================================================



