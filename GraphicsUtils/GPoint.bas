Attribute VB_Name = "GPoint"
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

Private Const ModuleName                            As String = "GPoint"

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

Public Function gLoadDimensionFromConfig( _
                ByVal pConfig As ConfigurationSection) As Dimension
Const ProcName As String = "gLoadDimensionFromConfig"
On Error GoTo Err

Set gLoadDimensionFromConfig = New Dimension
gLoadDimensionFromConfig.LoadFromConfig pConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gLoadPointFromConfig( _
                ByVal pConfig As ConfigurationSection) As Point
Const ProcName As String = "gLoadPointFromConfig"
On Error GoTo Err

Set gLoadPointFromConfig = New Point
gLoadPointFromConfig.LoadFromConfig pConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gLoadSizeFromConfig( _
                ByVal pConfig As ConfigurationSection) As size
Const ProcName As String = "gLoadSizeFromConfig"
On Error GoTo Err

Set gLoadSizeFromConfig = New size
gLoadSizeFromConfig.LoadFromConfig pConfig

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gNewDimension( _
                ByVal pLength As Double, _
                Optional ByVal pScaleUnit As ScaleUnits = ScaleUnitCm) As Dimension
Const ProcName As String = "gNewDimension"
Dim failpoint As String
On Error GoTo Err

Set gNewDimension = New Dimension
gNewDimension.Initialise pLength, pScaleUnit
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gNewPoint( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal coordSystemX As CoordinateSystems = CoordsLogical, _
                Optional ByVal coordSystemY As CoordinateSystems = CoordsLogical, _
                Optional ByVal pOffset As size) As Point
Const ProcName As String = "gNewPoint"
Dim failpoint As String
On Error GoTo Err

Set gNewPoint = New Point
gNewPoint.Initialise X, Y, coordSystemX, coordSystemY, pOffset

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gNewSize( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal pScaleUnitX As ScaleUnits = ScaleUnitCm, _
                Optional ByVal pScaleUnitY As ScaleUnits = ScaleUnitCm) As size
Const ProcName As String = "gNewSize"
Dim failpoint As String
On Error GoTo Err

Set gNewSize = New size
gNewSize.Initialise X, Y, pScaleUnitX, pScaleUnitY
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gTransformCoordX( _
                ByVal pValue As Double, _
                ByVal pFromCoordSys As CoordinateSystems, _
                ByVal pToCoordSys As CoordinateSystems, _
                ByVal pGraphics As Graphics) As Double
Const ProcName As String = "gTransformCoordX"
Dim failpoint As String
On Error GoTo Err

If pFromCoordSys = pToCoordSys Then
    gTransformCoordX = pValue
    Exit Function
End If

Select Case pFromCoordSys
Case CoordsLogical
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordX = pGraphics.ConvertLogicalToDistanceX(pGraphics.Boundary.Right - pValue)
    ElseIf pToCoordSys = CoordsCounterPixels Then
        gTransformCoordX = pGraphics.ConvertLogicalToPixelsX(pGraphics.Boundary.Right - pValue)
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordX = pGraphics.ConvertLogicalToDistanceX(pValue - pGraphics.Boundary.Left)
    ElseIf pToCoordSys = CoordsPixels Then
        gTransformCoordX = pGraphics.ConvertLogicalToPixelsX(pValue - pGraphics.Boundary.Left)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordX = pGraphics.ConvertLogicalToDistanceX(pValue - pGraphics.Boundary.Left) / pGraphics.Width
    End If
Case CoordsRelative
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordX = pGraphics.Width - pGraphics.ConvertRelativeToDistanceX(pValue)
    ElseIf pToCoordSys = CoordsCounterPixels Then
        gTransformCoordX = pGraphics.ConvertRelativeToPixelsX(1# - pValue)
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordX = pGraphics.ConvertRelativeToDistanceX(pValue)
    ElseIf pToCoordSys = CoordsLogical Then
        gTransformCoordX = pGraphics.ConvertRelativeToLogicalX(pValue) + pGraphics.Boundary.Left
    ElseIf pToCoordSys = CoordsPixels Then
        gTransformCoordX = pGraphics.ConvertRelativeToPixelsX(pValue)
    End If
Case CoordsDistance
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordX = pGraphics.WidthCm - pValue
    ElseIf pToCoordSys = CoordsCounterPixels Then
        gTransformCoordX = pGraphics.ConvertDistanceToPixelsX(pGraphics.WidthCm - pValue)
    ElseIf pToCoordSys = CoordsLogical Then
        gTransformCoordX = pGraphics.ConvertDistanceToLogicalX(pValue) + pGraphics.Boundary.Left
    ElseIf pToCoordSys = CoordsPixels Then
        gTransformCoordX = pGraphics.ConvertDistanceToPixelsX(pValue)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordX = pGraphics.ConvertDistanceToRelativeX(pValue)
    End If
Case CoordsCounterDistance
    If pToCoordSys = CoordsDistance Then
        gTransformCoordX = pGraphics.WidthCm - pValue
    ElseIf pToCoordSys = CoordsCounterPixels Then
        gTransformCoordX = pGraphics.ConvertDistanceToPixelsX(pValue)
    ElseIf pToCoordSys = CoordsLogical Then
        gTransformCoordX = pGraphics.Boundary.Right - pGraphics.ConvertDistanceToLogicalX(pGraphics.WidthCm - pValue)
    ElseIf pToCoordSys = CoordsPixels Then
        gTransformCoordX = pGraphics.ConvertDistanceToPixelsX(pGraphics.WidthCm - pValue)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordX = pGraphics.ConvertDistanceToRelativeX(pGraphics.WidthCm - pValue)
    End If
Case CoordsPixels
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordX = pGraphics.WidthCm - pGraphics.ConvertPixelsToDistanceX(pValue)
    ElseIf pToCoordSys = CoordsCounterPixels Then
        gTransformCoordX = pGraphics.WidthPixels - pValue
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordX = pGraphics.ConvertPixelsToDistanceX(pValue)
    ElseIf pToCoordSys = CoordsLogical Then
        gTransformCoordX = pGraphics.ConvertPixelsToLogicalX(pValue) + pGraphics.Boundary.Left
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordX = pValue / pGraphics.WidthPixels
    End If
Case CoordsCounterPixels
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordX = pGraphics.WidthCm - pGraphics.ConvertPixelsToDistanceX(pGraphics.WidthPixels - pValue)
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordX = pGraphics.ConvertPixelsToDistanceX(pGraphics.WidthPixels - pValue)
    ElseIf pToCoordSys = CoordsLogical Then
        gTransformCoordX = pGraphics.ConvertPixelsToLogicalX(pGraphics.WidthPixels - pValue) + pGraphics.Boundary.Left
    ElseIf pToCoordSys = CoordsPixels Then
        gTransformCoordX = pGraphics.WidthPixels - pValue
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordX = (pGraphics.WidthPixels - pValue) / pGraphics.WidthPixels
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gTransformCoordY( _
                ByVal pValue As Double, _
                ByVal pFromCoordSys As CoordinateSystems, _
                ByVal pToCoordSys As CoordinateSystems, _
                ByVal pGraphics As Graphics) As Double
Const ProcName As String = "gTransformCoordY"
Dim failpoint As String
On Error GoTo Err

If pFromCoordSys = pToCoordSys Then
    gTransformCoordY = pValue
    Exit Function
End If

Select Case pFromCoordSys
Case CoordsLogical
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordY = pGraphics.ConvertLogicalToDistanceY(pGraphics.Boundary.Top - pValue)
    ElseIf pToCoordSys = CoordsCounterPixels Then
        gTransformCoordY = pGraphics.ConvertLogicalToPixelsY(pGraphics.Boundary.Top - pValue)
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordY = pGraphics.ConvertLogicalToDistanceY(pValue - pGraphics.Boundary.Bottom)
    ElseIf pToCoordSys = CoordsPixels Then
        gTransformCoordY = pGraphics.ConvertLogicalToPixelsY(pValue - pGraphics.Boundary.Bottom)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordY = pGraphics.ConvertLogicalToDistanceY(pValue - pGraphics.Boundary.Bottom) / pGraphics.Height
    End If
Case CoordsRelative
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordY = pGraphics.Height - pGraphics.ConvertRelativeToDistanceY(pValue)
    ElseIf pToCoordSys = CoordsCounterPixels Then
        gTransformCoordY = pGraphics.ConvertRelativeToPixelsY(1# - pValue)
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordY = pGraphics.ConvertRelativeToDistanceY(pValue)
    ElseIf pToCoordSys = CoordsLogical Then
        gTransformCoordY = pGraphics.ConvertRelativeToLogicalY(pValue) + pGraphics.Boundary.Bottom
    ElseIf pToCoordSys = CoordsPixels Then
        gTransformCoordY = pGraphics.ConvertRelativeToPixelsY(pValue)
    End If
Case CoordsDistance
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordY = pGraphics.HeightCm - pValue
    ElseIf pToCoordSys = CoordsCounterPixels Then
        gTransformCoordY = pGraphics.ConvertDistanceToPixelsY(pGraphics.HeightCm - pValue)
    ElseIf pToCoordSys = CoordsLogical Then
        gTransformCoordY = pGraphics.ConvertDistanceToLogicalY(pValue) + pGraphics.Boundary.Bottom
    ElseIf pToCoordSys = CoordsPixels Then
        gTransformCoordY = pGraphics.ConvertDistanceToPixelsY(pValue)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordY = pGraphics.ConvertDistanceToRelativeY(pValue)
    End If
Case CoordsCounterDistance
    If pToCoordSys = CoordsDistance Then
        gTransformCoordY = pGraphics.HeightCm - pValue
    ElseIf pToCoordSys = CoordsCounterPixels Then
        gTransformCoordY = pGraphics.ConvertDistanceToPixelsY(pValue)
    ElseIf pToCoordSys = CoordsLogical Then
        gTransformCoordY = pGraphics.Boundary.Top - pGraphics.ConvertDistanceToLogicalY(pGraphics.HeightCm - pValue)
    ElseIf pToCoordSys = CoordsPixels Then
        gTransformCoordY = pGraphics.ConvertDistanceToPixelsY(pGraphics.HeightCm - pValue)
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordY = pGraphics.ConvertDistanceToRelativeY(pGraphics.HeightCm - pValue)
    End If
Case CoordsPixels
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordY = pGraphics.HeightCm - pGraphics.ConvertPixelsToDistanceY(pValue)
    ElseIf pToCoordSys = CoordsCounterPixels Then
        gTransformCoordY = pGraphics.HeightPixels - pValue
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordY = pGraphics.ConvertPixelsToDistanceY(pValue)
    ElseIf pToCoordSys = CoordsLogical Then
        gTransformCoordY = pGraphics.ConvertPixelsToLogicalY(pValue) + pGraphics.Boundary.Bottom
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordY = pValue / pGraphics.HeightPixels
    End If
Case CoordsCounterPixels
    If pToCoordSys = CoordsCounterDistance Then
        gTransformCoordY = pGraphics.HeightCm - pGraphics.ConvertPixelsToDistanceY(pGraphics.HeightPixels - pValue)
    ElseIf pToCoordSys = CoordsDistance Then
        gTransformCoordY = pGraphics.ConvertPixelsToDistanceY(pGraphics.HeightPixels - pValue)
    ElseIf pToCoordSys = CoordsLogical Then
        gTransformCoordY = pGraphics.ConvertPixelsToLogicalY(pGraphics.HeightPixels - pValue) + pGraphics.Boundary.Bottom
    ElseIf pToCoordSys = CoordsPixels Then
        gTransformCoordY = pGraphics.HeightPixels - pValue
    ElseIf pToCoordSys = CoordsRelative Then
        gTransformCoordY = (pGraphics.HeightPixels - pValue) / pGraphics.HeightPixels
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




