VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicsUtils"
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

Public Enum CoordinateSystems
    CoordsLogical = 0
    CoordsRelative
    CoordsDistance          ' Measured from Left or Bottom of region
    CoordsCounterDistance   ' Measured from Right or Top of region
    CoordsPixels            ' Measured from Left or Bottom of region
    CoordsCounterPixels     ' Measured from Right or Top of region
End Enum

Public Enum GraphicsChangeTypes
    GraphicsNoChange = 0
    GraphicsGaugeChangedX = 2
    GraphicsGaugeChangedY = 4
    GraphicsBoundsChangedX = 8
    GraphicsBoundsChangedY = 16
End Enum

Public Enum MixModes
    MixModeBlack = R2_BLACK
    MixModeCopyPen = R2_COPYPEN
    MixModeInvert = R2_NOT
    MixModeMaskNotPen = R2_MASKNOTPEN
    MixModeMaskPen = R2_MASKPEN
    MixModeMaskPenNot = R2_MASKPENNOT
    MixModeMergeNotPen = R2_MERGENOTPEN
    MixModeMergePen = R2_MERGEPEN
    MixModeMergePenNot = R2_MERGEPENNOT
    MixModeNop = R2_NOP
    MixModeNotCopyPen = R2_NOTCOPYPEN
    MixModeNotMaskPen = R2_NOTMASKPEN
    MixModeNotMergePen = R2_NOTMERGEPEN
    MixModeNotXorPen = R2_NOTXORPEN
    MixModeWhite = R2_WHITE
    MixModeXorPen = R2_XORPEN
End Enum

Public Enum EllipsisModes
    EllipsisNone = 0
    EllipsisEnd = DT_END_ELLIPSIS
    EllipsisPath = DT_PATH_ELLIPSIS
    EllipsisWord = DT_WORD_ELLIPSIS
End Enum

Public Enum Endcaps
    EndCapFlat = PS_ENDCAP_FLAT
    EndcapRound = PS_ENDCAP_ROUND
    EndcapSquare = PS_ENDCAP_SQUARE
End Enum

Public Enum FillModes
    FillModeAlternate = ALTERNATE
    FillModeWinding = WINDING
End Enum

Public Enum HatchStyles
    HatchNone = 0
    HatchHorizontal = &H80000000 Or HS_HORIZONTAL
    HatchVertical = &H80000000 Or HS_VERTICAL
    HatchUpwardDiagonal = &H80000000 Or HS_BDIAGONAL
    HatchDownwardDiagonal = &H80000000 Or HS_FDIAGONAL
    HatchCross = &H80000000 Or HS_CROSS
    HatchDiagonalCross = &H80000000 Or HS_DIAGCROSS
End Enum

Public Enum JoinModes
    JoinBevel = PS_JOIN_BEVEL
    JoinMiter = PS_JOIN_MITER
    JoinRound = PS_JOIN_ROUND
End Enum

Public Enum LinearGradientDirections
    LinearGradientDirectionHorizontal
    LinearGradientDirectionVertical
End Enum

Public Enum LineStyles
    LineSolid = PS_SOLID
    LineDash = PS_DASH
    LineDot = PS_DOT
    LineDashDot = PS_DASHDOT
    LineDashDotDot = PS_DASHDOTDOT
    LineInvisible = PS_NULL
    LineInsideSolid = PS_INSIDEFRAME
    LineAlternate = PS_ALTERNATE
    LineCustom = PS_USERSTYLE
End Enum

Public Enum ScaleUnits
    ScaleUnitLogical = 0
    ScaleUnitRelative
    ScaleUnitCm
    ScaleUnitPixels
End Enum

Public Enum TextJustifyModes
    JustifyLeft = DT_LEFT
    JustifyRight = DT_RIGHT
    JustifyCentre = DT_CENTER
    JustifyTop = DT_TOP
    JustifyVCentre = DT_VCENTER
    JustifyBottom = DT_BOTTOM
End Enum

Public Enum TileModes
    TileModeTile
    TileModeNone
    TileModeFlipX
    TileModeFlipY
    TileModeFlipXY
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type TInterval
    isValid         As Boolean
    startValue      As Double
    endValue        As Double
End Type

Public Type TPoint
    X               As Double
    Y               As Double
End Type

Public Type TRectangle
    Bottom          As Double
    Left            As Double
    Right           As Double
    Top             As Double
    isValid         As Boolean
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GraphicsUtils"

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

Public Function CreateBrush( _
                Optional ByVal pColor As Long = vbBlack) As IBrush
Const ProcName As String = "CreateBrush"
On Error GoTo Err

Set CreateBrush = gCreateBrush(pColor)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateGraphics( _
                ByVal phWnd As Long, _
                ByVal pLeft As Double, _
                ByVal pBottom As Double, _
                ByVal pRight As Double, _
                ByVal pTop As Double, _
                ByVal pBackGroundBrush As IBrush) As Graphics
Const ProcName As String = "CreateGraphics"
On Error GoTo Err

gLogger.Log "Creating new Graphics object", ProcName, ModuleName
Set CreateGraphics = New Graphics
CreateGraphics.Initialise phWnd, pLeft, pBottom, pRight, pTop, pBackGroundBrush

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateGraphicsMonitor( _
                ByVal pGraphics As Graphics) As GraphicsMonitor
Const ProcName As String = "CreateGraphicsMonitor"
On Error GoTo Err

Set CreateGraphicsMonitor = New GraphicsMonitor
CreateGraphicsMonitor.Initialise pGraphics

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateHatchedBrush( _
                Optional ByVal pColor As Long = vbBlack, _
                Optional ByVal pStyle As HatchStyles = HatchHorizontal) As IBrush
Const ProcName As String = "CreateHatchedBrush"
On Error GoTo Err

Set CreateHatchedBrush = gCreateHatchedBrush(pColor, pStyle)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidColor( _
                ByVal Value As Long) As Boolean
                
IsValidColor = gIsValidColor(Value)
End Function

Public Function CreateLogicalPen( _
                ByVal pGraphics As Graphics, _
                Optional ByVal pColor As Long = vbBlack, _
                Optional ByVal pWidth As Double = 1#, _
                Optional ByVal pLineStyle As LineStyles = LineSolid, _
                Optional ByVal pHatchStyle As HatchStyles = HatchNone) As Pen
Const ProcName As String = "CreateLogicalPen"
On Error GoTo Err

Set CreateLogicalPen = gCreateLogicalPen(pColor, pWidth, pLineStyle, pHatchStyle, pGraphics)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreatePixelPen( _
                Optional ByVal pColor As Long = vbBlack, _
                Optional ByVal pWidth As Double = 1#, _
                Optional ByVal pLineStyle As LineStyles = LineSolid, _
                Optional ByVal pHatchStyle As HatchStyles = HatchNone) As Pen
Const ProcName As String = "CreatePixelPen"
On Error GoTo Err

Set CreatePixelPen = gCreatePixelPen(pColor, pWidth, pLineStyle, pHatchStyle)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateRadialGradientBrush( _
                ByRef pGradientOrigin As TPoint, _
                ByRef pCentre As TPoint, _
                ByVal pRadiusX As Double, _
                ByVal pRadiusY As Double, _
                ByVal pPad As Boolean, _
                ByRef pColors() As Long) As IBrush
Const ProcName As String = "CreateRadialGradientBrush"
On Error GoTo Err

Set CreateRadialGradientBrush = gCreateRadialGradientBrush(pGradientOrigin, pCentre, pRadiusX, pRadiusY, pPad, pColors)
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTextureBrush( _
                ByVal pBitmap As Bitmap) As IBrush
Const ProcName As String = "CreateTextureBrush"
On Error GoTo Err

Set CreateTextureBrush = gCreateTextureBrush(pBitmap)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateTiledGradientBrush( _
                ByRef pBrushBoundary As TRectangle, _
                ByVal pDirection As LinearGradientDirections, _
                ByVal pPad As Boolean, _
                ByVal pTileMode As TileModes, _
                ByRef pColors() As Long, _
                ByRef pIntensities() As Double, _
                ByRef pPositions() As Double) As IBrush
Const ProcName As String = "CreateTiledGradientBrush"
On Error GoTo Err

Set CreateTiledGradientBrush = gCreateTiledGradientBrush(pBrushBoundary, pDirection, pPad, pTileMode, pColors, pIntensities, pPositions)
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function DegreesToRadians( _
                ByVal degrees As Double) As Double
DegreesToRadians = gDegreesToRadians(degrees)
End Function

Public Function GdiPointToString( _
                ByRef pPoint As GDI_POINT) As String
GdiPointToString = GGdi.GdiPointToString(pPoint)
End Function

Public Sub GdiRectangleSetFields( _
                ByRef pRect As GDI_RECT, _
                ByVal x1 As Long, _
                ByVal y1 As Long, _
                ByVal x2 As Long, _
                ByVal y2 As Long)
GGdi.GdiRectangleSetFields pRect, x1, y1, x2, y2
End Sub

Public Function GdiRectangleToString( _
                ByRef pRect As GDI_RECT) As String
GdiRectangleToString = GGdi.GdiRectangleToString(pRect)
End Function

Public Function GetBrush( _
                Optional ByVal pColor As Long = vbBlack) As SolidBrush
Const ProcName As String = "GetBrush"
On Error GoTo Err

Set GetBrush = GBrush.gGetBrush(pColor)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadBitmap(ByVal pFilename As String) As Bitmap
Const ProcName As String = "LoadBitmap"
On Error GoTo Err

Set LoadBitmap = gLoadBitmap(pFilename)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function LoadBitmapFromConfig(ByVal pConfig As ConfigurationSection) As Bitmap
Const ProcName As String = "LoadBitmapFromConfig"
On Error GoTo Err

Set LoadBitmapFromConfig = gLoadBitmapFromConfig(pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function LoadBitmapFromResource(ByVal pResource As StdPicture) As Bitmap
Const ProcName As String = "LoadBitmapFromResource"
On Error GoTo Err

Set LoadBitmapFromResource = gLoadBitmapFromResource(pResource)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadDimensionFromConfig( _
                ByVal pConfig As ConfigurationSection) As Dimension
Const ProcName As String = "LoadDimensionFromConfig"
On Error GoTo Err

Set LoadDimensionFromConfig = gLoadDimensionFromConfig(pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadIBrushFromConfig( _
                ByVal pConfig As ConfigurationSection) As IBrush
Const ProcName As String = "LoadBrushFromConfig"
On Error GoTo Err

Set LoadIBrushFromConfig = gLoadBrushFromConfig(pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadPenFromConfig( _
                ByVal pConfig As ConfigurationSection) As Pen
Const ProcName As String = "LoadPenFromConfig"
On Error GoTo Err

Set LoadPenFromConfig = gLoadPenFromConfig(pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadPointFromConfig( _
                ByVal pConfig As ConfigurationSection) As Point
Const ProcName As String = "LoadPointFromConfig"
On Error GoTo Err

Set LoadPointFromConfig = gLoadPointFromConfig(pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadSizeFromConfig( _
                ByVal pConfig As ConfigurationSection) As size
Const ProcName As String = "LoadSizeFromConfig"
On Error GoTo Err

Set LoadSizeFromConfig = gLoadSizeFromConfig(pConfig)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewDimension( _
                ByVal pLength As Double, _
                Optional ByVal pScale As ScaleUnits = ScaleUnitLogical) As Dimension
Const ProcName As String = "NewDimension"
Dim failpoint As String
On Error GoTo Err

Set NewDimension = gNewDimension(pLength, pScale)
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewPoint( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal pCoordSystemX As CoordinateSystems = CoordsLogical, _
                Optional ByVal pCoordSystemY As CoordinateSystems = CoordsLogical, _
                Optional ByVal pOffset As size) As Point
Const ProcName As String = "NewPoint"
On Error GoTo Err

Set NewPoint = gNewPoint(X, Y, pCoordSystemX, pCoordSystemY, pOffset)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewSize( _
                ByVal X As Double, _
                ByVal Y As Double, _
                Optional ByVal pScaleX As ScaleUnits = ScaleUnitLogical, _
                Optional ByVal pScaleY As ScaleUnits = ScaleUnitLogical) As size
Const ProcName As String = "NewSize"
Dim failpoint As String
On Error GoTo Err

Set NewSize = gNewSize(X, Y, pScaleX, pScaleY)
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function RadiansToDegrees( _
                ByVal radians As Double) As Double
RadiansToDegrees = gRadiansToDegrees(radians)
End Function

Public Function TIntervalContains( _
                ByRef pInterval As TInterval, _
                ByVal X As Double) As Boolean
TIntervalContains = GRectangle.TIntervalContains(pInterval, X)
End Function

Public Function TIntervalIntersection( _
                ByRef int1 As TInterval, _
                ByRef int2 As TInterval) As TInterval
TIntervalIntersection = GRectangle.TIntervalIntersection(int1, int2)
End Function

Public Function TIntervalOverlaps( _
                ByRef int1 As TInterval, _
                ByRef int2 As TInterval) As Boolean
                        
TIntervalOverlaps = GRectangle.TIntervalOverlaps(int1, int2)
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
TPointAdd = GRectangle.TPointAdd(pPoint1, pPoint2)
End Function

Public Function TPointEquals( _
                ByRef pPoint1 As TPoint, _
                ByRef pPoint2 As TPoint) As Boolean
TPointEquals = GRectangle.TPointEquals(pPoint1, pPoint2)
End Function

Public Sub TPointMultiply( _
                ByRef pPoint As TPoint, _
                ByVal pFactor As Double)
GRectangle.TPointMultiply pPoint, pFactor
End Sub

Public Function TPointSubtract( _
                ByRef pPoint1 As TPoint, _
                ByRef pPoint2 As TPoint) As TPoint
TPointSubtract = GRectangle.TPointSubtract(pPoint1, pPoint2)
End Function

Public Function TPointToString( _
                ByRef pPoint As TPoint) As String
TPointToString = GRectangle.TPointToString(pPoint)
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
        GRectangle.TRectangleValidate TRectangle, False
    End If
End With
End Function

Public Function TRectangleBottomCentre( _
                ByRef pRect As TRectangle) As TPoint
TRectangleBottomCentre = GRectangle.TRectangleBottomCentre(pRect)
End Function

Public Function TRectangleBottomLeft( _
                ByRef pRect As TRectangle) As TPoint
TRectangleBottomLeft = GRectangle.TRectangleBottomLeft(pRect)
End Function

Public Function TRectangleBottomRight( _
                ByRef pRect As TRectangle) As TPoint
TRectangleBottomRight = GRectangle.TRectangleBottomRight(pRect)
End Function

Public Function TRectangleCentreCentre( _
                ByRef pRect As TRectangle) As TPoint
TRectangleCentreCentre = GRectangle.TRectangleCentreCentre(pRect)
End Function

Public Function TRectangleCentreLeft( _
                ByRef pRect As TRectangle) As TPoint
TRectangleCentreLeft = GRectangle.TRectangleCentreLeft(pRect)
End Function

Public Function TRectangleCentreRight( _
                ByRef pRect As TRectangle) As TPoint
TRectangleCentreRight = GRectangle.TRectangleCentreRight(pRect)
End Function

Public Function TRectangleContainsPoint( _
                ByRef pRect As TRectangle, _
                ByVal X As Double, _
                ByVal Y As Double) As Boolean
TRectangleContainsPoint = GRectangle.TRectangleContainsPoint(pRect, X, Y)
End Function

Public Function TRectangleContainsRect( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
TRectangleContainsRect = GRectangle.TRectangleContainsRect(rect1, rect2)
End Function

Public Function TRectangleEquals( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
TRectangleEquals = GRectangle.TRectangleEquals(rect1, rect2)
End Function

Public Sub TRectangleExpand( _
                ByRef pRect As TRectangle, _
                ByVal xIncrement As Double, _
                ByVal yIncrement As Double)
GRectangle.TRectangleExpand pRect, xIncrement, yIncrement
End Sub


Public Sub TRectangleExpandByRotation( _
                ByRef pRect As TRectangle, _
                ByVal pAngle As Double, _
                ByVal pGraphics As Graphics)
GRectangle.TRectangleExpandByRotation pRect, pAngle, pGraphics
End Sub

Public Sub TRectangleExpandBySize( _
                ByRef pRect As TRectangle, _
                ByVal pSize As size, _
                ByVal pGraphics As Graphics)
GRectangle.TRectangleExpandBySize pRect, pSize, pGraphics
End Sub

Public Function TRectangleGetXInterval( _
                ByRef pRect As TRectangle) As TInterval
TRectangleGetXInterval = GRectangle.TRectangleGetXInterval(pRect)
End Function

Public Function TRectangleGetYInterval( _
                ByRef pRect As TRectangle) As TInterval
TRectangleGetYInterval = GRectangle.TRectangleGetYInterval(pRect)
End Function

Public Sub TRectangleInitialise( _
                ByRef pRect As TRectangle)
GRectangle.TRectangleInitialise pRect
End Sub

Public Function TRectangleIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TRectangle
TRectangleIntersection = GRectangle.TRectangleIntersection(rect1, rect2)
End Function

Public Function TRectangleOverlaps( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As Boolean
TRectangleOverlaps = GRectangle.TRectangleOverlaps(rect1, rect2)
End Function

Public Sub TRectangleSetFields( _
                ByRef pRect As TRectangle, _
                ByVal pLeft As Double, _
                ByVal pTop As Double, _
                ByVal pRight As Double, _
                ByVal pBottom As Double)
GRectangle.TRectangleSetFields pRect, pLeft, pTop, pRight, pBottom
End Sub

Public Sub TRectangleSetXInterval( _
                ByRef pRect As TRectangle, _
                ByRef pInterval As TInterval)
GRectangle.TRectangleSetXInterval pRect, pInterval
End Sub

Public Sub TRectangleSetYInterval( _
                ByRef pRect As TRectangle, _
                ByRef pInterval As TInterval)
GRectangle.TRectangleSetYInterval pRect, pInterval
End Sub

Public Function TRectangleTopCentre( _
                ByRef pRect As TRectangle) As TPoint
TRectangleTopCentre = GRectangle.TRectangleTopCentre(pRect)
End Function

Public Function TRectangleTopLeft( _
                ByRef pRect As TRectangle) As TPoint
TRectangleTopLeft = GRectangle.TRectangleTopLeft(pRect)
End Function

Public Function TRectangleTopRight( _
                ByRef pRect As TRectangle) As TPoint
TRectangleTopRight = GRectangle.TRectangleTopRight(pRect)
End Function

Public Function TRectangleToString( _
                ByRef pRect As TRectangle) As String
TRectangleToString = GRectangle.TRectangleToString(pRect)
End Function

Public Sub TRectangleTranslate( _
                ByRef pRect As TRectangle, _
                ByVal pdX As Double, _
                ByVal pdY As Double)
GRectangle.TRectangleTranslate pRect, pdX, pdY
End Sub

Public Sub TRectangleTranslateBySize( _
                ByRef pRect As TRectangle, _
                ByRef pOffset As size, _
                ByVal pGraphics As Graphics)
GRectangle.TRectangleTranslateBySize pRect, pOffset, pGraphics
End Sub

Public Sub TRectangleTranslatePoint( _
                ByRef pRect As TRectangle, _
                ByRef pOffset As TPoint)
GRectangle.TRectangleTranslateByPoint pRect, pOffset
End Sub

Public Function TRectangleUnion( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TRectangle
TRectangleUnion = GRectangle.TRectangleUnion(rect1, rect2)
End Function

Public Sub TRectangleValidate( _
                ByRef pRect As TRectangle, _
                Optional allowZeroDimensions As Boolean = False)
GRectangle.TRectangleValidate pRect, allowZeroDimensions
End Sub

Public Function TRectangleXIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TInterval
TRectangleXIntersection = GRectangle.TRectangleXIntersection(rect1, rect2)
End Function

Public Function TRectangleYIntersection( _
                ByRef rect1 As TRectangle, _
                ByRef rect2 As TRectangle) As TInterval
TRectangleYIntersection = GRectangle.TRectangleYIntersection(rect1, rect2)
End Function

'@================================================================================
' Helper Functions
'@================================================================================


