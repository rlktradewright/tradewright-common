VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TiledGradientBrush"
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

Implements IBrush
Implements ISyntheticBrush

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

Private Const ModuleName                            As String = "TiledGradientBrush"

Private Const ConfigSettingBrushBoundary            As String = "&BrushBoundary"
Private Const ConfigSettingDirection                As String = "&Direction"
Private Const ConfigSettingPad                      As String = "&Pad"
Private Const ConfigSettingTileMode                 As String = "&TileMode"
Private Const ConfigSettingColors                   As String = "&Colors"
Private Const ConfigSettingIntensities              As String = "&Intensities"
Private Const ConfigSettingPositions                As String = "&Positions"

'@================================================================================
' Member variables
'@================================================================================

Private mBrushBoundary                      As TRectangle
Private mBrushRect                          As GDI_RECT
Private mColors()                           As Long
Private mIntensities()                      As Double
Private mPositions()                        As Double
Private mVertices()                         As TRIVERTEX
Private mVerticesIndex                      As Long
Private mMeshRect()                         As GRADIENT_RECT
Private mMeshRectIndex                      As Long
Private mDirection                          As LinearGradientDirections
Private mTileMode                           As TileModes
Private mFillBoundary                       As GDI_RECT
Private mPad                                As Boolean

Private mGraphics                           As Graphics

Private mRectIndex                          As Long
Private mTileIndex                          As Long

Private mConfig                             As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

ReDim mColors(1) As Long
mColors(0) = vbRed
mColors(1) = &HFFFF00

ReDim mIntensities(1) As Double
mIntensities(0) = 1#
mIntensities(1) = 0#

ReDim mPositions(1) As Double
mPositions(0) = 0#
mPositions(1) = 1#

End Sub

'@================================================================================
' IBrush Interface Members
'@================================================================================

Private Property Let IBrush_ConfigurationSection(ByVal RHS As ConfigurationSection)
Const ProcName As String = "IBrush_ConfigurationSection"
On Error GoTo Err

ConfigurationSection = RHS

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IBrush_RemoveFromConfig()
Const ProcName As String = "IBrush_RemoveFromConfig"
On Error GoTo Err

RemoveFromConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ISyntheticBrush Interface Members
'@================================================================================

Private Sub ISyntheticBrush_Fill( _
                pRect As GDI_RECT, _
                ByVal pClipRegionHandle As Long, _
                ByVal pGraphics As Graphics)
Const ProcName As String = "ISyntheticBrush_Fill"
On Error GoTo Err

Clear

Set mGraphics = pGraphics

fillRect pRect

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

Private Property Let BrushBoundary( _
                ByRef Value As TRectangle)
Const ProcName As String = "BrushBoundary"
On Error GoTo Err

AssertArgument Value.isValid, "Not a valid TRectangle"

mBrushBoundary = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Colors() As Long()
Colors = mColors
End Property

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)

Const ProcName As String = "ConfigurationSection"
Dim failpoint As String
On Error GoTo Err

If Value Is Nothing Then
    RemoveFromConfig
    Exit Property
End If

Set mConfig = Value

gSetBrushClassInConfig Me, mConfig
mConfig.SetSetting ConfigSettingBrushBoundary, TRectangleToShortString(mBrushBoundary)
mConfig.SetSetting ConfigSettingDirection, CStr(mDirection)
mConfig.SetSetting ConfigSettingPad, CStr(mPad)
mConfig.SetSetting ConfigSettingTileMode, CStr(mTileMode)
mConfig.SetSetting ConfigSettingColors, gLongArrayToString(mColors)
mConfig.SetSetting ConfigSettingIntensities, gDoubleArrayToString(mIntensities)
mConfig.SetSetting ConfigSettingPositions, gDoubleArrayToString(mPositions)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Private Property Let Direction(ByVal Value As LinearGradientDirections)
Const ProcName As String = "Direction"
On Error GoTo Err

Select Case Value
Case LinearGradientDirectionHorizontal
Case LinearGradientDirectionVertical
Case Else
    AssertArgument False, "Invalid direction"
End Select
mDirection = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Direction() As LinearGradientDirections
Direction = mDirection
End Property

Public Property Get Intensities() As Double()
Intensities = mIntensities
End Property

Public Property Let Pad(ByVal Value As Boolean)
mPad = Value
End Property

Private Property Get Pad() As Boolean
Pad = mPad
End Property

Public Property Get Positions() As Double()
Positions = mPositions
End Property

Private Property Let TileMode(ByVal Value As TileModes)
Const ProcName As String = "TileMode"
On Error GoTo Err

Select Case Value
Case TileModeNone
Case TileModeTile
Case TileModeFlipX
Case TileModeFlipY
Case TileModeFlipXY
Case Else
    AssertArgument False, "Invalid tile mode"
End Select
mTileMode = Value
Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TileMode() As TileModes
TileMode = mTileMode
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByRef pBrushBoundary As TRectangle, _
                ByVal pDirection As LinearGradientDirections, _
                ByVal pPad As Boolean, _
                ByVal pTileMode As TileModes, _
                ByRef pColors() As Long, _
                ByRef pIntensities() As Double, _
                ByRef pPositions() As Double)
Const ProcName As String = "Initialise"
On Error GoTo Err

setColors pColors
setPositions pPositions
setIntensities pIntensities

Direction = pDirection
Pad = pPad
BrushBoundary = pBrushBoundary
TileMode = pTileMode

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName


End Sub

Friend Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
On Error GoTo Err

Set mConfig = pConfig

mBrushBoundary = TRectangleFromShortString(mConfig.GetSetting(ConfigSettingBrushBoundary), True)
mDirection = CLng(mConfig.GetSetting(ConfigSettingDirection))
mPad = CBool(mConfig.GetSetting(ConfigSettingPad))
mTileMode = CLng(mConfig.GetSetting(ConfigSettingTileMode))
mColors = gLongArrayFromString(mConfig.GetSetting(ConfigSettingColors))
mIntensities = gDoubleArrayFromString(mConfig.GetSetting(ConfigSettingIntensities))
mPositions = gDoubleArrayFromString(mConfig.GetSetting(ConfigSettingPositions))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As String
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove
Set mConfig = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addRect( _
                ByVal pUpperLeftVertexIndex As Long, _
                ByVal pLowerRightVertexIndex As Long)
If mMeshRectIndex > UBound(mMeshRect) Then ReDim Preserve mMeshRect(2 * (UBound(mMeshRect) + 1) - 1) As GRADIENT_RECT

mMeshRect(mMeshRectIndex).UpperLeft = pUpperLeftVertexIndex
mMeshRect(mMeshRectIndex).LowerRight = pLowerRightVertexIndex
mMeshRectIndex = mMeshRectIndex + 1
End Sub

Private Function addHorizontalRectangle() As Boolean
Dim lLeft As Long
Dim lRight As Long

Const ProcName As String = "addHorizontalRectangle"
On Error GoTo Err

nextRectHorizontal lLeft, lRight
If lLeft >= mFillBoundary.Right Then
    Exit Function
End If

If lRight > mFillBoundary.Left Then
    addVertex lLeft, _
            IIf(mTileMode = TileModeNone And Not mPad, mBrushRect.Top, mFillBoundary.Top), _
            getColor(mRectIndex - 1)
    
    
    addVertex lRight, _
            IIf(mTileMode = TileModeNone And Not mPad, mBrushRect.Bottom, mFillBoundary.Bottom), _
            getColor(mRectIndex)
    
    addRect mVerticesIndex - 2, mVerticesIndex - 1
End If

If reachedTileBoundary Then
    If mTileMode = TileModeNone Then
        addHorizontalRectangle = False
        Exit Function
    End If
    startNewTile
End If

addHorizontalRectangle = (lRight < mFillBoundary.Right)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub addVertex( _
                ByVal pX As Long, _
                ByVal pY As Long, _
                ByVal pColor As Long)
                
Const ProcName As String = "addVertex"
On Error GoTo Err

If mVerticesIndex > UBound(mVertices) Then ReDim Preserve mVertices(2 * (UBound(mVertices) + 1) - 1) As TRIVERTEX
mVertices(mVerticesIndex).X = pX
mVertices(mVerticesIndex).Y = pY
mVertices(mVerticesIndex).Red = pColor And &HFF&
mVertices(mVerticesIndex).Green = (pColor And &HFF00&) / &H100&
mVertices(mVerticesIndex).Blue = (pColor And &HFF0000) / &H10000
mVertices(mVerticesIndex).Alpha = 0

mVerticesIndex = mVerticesIndex + 1
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


Private Function addVerticalRectangle() As Boolean
Dim lBottom As Long
Dim lTop As Long

Const ProcName As String = "addVerticalRectangle"
On Error GoTo Err

nextRectVertical lBottom, lTop
If lBottom <= mFillBoundary.Top Then Exit Function

If lTop < mFillBoundary.Bottom Then
    addVertex IIf(mTileMode = TileModeNone And Not mPad, mBrushRect.Left, mFillBoundary.Left), _
            lBottom, _
            getColor(mRectIndex - 1)
    
    
    addVertex IIf(mTileMode = TileModeNone And Not mPad, mBrushRect.Right, mFillBoundary.Right), _
            lTop, _
            getColor(mRectIndex)
    
    addRect mVerticesIndex - 2, mVerticesIndex - 1
End If

If reachedTileBoundary Then
    If mTileMode = TileModeNone Then
        addVerticalRectangle = False
        Exit Function
    End If
    startNewTile
End If

addVerticalRectangle = (lTop > mFillBoundary.Top)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub calculateGradientFillParametersHorizontal()
Const ProcName As String = "calculateGradientFillParametersHorizontal"
On Error GoTo Err

setFirstHorizontalTileIndex

Do While addHorizontalRectangle
Loop

If mTileMode = TileModeNone Then
    padLeft
    padRight
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub calculateGradientFillParametersVertical()
Const ProcName As String = "calculateGradientFillParametersVertical"
On Error GoTo Err

setFirstVerticalTileIndex

Do While addVerticalRectangle
Loop

If mTileMode = TileModeNone Then
    padAbove
    padBelow
End If


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub Clear()
ReDim mVertices(3) As TRIVERTEX
mVerticesIndex = 0
ReDim mMeshRect(3) As GRADIENT_RECT
mMeshRectIndex = 0

mRectIndex = 0
mTileIndex = 0
End Sub

Private Sub fillRect( _
                ByRef pRect As GDI_RECT)
Const ProcName As String = "fillRect"
On Error GoTo Err

mFillBoundary = pRect
mBrushRect = mGraphics.ConvertTRectangleToGdiRect(mBrushBoundary)

If mDirection = LinearGradientDirectionHorizontal Then
    calculateGradientFillParametersHorizontal
Else
    calculateGradientFillParametersVertical
End If

GradientFill mGraphics.hDC, _
            mVertices(0), _
            mVerticesIndex, _
            VarPtr(mMeshRect(0)), _
            mMeshRectIndex, _
            IIf(mDirection = LinearGradientDirectionHorizontal, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getColor( _
                ByVal pPosition As Long) As Long
If isFlip Then pPosition = UBound(mPositions) - pPosition
If UBound(mColors) = 1 Then
    getColor = interpolateColors(gNormalizeColor(mColors(0)), gNormalizeColor(mColors(1)), mIntensities(pPosition))
Else
    getColor = gNormalizeColor(mColors(pPosition))
End If
End Function

Private Function interpolateColors( _
                ByVal pColor1 As Long, _
                ByVal pColor2 As Long, _
                ByVal pIntensity As Double) As Long
Dim pComplementaryIntensity As Double
pComplementaryIntensity = 1 - pIntensity
interpolateColors = pIntensity * (pColor1 And &HFF&) + pComplementaryIntensity * (pColor2 And &HFF&) + _
                    ((pIntensity * (pColor1 And &HFF00&)) And &HFF00&) + ((pComplementaryIntensity * (pColor2 And &HFF00&)) And &HFF00&) + _
                    ((pIntensity * (pColor1 And &HFF0000)) And &HFF0000) + ((pComplementaryIntensity * (pColor2 And &HFF0000)) And &HFF0000)
End Function

Private Function isFlip() As Boolean
If mTileMode = TileModeFlipX Or mTileMode = TileModeFlipY Or mTileMode = TileModeFlipXY Then isFlip = (mTileIndex Mod 2 <> 0)
End Function

Private Function nextX() As Long
If isFlip Then
    nextX = mBrushRect.Left + (mTileIndex + (1 - mPositions(UBound(mPositions) - mRectIndex))) * (mBrushRect.Right - mBrushRect.Left)
Else
    nextX = mBrushRect.Left + (mTileIndex + mPositions(mRectIndex)) * (mBrushRect.Right - mBrushRect.Left)
End If
End Function

Private Function nextY() As Long
If isFlip Then
    nextY = mBrushRect.Bottom + (mTileIndex + (1 - mPositions(UBound(mPositions) - mRectIndex))) * (mBrushRect.Top - mBrushRect.Bottom)
Else
    nextY = mBrushRect.Bottom + (mTileIndex + mPositions(mRectIndex)) * (mBrushRect.Top - mBrushRect.Bottom)
End If
End Function

Private Sub padAbove()
Const ProcName As String = "padAbove"
On Error GoTo Err

If mPad And mBrushRect.Top > mFillBoundary.Top Then
    addVertex mFillBoundary.Left, _
            mBrushRect.Top, _
            getColor(UBound(mPositions))
    
    addVertex mFillBoundary.Right, _
            mFillBoundary.Top, _
            getColor(UBound(mPositions))
    
    addRect mVerticesIndex - 2, mVerticesIndex - 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub padBelow()
Const ProcName As String = "padBelow"
On Error GoTo Err

If mRectIndex = 0 And mPad And mBrushRect.Bottom < mFillBoundary.Bottom Then
    addVertex mFillBoundary.Left, _
            mFillBoundary.Bottom, _
            getColor(0)
    
    addVertex mFillBoundary.Right, _
            mBrushRect.Bottom, _
            getColor(0)
    
    addRect mVerticesIndex - 2, mVerticesIndex - 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub padLeft()
Const ProcName As String = "padLeft"
On Error GoTo Err

If mRectIndex = 0 And mPad And mBrushRect.Left > mFillBoundary.Left Then
    addVertex mFillBoundary.Left, _
            mFillBoundary.Top, _
            getColor(0)
    
    addVertex mBrushRect.Left, _
            mFillBoundary.Bottom, _
            getColor(0)
    
    addRect mVerticesIndex - 2, mVerticesIndex - 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub padRight()
Const ProcName As String = "padRight"
On Error GoTo Err

If mPad And mBrushRect.Right < mFillBoundary.Right Then
    addVertex mBrushRect.Right, _
            mFillBoundary.Top, _
            getColor(UBound(mPositions))
    
    addVertex mFillBoundary.Right, _
            mFillBoundary.Bottom, _
            getColor(UBound(mPositions))
    
    addRect mVerticesIndex - 2, mVerticesIndex - 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function reachedTileBoundary() As Boolean
reachedTileBoundary = (mRectIndex = UBound(mPositions))
End Function

Private Sub startNewTile()
mRectIndex = 0
mTileIndex = mTileIndex + 1
End Sub

Private Sub nextRectHorizontal( _
                ByRef pLeft As Long, _
                ByRef pRight As Long)
pLeft = nextX
mRectIndex = mRectIndex + 1
pRight = nextX
End Sub

Private Sub nextRectVertical( _
                ByRef pBottom As Long, _
                ByRef pTop As Long)
pBottom = nextY
mRectIndex = mRectIndex + 1
pTop = nextY
End Sub

Private Sub setColors(ByRef pColors() As Long)
Const ProcName As String = "setColors"
On Error GoTo Err

On Error Resume Next
Dim l As Long

l = -1
l = UBound(pColors)
On Error GoTo Err

AssertArgument l >= 1, "pColors must have at least two elements"

Dim i As Long
For i = 0 To UBound(pColors)
    AssertArgument gIsValidColor(pColors(i)), "pColor(" & i & ") is not a valid color"
Next

mColors = pColors

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setFirstHorizontalTileIndex()
If mTileMode = TileModeNone Then
    mTileIndex = 0
Else
    mTileIndex = Int((mFillBoundary.Left - mBrushRect.Left) / (mBrushRect.Right - mBrushRect.Left))
End If
End Sub

Private Sub setFirstVerticalTileIndex()
If mTileMode = TileModeNone Then
    mTileIndex = 0
Else
    mTileIndex = Int((mFillBoundary.Bottom - mBrushRect.Bottom) / (mBrushRect.Top - mBrushRect.Bottom))
End If
End Sub

Private Sub setIntensities(ByRef pIntensities() As Double)
Const ProcName As String = "setIntensities"
On Error GoTo Err

If UBound(mColors) > 1 Then Exit Sub

On Error Resume Next
Dim l As Long
l = -1
l = UBound(pIntensities)
On Error GoTo Err

Dim i As Long
If l = -1 Then
    ReDim mIntensities(UBound(mPositions)) As Double
    For i = 0 To UBound(mPositions)
        mIntensities(i) = 1 - i / UBound(mPositions)
    Next
Else
    AssertArgument l = UBound(mPositions), "pIntensities must be null or have the same number of elements as pPositions"
    
    For i = 0 To UBound(pIntensities)
        AssertArgument pIntensities(i) >= 0# And pIntensities(i) <= 1#, "pIntensities(" & i & ") is not between 0.0 and 1.0"
    Next
    
    mIntensities = pIntensities
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPositions(ByRef pPositions() As Double)
Const ProcName As String = "setPositions"
On Error GoTo Err

On Error Resume Next
Dim l As Long
l = -1
l = UBound(pPositions)
On Error GoTo Err

AssertArgument l >= 1, "pPositions must have at least two elements"

Dim i As Long
For i = 0 To UBound(pPositions)
    AssertArgument pPositions(i) >= 0# And pPositions(i) <= 1#, "pPosition(" & i & ") is not between 0.0 and 1.0"
Next

mPositions = pPositions
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



