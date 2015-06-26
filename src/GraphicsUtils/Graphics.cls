VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Description
'================================================================================
'
'
'================================================================================
' Amendment history
'================================================================================
'
'
'
'

'================================================================================
' Interfaces
'================================================================================

'================================================================================
' Events
'================================================================================

Event Click()
Event DblClick()
Event Change(ByRef ev As ChangeEventData)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseWheel(Distance As Single)
Event Resize()

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                As String = "Graphics"

' Used to indicate that there was no currently defined clip region when
' SaveClippingRegion was called
Private Const NoClipRegion              As Long = -1

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' External FUnction Declarations
'================================================================================

'================================================================================
' Member variables
'================================================================================

' Text attributes
Private mTextColor                      As Long

Private mClippingRegionHandleStack      As LongStack

Private mBoundary                       As TRectangle
Private mGaugeX                         As Double
Private mGaugeY                         As Double

Private mTransform                      As XForm

Private mhWnd                           As Long
Private mTargethDC                      As Long
Private mhDC                            As Long
Private mhOrigBitmap                    As Long
Private mClientRect                     As GDI_RECT
Private mClientRectLogical              As GDI_RECT
Private mHeightCm                       As Double
Private mWidthCm                        As Double

Private mScaleFactorX                   As Double
Private mScaleFactorY                   As Double

Private mOrigWindowProcAddr             As Long

Private mMixMode                        As MixModes

Private mBackgroundBrush                As IBrush

Private mPrevBrushHandle                As Long
Private mPrevPenHandle                  As Long

Private WithEvents mFont                As StdFont
Attribute mFont.VB_VarHelpID = -1

Private mPrevFontHandle                 As Long
Private mCurrFontHandle                 As Long

Private mCurrX                          As Single
Private mCurrY                          As Single

Private mDeterminingBoundingRectangle   As Boolean

Private mIsInDev                        As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()
Debug.Print ProjectName & " in development environment: " & CStr(setInDev)
Set mClippingRegionHandleStack = New LongStack
End Sub

Private Sub Class_Terminate()
Finish
Debug.Print ModuleName & " terminated"
End Sub

'================================================================================
' XXXX Interface Members
'================================================================================

'================================================================================
' mFont Event Handlers
'================================================================================

Private Sub mFont_FontChanged(ByVal PropertyName As String)
Set mFont = Nothing
End Sub

'================================================================================
' Properties
'================================================================================

Friend Property Get AspectRatio() As Double
AspectRatio = Abs(mScaleFactorX / mScaleFactorY)
End Property

Public Property Let BackgroundBrush(ByVal Value As IBrush)
Set mBackgroundBrush = Value
End Property

Public Property Get BackgroundBrush() As IBrush
Set BackgroundBrush = mBackgroundBrush
End Property

Public Property Let Bottom(ByVal Value As Double)
Const ProcName As String = "Bottom"
Dim failpoint As String
On Error GoTo Err

setDimensions mBoundary.Left, Value, mBoundary.Right, mBoundary.Top

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Bottom() As Double
Const ProcName As String = "Bottom"
Dim failpoint As String
On Error GoTo Err

Bottom = mBoundary.Bottom

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Boundary() As TRectangle
Boundary = mBoundary
End Property

Friend Property Get GaugeX() As Double
GaugeX = mGaugeX
End Property

Friend Property Get GaugeY() As Double
GaugeY = mGaugeY
End Property

Friend Property Get hDC() As Long
hDC = mhDC
End Property

Public Property Let Height(ByVal Value As Double)
Const ProcName As String = "Height"
Dim failpoint As String
On Error GoTo Err

setDimensions mBoundary.Left, mBoundary.Top - Value, mBoundary.Right, mBoundary.Top

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Height() As Double
Const ProcName As String = "Height"
Dim failpoint As String
On Error GoTo Err

If mBoundary.Top = MaxDouble Or mBoundary.Bottom = MaxDouble Then
    Height = MaxDouble
ElseIf mBoundary.Top = MinDouble Or mBoundary.Bottom = MinDouble Then
    Height = MaxDouble
Else
    Height = mBoundary.Top - mBoundary.Bottom
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HeightCm() As Double
Const ProcName As String = "HeightCm"
Dim failpoint As String
On Error GoTo Err

HeightCm = mHeightCm

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HeightPixels() As Long
Const ProcName As String = "HeightPixels"
Dim failpoint As String
On Error GoTo Err

HeightPixels = mClientRect.Bottom - mClientRect.Top + 1

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Left(ByVal Value As Double)
Const ProcName As String = "Left"
Dim failpoint As String
On Error GoTo Err

setDimensions Value, mBoundary.Bottom, mBoundary.Right, mBoundary.Top

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Left() As Double
Const ProcName As String = "Left"
Dim failpoint As String
On Error GoTo Err

Left = mBoundary.Left

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let MixMode(ByVal Value As MixModes)
Const ProcName As String = "MixMode"
On Error GoTo Err

Select Case Value
Case MixModeBlack

Case MixModeCopyPen

Case MixModeInvert

Case MixModeMaskNotPen

Case MixModeMaskPen

Case MixModeMaskPenNot

Case MixModeMergeNotPen

Case MixModeMergePen

Case MixModeMergePenNot

Case MixModeNop

Case MixModeNotCopyPen

Case MixModeNotMaskPen

Case MixModeNotMergePen

Case MixModeNotXorPen

Case MixModeWhite

Case MixModeXorPen

Case Else
    AssertArgument False, "Value must be a member of the MixModes enum"
End Select

mMixMode = Value

If SetROP2(mhDC, mMixMode) = 0 Then gHandleWin32Error GetLastError

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get MixMode() As MixModes
MixMode = mMixMode
End Property

Public Property Get MousePosition() As Point
Const ProcName As String = "MousePosition"
Dim failpoint As String
On Error GoTo Err

Set MousePosition = gNewPoint(mCurrX, mCurrY)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Right(ByVal Value As Double)
Const ProcName As String = "Right"
Dim failpoint As String
On Error GoTo Err

setDimensions mBoundary.Left, mBoundary.Bottom, Value, mBoundary.Top

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Right() As Double
Const ProcName As String = "Right"
Dim failpoint As String
On Error GoTo Err

Right = mBoundary.Right

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Top(ByVal Value As Double)
Const ProcName As String = "Top"
Dim failpoint As String
On Error GoTo Err

setDimensions mBoundary.Left, mBoundary.Bottom, mBoundary.Right, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Top() As Double
Top = mBoundary.Top
End Property

Public Property Get Width() As Double
Const ProcName As String = "Width"
Dim failpoint As String
On Error GoTo Err

If mBoundary.Right = MaxDouble Or mBoundary.Left = MaxDouble Then
    Width = MaxDouble
ElseIf mBoundary.Right = MinDouble Or mBoundary.Left = MinDouble Then
    Width = MaxDouble
Else
    Width = mBoundary.Right - mBoundary.Left
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Width(ByVal Value As Double)
Const ProcName As String = "Width"
Dim failpoint As String
On Error GoTo Err

setDimensions mBoundary.Right - Value, mBoundary.Bottom, mBoundary.Right, mBoundary.Top

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get WidthCm() As Double
Const ProcName As String = "WidthCm"
Dim failpoint As String
On Error GoTo Err

WidthCm = mWidthCm

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get WidthPixels() As Long
Const ProcName As String = "WidthPixels"
Dim failpoint As String
On Error GoTo Err

WidthPixels = mClientRect.Right - mClientRect.Left + 1

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Methods
'================================================================================

Public Sub AddToClippingRegion( _
                ByRef pRect As TRectangle)
Const ProcName As String = "AddToClippingRegion"
On Error GoTo Err

If Not pRect.isValid Then Exit Sub

CombineClippingRegionGdi ConvertTRectangleToGdiRect(pRect), RegionSelectionModeOr

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub BeginBoundingRectangleDetermination()
Const ProcName As String = "BeginBoundingRectangleDetermination"
Dim failpoint As String
On Error GoTo Err

If BeginPath(mhDC) = 0 Then gHandleWin32Error GetLastError
mDeterminingBoundingRectangle = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearClippingRegion()
Const ProcName As String = "ClearClippingRegion"
Dim failpoint As String
On Error GoTo Err

If SelectClipRgn(mhDC, 0) = 0 Then gHandleWin32Error GetLastError

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ConvertDistanceToLogicalX(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertDistanceToLogicalX = MaxDouble
ElseIf Value = MinDouble Then
    ConvertDistanceToLogicalX = MinDouble
Else
    ConvertDistanceToLogicalX = (Value / WidthCm) * (mBoundary.Right - mBoundary.Left)
End If
End Function

Public Function ConvertDistanceToLogicalY(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertDistanceToLogicalY = MaxDouble
ElseIf Value = MinDouble Then
    ConvertDistanceToLogicalY = MinDouble
Else
    ConvertDistanceToLogicalY = (Value / HeightCm) * (mBoundary.Top - mBoundary.Bottom)
End If
End Function

Public Function ConvertDistanceToPixelsX(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertDistanceToPixelsX = MaxLong
ElseIf Value = MinDouble Then
    ConvertDistanceToPixelsX = MinLong
Else
    ConvertDistanceToPixelsX = (Value / WidthCm) * (mClientRect.Right - mClientRect.Left + 1)
End If
End Function

Public Function ConvertDistanceToPixelsY(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertDistanceToPixelsY = MaxLong
ElseIf Value = MinDouble Then
    ConvertDistanceToPixelsY = MinLong
Else
    ConvertDistanceToPixelsY = (Value / HeightCm) * (mClientRect.Bottom - mClientRect.Top + 1)
End If
End Function

Public Function ConvertDistanceToRelativeX(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertDistanceToRelativeX = MaxDouble
ElseIf Value = MinDouble Then
    ConvertDistanceToRelativeX = MinDouble
Else
    ConvertDistanceToRelativeX = 100# * Value / WidthCm
End If
End Function

Public Function ConvertDistanceToRelativeY(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertDistanceToRelativeY = MaxDouble
ElseIf Value = MinDouble Then
    ConvertDistanceToRelativeY = MinDouble
Else
    ConvertDistanceToRelativeY = 100# * Value / HeightCm
End If
End Function

Public Function ConvertLogicalToDistanceX(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertLogicalToDistanceX = MaxDouble
ElseIf Value = MinDouble Then
    ConvertLogicalToDistanceX = MinDouble
Else
    ConvertLogicalToDistanceX = Value / (mBoundary.Right - mBoundary.Left) * WidthCm
End If
End Function

Public Function ConvertLogicalToDistanceY(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertLogicalToDistanceY = MaxDouble
ElseIf Value = MinDouble Then
    ConvertLogicalToDistanceY = MinDouble
Else
    ConvertLogicalToDistanceY = Value / (mBoundary.Top - mBoundary.Bottom) * HeightCm
End If
End Function

Friend Function ConvertLogicalToGdiX(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertLogicalToGdiX = MaxLong
ElseIf Value = MinDouble Then
    ConvertLogicalToGdiX = MinLong
Else
    ConvertLogicalToGdiX = mScaleFactorX * Value
End If
End Function

Friend Function ConvertLogicalToGdiY(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertLogicalToGdiY = MaxLong
ElseIf Value = MinDouble Then
    ConvertLogicalToGdiY = MinLong
Else
    ConvertLogicalToGdiY = mScaleFactorY * Value
End If
End Function

Public Function ConvertLogicalToPixelsX(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertLogicalToPixelsX = MaxLong
ElseIf Value = MinDouble Then
    ConvertLogicalToPixelsX = MinLong
Else
    ConvertLogicalToPixelsX = Value / (mBoundary.Right - mBoundary.Left) * (mClientRect.Right - mClientRect.Left + 1)
End If
End Function

Public Function ConvertLogicalToPixelsY(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertLogicalToPixelsY = MaxLong
ElseIf Value = MinDouble Then
    ConvertLogicalToPixelsY = MinLong
Else
    ConvertLogicalToPixelsY = Value / (mBoundary.Top - mBoundary.Bottom) * (mClientRect.Bottom - mClientRect.Top + 1)
End If
End Function

Public Function ConvertLogicalToRelativeX(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertLogicalToRelativeX = MaxDouble
ElseIf Value = MinDouble Then
    ConvertLogicalToRelativeX = MinDouble
Else
    ConvertLogicalToRelativeX = Value / (mBoundary.Right - mBoundary.Left)
End If
End Function

Public Function ConvertLogicalToRelativeY(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertLogicalToRelativeY = MaxDouble
ElseIf Value = MinDouble Then
    ConvertLogicalToRelativeY = MinDouble
Else
    ConvertLogicalToRelativeY = Value / (mBoundary.Top - mBoundary.Bottom)
End If
End Function

Public Function ConvertLogicalXtoLogicalY(ByVal Value As Double) As Double
ConvertLogicalXtoLogicalY = ConvertGdiToLogicalY(ConvertLogicalToGdiX(Value))
End Function

Public Function ConvertLogicalYtoLogicalX(ByVal Value As Double) As Double
ConvertLogicalYtoLogicalX = ConvertGdiToLogicalX(ConvertLogicalToGdiY(Value))
End Function

Public Function ConvertPixelsToDistanceX(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertPixelsToDistanceX = MaxDouble
ElseIf Value = MinDouble Then
    ConvertPixelsToDistanceX = MinDouble
Else
    ConvertPixelsToDistanceX = Value / (mClientRect.Right - mClientRect.Left + 1) * WidthCm
End If
End Function

Public Function ConvertPixelsToDistanceY(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertPixelsToDistanceY = MaxDouble
ElseIf Value = MinDouble Then
    ConvertPixelsToDistanceY = MinDouble
Else
    ConvertPixelsToDistanceY = Value / (mClientRect.Bottom - mClientRect.Top + 1) * HeightCm
End If
End Function

Public Function ConvertPixelsToLogicalX(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertPixelsToLogicalX = MaxDouble
ElseIf Value = MinLong Then
    ConvertPixelsToLogicalX = MinDouble
Else
    ConvertPixelsToLogicalX = Value * LogicalUnitsPerPixel / mScaleFactorX
End If
End Function

Public Function ConvertPixelsToLogicalY(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertPixelsToLogicalY = MaxDouble
ElseIf Value = MinLong Then
    ConvertPixelsToLogicalY = MinDouble
Else
    ConvertPixelsToLogicalY = Value * LogicalUnitsPerPixel / mScaleFactorY
End If
End Function

Public Function ConvertPixelsToRelativeX(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertPixelsToRelativeX = MaxDouble
ElseIf Value = MinDouble Then
    ConvertPixelsToRelativeX = MinDouble
Else
    ConvertPixelsToRelativeX = Value / (mClientRect.Right - mClientRect.Left + 1)
End If
End Function

Public Function ConvertPixelsToRelativeY(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertPixelsToRelativeY = MaxDouble
ElseIf Value = MinDouble Then
    ConvertPixelsToRelativeY = MinDouble
Else
    ConvertPixelsToRelativeY = Value / (mClientRect.Bottom - mClientRect.Top + 1)
End If
End Function

Public Function ConvertRelativeToDistanceX(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertRelativeToDistanceX = MaxDouble
ElseIf Value = MinDouble Then
    ConvertRelativeToDistanceX = MinDouble
Else
    ConvertRelativeToDistanceX = Value * WidthCm
End If
End Function

Public Function ConvertRelativeToDistanceY(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertRelativeToDistanceY = MaxDouble
ElseIf Value = MinDouble Then
    ConvertRelativeToDistanceY = MinDouble
Else
    ConvertRelativeToDistanceY = Value * HeightCm
End If
End Function

Public Function ConvertRelativeToLogicalX(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertRelativeToLogicalX = MaxDouble
ElseIf Value = MinDouble Then
    ConvertRelativeToLogicalX = MinDouble
Else
    ConvertRelativeToLogicalX = Value * (mBoundary.Right - mBoundary.Left)
End If
End Function

Public Function ConvertRelativeToLogicalY(ByVal Value As Double) As Double
If Value = MaxDouble Then
    ConvertRelativeToLogicalY = MaxDouble
ElseIf Value = MinDouble Then
    ConvertRelativeToLogicalY = MinDouble
Else
    ConvertRelativeToLogicalY = Value * (mBoundary.Top - mBoundary.Bottom)
End If
End Function

Public Function ConvertRelativeToPixelsX(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertRelativeToPixelsX = MaxLong
ElseIf Value = MinDouble Then
    ConvertRelativeToPixelsX = MinLong
Else
    ConvertRelativeToPixelsX = Value * (mClientRect.Right - mClientRect.Left + 1)
End If
End Function

Public Function ConvertRelativeToPixelsY(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertRelativeToPixelsY = MaxLong
ElseIf Value = MinDouble Then
    ConvertRelativeToPixelsY = MinLong
Else
    ConvertRelativeToPixelsY = Value * (mClientRect.Bottom - mClientRect.Top + 1)
End If
End Function

Friend Function ConvertTRectangleToGdiRect( _
                ByRef pTRect As TRectangle) As GDI_RECT
GdiRectangleSetFieldsByPositionAndSize ConvertTRectangleToGdiRect, _
                    ConvertCoordsLogicalToGdiX(pTRect.Left), _
                    ConvertCoordsLogicalToGdiY(pTRect.Top), _
                    ConvertLogicalToGdiX(pTRect.Right - pTRect.Left), _
                    ConvertLogicalToGdiY(pTRect.Top - pTRect.Bottom)
End Function

Public Sub DrawCircle( _
                ByVal pPen As Pen, _
                ByVal pCentre As Point, _
                ByVal pRadius As Dimension)
Const ProcName As String = "DrawCircle"
Dim failpoint As String
On Error GoTo Err

DrawEllipseFromTRectangle pPen, circleToTRectangle(pCentre, pRadius, True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawClippedText( _
                ByVal pText As String, _
                ByVal pBrush As IBrush, _
                ByVal pFont As StdFont, _
                ByRef pClippingRect As TRectangle, _
                Optional ByVal pOutlinePen As Pen = Nothing)
Const ProcName As String = "DrawClippedText"
On Error GoTo Err

DrawFormattedText pText, pBrush, pFont, pClippingRect, pOutlinePen, JustifyLeft, False, EllipsisNone, False, 0, False, Nothing, Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawEllipse( _
                ByVal pPen As Pen, _
                ByVal pPoint1 As Point, _
                ByVal pPoint2 As Point)
Const ProcName As String = "DrawEllipse"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, _
                pPoint1.XLogical(Me), _
                pPoint1.YLogical(Me), _
                pPoint2.XLogical(Me), _
                pPoint2.YLogical(Me), _
                True

DrawEllipseFromTRectangle pPen, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawEllipseFromPositionAndSize( _
                ByVal pPen As Pen, _
                ByVal pPosition As Point, _
                ByVal pSize As size)
Const ProcName As String = "DrawEllipseFromPositionAndSize"
Dim failpoint As String
On Error GoTo Err

DrawEllipseFromTRectangle pPen, positionAndSizeToTRectangle(pPosition, pSize, True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawEllipseFromTRectangle( _
                ByVal pPen As Pen, _
                ByRef pRect As TRectangle)
Const ProcName As String = "DrawEllipseFromTRectangle"
Dim failpoint As String
On Error GoTo Err

AssertArgument pRect.isValid, "Argument is not a valid TRectangle"

DrawEllipseGdi pPen, ConvertTRectangleToGdiRect(pRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DrawEllipseGdi( _
                ByVal pPen As Pen, _
                ByRef pRect As GDI_RECT)
Const ProcName As String = "DrawEllipseGdi"
Dim failpoint As String
On Error GoTo Err

usePen pPen
useBrush Nothing

If Ellipse(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
invalidate
releaseBrush
ReleasePen

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawEllipseLogical( _
                ByVal pPen As Pen, _
                ByVal x1Logical As Double, _
                ByVal y1Logical As Double, _
                ByVal x2Logical As Double, _
                ByVal y2Logical As Double)
Const ProcName As String = "DrawEllipseLogical"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, x1Logical, y1Logical, x2Logical, y2Logical, True
DrawEllipseFromTRectangle pPen, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawFormattedText( _
                ByVal pText As String, _
                ByVal pBrush As IBrush, _
                ByVal pFont As StdFont, _
                ByRef pClippingRect As TRectangle, _
                Optional ByVal pOutlinePen As Pen = Nothing, _
                Optional ByVal pJustification As TextJustifyModes = TextJustifyModes.JustifyLeft, _
                Optional ByVal pMultiLine As Boolean = True, _
                Optional ByVal pEllipsis As EllipsisModes = EllipsisNone, _
                Optional ByVal pExpandTabs As Boolean = True, _
                Optional ByVal pTabWidth As Long = 8, _
                Optional ByVal pWordWrap As Boolean = True, _
                Optional ByVal pLeftMargin As Dimension = Nothing, _
                Optional ByVal pRightMargin As Dimension = Nothing)
Const ProcName As String = "DrawFormattedText"
Dim failpoint As String
On Error GoTo Err

Dim lClipRect As GDI_RECT

lClipRect = ConvertTRectangleToGdiRect(pClippingRect)

setFont pFont

If TypeOf pBrush Is SolidBrush Then
    Dim lSolidBrush As SolidBrush
    Set lSolidBrush = pBrush
    doDrawFormattedText pText, lSolidBrush.Color, pFont, lClipRect, pJustification, pMultiLine, pEllipsis, pExpandTabs, pTabWidth, pWordWrap, pLeftMargin, pRightMargin
ElseIf TypeOf pBrush Is ISyntheticBrush Then
    Dim lNullRect As GDI_RECT
    
    If Not mDeterminingBoundingRectangle Then
        
        If BeginPath(mhDC) = 0 Then gHandleWin32Error GetLastError
        doDrawFormattedText pText, vbBlack, pFont, lClipRect, pJustification, pMultiLine, pEllipsis, pExpandTabs, pTabWidth, pWordWrap, pLeftMargin, pRightMargin
        EndPath mhDC
        
        syntheticBrushFill pBrush, lClipRect
    End If
End If

If mDeterminingBoundingRectangle Then
    If pOutlinePen Is Nothing Then Set pOutlinePen = gPathPen
    DrawRectangleGdi pOutlinePen, lClipRect
Else
    If Not pOutlinePen Is Nothing Then
        usePen pOutlinePen
        
        If BeginPath(mhDC) = 0 Then gHandleWin32Error GetLastError
        doDrawFormattedText pText, vbBlack, pFont, lClipRect, pJustification, pMultiLine, pEllipsis, pExpandTabs, pTabWidth, pWordWrap, pLeftMargin, pRightMargin
        EndPath mhDC
        
        If StrokePath(mhDC) = 0 Then gHandleWin32Error GetLastError
        
        ReleasePen
End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawLine( _
                ByVal pPen As Pen, _
                ByVal pPoint1 As Point, _
                ByVal pPoint2 As Point)
Const ProcName As String = "DrawLine"
Dim failpoint As String
On Error GoTo Err

DrawLineGdi pPen, pointToGdiPoint(pPoint1), pointToGdiPoint(pPoint2)
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub DrawLineGdi( _
                ByVal pPen As Pen, _
                ByRef pPoint1 As GDI_POINT, _
                ByRef pPoint2 As GDI_POINT)
Const ProcName As String = "DrawLineGdi"
Dim failpoint As String
On Error GoTo Err

usePen pPen

If MoveToEx(mhDC, pPoint1.X, pPoint1.Y, 0) = 0 Then gHandleWin32Error GetLastError
If LineTo(mhDC, pPoint2.X, pPoint2.Y) = 0 Then gHandleWin32Error GetLastError

ReleasePen

invalidate

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawPoint( _
                ByVal pPen As Pen, _
                ByVal pPoint As Point)
Const ProcName As String = "DrawPoint"
Dim failpoint As String
On Error GoTo Err

DrawLine pPen, pPoint, pPoint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawPolygon( _
                ByVal pPen As Pen, _
                ByRef pPoints() As Point)
Const ProcName As String = "DrawPolygon"
Dim failpoint As String
On Error GoTo Err

DrawPolygonGdi pPen, ConvertPointsToGdiPoints(pPoints)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawPolygonFromPositionAndSize( _
                ByVal pPen As Pen, _
                ByVal pNumberOfSides As Long, _
                ByVal pPosition As Point, _
                ByVal pSize As size)
Const ProcName As String = "DrawPolygonFromPositionAndSize"
On Error GoTo Err

DrawPolygon pPen, positionAndSizeToPoints(pNumberOfSides, pPosition, pSize)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub DrawPolygonGdi( _
                ByVal pPen As Pen, _
                ByRef pPoints() As GDI_POINT)
Const ProcName As String = "DrawPolygonGdi"
Dim failpoint As String
On Error GoTo Err

usePen pPen
useBrush Nothing

If Polygon(mhDC, VarPtr(pPoints(0)), UBound(pPoints) + 1) = 0 Then gHandleWin32Error GetLastError
releaseBrush
ReleasePen

invalidate
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawRectangle( _
                ByVal pPen As Pen, _
                ByVal pPoint1 As Point, _
                ByVal pPoint2 As Point)
Const ProcName As String = "DrawRectangle"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, _
                pPoint1.XLogical(Me), _
                pPoint1.YLogical(Me), _
                pPoint2.XLogical(Me), _
                pPoint2.YLogical(Me), _
                True

DrawRectangleFromTRectangle pPen, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawRectangleFromPositionAndSize( _
                ByVal pPen As Pen, _
                ByVal pPosition As Point, _
                ByVal pSize As size)
Const ProcName As String = "DrawRectangleFromPositionAndSize"
Dim failpoint As String
On Error GoTo Err

DrawRectangleFromTRectangle pPen, positionAndSizeToTRectangle(pPosition, pSize, True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawRectangleFromTRectangle( _
                ByVal pPen As Pen, _
                ByRef pRect As TRectangle)
Const ProcName As String = "DrawRectangleFromTRectangle"
Dim failpoint As String
On Error GoTo Err

AssertArgument pRect.isValid, "Argument is not a valid TRectangle"

DrawRectangleGdi pPen, ConvertTRectangleToGdiRect(pRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub DrawRectangleGdi( _
                ByVal pPen As Pen, _
                ByRef pRect As GDI_RECT)
Const ProcName As String = "DrawRectangleGdi"
Dim failpoint As String
On Error GoTo Err

usePen pPen
useBrush Nothing

If Rectangle(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
releaseBrush
ReleasePen

invalidate
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawRectangleLogical( _
                ByVal pPen As Pen, _
                ByVal x1Logical As Double, _
                ByVal y1Logical As Double, _
                ByVal x2Logical As Double, _
                ByVal y2Logical As Double)
Const ProcName As String = "DrawRectangleLogical"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, x1Logical, y1Logical, x2Logical, y2Logical, True
DrawRectangleFromTRectangle pPen, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub DrawText( _
                ByVal pText As String, _
                ByVal pBrush As IBrush, _
                ByVal pFont As StdFont, _
                ByVal pPoint As Point, _
                Optional ByVal pOutlinePen As Pen = Nothing)
Dim lClip As TRectangle
Const ProcName As String = "DrawText"
On Error GoTo Err

lClip = mBoundary
lClip.Left = pPoint.XLogical(Me)
lClip.Top = pPoint.YLogical(Me)
DrawFormattedText pText, pBrush, pFont, lClip, pOutlinePen, JustifyLeft, False, EllipsisNone, False, 0, False, Nothing, Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function EndBoundingRectangleDetermination() As TRectangle
Const ProcName As String = "EndBoundingRectangleDetermination"
Dim failpoint As String
On Error GoTo Err

Dim lXForm As XForm
Dim lRegion As Long
Dim lRect As GDI_RECT

mDeterminingBoundingRectangle = False

EndPath mhDC
WidenPath mhDC

lRegion = PathToRegion(mhDC)
GetRgnBox lRegion, lRect
DeleteObject lRegion

GetWorldTransform mhDC, lXForm
ModifyWorldTransform mhDC, lXForm, MWT_IDENTITY

Dim lLogRect As GDI_RECT
lLogRect = GdiDeviceRectToLogicalRect(lRect)
EndBoundingRectangleDetermination = ConvertGdiRectToTRectangle(lLogRect)

lLogRect = ConvertTRectangleToGdiRect(EndBoundingRectangleDetermination)
If Not GdiRectangleCompare(gdiLogicalRectToDeviceRect(lLogRect), lRect) Then Stop

SetWorldTransform mhDC, lXForm

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub FillCircle( _
                ByVal pBrush As IBrush, _
                ByVal pCentre As Point, _
                ByVal pRadius As Dimension)
Const ProcName As String = "FillCircle"
Dim failpoint As String
On Error GoTo Err

FillEllipseFromTRectangle pBrush, circleToTRectangle(pCentre, pRadius, True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FillEllipse( _
                ByVal pBrush As IBrush, _
                ByVal pPointBottomLeft As Point, _
                ByVal pPointTopRight As Point)
Const ProcName As String = "FillEllipse"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, _
                    pPointBottomLeft.XLogical(Me), _
                    pPointTopRight.YLogical(Me), _
                    pPointTopRight.XLogical(Me), _
                    pPointBottomLeft.YLogical(Me), _
                    True

FillEllipseFromTRectangle pBrush, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FillEllipseFromPositionAndSize( _
                ByVal pBrush As IBrush, _
                ByVal pPosition As Point, _
                ByVal pSize As size)
Const ProcName As String = "FillEllipseFromPositionAndSize"
Dim failpoint As String
On Error GoTo Err

FillEllipseFromTRectangle pBrush, positionAndSizeToTRectangle(pPosition, pSize, True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FillEllipseFromTRectangle( _
                ByVal pBrush As IBrush, _
                ByRef pRect As TRectangle)
Const ProcName As String = "FillEllipseFromTRectangle"
Dim failpoint As String
On Error GoTo Err

AssertArgument pRect.isValid, "Argument is not a valid TRectangle"

FillEllipseGdi pBrush, ConvertTRectangleToGdiRect(pRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FillEllipseGdi( _
                ByVal pBrush As IBrush, _
                ByRef pRect As GDI_RECT)
Const ProcName As String = "FillEllipseGdi"
Dim failpoint As String
On Error GoTo Err

If mDeterminingBoundingRectangle Then Set pBrush = gPathBrush

If TypeOf pBrush Is INativeBrush Then
    UsePenHandle ghNullPen
    useBrush pBrush
    If Ellipse(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
    releaseBrush
    ReleasePen
ElseIf TypeOf pBrush Is ISyntheticBrush Then
    If BeginPath(mhDC) = 0 Then gHandleWin32Error GetLastError
    UsePenHandle ghPathPen
    If Ellipse(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
    ReleasePen
    EndPath mhDC
    syntheticBrushFill pBrush, pRect
End If

invalidate
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FillEllipseLogical( _
                ByVal pBrush As IBrush, _
                ByVal x1Logical As Double, _
                ByVal y1Logical As Double, _
                ByVal x2Logical As Double, _
                ByVal y2Logical As Double)
Const ProcName As String = "FillEllipseLogical"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, _
                    x1Logical, _
                    y1Logical, _
                    x2Logical, _
                    y2Logical, _
                    True

FillEllipseFromTRectangle pBrush, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FillPolygon( _
                ByVal pBrush As IBrush, _
                ByRef points() As Point, _
                Optional ByVal pFillMode As FillModes = FillModeAlternate)
Const ProcName As String = "FillPolygon"
Dim failpoint As String
On Error GoTo Err

FillPolygonGdi pBrush, ConvertPointsToGdiPoints(points), pFillMode

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FillPolygonFromPositionAndSize( _
                ByVal pBrush As IBrush, _
                ByVal pNumberOfSides As Long, _
                ByVal pPosition As Point, _
                ByVal pSize As size)
Const ProcName As String = "FillPolygonFromPositionAndSize"
On Error GoTo Err

FillPolygon pBrush, positionAndSizeToPoints(pNumberOfSides, pPosition, pSize)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub FillPolygonGdi( _
                ByVal pBrush As IBrush, _
                ByRef pPoints() As GDI_POINT, _
                Optional ByVal pFillMode As FillModes = FillModeAlternate)
Const ProcName As String = "FillPolygonGdi"
Dim failpoint As String
On Error GoTo Err

If mDeterminingBoundingRectangle Then Set pBrush = gPathBrush

If TypeOf pBrush Is INativeBrush Then
    UsePenHandle ghNullPen
    useBrush pBrush
    If Polygon(mhDC, VarPtr(pPoints(0)), UBound(pPoints) + 1) = 0 Then gHandleWin32Error GetLastError
    releaseBrush
    ReleasePen
ElseIf TypeOf pBrush Is ISyntheticBrush Then
    If BeginPath(mhDC) = 0 Then gHandleWin32Error GetLastError
    UsePenHandle ghPathPen
    If Polygon(mhDC, VarPtr(pPoints(0)), UBound(pPoints) + 1) = 0 Then gHandleWin32Error GetLastError
    ReleasePen
    EndPath mhDC
    syntheticBrushFill pBrush, GdiPolygonToRectangle(pPoints)
End If

invalidate
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FillRectangle( _
                ByVal pBrush As IBrush, _
                ByVal pPointBottomLeft As Point, _
                ByVal pPointTopRight As Point)
Const ProcName As String = "FillRectangle"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, _
                    pPointBottomLeft.XLogical(Me), _
                    pPointTopRight.YLogical(Me), _
                    pPointTopRight.XLogical(Me), _
                    pPointBottomLeft.YLogical(Me), _
                    True

FillRectangleFromTRectangle pBrush, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FillRectangleFromPositionAndSize( _
                ByVal pBrush As IBrush, _
                ByVal pPosition As Point, _
                ByVal pSize As size)
Const ProcName As String = "FillRectangleFromPositionAndSize"
Dim failpoint As String
On Error GoTo Err

FillRectangleFromTRectangle pBrush, positionAndSizeToTRectangle(pPosition, pSize, True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FillRectangleFromTRectangle( _
                ByVal pBrush As IBrush, _
                ByRef pRect As TRectangle)
Const ProcName As String = "FillRectangleFromTRectangle"
Dim failpoint As String
On Error GoTo Err

AssertArgument pRect.isValid, "Argument is not a valid TRectangle"

FillRectangleGdi pBrush, ConvertTRectangleToGdiRect(pRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FillRectangleGdi( _
                ByVal pBrush As IBrush, _
                ByRef pRect As GDI_RECT)
Const ProcName As String = "FillRectangleGdi"
Dim failpoint As String
On Error GoTo Err

If mDeterminingBoundingRectangle Then Set pBrush = gPathBrush

If TypeOf pBrush Is INativeBrush Then
    UsePenHandle ghNullPen
    useBrush pBrush
    If Rectangle(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
    releaseBrush
    ReleasePen
ElseIf TypeOf pBrush Is ISyntheticBrush Then
    If BeginPath(mhDC) = 0 Then gHandleWin32Error GetLastError
    UsePenHandle ghPathPen
    If Rectangle(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
    ReleasePen
    EndPath mhDC
    syntheticBrushFill pBrush, pRect
End If

invalidate
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub FillRectangleLogical( _
                ByVal pBrush As IBrush, _
                ByVal x1Logical As Double, _
                ByVal y1Logical As Double, _
                ByVal x2Logical As Double, _
                ByVal y2Logical As Double)
Const ProcName As String = "FillRectangleLogical"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, _
                    x1Logical, _
                    y1Logical, _
                    x2Logical, _
                    y2Logical, _
                    True

FillRectangleFromTRectangle pBrush, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

gUnRegister mhWnd, mOrigWindowProcAddr
Set mBackgroundBrush = Nothing
Set mFont = Nothing

gLogger.Log "Graphics finished", ProcName, ModuleName, LogLevelHighDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GdiDeviceRectToLogicalRect( _
                ByRef pRect As GDI_RECT) As GDI_RECT
Dim points(3) As GDI_POINT

points(0).X = pRect.Left
points(0).Y = pRect.Top
points(1).X = pRect.Right
points(1).Y = pRect.Top
points(2).X = pRect.Left
points(2).Y = pRect.Bottom
points(3).X = pRect.Right
points(3).Y = pRect.Bottom
DPtoLP mhDC, points(0), 4

GdiDeviceRectToLogicalRect.Bottom = max4(points(0).Y, points(1).Y, points(2).Y, points(3).Y)
GdiDeviceRectToLogicalRect.Left = min4(points(0).X, points(1).X, points(2).X, points(3).X)
GdiDeviceRectToLogicalRect.Top = min4(points(0).Y, points(1).Y, points(2).Y, points(3).Y)
GdiDeviceRectToLogicalRect.Right = max4(points(0).X, points(1).X, points(2).X, points(3).X)

End Function

Public Function GetFormattedTextSize( _
                ByVal pText As String, _
                ByVal pFont As StdFont, _
                ByRef pSize As size, _
                ByVal pJustification As TextJustifyModes, _
                ByVal pMultiLine As Boolean, _
                ByVal pEllipsis As EllipsisModes, _
                ByVal pExpandTabs As Boolean, _
                ByVal pTabWidth As Long, _
                ByVal pWordWrap As Boolean, _
                ByVal pLeftMargin As Dimension, _
                ByVal pRightMargin As Dimension) As size
Dim textLines() As String
Dim i As Long
Dim clipRect As GDI_RECT
Dim dtp As DRAWTEXTPARAMS
Dim sizeHeight As Long
Dim sizeWidth As Long

Const ProcName As String = "GetFormattedTextSize"
Dim failpoint As String
On Error GoTo Err

setFont pFont
If SetTextColor(mhDC, vbBlack) = CLR_INVALID Then gHandleWin32Error GetLastError

If Not pSize Is Nothing Then
    sizeWidth = ConvertLogicalToGdiX(pSize.WidthLogical(Me))
    sizeHeight = ConvertLogicalToGdiY(pSize.HeightLogical(Me))
    clipRect.Right = sizeWidth
    clipRect.Bottom = sizeHeight
Else
    clipRect.Right = &H7FFFFFF
    clipRect.Bottom = 0
End If
    
dtp.cbSize = Len(dtp)
dtp.iTabLength = pTabWidth
If Not pLeftMargin Is Nothing Then dtp.iLeftMargin = ConvertLogicalToGdiX(pLeftMargin.LengthLogicalX(Me))
If Not pLeftMargin Is Nothing Then dtp.iRightMargin = ConvertLogicalToGdiX(pRightMargin.LengthLogicalX(Me))

If User32.DrawTextEx(mhDC, _
                        StrPtr(pText), _
                        Len(pText), _
                        clipRect, _
                        DT_CALCRECT + _
                            pJustification + _
                            pEllipsis + _
                            ((Not pMultiLine) And DT_SINGLELINE) + _
                            (pExpandTabs And DT_EXPANDTABS) + _
                            (pWordWrap And DT_WORDBREAK), _
                        dtp) = 0 Then gHandleWin32Error GetLastError
                        
If Not pSize Is Nothing Then
    If sizeHeight <> 0 Then clipRect.Bottom = sizeHeight
    If sizeWidth <> 0 Then clipRect.Right = sizeWidth
End If

Set GetFormattedTextSize = gNewSize(ConvertGdiToLogicalX(clipRect.Right), _
                                            -ConvertGdiToLogicalY(clipRect.Bottom), _
                                            ScaleUnitLogical, _
                                            ScaleUnitLogical)
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetGraphicsChanges( _
                ByVal pPrevGaugeX As Double, _
                ByVal pPrevGaugeY As Double, _
                ByRef pPrevBoundary As TRectangle) As GraphicsChangeTypes
Const ProcName As String = "GetGraphicsChanges"
Dim failpoint As String
On Error GoTo Err

If Not pPrevBoundary.isValid Or Not mBoundary.isValid Then
    GetGraphicsChanges = GraphicsChangeTypes.GraphicsNoChange
    Exit Function
End If
If pPrevGaugeX <> mGaugeX Then
    GetGraphicsChanges = GraphicsChangeTypes.GraphicsGaugeChangedX
End If
If pPrevGaugeY <> mGaugeY Then
    GetGraphicsChanges = GetGraphicsChanges Or GraphicsChangeTypes.GraphicsGaugeChangedY
End If
If pPrevBoundary.Left <> mBoundary.Left Or _
    pPrevBoundary.Right <> mBoundary.Right _
Then
    GetGraphicsChanges = GetGraphicsChanges Or GraphicsChangeTypes.GraphicsBoundsChangedX
End If
If pPrevBoundary.Bottom <> mBoundary.Bottom Or _
    pPrevBoundary.Top <> mBoundary.Top _
Then
    GetGraphicsChanges = GetGraphicsChanges Or GraphicsChangeTypes.GraphicsBoundsChangedY
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetTextSize( _
                ByVal pText As String, _
                ByVal pFont As StdFont) As size
Const ProcName As String = "GetTextSize"
Dim failpoint As String
On Error GoTo Err

Dim lSize As GDI_SIZE

setFont pFont
If GetTextExtentPoint32(mhDC, StrPtr(pText), Len(pText), lSize) = 0 Then gHandleWin32Error GetLastError

Set GetTextSize = gNewSize(lSize.cx / mScaleFactorX, -lSize.cy / mScaleFactorY, ScaleUnitLogical, ScaleUnitLogical)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal phWnd As Long, _
                ByVal pLeft As Double, _
                ByVal pBottom As Double, _
                ByVal pRight As Double, _
                ByVal pTop As Double, _
                ByVal pBackGroundBrush As IBrush)
Const ProcName As String = "Initialise"
On Error GoTo Err

gLogger.Log "Initialising new Graphics object", ProcName, ModuleName

mhWnd = phWnd

mTargethDC = GetDC(phWnd)
If mTargethDC = 0 Then gHandleWin32Error GetLastError

SetScales pLeft, pBottom, pRight, pTop

BackgroundBrush = pBackGroundBrush

mOrigWindowProcAddr = gRegister(Me, phWnd)

'PaintBackground
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub IntersectClippingRegion( _
                ByRef pRect As TRectangle)
Const ProcName As String = "IntersectClippingRegion"
On Error GoTo Err

If Not pRect.isValid Then Exit Sub

CombineClippingRegionGdi ConvertTRectangleToGdiRect(pRect), RegionSelectionModeAnd

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintBackground()
Const ProcName As String = "PaintBackground"
Dim failpoint As String
On Error GoTo Err

FillRectangleGdi mBackgroundBrush, mClientRectLogical

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintCircle( _
                ByVal pPen As Pen, _
                ByVal pBrush As IBrush, _
                ByVal pCentre As Point, _
                ByVal pRadius As Dimension)
Const ProcName As String = "PaintCircle"
Dim failpoint As String
On Error GoTo Err

PaintEllipseFromTRectangle pBrush, pPen, circleToTRectangle(pCentre, pRadius, True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintEllipse( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByVal pPoint1 As Point, _
                ByVal pPoint2 As Point)
Const ProcName As String = "PaintEllipse"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, _
                pPoint1.XLogical(Me), _
                pPoint1.YLogical(Me), _
                pPoint2.XLogical(Me), _
                pPoint2.YLogical(Me), _
                True

PaintEllipseFromTRectangle pBrush, pPen, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintEllipseFromPositionAndSize( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByVal pPosition As Point, _
                ByVal pSize As size)
Const ProcName As String = "PaintEllipseFromPositionAndSize"
Dim failpoint As String
On Error GoTo Err

PaintEllipseFromTRectangle pBrush, pPen, positionAndSizeToTRectangle(pPosition, pSize, True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintEllipseFromTRectangle( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByRef pRect As TRectangle)
Const ProcName As String = "PaintEllipseFromTRectangle"
Dim failpoint As String
On Error GoTo Err

AssertArgument pRect.isValid, "Argument is not a valid TRectangle"

PaintEllipseGdi pBrush, pPen, ConvertTRectangleToGdiRect(pRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PaintEllipseGdi( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByRef pRect As GDI_RECT)
Const ProcName As String = "PaintEllipseGdi"
Dim failpoint As String
On Error GoTo Err

If mDeterminingBoundingRectangle Then Set pBrush = gPathBrush

If TypeOf pBrush Is INativeBrush Then
    useBrush pBrush
ElseIf TypeOf pBrush Is ISyntheticBrush Then
    If BeginPath(mhDC) = 0 Then gHandleWin32Error GetLastError
    UsePenHandle ghPathPen
    If Ellipse(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
    ReleasePen
    EndPath mhDC
    syntheticBrushFill pBrush, pRect
End If

usePen pPen
If Ellipse(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
releaseBrush
ReleasePen

invalidate

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintEllipseLogical( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByVal x1Logical As Double, _
                ByVal y1Logical As Double, _
                ByVal x2Logical As Double, _
                ByVal y2Logical As Double)
Const ProcName As String = "PaintEllipseLogical"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, x1Logical, y1Logical, x2Logical, y2Logical, True
PaintEllipseFromTRectangle pBrush, pPen, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintPolygon( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByRef pPoints() As Point)
Const ProcName As String = "PaintPolygon"
Dim failpoint As String
On Error GoTo Err

PaintPolygonGdi pBrush, pPen, ConvertPointsToGdiPoints(pPoints)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintPolygonFromPositionAndSize( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByVal pNumberOfSides As Long, _
                ByVal pPosition As Point, _
                ByVal pSize As size)
Const ProcName As String = "PaintPolygonFromPositionAndSize"
On Error GoTo Err

PaintPolygon pBrush, pPen, positionAndSizeToPoints(pNumberOfSides, pPosition, pSize)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub PaintPolygonGdi( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByRef pPoints() As GDI_POINT)
Const ProcName As String = "PaintPolygonGdi"
Dim failpoint As String
On Error GoTo Err

If mDeterminingBoundingRectangle Then Set pBrush = gPathBrush

If TypeOf pBrush Is INativeBrush Then
    useBrush pBrush
ElseIf TypeOf pBrush Is ISyntheticBrush Then
    If BeginPath(mhDC) = 0 Then gHandleWin32Error GetLastError
    UsePenHandle ghPathPen
    If Polygon(mhDC, VarPtr(pPoints(0)), UBound(pPoints) + 1) = 0 Then gHandleWin32Error GetLastError
    ReleasePen
    EndPath mhDC
    syntheticBrushFill pBrush, GdiPolygonToRectangle(pPoints)
End If

usePen pPen
If Polygon(mhDC, VarPtr(pPoints(0)), UBound(pPoints) + 1) = 0 Then gHandleWin32Error GetLastError
releaseBrush
ReleasePen

invalidate
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintRectangle( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByVal pPoint1 As Point, _
                ByVal pPoint2 As Point)
Const ProcName As String = "PaintRectangle"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, _
                pPoint1.XLogical(Me), _
                pPoint1.YLogical(Me), _
                pPoint2.XLogical(Me), _
                pPoint2.YLogical(Me), _
                True

PaintRectangleFromTRectangle pBrush, pPen, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintRectangleFromPositionAndSize( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByVal pPosition As Point, _
                ByVal pSize As size)
Const ProcName As String = "PaintRectangleFromPositionAndSize"
Dim failpoint As String
On Error GoTo Err

PaintRectangleFromTRectangle pBrush, pPen, positionAndSizeToTRectangle(pPosition, pSize, True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintRectangleFromTRectangle( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByRef pRect As TRectangle)
Const ProcName As String = "PaintRectangleFromTRectangle"
Dim failpoint As String
On Error GoTo Err

AssertArgument pRect.isValid, "Argument is not a valid TRectangle"

PaintRectangleGdi pBrush, pPen, ConvertTRectangleToGdiRect(pRect)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub PaintRectangleGdi( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByRef pRect As GDI_RECT)
Const ProcName As String = "PaintRectangleGdi"
Dim failpoint As String
On Error GoTo Err

If mDeterminingBoundingRectangle Then Set pBrush = gPathBrush

If TypeOf pBrush Is INativeBrush Then
    useBrush pBrush
ElseIf TypeOf pBrush Is ISyntheticBrush Then
    If BeginPath(mhDC) = 0 Then gHandleWin32Error GetLastError
    UsePenHandle ghPathPen
    If Rectangle(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
    ReleasePen
    EndPath mhDC
    syntheticBrushFill pBrush, pRect
End If

usePen pPen
If Rectangle(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
releaseBrush
ReleasePen

invalidate
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub PaintRectangleLogical( _
                ByVal pBrush As IBrush, _
                ByVal pPen As Pen, _
                ByVal x1Logical As Double, _
                ByVal y1Logical As Double, _
                ByVal x2Logical As Double, _
                ByVal y2Logical As Double)
Const ProcName As String = "PaintRectangleLogical"
Dim lTRect As TRectangle
Dim failpoint As String
On Error GoTo Err

TRectangleSetFields lTRect, x1Logical, y1Logical, x2Logical, y2Logical, True
PaintRectangleFromTRectangle pBrush, pPen, lTRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Refresh()
Const ProcName As String = "Refresh"
On Error GoTo Err

BitBlt mTargethDC, _
            0, _
            0, _
            (mBoundary.Right - mBoundary.Left) * mScaleFactorX, _
            (mBoundary.Top - mBoundary.Bottom) * mScaleFactorY, _
            mhDC, _
            0, _
            0, _
            SRCCOPY

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ReleasePen()
Const ProcName As String = "releasePen"
On Error GoTo Err

If mPrevPenHandle <> 0 Then
    If SelectObject(mhDC, mPrevPenHandle) = 0 Then gHandleWin32Error GetLastError
    mPrevPenHandle = 0
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Reset()
If ModifyWorldTransform(mhDC, mTransform, MWT_IDENTITY) = 0 Then gHandleWin32Error GetLastError
mTransform.eM11 = 1
mTransform.eM12 = 0
mTransform.eM21 = 0
mTransform.eM22 = 1
mTransform.eDx = 0
mTransform.eDy = 0
End Sub

Public Sub RestoreClippingRegion()

Const ProcName As String = "RestoreClippingRegion"
Dim failpoint As String
On Error GoTo Err

Dim lHandle As Long
lHandle = mClippingRegionHandleStack.Pop
If lHandle = NoClipRegion Then
    ClearClippingRegion
Else
    If SelectClipRgn(mhDC, lHandle) = 0 Then gHandleWin32Error GetLastError
    If DeleteObject(lHandle) = 0 Then gHandleWin32Error GetLastError
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Rotate(ByVal angleRadians As Double)
Dim pt As TPoint
RotateAboutPoint angleRadians, pt
End Sub

Public Sub RotateAboutPoint(ByVal pAngleRadians As Double, ByRef pPoint As TPoint)
Dim cosAngle As Double: cosAngle = Cos(-pAngleRadians)
Dim sinAngle As Double: sinAngle = Sin(-pAngleRadians)
Dim dx As Long: dx = ConvertCoordsLogicalToGdiX(pPoint.X)
Dim dy As Long: dy = ConvertCoordsLogicalToGdiY(pPoint.Y)

mTransform.eM11 = cosAngle
mTransform.eM12 = sinAngle
mTransform.eM21 = -sinAngle
mTransform.eM22 = cosAngle
mTransform.eDx = dx * (1 - cosAngle) + dy * sinAngle
mTransform.eDy = dy * (1 - cosAngle) - dx * sinAngle
If SetWorldTransform(mhDC, mTransform) = 0 Then gHandleWin32Error GetLastError
End Sub

Public Sub SaveClippingRegion()

Const ProcName As String = "SaveClippingRegion"
Dim failpoint As String
On Error GoTo Err

Dim res As Long
Dim lHandle As Long

lHandle = CreateRectRgn(0, 0, 2, 2)
res = GetClipRgn(mhDC, lHandle)

If res = -1 Then gHandleWin32Error GetLastError

If res = 0 Then
    ' means there is no current clipping region - see documentation for GetClipRgn
    If DeleteObject(lHandle) = 0 Then gHandleWin32Error GetLastError
    lHandle = NoClipRegion
End If
mClippingRegionHandleStack.Push lHandle

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetClippingRegion( _
                ByRef pRect As TRectangle)
Const ProcName As String = "SetClippingRegion"
On Error GoTo Err

If Not pRect.isValid Then Exit Sub

CombineClippingRegionGdi ConvertTRectangleToGdiRect(pRect), RegionSelectionModeCopy

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetClippingRegionLogical( _
                ByVal x1Logical As Double, _
                ByVal y1Logical As Double, _
                ByVal x2Logical As Double, _
                ByVal y2Logical As Double)
Const ProcName As String = "SetClippingRegionLogical"
Dim lTRect As TRectangle
Dim failpoint As String

On Error GoTo Err

SetClippingRegion TRectangle(x1Logical, y1Logical, x2Logical, y2Logical, True)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetScales( _
                ByVal pLeft As Double, _
                ByVal pBottom As Double, _
                ByVal pRight As Double, _
                ByVal pTop As Double)
Const ProcName As String = "SetScales"
On Error GoTo Err

setDimensions pLeft, pBottom, pRight, pTop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub UndrawArea( _
                ByRef undrawnArea As TRectangle)
Const ProcName As String = "UndrawArea"
Dim failpoint As String
On Error GoTo Err

SetClippingRegion undrawnArea
PaintBackground
ClearClippingRegion

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub UsePenHandle(ByVal phPen As Long)
Const ProcName As String = "UsePenHandle"
Dim failpoint As String
On Error GoTo Err

mPrevPenHandle = SelectObject(mhDC, phPen)
If mPrevPenHandle = 0 Then gHandleWin32Error GetLastError

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function WindowProc( _
                ByVal hwnd As Long, _
                ByVal iMsg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long
Const ProcName As String = "gWindowProc"

Dim lUpdateRect As GDI_RECT
Dim lPaintStruct As PAINTSTRUCT
Dim lShift As Integer
Dim lKeyCode As Integer
Static lDblClickRaised As Boolean
Dim lButton As Integer

Select Case iMsg
Case WM_CHAR
    Dim lKeyAscii As Integer
    lKeyAscii = wParam
    RaiseEvent KeyPress(lKeyAscii)
    If Not lKeyAscii = 0 Then WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
Case WM_KEYDOWN
    lShift = getShift
    lKeyCode = wParam
    RaiseEvent KeyDown(lKeyCode, lShift)
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
Case WM_KEYUP
    lShift = getShift
    lKeyCode = wParam
    RaiseEvent KeyUp(lKeyCode, lShift)
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
Case WM_LBUTTONDBLCLK
    lDblClickRaised = True
    RaiseEvent DblClick
Case WM_LBUTTONDOWN
    getXY lParam, mCurrX, mCurrY
    RaiseEvent MouseDown(MouseButtonConstants.vbLeftButton, getShift, mCurrX, mCurrY)
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
Case WM_LBUTTONUP
    getXY lParam, mCurrX, mCurrY
    RaiseEvent MouseUp(MouseButtonConstants.vbLeftButton, getShift, mCurrX, mCurrY)
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
    If lDblClickRaised Then
        lDblClickRaised = False
    Else
        RaiseEvent Click
    End If
Case WM_MBUTTONDBLCLK
    lDblClickRaised = True
    RaiseEvent DblClick
Case WM_MBUTTONDOWN
    getXY lParam, mCurrX, mCurrY
    RaiseEvent MouseDown(MouseButtonConstants.vbMiddleButton, getShift, mCurrX, mCurrY)
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
Case WM_MBUTTONUP
    getXY lParam, mCurrX, mCurrY
    RaiseEvent MouseUp(MouseButtonConstants.vbMiddleButton, getShift, mCurrX, mCurrY)
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
    If lDblClickRaised Then
        lDblClickRaised = False
    Else
        RaiseEvent Click
    End If
Case WM_MOUSEMOVE
    lButton = getButton(wParam)
    getXY lParam, mCurrX, mCurrY
    RaiseEvent MouseMove(lButton, getShift, mCurrX, mCurrY)
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
Case WM_MOUSEWHEEL
    Dim lDistance As Single
    lDistance = ((wParam And &HFFFF0000) / &H80000) / WHEEL_DELTA
    RaiseEvent MouseWheel(lDistance)
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
Case WM_RBUTTONDBLCLK
    lDblClickRaised = True
    RaiseEvent DblClick
Case WM_RBUTTONDOWN
    getXY lParam, mCurrX, mCurrY
    RaiseEvent MouseDown(MouseButtonConstants.vbRightButton, getShift, mCurrX, mCurrY)
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
Case WM_RBUTTONUP
    getXY lParam, mCurrX, mCurrY
    RaiseEvent MouseUp(MouseButtonConstants.vbRightButton, getShift, mCurrX, mCurrY)
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
    If lDblClickRaised Then
        lDblClickRaised = False
    Else
        RaiseEvent Click
    End If
Case WM_PAINT
    If GetUpdateRect(hwnd, lUpdateRect, 1) = 0 Then
        'Debug.Print "Repaint all"
        Refresh
    Else
        If mIsInDev Then countPaints
        
        Dim lhDC As Long
        lhDC = BeginPaint(hwnd, lPaintStruct)
        If lhDC = mTargethDC Then
            ' the window has a private device context - no need to set it up
            'Debug.Print "Repaint: " & gdiLogicalToAppLogicalX(lUpdateRect.Left) & ", " & gdiLogicalToAppLogicalY(lUpdateRect.Top) & ", " & gdiLogicalToAppLogicalX(lUpdateRect.Right) & ", " & gdiLogicalToAppLogicalY(lUpdateRect.Bottom)
            BitBlt lhDC, lUpdateRect.Left, lUpdateRect.Top, lUpdateRect.Right - lUpdateRect.Left, lUpdateRect.Bottom - lUpdateRect.Top, mhDC, lUpdateRect.Left, lUpdateRect.Top, SRCCOPY
        Else
            ' not a private device context so we have to set up the device context each time
            setupDeviceContext lhDC

            ReDim points(1) As GDI_POINT
            points(0).X = lUpdateRect.Left
            points(0).Y = lUpdateRect.Top
            points(1).X = lUpdateRect.Right
            points(1).Y = lUpdateRect.Bottom
            DPtoLP lhDC, points(0), 2
            'Debug.Print "Repaint: " & lUpdateRect.Left & ", " & lUpdateRect.Top & ", " & (lUpdateRect.Right - lUpdateRect.Left) & ", " & (lUpdateRect.Top - lUpdateRect.Bottom)
            BitBlt lhDC, points(0).X, points(0).Y, points(1).X - points(0).X, points(1).Y - points(0).Y, mhDC, points(0).X, points(0).Y, SRCCOPY
        End If
        EndPaint hwnd, lPaintStruct
    End If
    WindowProc = 1
Case WM_SIZE
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
    RaiseEvent Resize
Case Else
    WindowProc = CallWindowProc(mOrigWindowProcAddr, hwnd, iMsg, wParam, lParam)
End Select


End Function

'================================================================================
' Helper Functions
'================================================================================

Private Sub adjustGdiDeviceRectForPenWidth( _
                ByRef pRect As GDI_RECT, _
                ByVal pPen As Pen)
If pPen.LineStyle = LineStyles.LineInsideSolid Then Exit Sub

Dim lWidth As Long
If pPen.IsPixelPen Then
    lWidth = pPen.Width
Else
    lWidth = LogicalUnitsPerPixel * ConvertLogicalToGdiX(pPen.Width)
End If
pRect.Left = pRect.Left - Int(lWidth / 2) - 1
pRect.Top = pRect.Top - Int(lWidth / 2) - 1
pRect.Right = pRect.Right + Int((lWidth + 1) / 2)
pRect.Bottom = pRect.Bottom + Int((lWidth + 1) / 2)
End Sub

Private Sub adjustGdiLogicalRectForPenWidth( _
                ByRef pRect As GDI_RECT, _
                ByVal pPen As Pen)
Dim lXForm As XForm

GetWorldTransform mhDC, lXForm
ModifyWorldTransform mhDC, lXForm, MWT_IDENTITY
pRect = gdiLogicalRectToDeviceRect(pRect)
pRect.Right = pRect.Right + 1
pRect.Bottom = pRect.Bottom + 1
If Not pPen Is Nothing Then adjustGdiDeviceRectForPenWidth pRect, pPen
pRect = GdiDeviceRectToLogicalRect(pRect)
SetWorldTransform mhDC, lXForm
End Sub

Private Sub adjustValuesForNonZeroSingleDifference( _
                ByRef lLowValue As Double, _
                ByRef lHighValue As Double)
Dim lAdjustBy As Double
Do While CSng(lHighValue) - CSng(lLowValue) = 0
    lAdjustBy = (lHighValue - lLowValue) * 2
    lHighValue = lHighValue + lAdjustBy
    lLowValue = lLowValue - lAdjustBy
Loop
End Sub

Private Function appLogicalToGdiLogicalX(ByVal pX As Double) As Long
appLogicalToGdiLogicalX = (pX - mBoundary.Left) * mScaleFactorX
End Function

Private Function appLogicalToGdiLogicalY(ByVal pY As Double) As Long
appLogicalToGdiLogicalY = (pY - mBoundary.Bottom) * mScaleFactorY
End Function

Private Sub checkValidDimensions(ByRef lBoundary As TRectangle)
Dim midPoint As Double
Dim range As Double

Const ProcName As String = "checkValidDimensions"
Dim failpoint As String
On Error GoTo Err

If lBoundary.isValid Then
    adjustValuesForNonZeroSingleDifference lBoundary.Bottom, lBoundary.Top
    adjustValuesForNonZeroSingleDifference lBoundary.Left, lBoundary.Right
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function circleToTRectangle( _
                ByVal pCentre As Point, _
                ByVal pRadius As Dimension, _
                Optional allowZeroDimensions As Boolean = False) As TRectangle
Const ProcName As String = "circleToTRectangle"
Dim failpoint As String
On Error GoTo Err

Dim lCentreX As Double
Dim lCentreY As Double
Dim lRadiusX As Double
Dim lRadiusY As Double

lCentreX = pCentre.XLogical(Me)
lCentreY = pCentre.YLogical(Me)
lRadiusX = pRadius.LengthLogicalX(Me)
lRadiusY = ConvertLogicalXtoLogicalY(lRadiusX)

TRectangleSetFields circleToTRectangle, _
                lCentreX - lRadiusX, _
                lCentreY - lRadiusY, _
                lCentreX + lRadiusX, _
                lCentreY + lRadiusY, _
                allowZeroDimensions

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub CombineClippingRegionGdi( _
                ByRef pRect As GDI_RECT, _
                ByVal pMode As RegionSelectionModes)
Const ProcName As String = "CombineClippingRegionGdi"
Dim failpoint As String
On Error GoTo Err

'If BeginPath(mhDC) = 0 Then gHandleWin32Error GetLastError
'UsePenHandle ghPathPen
'If Rectangle(mhDC, pRect.Left, pRect.Top, pRect.Right, pRect.Bottom) = 0 Then gHandleWin32Error GetLastError
'ReleasePen
'EndPath mhDC
'
'If SelectClipPath(mhDC, pMode) = 0 Then gHandleWin32Error GetLastError

Dim lRegionHandle As Long
Dim devClipRect As GDI_RECT

devClipRect = gdiLogicalRectToDeviceRect(pRect)
lRegionHandle = CreateRectRgn(devClipRect.Left, _
                        devClipRect.Top, _
                        devClipRect.Right, _
                        devClipRect.Bottom)
If lRegionHandle = 0 Then gHandleWin32Error GetLastError

If ExtSelectClipRgn(mhDC, lRegionHandle, pMode) = 0 Then gHandleWin32Error GetLastError

If DeleteObject(lRegionHandle) = 0 Then gHandleWin32Error GetLastError
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ConvertCoordsGdiToLogicalX(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertCoordsGdiToLogicalX = MaxDouble
ElseIf Value = MinLong Then
    ConvertCoordsGdiToLogicalX = MinDouble
Else
    ConvertCoordsGdiToLogicalX = ConvertGdiToLogicalX(Value) + mBoundary.Left
End If
End Function

Private Function ConvertCoordsGdiToLogicalY(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertCoordsGdiToLogicalY = MaxDouble
ElseIf Value = MinLong Then
    ConvertCoordsGdiToLogicalY = MinDouble
Else
    ConvertCoordsGdiToLogicalY = ConvertGdiToLogicalY(Value) + mBoundary.Bottom
End If
End Function

Private Function ConvertCoordsLogicalToGdiX(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertCoordsLogicalToGdiX = MaxLong
ElseIf Value = MinDouble Then
    ConvertCoordsLogicalToGdiX = MinLong
Else
    ConvertCoordsLogicalToGdiX = normaliseGdiLogicalLength(ConvertLogicalToGdiX(Value - mBoundary.Left))
End If
End Function

Private Function ConvertCoordsLogicalToGdiY(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertCoordsLogicalToGdiY = MaxLong
ElseIf Value = MinDouble Then
    ConvertCoordsLogicalToGdiY = MinLong
Else
    ConvertCoordsLogicalToGdiY = normaliseGdiLogicalLength(ConvertLogicalToGdiY(Value - mBoundary.Bottom))
End If
End Function

Private Function ConvertDistanceToGdiX(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertDistanceToGdiX = MaxLong
ElseIf Value = MinDouble Then
    ConvertDistanceToGdiX = MinLong
Else
    ConvertDistanceToGdiX = Round(LogicalUnitsPerPixel * (Value * TwipsPerCm) / Screen.TwipsPerPixelX, 0)
End If
End Function

Private Function ConvertDistanceToGdiY(ByVal Value As Double) As Long
If Value = MaxDouble Then
    ConvertDistanceToGdiY = MaxLong
ElseIf Value = MinDouble Then
    ConvertDistanceToGdiY = MinLong
Else
    ConvertDistanceToGdiY = Round(LogicalUnitsPerPixel * (Value * TwipsPerCm) / Screen.TwipsPerPixelY, 0)
End If
End Function

Private Function ConvertGdiRectToTRectangle( _
                ByRef pRect As GDI_RECT) As TRectangle
TRectangleSetFields ConvertGdiRectToTRectangle, _
                    ConvertCoordsGdiToLogicalX(pRect.Left), _
                    ConvertCoordsGdiToLogicalY(pRect.Top), _
                    ConvertCoordsGdiToLogicalX(pRect.Right), _
                    ConvertCoordsGdiToLogicalY(pRect.Bottom)
End Function

Private Function ConvertGdiToDistanceX(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertGdiToDistanceX = MaxDouble
ElseIf Value = MinLong Then
    ConvertGdiToDistanceX = MinDouble
Else
    ConvertGdiToDistanceX = Value / LogicalUnitsPerPixel * Screen.TwipsPerPixelX / TwipsPerCm
End If
End Function

Private Function ConvertGdiToDistanceY(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertGdiToDistanceY = MaxDouble
ElseIf Value = MinLong Then
    ConvertGdiToDistanceY = MinDouble
Else
    ConvertGdiToDistanceY = Value / LogicalUnitsPerPixel * Screen.TwipsPerPixelY / TwipsPerCm
End If
End Function

Private Function ConvertGdiToLogicalX(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertGdiToLogicalX = MaxDouble
ElseIf Value = MinLong Then
    ConvertGdiToLogicalX = MinDouble
Else
    ConvertGdiToLogicalX = Value / mScaleFactorX
End If
End Function

Private Function ConvertGdiToLogicalY(ByVal Value As Long) As Double
If Value = MaxLong Then
    ConvertGdiToLogicalY = MaxDouble
ElseIf Value = MinLong Then
    ConvertGdiToLogicalY = MinDouble
Else
    ConvertGdiToLogicalY = Value / mScaleFactorY
End If
End Function

Private Function ConvertPointsToGdiPoints( _
                ByRef points() As Point) As GDI_POINT()
ReDim GdiPoints(UBound(points)) As GDI_POINT
Dim i As Long

For i = 0 To UBound(points)
    GdiPoints(i).X = ConvertCoordsLogicalToGdiX(points(i).XLogical(Me))
    GdiPoints(i).Y = ConvertCoordsLogicalToGdiY(points(i).YLogical(Me))
Next

ConvertPointsToGdiPoints = GdiPoints
End Function

Private Sub countPaints()
Static lElapsedTimer As ElapsedTimer
Static lPaintCount As Single
If lElapsedTimer Is Nothing Then
    Set lElapsedTimer = New ElapsedTimer
    lElapsedTimer.StartTiming
Else
    lPaintCount = lPaintCount + 1
    Dim et As Single
    et = lElapsedTimer.ElapsedTimeMicroseconds
    If et >= 1000000# Then
        Debug.Print "Paints per second: " & lPaintCount / et
        lElapsedTimer.StartTiming
        lPaintCount = 0
    End If
End If
End Sub

Private Sub doDrawFormattedText( _
                ByVal pText As String, _
                ByVal pColor As Long, _
                ByVal pFont As StdFont, _
                ByRef pClippingRect As GDI_RECT, _
                Optional ByVal pJustification As TextJustifyModes = TextJustifyModes.JustifyLeft, _
                Optional ByVal pMultiLine As Boolean = True, _
                Optional ByVal pEllipsis As EllipsisModes = EllipsisNone, _
                Optional ByVal pExpandTabs As Boolean = True, _
                Optional ByVal pTabWidth As Long = 8, _
                Optional ByVal pWordWrap As Boolean = True, _
                Optional ByVal pLeftMargin As Dimension = Nothing, _
                Optional ByVal pRightMargin As Dimension = Nothing)
Const ProcName As String = "doDrawFormattedText"
On Error GoTo Err

Dim dtp As DRAWTEXTPARAMS

If SetTextColor(mhDC, pColor) = CLR_INVALID Then gHandleWin32Error GetLastError

dtp.cbSize = Len(dtp)
dtp.iTabLength = pTabWidth
If Not pLeftMargin Is Nothing Then dtp.iLeftMargin = ConvertLogicalToGdiX(pLeftMargin.LengthLogicalX(Me))
If Not pLeftMargin Is Nothing Then dtp.iRightMargin = ConvertLogicalToGdiX(pRightMargin.LengthLogicalX(Me))

If User32.DrawTextEx(mhDC, _
                        StrPtr(pText), _
                        Len(pText), _
                        pClippingRect, _
                        pJustification + _
                            pEllipsis + _
                            ((Not pMultiLine) And DT_SINGLELINE) + _
                            (pExpandTabs And DT_EXPANDTABS) + _
                            (pWordWrap And DT_WORDBREAK), _
                        dtp) = 0 Then gHandleWin32Error GetLastError

invalidate
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function fontsEqual(ByVal pFont1 As StdFont, ByVal pFont2 As StdFont) As Boolean
If pFont1 Is Nothing Or pFont2 Is Nothing Then Exit Function

If pFont1 Is pFont2 Then
    fontsEqual = True
    Exit Function
End If

If pFont1.Bold <> pFont2.Bold Then Exit Function
If pFont1.Charset <> pFont2.Charset Then Exit Function
If pFont1.Italic <> pFont2.Italic Then Exit Function
If pFont1.Name <> pFont2.Name Then Exit Function
If pFont1.size <> pFont2.size Then Exit Function
If pFont1.Strikethrough <> pFont2.Strikethrough Then Exit Function
If pFont1.Underline <> pFont2.Underline Then Exit Function
If pFont1.Weight <> pFont2.Weight Then Exit Function

fontsEqual = True
End Function

Private Function gdiLogicalToAppLogicalX(ByVal pX As Long) As Double
gdiLogicalToAppLogicalX = pX / mScaleFactorX + mBoundary.Left
End Function

Private Function gdiLogicalToAppLogicalY(ByVal pY As Long) As Double
gdiLogicalToAppLogicalY = pY / mScaleFactorY + mBoundary.Bottom
End Function

Private Function gdiLogicalRectToDeviceRect( _
                ByRef pRect As GDI_RECT) As GDI_RECT
Dim points(3) As GDI_POINT

points(0).X = pRect.Left
points(0).Y = pRect.Top
points(1).X = pRect.Right
points(1).Y = pRect.Top
points(2).X = pRect.Left
points(2).Y = pRect.Bottom
points(3).X = pRect.Right
points(3).Y = pRect.Bottom
LPtoDP mhDC, points(0), 4

gdiLogicalRectToDeviceRect.Bottom = max4(points(0).Y, points(1).Y, points(2).Y, points(3).Y)
gdiLogicalRectToDeviceRect.Left = min4(points(0).X, points(1).X, points(2).X, points(3).X)
gdiLogicalRectToDeviceRect.Top = min4(points(0).Y, points(1).Y, points(2).Y, points(3).Y)
gdiLogicalRectToDeviceRect.Right = max4(points(0).X, points(1).X, points(2).X, points(3).X)

End Function

Private Function gdiPointsToGdiRect( _
                ByRef pPoint1 As GDI_POINT, _
                ByRef pPoint2 As GDI_POINT) As GDI_RECT
Const ProcName As String = "gdiPointsToGdiRect"
On Error GoTo Err

gdiPointsToGdiRect.Left = pPoint1.X
gdiPointsToGdiRect.Bottom = pPoint1.Y
gdiPointsToGdiRect.Right = pPoint2.X
gdiPointsToGdiRect.Top = pPoint2.Y

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function gdiPolygonToGdiRect(ByRef pPoints() As GDI_POINT) As GDI_RECT
Dim i As Long
Dim lRect As GDI_RECT
Dim pt As GDI_POINT

lRect.Bottom = MinLong
lRect.Left = MaxLong
lRect.Right = MinLong
lRect.Top = MaxLong

For i = 0 To UBound(pPoints)
    pt = pPoints(i)
    If pt.X < lRect.Left Then lRect.Left = pt.X
    If pt.X > lRect.Right Then lRect.Right = pt.X
    If pt.Y < lRect.Top Then lRect.Top = pt.Y
    If pt.Y > lRect.Bottom Then lRect.Bottom = pt.Y
Next
gdiPolygonToGdiRect = lRect
End Function

Private Function getButton(ByVal pwParam As Long) As Long
If (pwParam And MK_LBUTTON) = MK_LBUTTON Then getButton = MouseButtonConstants.vbLeftButton
If (pwParam And MK_MBUTTON) = MK_MBUTTON Then getButton = getButton Or MouseButtonConstants.vbMiddleButton
If (pwParam And MK_RBUTTON) = MK_RBUTTON Then getButton = getButton Or MouseButtonConstants.vbRightButton
End Function

Private Function getShift() As Integer
getShift = IIf(isKeyPressed(VK_SHIFT), ShiftConstants.vbShiftMask, 0)
getShift = getShift Or IIf(isKeyPressed(VK_CONTROL), ShiftConstants.vbCtrlMask, 0)
getShift = getShift Or IIf(isKeyPressed(VK_MENU), ShiftConstants.vbAltMask, 0)
End Function

Private Sub getXY( _
                ByVal plParam As Long, _
                ByRef lX As Single, _
                ByRef lY As Single)
lX = mBoundary.Left + ConvertGdiToLogicalX(LogicalUnitsPerPixel * (plParam And &HFFFF&))
lY = mBoundary.Top + ConvertGdiToLogicalY(LogicalUnitsPerPixel * (((plParam And &HFFFF0000) / &H10000) And &HFFFF))
End Sub

Private Sub invalidate()
Static sClipRegionHandle As Long
Dim lClipRect As GDI_RECT
Dim lRect As GDI_RECT
Dim lResult As Long

If mDeterminingBoundingRectangle Then Exit Sub

If sClipRegionHandle = 0 Then
    sClipRegionHandle = CreateRectRgn(0, 0, 1, 1)
    If sClipRegionHandle = 0 Then gHandleWin32Error GetLastError
End If

lResult = GetClipRgn(mhDC, sClipRegionHandle)

If lResult = -1 Then gHandleWin32Error GetLastError
If lResult <> 0 Then
    GetRgnBox sClipRegionHandle, lClipRect
    InvalidateRect mhWnd, lClipRect, 0
Else
    InvalidateRect mhWnd, mClientRect, 0
End If

End Sub

Private Function isKeyPressed(ByVal pVKeyCode As Long) As Boolean
isKeyPressed = (GetKeyState(pVKeyCode) And &H8000) = &H8000
End Function

Private Function max4( _
                ByVal v1 As Long, _
                ByVal v2 As Long, _
                ByVal v3 As Long, _
                ByVal v4 As Long) As Long
max4 = v1
If v2 > max4 Then max4 = v2
If v3 > max4 Then max4 = v3
If v4 > max4 Then max4 = v4
End Function

Private Function min4( _
                ByVal v1 As Long, _
                ByVal v2 As Long, _
                ByVal v3 As Long, _
                ByVal v4 As Long) As Long
min4 = v1
If v2 < min4 Then min4 = v2
If v3 < min4 Then min4 = v3
If v4 < min4 Then min4 = v4
End Function

Private Function normaliseGdiLogicalLength(ByVal pValue As Long) As Long
'normaliseGdiLogicalLength = LogicalUnitsPerPixel * Int((pValue + LogicalUnitsPerPixel - 1) / LogicalUnitsPerPixel)
normaliseGdiLogicalLength = LogicalUnitsPerPixel * Int((pValue + 1) / LogicalUnitsPerPixel)
End Function

Private Function pointToGdiPoint(ByVal pPoint As Point) As GDI_POINT
Const ProcName As String = "pointToGdiPoint"
On Error GoTo Err

pointToGdiPoint.X = appLogicalToGdiLogicalX(pPoint.XLogical(Me))
pointToGdiPoint.Y = appLogicalToGdiLogicalY(pPoint.YLogical(Me))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function positionAndSizeToPoints( _
                ByVal pNumberOfSides As Long, _
                ByVal pPosition As Point, _
                ByVal pSize As size) As Point()
Dim i As Long
Dim lSemiWidth As Double
Dim lSemiHeight As Double
Dim lX As Double
Dim lY As Double
Dim lAngleIncr As Double
Dim lAngle As Double

Const ProcName As String = "positionAndSizeToPoints"
On Error GoTo Err

lX = pPosition.XLogical(Me)
lY = pPosition.YLogical(Me)
lSemiWidth = pSize.WidthLogical(Me) / 2#
lSemiHeight = pSize.HeightLogical(Me) / 2#
lAngleIncr = 2# * Pi / pNumberOfSides

ReDim lVertices(pNumberOfSides - 1) As Point

lAngle = (Pi - lAngleIncr) / 2#
For i = 0 To pNumberOfSides - 1
    Set lVertices(i) = gNewPoint(lX + lSemiWidth * Cos(lAngle), lY + lSemiHeight * Sin(lAngle))
    lAngle = lAngle + lAngleIncr
Next

positionAndSizeToPoints = lVertices

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function positionAndSizeToTRectangle( _
                ByVal pPosition As Point, _
                ByVal pSize As size, _
                Optional ByVal pAllowZeroDimensions As Boolean = False) As TRectangle
                
Const ProcName As String = "positionAndSizeToTRectangle"
On Error GoTo Err

Dim lX As Double
Dim lY As Double
Dim lSemiWidth As Double
Dim lSemiHeight As Double

lX = pPosition.XLogical(Me)
lY = pPosition.YLogical(Me)
lSemiWidth = pSize.WidthLogical(Me) / 2#
lSemiHeight = pSize.HeightLogical(Me) / 2#

positionAndSizeToTRectangle = TRectangle(lX - lSemiWidth, lY - lSemiHeight, lX + lSemiWidth, lY + lSemiHeight, pAllowZeroDimensions)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub releaseBrush()
Const ProcName As String = "releaseBrush"
On Error GoTo Err

If mPrevBrushHandle <> 0 Then
    If SelectObject(mhDC, mPrevBrushHandle) = 0 Then gHandleWin32Error GetLastError
    mPrevBrushHandle = 0
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub releaseFont()
Const ProcName As String = "releaseFont"
On Error GoTo Err

If mPrevFontHandle = 0 Then Exit Sub

If SelectObject(mhDC, mPrevFontHandle) = 0 Then gHandleWin32Error GetLastError
If DeleteObject(mCurrFontHandle) = 0 Then gHandleWin32Error GetLastError
mPrevFontHandle = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setBoundary( _
                ByVal pLeft As Double, _
                ByVal pBottom As Double, _
                ByVal pRight As Double, _
                ByVal pTop As Double) As GraphicsChangeTypes
Dim lBoundary As TRectangle

lBoundary = TRectangle(pLeft, pBottom, pRight, pTop, False)
If Not lBoundary.isValid Then Exit Function
checkValidDimensions lBoundary
If (lBoundary.Left <> mBoundary.Left) Or _
    (lBoundary.Right <> mBoundary.Right) Then setBoundary = GraphicsBoundsChangedX
If (lBoundary.Bottom <> mBoundary.Bottom) Or _
    (lBoundary.Top <> mBoundary.Top) Then setBoundary = GraphicsBoundsChangedY
mBoundary = lBoundary
End Function

Private Sub setDimensions( _
                ByVal pLeft As Double, _
                ByVal pBottom As Double, _
                ByVal pRight As Double, _
                ByVal pTop As Double)
Const ProcName As String = "setDimensions"
Dim failpoint As String
On Error GoTo Err

Dim lChanges As GraphicsChangeTypes

lChanges = setBoundary(pLeft, pBottom, pRight, pTop)

setupDeviceContexts

lChanges = lChanges Or setupScaleFactors

If lChanges <> GraphicsNoChange Then
    Dim ev As ChangeEventData
    Set ev.Source = Me
    ev.ChangeType = lChanges
    RaiseEvent Change(ev)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setFont(ByVal pFont As StdFont)
Const ProcName As String = "setFont"
Dim failpoint As String
On Error GoTo Err

If pFont Is Nothing Then Exit Sub

If fontsEqual(pFont, mFont) Then Exit Sub

Set mFont = pFont
    
releaseFont

mCurrFontHandle = CreateFont(LogicalUnitsPerPixel * mFont.size * TwipsPerPoint / Screen.TwipsPerPixelX, _
                            0, _
                            0, _
                            0, _
                            IIf(mFont.Bold, FW_BOLD, FW_NORMAL), _
                            mFont.Italic, _
                            mFont.Underline, _
                            mFont.Strikethrough, _
                            DEFAULT_CHARSET, _
                            OUT_DEFAULT_PRECIS, _
                            CLIP_DEFAULT_PRECIS, _
                            DEFAULT_QUALITY, _
                            DEFAULT_PITCH Or FF_DONTCARE, _
                            StrPtr(mFont.Name))
If mCurrFontHandle = 0 Then gHandleWin32Error GetLastError

mPrevFontHandle = SelectObject( _
                        mhDC, _
                        mCurrFontHandle)
If mPrevFontHandle = 0 Then gHandleWin32Error GetLastError

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setInDev() As Boolean
mIsInDev = True
setInDev = True
End Function

Private Sub setupBufferDeviceContext()
Const ProcName As String = "setupBufferDeviceContext"
On Error GoTo Err

Dim prevWindowExt As GDI_SIZE
Dim prevWindowOrg As GDI_POINT
Dim prevViewportExt As GDI_SIZE

Dim failpoint As String

If mhOrigBitmap <> 0 Then
    failpoint = "Delete existing bitmap"
    mhOrigBitmap = SelectObject(mhDC, mhOrigBitmap)
    If DeleteObject(mhOrigBitmap) = 0 Then gHandleWin32Error GetLastError
    mhOrigBitmap = 0
End If

If mhDC <> 0 Then
    failpoint = "Delete existing DC"
    If DeleteDC(mhDC) = 0 Then gHandleWin32Error GetLastError
    mhDC = 0
End If

failpoint = "Create new DC"
mhDC = CreateCompatibleDC(mTargethDC)
If mhDC = 0 Then gHandleWin32Error GetLastError

failpoint = "Create bitmap"
mhOrigBitmap = SelectObject(mhDC, CreateCompatibleBitmap(mTargethDC, mClientRect.Right, mClientRect.Bottom))
If mhOrigBitmap = 0 Then gHandleWin32Error GetLastError

failpoint = "Set graphics and map modes"
If SetGraphicsMode(mhDC, GM_ADVANCED) = 0 Then gHandleWin32Error GetLastError
If SetMapMode(mhDC, MM_ANISOTROPIC) = 0 Then gHandleWin32Error GetLastError

failpoint = "Set background color and mode"
If SetBkColor(mhDC, vbWhite) = 0 Then gHandleWin32Error GetLastError
If SetBkMode(mhDC, TRANSPARENT) = 0 Then gHandleWin32Error GetLastError

failpoint = "Set window extents and origins"
If SetWindowExtEx(mhDC, LogicalUnitsPerPixel * mClientRect.Right, LogicalUnitsPerPixel * mClientRect.Bottom, prevWindowExt) = 0 Then gHandleWin32Error GetLastError
If SetWindowOrgEx(mhDC, 0, -LogicalUnitsPerPixel * mClientRect.Bottom, prevWindowOrg) = 0 Then gHandleWin32Error GetLastError

failpoint = "Set viewport extents"
If SetViewportExtEx(mhDC, mClientRect.Right, mClientRect.Bottom, prevViewportExt) = 0 Then gHandleWin32Error GetLastError

failpoint = "Set default pen, brush and font"
If SelectObject(mhDC, GetStockObject(BLACK_PEN)) = 0 Then gHandleWin32Error GetLastError
If SelectObject(mhDC, GetStockObject(NULL_BRUSH)) = 0 Then gHandleWin32Error GetLastError
If SelectObject(mhDC, GetStockObject(SYSTEM_FONT)) = 0 Then gHandleWin32Error GetLastError

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, pFailpoint:=failpoint
End Sub

Private Sub setupDeviceContext(ByVal phDC As Long)
Dim prevWindowExt As GDI_SIZE
Dim prevWindowOrg As GDI_POINT
Dim prevViewportExt As GDI_SIZE

Const ProcName As String = "setupDeviceContext"
On Error GoTo Err

If SetGraphicsMode(phDC, GM_ADVANCED) = 0 Then gHandleWin32Error GetLastError
If SetMapMode(phDC, MM_ANISOTROPIC) = 0 Then gHandleWin32Error GetLastError

If SetWindowExtEx(phDC, LogicalUnitsPerPixel * mClientRect.Right, LogicalUnitsPerPixel * mClientRect.Bottom, prevWindowExt) = 0 Then gHandleWin32Error GetLastError
If SetWindowOrgEx(phDC, 0, -LogicalUnitsPerPixel * mClientRect.Bottom, prevWindowOrg) = 0 Then gHandleWin32Error GetLastError

If SetViewportExtEx(phDC, mClientRect.Right, mClientRect.Bottom, prevViewportExt) = 0 Then gHandleWin32Error GetLastError

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupDeviceContexts()
Const ProcName As String = "setupDeviceContexts"
Dim failpoint As String
On Error GoTo Err

GetClientRect mhWnd, mClientRect
mHeightCm = ((mClientRect.Bottom - mClientRect.Top + 1) * Screen.TwipsPerPixelY) / TwipsPerCm
mWidthCm = ((mClientRect.Right - mClientRect.Left + 1) * Screen.TwipsPerPixelX) / TwipsPerCm

Debug.Print "Client rect is: " & mClientRect.Right & ", " & mClientRect.Bottom

setupDeviceContext mTargethDC
setupBufferDeviceContext

mClientRectLogical = GdiDeviceRectToLogicalRect(mClientRect)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function setupScaleFactors() As GraphicsChangeTypes
mScaleFactorX = (mClientRectLogical.Right - mClientRectLogical.Left + LogicalUnitsPerPixel) / (mBoundary.Right - mBoundary.Left)
mScaleFactorY = (mClientRectLogical.Top - mClientRectLogical.Bottom + LogicalUnitsPerPixel) / (mBoundary.Top - mBoundary.Bottom)

Dim lNewGaugeX As Double
Dim lNewGaugeY As Double
lNewGaugeX = (mBoundary.Right - mBoundary.Left) / (mClientRect.Right - mClientRect.Left + 1)
lNewGaugeY = (mBoundary.Bottom - mBoundary.Top) / (mClientRect.Top - mClientRect.Bottom + 1)
If lNewGaugeX <> mGaugeX Then setupScaleFactors = GraphicsGaugeChangedX
If lNewGaugeY <> mGaugeY Then setupScaleFactors = setupScaleFactors Or GraphicsGaugeChangedY
End Function

Private Sub syntheticBrushFill( _
                ByVal pBrush As ISyntheticBrush, _
                ByRef pRect As GDI_RECT)
Const ProcName As String = "syntheticBrushFill"
On Error GoTo Err

SaveClippingRegion
SelectClipPath mhDC, RGN_AND
pBrush.Fill pRect, 0, Me
RestoreClippingRegion

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub useBrush(ByVal pBrush As INativeBrush)
Const ProcName As String = "useBrush"
Dim failpoint As String
On Error GoTo Err

If pBrush Is Nothing Then
    useBrushHandle ghNullBrush
Else
    useBrushHandle pBrush.Handle
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub usePen(ByVal pPen As Pen)
Const ProcName As String = "usePen"
Dim failpoint As String
On Error GoTo Err

If pPen Is Nothing Then
    UsePenHandle ghNullPen
Else
    UsePenHandle pPen.Handle
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub useBrushHandle(ByVal pBrushHandle As Long)
Const ProcName As String = "useBrushHandle"
Dim failpoint As String
On Error GoTo Err

mPrevBrushHandle = SelectObject(mhDC, pBrushHandle)
If mPrevBrushHandle = 0 Then gHandleWin32Error GetLastError

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



