VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8265
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8265
   ScaleWidth      =   14880
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox GraphicsPicture 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7095
      Left            =   0
      ScaleHeight     =   7095
      ScaleWidth      =   14880
      TabIndex        =   4
      Top             =   0
      Width           =   14880
   End
   Begin VB.CommandButton PaintAllButton 
      Caption         =   "Paint all"
      Default         =   -1  'True
      Height          =   735
      Left            =   12240
      TabIndex        =   3
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton PaintMoreButton 
      Caption         =   "Paint more"
      Height          =   735
      Left            =   13560
      TabIndex        =   2
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton MiscButton 
      Caption         =   "Misc"
      Height          =   735
      Left            =   10920
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin VB.TextBox LogText 
      Height          =   975
      Left            =   3840
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7080
      Width           =   6975
   End
   Begin VB.Label BackgroundFillTimeLabel 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   7200
      Width           =   3615
   End
   Begin VB.Label TiledGradientTextTimeLabel 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   7560
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "Form1"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mGraphics                        As Graphics
Attribute mGraphics.VB_VarHelpID = -1

Private mPainted                                    As Boolean

Private mLogger                                     As FormattingLogger

Private WithEvents mUnhandledErrorHandler           As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1

Private mIsInDev                                    As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Form_Initialize()
Const ProcName As String = "Form_Initialize"
On Error GoTo Err

Debug.Print "Running in development environment: " & CStr(inDev)
Set mUnhandledErrorHandler = UnhandledErrorHandler

InitialiseTWUtilities

ApplicationGroupName = "TradeWright"
ApplicationName = "TestGraphicsUtils2"
SetupDefaultLogging Command
Set mLogger = CreateFormattingLogger("log", App.Title)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Resize()
Const ProcName As String = "Form_Resize"
On Error GoTo Err

BackgroundFillTimeLabel.Left = 0
BackgroundFillTimeLabel.Top = Me.ScaleHeight - BackgroundFillTimeLabel.Height - 360

TiledGradientTextTimeLabel.Left = 0
TiledGradientTextTimeLabel.Top = Me.ScaleHeight - BackgroundFillTimeLabel.Height - 120

PaintMoreButton.Left = Me.ScaleWidth - PaintMoreButton.Width - 120
PaintMoreButton.Top = Me.ScaleHeight - PaintMoreButton.Height - 120

PaintAllButton.Left = PaintMoreButton.Left - PaintAllButton.Width - 120
PaintAllButton.Top = Me.ScaleHeight - PaintAllButton.Height - 120

MiscButton.Left = PaintAllButton.Left - MiscButton.Width - 120
MiscButton.Top = Me.ScaleHeight - MiscButton.Height - 120

LogText.Top = Me.ScaleHeight - LogText.Height
GraphicsPicture.Height = LogText.Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub Form_Terminate()
TerminateTWUtilities
End Sub

Private Sub Form_Unload(Cancel As Integer)
Const ProcName As String = "Form_Unload"
On Error GoTo Err

If Not mGraphics Is Nothing Then mGraphics.Finish

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub MiscButton_Click()
Const ProcName As String = "MiscButton_Click"
On Error GoTo Err

miscTests

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub PaintAllButton_Click()
Const ProcName As String = "PaintAllButton_Click"
On Error GoTo Err

Paint

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub PaintMoreButton_Click()
Const ProcName As String = "PaintMoreButton_Click"
On Error GoTo Err

paintMore

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub GraphicsPicture_Resize()
Debug.Print "GraphicsPicture size: " & _
            GraphicsPicture.Width / Screen.TwipsPerPixelX & _
            ", " & _
            GraphicsPicture.Height / Screen.TwipsPerPixelY
End Sub

'@================================================================================
' mGraphics Event Handlers
'@================================================================================

Private Sub mGraphics_Click()
logMessage "Click"
End Sub

Private Sub mGraphics_DblClick()
logMessage "DblClick"
End Sub

Private Sub mGraphics_KeyDown(KeyCode As Integer, Shift As Integer)
logMessage "KeyDown: Keycode=" & KeyCode & "; Shift=" & Hex(Shift)
End Sub

Private Sub mGraphics_KeyPress(KeyAscii As Integer)
logMessage "KeyPress: KeyAscii=" & KeyAscii & " (" & Chr$(KeyAscii) & ")"
End Sub

Private Sub mGraphics_KeyUp(KeyCode As Integer, Shift As Integer)
logMessage "KeyUp: Keycode=" & KeyCode & "; Shift=" & Hex(Shift)
End Sub

Private Sub mGraphics_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
logMessage "MouseDown: Button=" & Hex(Button) & "; Shift=" & Hex(Shift) & "; X=" & X & "; Y=" & Y
End Sub

Private Sub mGraphics_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
logMessage "MouseMove: Button=" & Hex(Button) & "; Shift=" & Hex(Shift) & "; X=" & X & "; Y=" & Y
End Sub

Private Sub mGraphics_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
logMessage "MouseUp: Button=" & Hex(Button) & "; Shift=" & Hex(Shift) & "; X=" & X & "; Y=" & Y
End Sub

Private Sub mGraphics_MouseWheel(Distance As Single)
logMessage "MouseWheel: Distance=" & Distance
End Sub

Private Sub mGraphics_Resize()
Const ProcName As String = "mGraphics_Resize"
On Error GoTo Err

mGraphics.SetScales 42, 4352, 110, 4710
If mPainted Then Paint

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As ErrorEventData)

HandleFatalError

' Tell TWUtilities that we've now handled this unhandled error. Not actually
' needed here because HandleFatalError never returns anyway
UnhandledErrorHandler.Handled = True
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub HandleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly.
'
' Note that normally one would use the End statement to terminate a VB6 program abruptly. But
' the TWUtilities component interferes with the End statement's processing and may prevent
' proper shutdown, so we use the TWUtilities component's EndProcess method instead.
'
' However if we are running in the development environment, then we call End because the
' EndProcess method kills the entire development environment as well which can have undesirable
' side effects if other components are also loaded.

If mIsInDev Then
    End
Else
    EndProcess
End If

End Sub

Private Function inDev() As Boolean
Const ProcName As String = "inDev"
On Error GoTo Err

mIsInDev = True
inDev = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub logMessage(ByVal pMsg As String)

Const ProcName As String = "logMessage"
On Error GoTo Err

If Len(LogText.text) >= 32767 Then
    ' clear some space at the start of the textbox
    LogText.SelStart = 0
    LogText.SelLength = 16384
    LogText.SelText = ""
End If

LogText.SelStart = Len(LogText.text)
LogText.SelLength = 0
If Len(LogText.text) > 0 Then LogText.SelText = vbCrLf
LogText.SelText = pMsg
LogText.SelStart = InStrRev(LogText.text, vbCrLf) + 2

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub miscTests()
'setupGraphics
'
'Dim lTRect1 As TRectangle
'Dim lTRect2 As TRectangle
'Dim lGdiRect As GDI_RECT
'
'TRectangleSetFields lTRect1, 50, 4400, 70, 4500
'lGdiRect = mGraphics.ConvertTRectangleToGdiRect(lTRect1)
'
'lGdiRect.Right = lGdiRect.Right + 1000
'lGdiRect.Bottom = lGdiRect.Bottom + 1000
'
'lTRect2 = mGraphics.ConvertGdiRectToTRectangle(lGdiRect)
'
'Debug.Print TRectangleToString(lTRect1)
'Debug.Print TRectangleToString(lTRect2)
End Sub

Private Sub Paint()
Dim rotPoint As TPoint
Dim lPen As Pen
Dim lBrush As IBrush

Const ProcName As String = "Paint"
On Error GoTo Err

setupGraphics

ReDim lcolors(1) As Long
ReDim points(4) As Point

lcolors(0) = vbWhite
lcolors(1) = &H3F

mGraphics.BackgroundBrush = CreateRadialGradientBrush(TPoint(0.25, 0.75), TPoint(0.5, 0.5), 0.5, 0.5, True, lcolors)

Dim et As New ElapsedTimer
et.StartTiming
mGraphics.PaintBackground
BackgroundFillTimeLabel.Caption = et.ElapsedTimeMicroseconds

rotPoint.X = 51.2
rotPoint.Y = 4551
mGraphics.RotateAboutPoint DegreesToRadians(15), rotPoint

Set lBrush = CreateBrush(vbGreen)
mGraphics.FillRectangleLogical lBrush, 51.2, 4400, 76.95, 4551

Set lPen = CreatePixelPen(vbYellow, 5, LineSolid)
mGraphics.DrawRectangleLogical lPen, 51.2, 4400, 76.95, 4551

mGraphics.Reset

Set lBrush = CreateBrush(vbBlue)
mGraphics.FillCircle lBrush, NewPoint(63, 4600), NewDimension(1.27, ScaleUnitCm)

Set lBrush = CreateHatchedBrush(&HC0C0C0, HatchDiagonalCross)
mGraphics.FillCircle lBrush, NewPoint(63, 4600), NewDimension(1.27, ScaleUnitCm)

Set lPen = CreatePixelPen(vbWhite, 2, LineSolid)
mGraphics.DrawCircle lPen, NewPoint(63, 4600), NewDimension(1.27, ScaleUnitCm)

mGraphics.DrawEllipse lPen, _
                    NewPoint(63, 4400, CoordsLogical, CoordsLogical), _
                    NewPoint(63, 4400, CoordsLogical, CoordsLogical, NewSize(5, 2))

Set points(0) = NewPoint(74, 4360)
Set points(1) = NewPoint(83, 4425)
Set points(2) = NewPoint(87, 4500)
Set points(3) = NewPoint(77, 4550)
Set points(4) = NewPoint(68, 4450)

Set lBrush = CreateBrush(&HC0C0FF)
mGraphics.FillPolygon lBrush, points

Set lPen = CreatePixelPen(vbBlue, 1, LineCustom)
ReDim lDashPattern(7) As Single
lDashPattern(0) = 5
lDashPattern(1) = 2
lDashPattern(2) = 4
lDashPattern(3) = 3
lDashPattern(4) = 3
lDashPattern(5) = 4
lDashPattern(6) = 2
lDashPattern(7) = 5
lPen.CustomDashPattern = lDashPattern
mGraphics.DrawPolygon lPen, points

Dim lLogPen As Pen
Set lLogPen = CreateLogicalPen(mGraphics, vbBlue, 0.5, LineCustom)
ReDim lDashPattern(7) As Single
lDashPattern(0) = 5
lDashPattern(1) = 2
lDashPattern(2) = 4
lDashPattern(3) = 3
lDashPattern(4) = 3
lDashPattern(5) = 4
lDashPattern(6) = 2
lDashPattern(7) = 5
lLogPen.CustomDashPattern = lDashPattern
lLogPen.Endcaps = EndCapFlat

mGraphics.MixMode = MixModeXorPen
mGraphics.DrawRectangle lLogPen, NewPoint(50, 4400), NewPoint(85, 4600)
mGraphics.MixMode = MixModeCopyPen

Dim text As String
Dim textSize As Size
Dim locn As TRectangle
Dim aFont As StdFont

text = "This is some outlined text!!! "
Set aFont = Me.Font
aFont.Name = "Lucida Console"
aFont.Size = 24
aFont.Bold = True
aFont.Italic = True
Set lBrush = CreateBrush(vbYellow)
Set textSize = mGraphics.GetFormattedTextSize(text, _
                                            aFont, _
                                            Nothing, _
                                            JustifyLeft, _
                                            False, _
                                            EllipsisNone, _
                                            True, _
                                            8, _
                                            False, _
                                            Nothing, _
                                            Nothing)
With locn
    .Left = 48
    .Right = .Left + textSize.WidthLogical(mGraphics)
    .Top = 4420
    .Bottom = .Top - textSize.HeightLogical(mGraphics)
    .isValid = True
End With

rotPoint.X = 48
rotPoint.Y = 4420

mGraphics.RotateAboutPoint DegreesToRadians(30), rotPoint
mGraphics.DrawFormattedText text, lBrush, aFont, locn, lPen, JustifyLeft, False, EllipsisNone, True, 8, False, Nothing, Nothing
mGraphics.Reset

Dim pt As Point
Set pt = NewPoint(75, 4680)
aFont.Size = 16
aFont.Italic = True
aFont.Bold = False
Set lBrush = CreateBrush(vbMagenta)
mGraphics.DrawText "Another piece of text", lBrush, aFont, pt

Dim clip As TRectangle
TRectangleSetFields clip, 50, 4600, 85, 4590
aFont.Italic = False
aFont.Bold = False
aFont.Size = 20
Set lBrush = CreateBrush(vbWhite)
mGraphics.DrawClippedText "Yet another rather longer piece of text that should be clipped if everything is working ok", lBrush, aFont, clip

TRectangleSetFields clip, 80, 4600, 95, 4300

Set lBrush = CreateBrush(vbWhite)
mGraphics.FillRectangleFromTRectangle lBrush, clip

Set lPen = CreatePixelPen(vbBlue, 2, LineSolid)
mGraphics.DrawRectangleFromTRectangle lPen, clip

aFont.Name = "Courier New"
aFont.Italic = False
aFont.Bold = False
aFont.Size = 10
Set lBrush = CreateBrush(vbRed)
mGraphics.DrawFormattedText "An even longer piece of text that should be wordwrapped in the available space if everything is working ok." & vbCrLf & vbCrLf & _
                        "It also contains several newlines, so it should appear as more than one paragraph." & vbCrLf & vbCrLf & _
                        "Not only that, its justification is centred." & vbCrLf & vbCrLf & _
                        "And" & vbTab & "the" & vbTab & "last" & vbTab & "paragraph" & vbTab & "has" & vbTab & "tabs.", _
                        lBrush, _
                        aFont, _
                        clip, _
                        , _
                        JustifyCentre, _
                        pLeftMargin:=NewDimension(0.5), _
                        pRightMargin:=NewDimension(0.8)

Dim gradBound As TRectangle

ReDim lcolors(1) As Long
ReDim lpositions(3) As Double
ReDim lintensities(3) As Double

lcolors(0) = &H7F0000
lcolors(1) = vbWhite

lpositions(0) = 0
lpositions(1) = 0.15
lpositions(2) = 0.5
lpositions(3) = 1

lintensities(0) = 1
lintensities(1) = 0.5
lintensities(2) = 0.25
lintensities(3) = 0

TRectangleSetFields gradBound, 80, 4680, 85, 4620

mGraphics.FillRectangle CreateTiledGradientBrush(gradBound, LinearGradientDirectionHorizontal, True, TileModeFlipXY, lcolors, lintensities, lpositions), _
                        NewPoint(0.25, 0.8, CoordsRelative, CoordsRelative), _
                        NewPoint(0.75, 0.85, CoordsRelative, CoordsRelative)

Dim tgb As TiledGradientBrush
Set tgb = CreateTiledGradientBrush(gradBound, LinearGradientDirectionVertical, True, TileModeFlipXY, lcolors, lintensities, lpositions)

mGraphics.RotateAboutPoint DegreesToRadians(-30), NewPoint(0.89, 0.55, CoordsRelative, CoordsRelative).ToTPoint(mGraphics)
mGraphics.FillRectangle tgb, _
                        NewPoint(0.8, 0.1, CoordsRelative, CoordsRelative), _
                        NewPoint(0.98, 0.99, CoordsRelative, CoordsRelative)
mGraphics.Reset

mGraphics.FillEllipse tgb, _
                        NewPoint(0.4, 0.5, CoordsRelative, CoordsRelative), _
                        NewPoint(0.1, 0.55, CoordsRelative, CoordsRelative)

Set points(0) = NewPoint(0.15, 0.85, CoordsRelative, CoordsRelative)
Set points(1) = NewPoint(0.2, 0.5, CoordsRelative, CoordsRelative)
Set points(2) = NewPoint(0.5, 0.7, CoordsRelative, CoordsRelative)
Set points(3) = NewPoint(0.25, 0.7, CoordsRelative, CoordsRelative)
Set points(4) = NewPoint(0.1, 0.5, CoordsRelative, CoordsRelative)
mGraphics.FillPolygon tgb, points, FillModeWinding

ReDim lcolors(1) As Long
ReDim lpositions(2) As Double
ReDim lintensities(2) As Double

lcolors(0) = vbRed
lcolors(1) = vbBlue

lpositions(0) = 0
lpositions(1) = 0.5
lpositions(2) = 1

lintensities(0) = 1
lintensities(1) = 0
lintensities(2) = 1

aFont.Name = "Arial"
aFont.Size = 48
aFont.Bold = True

et.StartTiming
mGraphics.RotateAboutPoint DegreesToRadians(15), NewPoint(1#, 4#, CoordsDistance, CoordsDistance).ToTPoint(mGraphics)
mGraphics.DrawText "Big outlined text! With a tiled gradient fill!!", _
                    CreateTiledGradientBrush(TRectangle(mGraphics.Boundary.Left + mGraphics.ConvertDistanceToLogicalX(1#), _
                                                        mGraphics.Boundary.Bottom + mGraphics.ConvertDistanceToLogicalY(4#), _
                                                        mGraphics.Boundary.Left + mGraphics.ConvertDistanceToLogicalX(2#), _
                                                        mGraphics.Boundary.Bottom + mGraphics.ConvertDistanceToLogicalY(4#) - mGraphics.GetTextSize("Ag", aFont).HeightLogical(mGraphics)), _
                                            LinearGradientDirectionVertical, _
                                            True, _
                                            TileModeTile, _
                                            lcolors, _
                                            lintensities, _
                                            lpositions), _
                    aFont, _
                    NewPoint(1#, 4#, CoordsDistance, CoordsDistance), _
                    CreatePixelPen(ColorConstants.vbYellow, 3, LineSolid)
TiledGradientTextTimeLabel.Caption = et.ElapsedTimeMicroseconds
mGraphics.Reset

mGraphics.RotateAboutPoint DegreesToRadians(-30), TPoint(55, 4500)

mGraphics.FillRectangle CreateRadialGradientBrush(TPoint(0.3, 0.75), _
                                                    TPoint(0.5, 0.5), _
                                                    0.5, _
                                                    0.5, _
                                                    True, _
                                                    lcolors), _
                        NewPoint(55, 4425), _
                        NewPoint(80, 4500)


mGraphics.Reset

Dim rgb As RadialGradientBrush
lcolors(0) = vbWhite
lcolors(1) = &HD0D0D0
Set rgb = CreateRadialGradientBrush(TPoint(0.6, 0.7), _
                                    TPoint(0.5, 0.5), _
                                    0.25, _
                                    0.5, _
                                    True, _
                                    lcolors)

mGraphics.RotateAboutPoint DegreesToRadians(-30), TPoint(45, 4680)

mGraphics.FillEllipse rgb, _
                        NewPoint(45, 4550), _
                        NewPoint(80, 4680)
mGraphics.Reset

mGraphics.DrawText "ZAP!!", rgb, aFont, NewPoint(0.3, 0.1, CoordsRelative, CoordsRelative)

mPainted = True

mGraphics.RotateAboutPoint DegreesToRadians(-60), TPoint(75, 4600)
mGraphics.SetClippingRegion TRectangle(75, 4600, 100, 4560)
mGraphics.BackgroundBrush = CreateBrush(&H7F00&)
mGraphics.PaintBackground
mGraphics.FillRectangleFromTRectangle CreateBrush(vbRed), TRectangle(75, 4600, 100, 4560)
mGraphics.ClearClippingRegion
mGraphics.Reset
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub paintMore()

Const ProcName As String = "paintMore"
On Error GoTo Err

setupGraphics

Dim lPen As Pen
Set lPen = CreateLogicalPen(mGraphics, vbRed, 2, , HatchDiagonalCross)
mGraphics.DrawCircle lPen, NewPoint(50, 4500), NewDimension(2.5, ScaleUnitCm)

mGraphics.FillRectangle CreateTextureBrush(LoadBitmapFromResource(LoadResPicture(1, 0))), _
                        NewPoint(70, 4400), _
                        NewPoint(105, 4700)
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub setupGraphics()
Const ProcName As String = "setupGraphics"
On Error GoTo Err

If mGraphics Is Nothing Then
    Set mGraphics = CreateGraphics(GraphicsPicture.hWnd, 42, 4352, 110, 4710, CreateBrush(&HD0D0D0))
    mLogger.Log "Created Graphics object", ProcName, ModuleName
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub


