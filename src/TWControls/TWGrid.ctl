VERSION 5.00
Begin VB.UserControl TWGrid 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   ControlContainer=   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   2790
   ScaleWidth      =   6210
   Begin VB.PictureBox FontPicture 
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   6075
      TabIndex        =   9
      Top             =   2160
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.HScrollBar HScroll 
      Height          =   210
      LargeChange     =   2
      Left            =   120
      TabIndex        =   7
      Top             =   2595
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1335
      LargeChange     =   2
      Left            =   5985
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.PictureBox GridPicture 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   0
      ScaleHeight     =   2175
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      Begin VB.Label HottrackLabel 
         BackColor       =   &H0000A9F8&
         Caption         =   "Label1"
         Height          =   30
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Line RowResizeLine 
         BorderWidth     =   2
         DrawMode        =   2  'Blackness
         Visible         =   0   'False
         X1              =   0
         X2              =   6240
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line ColResizeLine 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         DrawMode        =   2  'Blackness
         Visible         =   0   'False
         X1              =   4920
         X2              =   4920
         Y1              =   0
         Y2              =   2760
      End
      Begin VB.Label ValueLabel 
         Height          =   315
         Index           =   0
         Left            =   1800
         MouseIcon       =   "TWGrid.ctx":0000
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label RightBorder 
         BackColor       =   &H80000015&
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label TopBorder 
         BackColor       =   &H80000014&
         Height          =   15
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label LeftBorder 
         BackColor       =   &H80000014&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   15
      End
      Begin VB.Label BottomBorder 
         BackColor       =   &H80000015&
         Height          =   15
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Shape FocusBox 
         BorderStyle     =   3  'Dot
         DrawMode        =   2  'Blackness
         Height          =   375
         Left            =   2520
         Top             =   1200
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label Cell 
         Height          =   315
         Index           =   0
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
End
Attribute VB_Name = "TWGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@===============================================================================+
' Interfaces
'@================================================================================

Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event Click()
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"

Event ColMoved( _
                ByVal pFromCol As Long, _
                ByVal pToCol As Long)

Event ColMoving( _
                ByVal pFromCol As Long, _
                ByVal pToCol As Long, _
                ByRef pCancel As Boolean)

Event DblClick()
Attribute DblClick.VB_UserMemId = -601

Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602

Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603

Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604

Event MouseDown( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseDown.VB_UserMemId = -605

Event MouseMove( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseMove.VB_UserMemId = -606

Event MouseUp( _
                Button As Integer, _
                Shift As Integer, _
                X As Single, _
                Y As Single)
Attribute MouseUp.VB_UserMemId = -607

Event Paint()

Event RowMoved( _
                ByVal pFromRow As Long, _
                ByVal pToRow As Long)

Event RowMoving( _
                ByVal pFromRow As Long, _
                ByVal pToRow As Long, _
                ByRef pCancel As Boolean)

Event SelectionChanged( _
                ByVal pRow1 As Long, _
                ByVal pCol1 As Long, _
                ByVal pRow2 As Long, _
                ByVal pCol2 As Long)
                
'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type ColTableEntry
    Align           As AlignmentSettings
    FixedAlign      As AlignmentSettings
    Width           As Long
    Data            As Long
    Id              As String
End Type

Private Type RowTableEntry
    Align           As AlignmentSettings
    Height          As Long
    Pos             As Long
    Cells()         As GridCell
    Data            As Long
End Type

Private Type SelectionSpecifier
    ColMin As Long
    ColMax As Long
    RowMin As Long
    RowMax As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                                As String = "TWGrid"

Private Const ConfigSectionColumns                      As String = "Columns"
Private Const ConfigSectionColumn                       As String = "Column"
Private Const ConfigSectionFontFixed                    As String = "FontFixed"
Private Const ConfigSectionFont                         As String = "Font"

Private Const ConfigSettingSelectionMode                As String = "&SelectionMode"
Private Const ConfigSettingScrollBars                   As String = "&ScrollBars"
Private Const ConfigSettingRowSizingMode                As String = "&RowSizingMode"
Private Const ConfigSettingRowHeightMin                 As String = "&RowHeightMin"
Private Const ConfigSettingRowForeColorOdd              As String = "&RowForeColorOdd"
Private Const ConfigSettingRowForeColorEven             As String = "&RowForeColorEven"
Private Const ConfigSettingRowBackColorOdd              As String = "&RowBackColorOdd"
Private Const ConfigSettingRowBackColorEven             As String = "&RowBackColorEven"
Private Const ConfigSettingHighlight                    As String = "&Highlight"
Private Const ConfigSettingGridLines                    As String = "&GridLines"
Private Const ConfigSettingGridLinesFixed               As String = "&GridLinesFixed"
Private Const ConfigSettingGridLineWidth                As String = "&GridLineWidth"
Private Const ConfigSettingGridColorFixed               As String = "&GridColorFixed"
Private Const ConfigSettingGridColor                    As String = "&GridColor"
Private Const ConfigSettingForeColorFixed               As String = "&ForeColorFixed"
Private Const ConfigSettingForeColor                    As String = "&ForeColor"
Private Const ConfigSettingFontBold                     As String = "&Bold"
Private Const ConfigSettingFontName                     As String = "&Name"
Private Const ConfigSettingFontItalic                   As String = "&Italic"
Private Const ConfigSettingFontSize                     As String = "&Size"
Private Const ConfigSettingFontStrikethrough            As String = "&Strikethrough"
Private Const ConfigSettingFontUnderline                As String = "&Underline"
Private Const ConfigSettingColumnWidth                  As String = "&Width"
Private Const ConfigSettingColumnTitle                  As String = "&Title"
Private Const ConfigSettingColumnFixedAlignment         As String = "&FixedAlignment"
Private Const ConfigSettingColumnAlignment              As String = "&Alignment"
Private Const ConfigSettingBorderStyle                  As String = "&BorderStyle"
Private Const ConfigSettingBackColorFixed               As String = "&BackColorFixed"
Private Const ConfigSettingBackColorBkg                 As String = "&BackColorBkg"
Private Const ConfigSettingBackColor                    As String = "&BackColor"
Private Const ConfigSettingAllowUserResizing            As String = "&AllowUserResizing"
Private Const ConfigSettingAllowUserReordering          As String = "&AllowUserReordering"

Private Const SampleCellText                            As String = "Cell value"
Private Const SampleFixedText                           As String = "Fixed value"

Private Const ScrollLargeChange                         As Long = 2
Private Const ScrollSmallChange                         As Long = 1

Private Const ScrollBarHideTimeSecs                     As Long = 1

'@================================================================================
' Member variables
'@================================================================================

Private mRows                               As Long
Private mCols                               As Long

Private mFixedRows                          As Long
Private mFixedCols                          As Long

' selected row(s) and column(s)
Private mRow                                As Long
Private mRowSel                             As Long
Private mCol                                As Long
Private mColSel                             As Long

Private mNextCellIndex                      As Long
Private mNextBordersIndex                   As Long

Private mMappedCells                        As Collection

Private mRowTable()                         As RowTableEntry
Private mColTable()                         As ColTableEntry

Private mGridColorFixed                     As Long
Private mGridColor                          As Long

Private mGridLines                          As GridLineSettings
Private mGridLinesFixed                     As GridLineSettings

Private mGridLineWidth                      As Long

Private mGridLineWidthTwipsX                As Long
Private mGridLineWidthTwipsY                As Long
Private mGridLineWidthTwipsXFixed           As Long
Private mGridLineWidthTwipsYFixed           As Long

Private mFillStyle                          As FillStyleSettings

Private mRedraw                             As Boolean

Private mForeColor                          As Long
Private mForeColorFixed                     As Long
Private mForeColorSel                       As Long
Private mBackColor                          As Long
Private mBackColorFixed                     As Long
Private mBackColorSel                       As Long

Private mControlDown                        As Boolean
Private mShiftDown                          As Boolean
Private mAltDown                            As Boolean

Private mLeftMouseDown                      As Boolean
Private mRightMouseDown                     As Boolean
Private mMiddleMouseDown                    As Boolean

Private mFocusRect                          As FocusRectSettings
Private mFocusRectColor                     As Long

Private mAllowBigSelection                  As Boolean
Private mSelectionMode                      As SelectionModeSettings

' these fields identify the visible cells (excluding fixed ones)
Private mTopRow                             As Long
Private mBottomRow                          As Long
Private mLeftCol                            As Long
Private mRightCol                           As Long

Private mRowNextLeft                        As Long
Private mColNextTop                         As Long

Private mAppearance                         As AppearanceSettings
Private mBorderStyle                        As BorderStyleSettings
Private mBackColorBkg                       As Long

Private mAllowUserResizing                  As AllowUserResizeSettings
Private mRowSizingMode                      As RowSizingSettings

Private mColResizer                         As ColumnResizer
Private mRowResizer                         As RowResizer

Private mAllowUserReordering                As AllowUserReorderSettings

Private mColMover                           As ColumnMover
Private mRowMover                           As RowMover

Private mHighlight                          As HighLightSettings
Private mInFocus                            As Boolean

Private mScrollBars                         As ScrollBarsSettings

Private mHScrollActive                      As Boolean
Private mVScrollActive                      As Boolean
Private mPopupScrollbars                    As Boolean

Private mFixedRowsHeight                    As Long
Private mFixedColsWidth                     As Long

' indicates that the user is scrolling by dragging the scroll bar thumb
Private mUserScrolling                      As Boolean

Private WithEvents mDefaultCellFont         As StdFont
Attribute mDefaultCellFont.VB_VarHelpID = -1
Private WithEvents mDefaultFixedFont        As StdFont
Attribute mDefaultFixedFont.VB_VarHelpID = -1

Private mDefaultCellTextHeight              As Long
Private mDefaultCellTextWidth               As Long
Private mDefaultCellWidth                   As Long

Private mDefaultFixedTextHeight             As Long
Private mDefaultFixedTextWidth              As Long
Private mDefaultFixedWidth                  As Long

Private mRowHeightMin                       As Long
Private mDefaultRowHeight                   As Long

Private mCurrMouseX                         As Long
Private mCurrMouseY                         As Long
Private mCurrMouseRow                       As Long
Private mCurrMouseCol                       As Long

Private mHighlightedCell                    As GridCell

Private mGotMouse                           As Boolean

Private mRowBackColorEven                   As Long
Private mRowBackColorOdd                    As Long
Private mRowForeColorEven                   As Long
Private mRowForeColorOdd                    As Long

Private WithEvents mMouseTracker            As MouseTracker
Attribute mMouseTracker.VB_VarHelpID = -1

Private mConfig                             As ConfigurationSection

Private mEditingCell                        As Boolean
Private mSavedRow                           As Long
Private mSavedCol                           As Long

Private WithEvents mScrollbarHideTLI        As TimerListItem
Attribute mScrollbarHideTLI.VB_VarHelpID = -1

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Click()
RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
RaiseEvent DblClick
End Sub

Private Sub UserControl_EnterFocus()
Const ProcName As String = "UserControl_EnterFocus"
On Error GoTo Err

mInFocus = True
If mHighlight = TwGridHighlightWithFocus Then showSelection

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_ExitFocus()
Const ProcName As String = "UserControl_ExitFocus"
On Error GoTo Err

If mHighlight = TwGridHighlightWithFocus Then hideSelection
mInFocus = False

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Initialize()
Initialise
mRedraw = True
End Sub

Private Sub UserControl_InitProperties()
Const ProcName As String = "UserControl_InitProperties"
On Error GoTo Err

'have to do this here because it fails in UserControl_Initialize
On Error Resume Next
setInitialFonts
On Error GoTo Err

Rows = 2
Cols = 2
FixedRows = 1
FixedCols = 1
Row = -1
Col = -1

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "UserControl_KeyDown"
On Error GoTo Err

If KeyCode = KeyCodeConstants.vbKeyPageDown Then
    pageDown
ElseIf KeyCode = KeyCodeConstants.vbKeyPageUp Then
    pageUp
End If
RaiseEvent KeyDown(KeyCode, Shift)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "UserControl_MouseDown"
On Error GoTo Err

MouseDown Button, Shift, X, Y

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "UserControl_MouseMove"
On Error GoTo Err

MouseMove UserControl.hWnd, Button, Shift, X, Y

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "UserControl_MouseUp"
On Error GoTo Err

MouseUp Button, Shift, X, Y

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'Private Sub UserControl_Paint()
'Const ProcName As String = "UserControl_Paint"
'On Error GoTo Err
'
'TransPanel.Refresh
'
'Exit Sub
'
'Err:
'gNotifyUnhandledError ProcName, ModuleName, ProjectName
'End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_ReadProperties"
On Error GoTo Err

'have to do this here because it fails in UserControl_Initialize
On Error Resume Next
setInitialFonts
On Error GoTo Err

Rows = 2
Cols = 2
FixedRows = 1
FixedCols = 1
Row = -1
Col = -1

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Show()
Const ProcName As String = "UserControl_Show"
On Error GoTo Err

paintView

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'
End Sub

'@================================================================================
' IThemeable Interface Members
'@================================================================================

Private Property Get IThemeable_Theme() As ITheme
Set IThemeable_Theme = Theme
End Property

Private Property Let IThemeable_Theme(ByVal Value As ITheme)
Const ProcName As String = "IThemeable_Theme"
On Error GoTo Err

Theme = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Control Event Handlers
'@================================================================================

Private Sub BottomBorder_Click(index As Integer)
Const ProcName As String = "BottomBorder_Click"
On Error GoTo Err

RaiseEvent Click
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub BottomBorder_DblClick(index As Integer)
Const ProcName As String = "BottomBorder_DblClick"
On Error GoTo Err

RaiseEvent DblClick
releaseMouse

Exit Sub

Err:
End Sub

Private Sub BottomBorder_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "BottomBorder_MouseDown"
On Error GoTo Err

MouseDown Button, Shift, X + BottomBorder(index).Left, Y + BottomBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub BottomBorder_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "BottomBorder_MouseMove"
On Error GoTo Err

MouseMove GridPicture.hWnd, Button, Shift, X + BottomBorder(index).Left, Y + BottomBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub BottomBorder_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "BottomBorder_MouseUp"
On Error GoTo Err

MouseUp Button, Shift, X + BottomBorder(index).Left, Y + BottomBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Cell_Click(index As Integer)
Const ProcName As String = "Cell_Click"
On Error GoTo Err

RaiseEvent Click
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Cell_DblClick(index As Integer)
Const ProcName As String = "Cell_DblClick"
On Error GoTo Err

RaiseEvent DblClick
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Cell_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "Cell_MouseDown"
On Error GoTo Err

MouseDown Button, Shift, X + Cell(index).Left, Y + Cell(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Cell_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "Cell_MouseMove"
On Error GoTo Err

MouseMove GridPicture.hWnd, Button, Shift, X + Cell(index).Left, Y + Cell(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Cell_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "Cell_MouseUp"
On Error GoTo Err

MouseUp Button, Shift, X + Cell(index).Left, Y + Cell(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub GridPicture_Click()
Const ProcName As String = "GridPicture_Click"
On Error GoTo Err

RaiseEvent Click
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub GridPicture_DblClick()
Const ProcName As String = "GridPicture_DblClick"
On Error GoTo Err

RaiseEvent DblClick
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub GridPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "GridPicture_MouseDown"
On Error GoTo Err

MouseDown Button, Shift, X + GridPicture.Left, Y + GridPicture.Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub GridPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "GridPicture_MouseMove"
On Error GoTo Err

MouseMove GridPicture.hWnd, Button, Shift, X + GridPicture.Left, Y + GridPicture.Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub GridPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "GridPicture_MouseUp"
On Error GoTo Err

MouseUp Button, Shift, X + GridPicture.Left, Y + GridPicture.Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub HottrackLabel_Click()
Const ProcName As String = "HottrackLabel_Click"
On Error GoTo Err

RaiseEvent Click
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub HottrackLabel_DblClick()
Const ProcName As String = "HottrackLabel_DblClick"
On Error GoTo Err

RaiseEvent Click
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub HottrackLabel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "HottrackLabel_MouseDown"
On Error GoTo Err

MouseDown Button, Shift, X + HottrackLabel.Left, Y + HottrackLabel.Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub HottrackLabel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "HottrackLabel_MouseMove"
On Error GoTo Err

MouseMove GridPicture.hWnd, Button, Shift, X + HottrackLabel.Left, Y + HottrackLabel.Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub HottrackLabel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "HottrackLabel_MouseUp"
On Error GoTo Err

MouseUp Button, Shift, X + HottrackLabel.Left, Y + HottrackLabel.Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub HScroll_Change()
Const ProcName As String = "HScroll_Change"
On Error GoTo Err

Static adjustingScrollbar  As Boolean

If adjustingScrollbar Then
    adjustingScrollbar = False
    Exit Sub
End If

If HScroll.Value = mLeftCol Then Exit Sub

If mUserScrolling Then
    ScrollToCol HScroll.Value
    mUserScrolling = False
ElseIf Abs(HScroll.Value - mLeftCol) = ScrollSmallChange Then
    ScrollToCol HScroll.Value
ElseIf (HScroll.Value - mLeftCol) > 0 Then
    adjustingScrollbar = True
    pageRight
    HScroll.Value = mLeftCol
    adjustingScrollbar = False
Else
    adjustingScrollbar = True
    pageLeft
    HScroll.Value = mLeftCol
    adjustingScrollbar = False
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub HScroll_Scroll()
mUserScrolling = True
End Sub

Private Sub LeftBorder_Click(index As Integer)
Const ProcName As String = "LeftBorder_Click"
On Error GoTo Err

RaiseEvent Click
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LeftBorder_DblClick(index As Integer)
Const ProcName As String = "LeftBorder_DblClick"
On Error GoTo Err

RaiseEvent DblClick
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LeftBorder_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "LeftBorder_MouseDown"
On Error GoTo Err

MouseDown Button, Shift, X + LeftBorder(index).Left, Y + LeftBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LeftBorder_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "LeftBorder_MouseMove"
On Error GoTo Err

MouseMove GridPicture.hWnd, Button, Shift, X + LeftBorder(index).Left, Y + LeftBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub LeftBorder_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "LeftBorder_MouseUp"
On Error GoTo Err

MouseUp Button, Shift, X + LeftBorder(index).Left, Y + LeftBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RightBorder_Click(index As Integer)
Const ProcName As String = "RightBorder_Click"
On Error GoTo Err

RaiseEvent Click
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RightBorder_DblClick(index As Integer)
Const ProcName As String = "RightBorder_DblClick"
On Error GoTo Err

RaiseEvent DblClick
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RightBorder_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "RightBorder_MouseDown"
On Error GoTo Err

MouseDown Button, Shift, X + RightBorder(index).Left, Y + RightBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RightBorder_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "RightBorder_MouseMove"
On Error GoTo Err

MouseMove GridPicture.hWnd, Button, Shift, X + RightBorder(index).Left, Y + RightBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub RightBorder_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "RightBorder_MouseUp"
On Error GoTo Err

MouseUp Button, Shift, X + RightBorder(index).Left, Y + RightBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TopBorder_Click(index As Integer)
Const ProcName As String = "TopBorder_Click"
On Error GoTo Err

releaseMouse
RaiseEvent Click

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TopBorder_DblClick(index As Integer)
Const ProcName As String = "TopBorder_DblClick"
On Error GoTo Err

RaiseEvent DblClick
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TopBorder_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "TopBorder_MouseDown"
On Error GoTo Err

MouseDown Button, Shift, X + TopBorder(index).Left, Y + TopBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TopBorder_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "TopBorder_MouseMove"
On Error GoTo Err

MouseMove GridPicture.hWnd, Button, Shift, X + TopBorder(index).Left, Y + TopBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub TopBorder_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "TopBorder_MouseUp"
On Error GoTo Err

MouseUp Button, Shift, X + TopBorder(index).Left, Y + TopBorder(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ValueLabel_Click(index As Integer)
Const ProcName As String = "ValueLabel_Click"
On Error GoTo Err

RaiseEvent Click
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ValueLabel_DblClick(index As Integer)
Const ProcName As String = "ValueLabel_DblClick"
On Error GoTo Err

RaiseEvent DblClick
releaseMouse

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ValueLabel_MouseDown(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "ValueLabel_MouseDown"
On Error GoTo Err

MouseDown Button, Shift, X + ValueLabel(index).Left, Y + ValueLabel(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ValueLabel_MouseMove(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "ValueLabel_MouseMove"
On Error GoTo Err

MouseMove GridPicture.hWnd, Button, Shift, X + ValueLabel(index).Left, Y + ValueLabel(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub ValueLabel_MouseUp(index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Const ProcName As String = "ValueLabel_MouseUp"
On Error GoTo Err

MouseUp Button, Shift, X + ValueLabel(index).Left, Y + ValueLabel(index).Top

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub VScroll_Change()
Const ProcName As String = "VScroll_Change"
On Error GoTo Err

Static adjustingScrollbar  As Boolean

If adjustingScrollbar Then
    adjustingScrollbar = False
    Exit Sub
End If

If VScroll.Value = mTopRow Then Exit Sub

If mUserScrolling Then
    ScrollToRow VScroll.Value
    mUserScrolling = False
ElseIf Abs(VScroll.Value - mTopRow) = ScrollSmallChange Then
    ScrollToRow VScroll.Value
ElseIf (VScroll.Value - mTopRow) > 0 Then
    adjustingScrollbar = True
    pageDown
    VScroll.Value = mTopRow
    adjustingScrollbar = False
Else
    adjustingScrollbar = True
    pageUp
    VScroll.Value = mTopRow
    adjustingScrollbar = False
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub VScroll_Scroll()
Const ProcName As String = "VScroll_Scroll"
On Error GoTo Err

mUserScrolling = True

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mDefaultCellFont Event Handlers
'@================================================================================

Private Sub mDefaultCellFont_FontChanged(ByVal PropertyName As String)
Const ProcName As String = "mDefaultCellFont_FontChanged"
On Error GoTo Err

setDefaultCellFont
storeDefaultCellFontSettings
repaintView

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mDefaultFixedFont Event Handlers
'@================================================================================

Private Sub mDefaultFixedFont_FontChanged(ByVal PropertyName As String)
Const ProcName As String = "mDefaultFixedFont_FontChanged"
On Error GoTo Err

setDefaultFixedFont
storeDefaultFixedFontSettings
repaintView

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mMouseTracker Event Handlers
'@================================================================================

Private Sub mMouseTracker_MouseLeave()
Const ProcName As String = "mMouseTracker_MouseLeave"
On Error GoTo Err

Set mMouseTracker = Nothing

If Not isMouseDown Then
    clearHottracking
    Screen.MousePointer = vbDefault
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' mScrollbarHideTLI Event Handlers
'@================================================================================

Private Sub mScrollbarHideTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mScrollbarHideTLI_StateChange"
On Error GoTo Err

If Not ev.State = TimerListItemStates.TimerListItemStateExpired Then Exit Sub
HScroll.Visible = False
VScroll.Visible = False
'FillerPicture.Visible = False
Set mScrollbarHideTLI = Nothing

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

' Returns or sets whether clicking on a column or row header causes the entire column or row to be selected.
Public Property Get AllowBigSelection() As Boolean
Const ProcName As String = "AllowBigSelection"
On Error GoTo Err

AllowBigSelection = mAllowBigSelection

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets whether clicking on a column or row header causes the entire column or row to be selected.
Public Property Let AllowBigSelection(ByVal Value As Boolean)
Const ProcName As String = "AllowBigSelection"
On Error GoTo Err

mAllowBigSelection = Value
PropertyChanged "AllowBigSelection"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'Public Property Get AllowOwnerDrawing() As Boolean
'Const ProcName As String = "AllowOwnerDrawing"
'On Error GoTo Err
'
'AllowOwnerDrawing = TransPanel.Visible
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property
'
'Public Property Let AllowOwnerDrawing( _
'                ByVal Value As Boolean)
'Const ProcName As String = "AllowOwnerDrawing"
'On Error GoTo Err
'
'TransPanel.Visible = Value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

' Returns or sets whether the user is allowed to reorder rows and columns with the mouse.
Public Property Get AllowUserReordering() As AllowUserReorderSettings
Const ProcName As String = "AllowUserReordering"
On Error GoTo Err

AllowUserReordering = mAllowUserReordering

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets whether the user is allowed to reorder rows and columns with the mouse.
Public Property Let AllowUserReordering(ByVal Value As AllowUserReorderSettings)
Const ProcName As String = "AllowUserReordering"
On Error GoTo Err

mAllowUserReordering = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingAllowUserReordering, mAllowUserReordering

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets whether the user is allowed to resize rows and columns with the mouse.
Public Property Get AllowUserResizing() As AllowUserResizeSettings
Const ProcName As String = "AllowUserResizing"
On Error GoTo Err

AllowUserResizing = mAllowUserResizing

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets whether the user is allowed to resize rows and columns with the mouse.
Public Property Let AllowUserResizing(ByVal Value As AllowUserResizeSettings)
Const ProcName As String = "AllowUserResizing"
On Error GoTo Err

mAllowUserResizing = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingAllowUserResizing, mAllowUserResizing

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets whether a control should be painted with 3-D effects.
Public Property Get Appearance() As AppearanceSettings
Const ProcName As String = "Appearance"
On Error GoTo Err

Appearance = UserControl.Appearance

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets whether a control should be painted with 3-D effects.
Public Property Let Appearance(ByVal Value As AppearanceSettings)
Const ProcName As String = "Appearance"
On Error GoTo Err

mAppearance = Value
setUserControlAppearance

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the background color of various elements of the Hierarchical FlexGrid.
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
Const ProcName As String = "BackColor"
On Error GoTo Err

BackColor = mBackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the background color of various elements of the Hierarchical FlexGrid.
Public Property Let BackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "BackColor"
On Error GoTo Err

mBackColor = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the background color of various elements of the Hierarchical FlexGrid.
Public Property Get BackColorBkg() As OLE_COLOR
Const ProcName As String = "BackColorBkg"
On Error GoTo Err

BackColorBkg = UserControl.BackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the background color of various elements of the Hierarchical FlexGrid.
Public Property Let BackColorBkg(ByVal Value As OLE_COLOR)
Const ProcName As String = "BackColorBkg"
On Error GoTo Err

mBackColorBkg = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBackColorBkg, mBackColorBkg
setUserControlAppearance

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the background color of various elements of the Hierarchical FlexGrid.
Public Property Get BackColorFixed() As OLE_COLOR
Const ProcName As String = "BackColorFixed"
On Error GoTo Err

BackColorFixed = mBackColorFixed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the background color of various elements of the Hierarchical FlexGrid.
Public Property Let BackColorFixed(ByVal Value As OLE_COLOR)
Const ProcName As String = "BackColorFixed"
On Error GoTo Err

mBackColorFixed = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBackColorFixed, mBackColorFixed
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the background color of various elements of the Hierarchical FlexGrid.
Public Property Get BackColorSel() As OLE_COLOR
Const ProcName As String = "BackColorSel"
On Error GoTo Err

BackColorSel = mBackColorSel

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the background color of various elements of the Hierarchical FlexGrid.
Public Property Let BackColorSel(ByVal Value As OLE_COLOR)
Const ProcName As String = "BackColorSel"
On Error GoTo Err

mBackColorSel = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the border style for an object.
Public Property Get BorderStyle() As BorderStyleSettings
Attribute BorderStyle.VB_UserMemId = -504
Const ProcName As String = "BorderStyle"
On Error GoTo Err

BorderStyle = UserControl.BorderStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the border style for an object.
Public Property Let BorderStyle(ByVal Value As BorderStyleSettings)
Const ProcName As String = "BorderStyle"
On Error GoTo Err

mBorderStyle = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingBorderStyle, mBorderStyle
setUserControlAppearance

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the alignment of data in a cell or range of selected cells. Not available at design time.
Public Property Get CellAlignment() As AlignmentSettings
Const ProcName As String = "CellAlignment"
On Error GoTo Err

CellAlignment = GetCellAlignment(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the alignment of data in a cell or range of selected cells. Not available at design time.
Public Property Let CellAlignment(ByVal Value As AlignmentSettings)
Const ProcName As String = "CellAlignment"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    setCellAlignment mRow, mCol, Value
    If Not mEditingCell Then paintCell mRow, mCol
Else
    sel = getSelection
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            setCellAlignment i, j, Value
            paintCell i, j
        Next
    Next
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the background color of individual cells or ranges of cells.
Public Property Get CellBackColor() As OLE_COLOR
Const ProcName As String = "CellBackColor"
On Error GoTo Err

CellBackColor = GetCellBackcolor(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the background color of individual cells or ranges of cells.
Public Property Let CellBackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "CellBackColor"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    setCellBackColor mRow, mCol, Value
    If Not mEditingCell Then paintCell mRow, mCol
Else
    sel = getSelection
    
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            setCellBackColor i, j, Value
            paintCell i, j
        Next
    Next
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the bold style for the current cell text.
Public Property Get CellFontBold() As Boolean
Const ProcName As String = "CellFontBold"
On Error GoTo Err

CellFontBold = getCellFontBold(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the bold style for the current cell text.
Public Property Let CellFontBold(ByVal Value As Boolean)
Const ProcName As String = "CellFontBold"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    setCellFontBold mRow, mCol, Value
    If Not mEditingCell Then paintCell mRow, mCol
Else
    sel = getSelection
    
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            setCellFontBold i, j, Value
            paintCell i, j
        Next
    Next
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the italic style for the current cell text.
Public Property Get CellFontItalic() As Boolean
Const ProcName As String = "CellFontItalic"
On Error GoTo Err

CellFontItalic = getCellFontItalic(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the italic style for the current cell text.
Public Property Let CellFontItalic(ByVal Value As Boolean)
Const ProcName As String = "CellFontItalic"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    setCellFontItalic mRow, mCol, Value
    If Not mEditingCell Then paintCell mRow, mCol
Else
    sel = getSelection
    
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            setCellFontItalic i, j, Value
            paintCell i, j
        Next
    Next
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the font used in individual cells or ranges of cells.
Public Property Get CellFontName() As String
Const ProcName As String = "CellFontName"
On Error GoTo Err

CellFontName = getCellFontName(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the font used in individual cells or ranges of cells.
Public Property Let CellFontName(ByVal Value As String)
Const ProcName As String = "CellFontName"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    setCellFontName mRow, mCol, Value
    If Not mEditingCell Then paintCell mRow, mCol
Else
    sel = getSelection
    
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            setCellFontName i, j, Value
            paintCell i, j
        Next
    Next
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the size, in points, for the current cell text.
Public Property Get CellFontSize() As Single
Const ProcName As String = "CellFontSize"
On Error GoTo Err

CellFontSize = getCellFontSize(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the size, in points, for the current cell text.
Public Property Let CellFontSize(ByVal Value As Single)
Const ProcName As String = "CellFontSize"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    setCellFontSize mRow, mCol, Value
    If Not mEditingCell Then paintCell mRow, mCol
Else
    sel = getSelection
    
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            setCellFontSize i, j, Value
            paintCell i, j
        Next
    Next
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the strike through style for the current cell text.
Public Property Get CellFontStrikeThrough() As Boolean
Const ProcName As String = "CellFontStrikeThrough"
On Error GoTo Err

CellFontStrikeThrough = getCellFontStrikethrough(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the strike through style for the current cell text.
Public Property Let CellFontStrikeThrough(ByVal Value As Boolean)
Const ProcName As String = "CellFontStrikeThrough"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    setCellFontStrikethrough mRow, mCol, Value
    If Not mEditingCell Then paintCell mRow, mCol
Else
    sel = getSelection
    
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            setCellFontStrikethrough i, j, Value
            paintCell i, j
        Next
    Next
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the underline style for the current cell text.
Public Property Get CellFontUnderline() As Boolean
Const ProcName As String = "CellFontUnderline"
On Error GoTo Err

CellFontUnderline = getCellFontUnderline(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the underline style for the current cell text.
Public Property Let CellFontUnderline(ByVal Value As Boolean)
Const ProcName As String = "CellFontUnderline"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    setCellFontUnderline mRow, mCol, Value
    If Not mEditingCell Then paintCell mRow, mCol
Else
    sel = getSelection
    
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            setCellFontUnderline i, j, Value
            paintCell i, j
        Next
    Next
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the font width for the current cell text.
Public Property Get CellFontWidth() As Single
Attribute CellFontWidth.VB_MemberFlags = "400"
Const ProcName As String = "CellFontWidth"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the font width for the current cell text.
Public Property Let CellFontWidth(ByVal Value As Single)
Const ProcName As String = "CellFontWidth"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the foreground color of individual cells or ranges of cells.
Public Property Get CellForeColor() As OLE_COLOR
Const ProcName As String = "CellForeColor"
On Error GoTo Err

CellForeColor = GetCellForecolor(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the foreground color of individual cells or ranges of cells.
Public Property Let CellForeColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "CellForeColor"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    setCellForeColor mRow, mCol, Value
    If Not mEditingCell Then paintCell mRow, mCol
Else
    sel = getSelection
    
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            setCellForeColor i, j, Value
            paintCell i, j
        Next
    Next
End If
Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns the height of the current cell, in Twips.
Public Property Get CellHeight() As Long
Const ProcName As String = "CellHeight"
On Error GoTo Err

CellHeight = getCellHeight(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns the left position of the current cell, in twips.
Public Property Get CellLeft() As Long
Const ProcName As String = "CellLeft"
On Error GoTo Err

ensureCellVisible mRow, mCol
CellLeft = getCellLeft(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets an image that displays in the current cell or in a range of cells.
Public Property Get CellPicture() As Picture
Const ProcName As String = "CellPicture"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets an image that displays in the current cell or in a range of cells.
Public Property Let CellPicture(ByVal Value As Picture)
Const ProcName As String = "CellPicture"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the alignment of pictures in a cell or range of selected cells. Not available at design time.
Public Property Get CellPictureAlignment() As AlignmentSettings
Const ProcName As String = "CellPictureAlignment"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the alignment of pictures in a cell or range of selected cells. Not available at design time.
Public Property Let CellPictureAlignment(ByVal Value As AlignmentSettings)
Const ProcName As String = "CellPictureAlignment"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets 3-D effects for text in a specific cell or range of cells.
Public Property Get CellTextStyle() As TextStyleSettings
Const ProcName As String = "CellTextStyle"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets 3-D effects for text in a specific cell or range of cells.
Public Property Let CellTextStyle(ByVal Value As TextStyleSettings)
Const ProcName As String = "CellTextStyle"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns the top position of the current cell, in twips.
Public Property Get Celltop() As Long
Const ProcName As String = "Celltop"
On Error GoTo Err

ensureCellVisible mRow, mCol
Celltop = getCellTop(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns the width of the current cell, in twips.
Public Property Get CellWidth() As Long
Const ProcName As String = "CellWidth"
On Error GoTo Err

ensureCellVisible mRow, mCol
CellWidth = getCellWidth(mRow, mCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the contents of the cells in a selected region of a Hierarchical FlexGrid. Not available at design time.
Public Property Get Clip() As String
Const ProcName As String = "Clip"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

sel = getSelection

For i = sel.RowMin To sel.RowMax
    If i <> sel.RowMin Then Clip = Clip & vbCr
    For j = sel.ColMin To sel.ColMax
        If j <> sel.ColMin Then Clip = Clip & vbTab
        Clip = Clip & getCellValue(i, j)
    Next
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the contents of the cells in a selected region of a Hierarchical FlexGrid. Not available at design time.
Public Property Let Clip(ByVal Value As String)
Const ProcName As String = "Clip"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim lines() As String
Dim values() As String
Dim i As Long
Dim j As Long

sel = getSelection

lines = Split(Value, vbCr)

Do While i <= UBound(lines) And i <= sel.RowMax
    values = Split(lines(i), vbTab)
    j = 0
    Do While j <= UBound(values) And i <= sel.ColMax
        TextMatrix(sel.RowMin + i, sel.ColMin + j) = values(j)
    Loop
Loop

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the active cell in a Hierarchical FlexGrid. Not available at design time.
Public Property Get Col() As Long
Const ProcName As String = "Col"
On Error GoTo Err

Col = mCol

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the active cell in a Hierarchical FlexGrid. Not available at design time.
Public Property Let Col(ByVal Value As Long)
Const ProcName As String = "Col"
On Error GoTo Err

SelectCell mRow, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString Public Property erty).
Public Property Get ColAlignment( _
                    ByVal index As Long) As AlignmentSettings
Const ProcName As String = "ColAlignment"
On Error GoTo Err

ColAlignment = mColTable(index).Align

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString Public Property erty).
Public Property Let ColAlignment( _
                    ByVal index As Long, _
                    ByVal Value As AlignmentSettings)
Const ProcName As String = "ColAlignment"
On Error GoTo Err

mColTable(index).Align = Value
If Not mConfig Is Nothing Then storeColSettings index
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString Public Property erty).
Public Property Get ColAlignmentFixed( _
                    ByVal index As Long) As AlignmentSettings
Const ProcName As String = "ColAlignmentFixed"
On Error GoTo Err

ColAlignmentFixed = mColTable(index).FixedAlign

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the alignment of data in a column. Not available at design time (except indirectly through the FormatString Public Property erty).
Public Property Let ColAlignmentFixed( _
                    ByVal index As Long, _
                    ByVal Value As AlignmentSettings)
Const ProcName As String = "ColAlignmentFixed"
On Error GoTo Err

mColTable(index).FixedAlign = Value
If Not mConfig Is Nothing Then storeColSettings index
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Array of long integer values with one item for each row (RowData) and for each column (ColData) of the Hierarchical FlexGrid. Not available at design time.
Public Property Get ColData( _
                    ByVal index As Long) As Long
Const ProcName As String = "ColData"
On Error GoTo Err

ColData = mColTable(index).Data

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Array of long integer values with one item for each row (RowData) and for each column (ColData) of the Hierarchical FlexGrid. Not available at design time.
Public Property Let ColData( _
                    ByVal index As Long, _
                    ByVal Value As Long)
Const ProcName As String = "ColData"
On Error GoTo Err

mColTable(index).Data = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns True if the specified column is visible.
Public Property Get ColIsVisible( _
                    ByVal index As Long) As Boolean
Const ProcName As String = "ColIsVisible"
On Error GoTo Err

If index < mFixedCols Then
    ColIsVisible = True
ElseIf index < mLeftCol Then
    ColIsVisible = False
ElseIf index < mRightCol Then
    ColIsVisible = True
ElseIf index > mRightCol Then
    ColIsVisible = False
ElseIf mRowNextLeft <= UserControl.ScaleWidth Then
    ColIsVisible = True
Else
    ColIsVisible = False
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns the distance, in Twips, between the upper left corner of the control and the upper left corner of a specified column.
Public Property Get ColPos( _
                    ByVal index As Long) As Long
Const ProcName As String = "ColPos"
On Error GoTo Err

ensureCellVisible mTopRow, index
ColPos = getCellLeft(mTopRow, index)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Sets the position in order of the specified column (ie changes the column number).
Public Property Let ColPosition( _
                    ByVal index As Long, _
                    ByVal Value As Long)
Const ProcName As String = "ColPosition"
On Error GoTo Err

MoveColumn index, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the total number of columns or rows in the Hierarchical FlexGrid
Public Property Get Cols() As Long
Const ProcName As String = "Cols"
On Error GoTo Err

Cols = mCols

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the total number of columns or rows in the Hierarchical FlexGrid.
Public Property Let Cols( _
                    ByVal Value As Long)
Const ProcName As String = "Cols"
On Error GoTo Err

Dim i As Long

AssertArgument Value >= 0, "Cols must be >= 0"
AssertArgument mFixedCols = 0 Or Value > mFixedCols, "Cols must be at least one greater than FixedCols"

enableDrawing False
clearView

If Value > mCols Then
    For i = mCols To Value - 1
        addCol ""
    Next
Else
    For i = Value To mCols - 1
        removeLastCol
    Next
End If

mCols = Value

paintView
enableDrawing True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the starting or ending row or column for a range of cells. Not available at design time.
Public Property Get ColSel() As Long
Const ProcName As String = "ColSel"
On Error GoTo Err

ColSel = mColSel

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the starting or ending row or column for a range of cells. Not available at design time.
Public Property Let ColSel(ByVal Value As Long)
Const ProcName As String = "ColSel"
On Error GoTo Err

AssertArgument Value >= 0 And Value <= mCols - 1, "Value out of bounds"

SelectCells mRow, mCol, mRowSel, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the width of the specified column, in Twips. Not available at design time.
Public Property Get ColWidth( _
                    ByVal index As Long) As Long
Const ProcName As String = "ColWidth"
On Error GoTo Err

ColWidth = IIf(mColTable(index).Width >= 0, _
                mColTable(index).Width, _
                getDefaultColWidth(index))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the width of the specified column, in Twips. Not available at design time.
Public Property Let ColWidth( _
                    ByVal index As Long, _
                    ByVal Value As Long)
Const ProcName As String = "ColWidth"
On Error GoTo Err

mColTable(index).Width = Value
If Not mConfig Is Nothing Then storeColSettings index

If index < mFixedCols Then calcFixedColsWidth

calcHScroll

' check whether col is visible and if so, repaint the view
If index < mFixedCols Or _
    (index >= mLeftCol And index <= mRightCol) Then repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Let ConfigurationSection( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

Set mConfig = config
If mConfig Is Nothing Then Exit Property

mConfig.SetSetting ConfigSettingSelectionMode, mSelectionMode
mConfig.SetSetting ConfigSettingScrollBars, mScrollBars
mConfig.SetSetting ConfigSettingRowSizingMode, mRowSizingMode
mConfig.SetSetting ConfigSettingRowHeightMin, mRowHeightMin
mConfig.SetSetting ConfigSettingRowBackColorOdd, mRowBackColorOdd
mConfig.SetSetting ConfigSettingRowBackColorEven, mRowBackColorEven
mConfig.SetSetting ConfigSettingHighlight, mHighlight
mConfig.SetSetting ConfigSettingGridLines, mGridLines
mConfig.SetSetting ConfigSettingGridLinesFixed, mGridLinesFixed
mConfig.SetSetting ConfigSettingGridLineWidth, mGridLineWidth
mConfig.SetSetting ConfigSettingGridColorFixed, mGridColorFixed
mConfig.SetSetting ConfigSettingGridColor, mGridColor
mConfig.SetSetting ConfigSettingForeColorFixed, mForeColorFixed
mConfig.SetSetting ConfigSettingForeColor, mForeColor
storeFontSettings mConfig.AddConfigurationSection(ConfigSectionFontFixed), mDefaultFixedFont
storeFontSettings mConfig.AddConfigurationSection(ConfigSectionFont), mDefaultCellFont
mConfig.SetSetting ConfigSettingBackColorFixed, mBackColorFixed
mConfig.SetSetting ConfigSettingBackColorBkg, mBackColorBkg
mConfig.SetSetting ConfigSettingBackColor, mBackColor
mConfig.SetSetting ConfigSettingAllowUserResizing, mAllowUserResizing
mConfig.SetSetting ConfigSettingAllowUserReordering, mAllowUserReordering

storeColumnSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines whether setting the Text property or one of the cell formatting properties of a Hierarchical FlexGrid applies the change to all selected cells.
Public Property Get FillStyle() As FillStyleSettings
Attribute FillStyle.VB_UserMemId = -511
Const ProcName As String = "FillStyle"
On Error GoTo Err

FillStyle = mFillStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines whether setting the Text property or one of the cell formatting properties of a Hierarchical FlexGrid applies the change to all selected cells.
Public Property Let FillStyle(ByVal Value As FillStyleSettings)
Const ProcName As String = "FillStyle"
On Error GoTo Err

mFillStyle = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the total number of fixed (non-scrollable) columns or rows for a Hierarchical FlexGrid.
Public Property Get FixedCols() As Long
Const ProcName As String = "FixedCols"
On Error GoTo Err

FixedCols = mFixedCols

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the total number of fixed (non-scrollable) columns or rows for a Hierarchical FlexGrid.
Public Property Let FixedCols(ByVal Value As Long)
Const ProcName As String = "FixedCols"
On Error GoTo Err

AssertArgument Value >= 0, "FixedCols must be >= 0"

enableDrawing False
clearView

mFixedCols = Value
If mLeftCol < mFixedCols Then
    mLeftCol = mFixedCols
ElseIf mFixedCols < mLeftCol And Not (mScrollBars And TwGridScrollBarHorizontal) Then
    mLeftCol = mFixedCols
End If

calcFixedColsWidth
calcHScroll

paintView
enableDrawing True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the total number of fixed (non-scrollable) columns or rows for a Hierarchical FlexGrid.
Public Property Get FixedRows() As Long
Const ProcName As String = "FixedRows"
On Error GoTo Err

FixedRows = mFixedRows

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the total number of fixed (non-scrollable) columns or rows for a Hierarchical FlexGrid.
Public Property Let FixedRows(ByVal Value As Long)
Const ProcName As String = "FixedRows"
On Error GoTo Err

AssertArgument Value >= 0, "FixedRows must be >= 0"

enableDrawing False
clearView

mFixedRows = Value
If mTopRow < mFixedRows Then
    mTopRow = mFixedRows
ElseIf mFixedRows < mTopRow And Not (mScrollBars And TwGridScrollBarVertical) Then
    mTopRow = mFixedRows
End If

calcFixedRowsHeight
calcVScroll

paintView
enableDrawing True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FocusRectColor() As OLE_COLOR
Const ProcName As String = "FocusRectColor"
On Error GoTo Err

FocusRectColor = mFocusRectColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let FocusRectColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "FocusRectColor"
On Error GoTo Err

mFocusRectColor = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines whether the Hierarchical FlexGrid control should draw a focus rectangle around the current cell.
Public Property Get FocusRect() As FocusRectSettings
Const ProcName As String = "FocusRect"
On Error GoTo Err

FocusRect = mFocusRect

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines whether the Hierarchical FlexGrid control should draw a focus rectangle around the current cell.
Public Property Let FocusRect(ByVal Value As FocusRectSettings)
Const ProcName As String = "FocusRect"
On Error GoTo Err

mFocusRect = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the default font or the font for individual cells.
Public Property Get Font() As StdFont
Attribute Font.VB_UserMemId = -512
Const ProcName As String = "Font"
On Error GoTo Err

Set Font = mDefaultCellFont

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the default font or the font for individual cells.
Public Property Let Font(ByVal Value As StdFont)
Const ProcName As String = "Font"
On Error GoTo Err

Set mDefaultCellFont = Value
setDefaultCellFont
If Not mConfig Is Nothing Then storeDefaultCellFontSettings
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Set Font(ByVal Value As StdFont)
Const ProcName As String = "Font"
On Error GoTo Err

Set mDefaultCellFont = Value
setDefaultCellFont
If Not mConfig Is Nothing Then storeDefaultCellFontSettings
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the default font or the font for individual cells.
Public Property Get FontFixed() As StdFont
Const ProcName As String = "FontFixed"
On Error GoTo Err

Set FontFixed = mDefaultFixedFont

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the default font or the font for individual cells.
Public Property Let FontFixed(ByVal Value As StdFont)
Const ProcName As String = "FontFixed"
On Error GoTo Err

Set mDefaultFixedFont = Value
setDefaultFixedFont
If Not mConfig Is Nothing Then storeDefaultFixedFontSettings
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Set FontFixed(ByVal Value As StdFont)
Const ProcName As String = "FontFixed"
On Error GoTo Err

Set mDefaultFixedFont = Value
setDefaultFixedFont
If Not mConfig Is Nothing Then storeDefaultFixedFontSettings
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the width, in points, for the current cell text.
Public Property Get FontWidth() As Single
Const ProcName As String = "FontWidth"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the width, in points, for the current cell text.
Public Property Let FontWidth(ByVal Value As Single)
Const ProcName As String = "FontWidth"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the width, in points, for the current cell text.
Public Property Get FontWidthFixed() As Single
Const ProcName As String = "FontWidthFixed"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the width, in points, for the current cell text.
Public Property Let FontWidthFixed(ByVal Value As Single)
Const ProcName As String = "FontWidthFixed"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the color used to draw text on each part of the Hierarchical FlexGrid.
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
Const ProcName As String = "ForeColor"
On Error GoTo Err

ForeColor = mForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the color used to draw text on each part of the Hierarchical FlexGrid.
Public Property Let ForeColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "ForeColor"
On Error GoTo Err

mForeColor = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingForeColor, mForeColor
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the color used to draw text on each part of the Hierarchical FlexGrid.
Public Property Get ForeColorFixed() As OLE_COLOR
Const ProcName As String = "ForeColorFixed"
On Error GoTo Err

ForeColorFixed = mForeColorFixed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the color used to draw text on each part of the Hierarchical FlexGrid.
Public Property Let ForeColorFixed(ByVal Value As OLE_COLOR)
Const ProcName As String = "ForeColorFixed"
On Error GoTo Err

mForeColorFixed = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingForeColorFixed, mForeColorFixed
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the color used to draw text on each part of the Hierarchical FlexGrid.
Public Property Get ForeColorSel() As OLE_COLOR
Const ProcName As String = "ForeColorSel"
On Error GoTo Err

ForeColorSel = mForeColorSel

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the color used to draw text on each part of the Hierarchical FlexGrid.
Public Property Let ForeColorSel(ByVal Value As OLE_COLOR)
Const ProcName As String = "ForeColorSel"
On Error GoTo Err

mForeColorSel = Value
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Allows you to set up column widths, alignments, and fixed row and column text for a Hierarchical FlexGrid at design time. See Help for more information.
Public Property Get FormatString() As String
Const ProcName As String = "FormatString"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Allows you to set up column widths, alignments, and fixed row and column text for a Hierarchical FlexGrid at design time. See Help for more information.
Public Property Let FormatString(ByVal Value As String)
Const ProcName As String = "FormatString"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the color used to draw the lines between Hierarchical FlexGrid cells.
Public Property Get GridColor() As OLE_COLOR
Const ProcName As String = "GridColor"
On Error GoTo Err

GridColor = mGridColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the color used to draw the lines between Hierarchical FlexGrid cells.
Public Property Let GridColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "GridColor"
On Error GoTo Err

mGridColor = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingGridColor, mGridColor
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the color used to draw the lines between Hierarchical FlexGrid cells.
Public Property Get GridColorFixed() As OLE_COLOR
GridColorFixed = mGridColorFixed
End Property

' Returns or sets the color used to draw the lines between Hierarchical FlexGrid cells.
Public Property Let GridColorFixed(ByVal Value As OLE_COLOR)
Const ProcName As String = "GridColorFixed"
On Error GoTo Err

mGridColorFixed = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingGridColorFixed, mGridColorFixed
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the type of lines that are drawn between Hierarchical FlexGrid cells.
Public Property Get GridLines() As GridLineSettings
Const ProcName As String = "GridLines"
On Error GoTo Err

GridLines = mGridLines

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the type of lines that are drawn between Hierarchical FlexGrid cells.
Public Property Let GridLines(ByVal Value As GridLineSettings)
Const ProcName As String = "GridLines"
On Error GoTo Err

Select Case Value
Case TwGridGridNone, TwGridGridFlat, TwGridGridInset, TwGridGridRaised
Case Else
    Assert False, "Invalid property value", VBErrorCodes.VbErrInvalidPropertyValue
End Select

mGridLines = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingGridLines, mGridLines

mGridLineWidthTwipsX = calcGridLineWidthTwipsX(mGridLines, mGridLineWidth)
mGridLineWidthTwipsY = calcGridLineWidthTwipsY(mGridLines, mGridLineWidth)

repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the type of lines that are drawn between Hierarchical FlexGrid cells.
Public Property Get GridLinesFixed() As GridLineSettings
Const ProcName As String = "GridLinesFixed"
On Error GoTo Err

GridLinesFixed = mGridLinesFixed

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the type of lines that are drawn between Hierarchical FlexGrid cells.
Public Property Let GridLinesFixed(ByVal Value As GridLineSettings)
Const ProcName As String = "GridLinesFixed"
On Error GoTo Err

Select Case Value
Case TwGridGridNone, TwGridGridFlat, TwGridGridInset, TwGridGridRaised
Case Else
    Assert False, "Invalid property value", VBErrorCodes.VbErrInvalidPropertyValue
End Select

mGridLinesFixed = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingGridLinesFixed, mGridLinesFixed

mGridLineWidthTwipsXFixed = calcGridLineWidthTwipsX(mGridLinesFixed, mGridLineWidth)
mGridLineWidthTwipsYFixed = calcGridLineWidthTwipsY(mGridLinesFixed, mGridLineWidth)

repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the width, in Pixels, of the gridlines.
Public Property Get GridLineWidth() As Long
Const ProcName As String = "GridLineWidth"
On Error GoTo Err

GridLineWidth = mGridLineWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the width, in Pixels, of the gridlines.
Public Property Let GridLineWidth(ByVal Value As Long)
Const ProcName As String = "GridLineWidth"
On Error GoTo Err

Assert IsInteger(Value, 1, 10), "Invalid property value", VBErrorCodes.VbErrInvalidPropertyValue

mGridLineWidth = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingGridLineWidth, mGridLineWidth

mGridLineWidthTwipsX = calcGridLineWidthTwipsX(mGridLines, mGridLineWidth)
mGridLineWidthTwipsY = calcGridLineWidthTwipsY(mGridLines, mGridLineWidth)

mGridLineWidthTwipsXFixed = calcGridLineWidthTwipsX(mGridLinesFixed, mGridLineWidth)
mGridLineWidthTwipsYFixed = calcGridLineWidthTwipsY(mGridLinesFixed, mGridLineWidth)

repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get GridLineWidthTwipsX() As Long
GridLineWidthTwipsX = mGridLineWidthTwipsX
End Property

Friend Property Get GridLineWidthTwipsY() As Long
GridLineWidthTwipsY = mGridLineWidthTwipsY
End Property

Friend Property Get GridLineWidthTwipsXFixed() As Long
GridLineWidthTwipsXFixed = mGridLineWidthTwipsXFixed
End Property

Friend Property Get GridLineWidthTwipsYFixed() As Long
GridLineWidthTwipsYFixed = mGridLineWidthTwipsYFixed
End Property

'Public Property Get hDC() As Long
'Const ProcName As String = "hDC"
'On Error GoTo Err
'
'hDC = TransPanel.hDC
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

' Returns or sets whether selected cells appear highlighted.
Public Property Get HighLight() As HighLightSettings
Const ProcName As String = "HighLight"
On Error GoTo Err

HighLight = mHighlight

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets whether selected cells appear highlighted.
Public Property Let HighLight(ByVal Value As HighLightSettings)
Const ProcName As String = "HighLight"
On Error GoTo Err

mHighlight = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHighlight, mHighlight

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns a handle to a form or control.
Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
Const ProcName As String = "hWnd"
On Error GoTo Err

hWnd = UserControl.hWnd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the left-most visible column (other than a fixed column) in the Hierarchical FlexGrid. Not available at design time.
Public Property Get LeftCol() As Long
Const ProcName As String = "LeftCol"
On Error GoTo Err

LeftCol = mLeftCol

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the left-most visible column (other than a fixed column) in the Hierarchical FlexGrid. Not available at design time.
Public Property Let LeftCol(ByVal Value As Long)
Const ProcName As String = "LeftCol"
On Error GoTo Err

ScrollToCol Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the row or column over which the mouse pointer is positioned. Not available at design time.
Public Property Get MouseCol() As Long
Const ProcName As String = "MouseCol"
On Error GoTo Err

MouseCol = mCurrMouseCol

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'' Returns or sets a custom mouse icon.
'Public Property Get MouseIcon() As Picture
'End Property
'
'' Returns or sets a custom mouse icon.
'Public Property Let MouseIcon(ByVal Value As Picture)
'End Property
'
'' Returns or sets the type of mouse pointer displayed when over part of an object.
'Public Property Get MousePointer() As MousePointerSettings
'End Property
'
'' Returns or sets the type of mouse pointer displayed when over part of an object.
'Public Property Let MousePointer(ByVal Value As MousePointerSettings)
'End Property

' Returns or sets the row or column over which the mouse pointer is positioned. Not available at design time.
Public Property Get MouseRow() As Long
Const ProcName As String = "MouseRow"
On Error GoTo Err

MouseRow = mCurrMouseRow

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let PopupScrollbars(ByVal Value As Boolean)
Const ProcName As String = "PopupScrollbars"
On Error GoTo Err

mPopupScrollbars = Value
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property
    
Public Property Get PopupScrollbars() As Boolean
PopupScrollbars = mPopupScrollbars
End Property
    
'' Returns or sets whether this control acts as an OLE drop target.
'Public Property Get OLEDropMode() As OLEDropConstants
'End Property
'
'' Returns or sets whether this control acts as an OLE drop target.
'Public Property Let OLEDropMode(ByVal psOLEDropMode As OLEDropConstants)
'End Property

'' Returns a picture of the Hierarchical FlexGrid, suitable for printing, saving to disk, copying to the clipboard, or assigning to a different control.
'Public Property Get Picture() As Picture
'End Property
'
'' Returns or sets the type of picture that is generated by the Picture Public Property erty.
'Public Property Get PictureType() As PictureTypeSettings
'End Property
'
'' Returns or sets the type of picture that is generated by the Picture Public Property erty.
'Public Property Let PictureType(ByVal Value As PictureTypeSettings)
'End Property

' Enables or disables redrawing of the Hierarchical FlexGrid control.
Public Property Get Redraw() As Boolean
Const ProcName As String = "Redraw"
On Error GoTo Err

Redraw = mRedraw

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Enables or disables redrawing of the Hierarchical FlexGrid control.
Public Property Let Redraw(ByVal Value As Boolean)
Const ProcName As String = "Redraw"
On Error GoTo Err

mRedraw = Value
If mRedraw Then repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property
    
'Public Property Get RightToLeft() As Boolean
'End Property
'
'' Determines text display direction and controls visual appearance on a bidirectional system.
'Public Property Let RightToLeft(ByVal Value As Boolean)
'End Property

' Returns or sets the active cell in a Hierarchical FlexGrid. Not available at design time.
Public Property Get Row() As Long
Const ProcName As String = "Row"
On Error GoTo Err

Row = mRow

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the active cell in a Hierarchical FlexGrid. Not available at design time.
Public Property Let Row(ByVal Value As Long)
Const ProcName As String = "Row"
On Error GoTo Err

SelectCell Value, mCol

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowBackColorEven() As OLE_COLOR
Const ProcName As String = "RowBackColorEven"
On Error GoTo Err

RowBackColorEven = mRowBackColorEven

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowBackColorEven( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "RowBackColorEven"
On Error GoTo Err

mRowBackColorEven = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingRowBackColorEven, mRowBackColorEven
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowBackColorOdd() As OLE_COLOR
Const ProcName As String = "RowBackColorOdd"
On Error GoTo Err

RowBackColorOdd = mRowBackColorOdd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowBackColorOdd( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "RowBackColorOdd"
On Error GoTo Err

mRowBackColorOdd = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingRowBackColorOdd, mRowBackColorOdd
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowForeColorEven() As OLE_COLOR
Const ProcName As String = "RowForeColorEven"
On Error GoTo Err

RowForeColorEven = mRowForeColorEven

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowForeColorEven( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "RowForeColorEven"
On Error GoTo Err

mRowForeColorEven = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingRowForeColorEven, mRowForeColorEven
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RowForeColorOdd() As OLE_COLOR
Const ProcName As String = "RowForeColorOdd"
On Error GoTo Err

RowForeColorOdd = mRowForeColorOdd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RowForeColorOdd( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "RowForeColorOdd"
On Error GoTo Err

mRowForeColorOdd = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingRowForeColorOdd, mRowForeColorOdd
repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Array of long integer values with one item for each row (RowData) and for each column (ColData) of the Hierarchical FlexGrid. Not available at design time.
Public Property Get RowData( _
                ByVal index As Long) As Long
Const ProcName As String = "RowData"
On Error GoTo Err

RowData = mRowTable(index).Data

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Array of long integer values with one item for each row (RowData) and for each column (ColData) of the Hierarchical FlexGrid. Not available at design time.
Public Property Let RowData( _
                ByVal index As Long, _
                ByVal Value As Long)
Const ProcName As String = "RowData"
On Error GoTo Err

mRowTable(index).Data = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the height of the specified row, in Twips. Not available at design time.
Public Property Get RowHeight( _
                ByVal index As Long) As Long
Const ProcName As String = "RowHeight"
On Error GoTo Err

If mRowTable(index).Height >= 0 Then
    RowHeight = mRowTable(index).Height
ElseIf mDefaultRowHeight <> 0 And Not isCellColHeader(index) Then
    RowHeight = mDefaultRowHeight
Else
    RowHeight = getDefaultRowHeight(index)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the height of the specified row, in Twips. Not available at design time.
Public Property Let RowHeight( _
                ByVal index As Long, _
                ByVal Value As Long)
Const ProcName As String = "RowHeight"
On Error GoTo Err

If Value < mRowHeightMin Then Value = mRowHeightMin

mRowTable(index).Height = Value

If index < mFixedRows Then calcFixedRowsHeight

calcVScroll

' check whether row is visible and if so, repaint the view
If index < mFixedRows Or _
    (index >= mTopRow And index <= mBottomRow) Then repaintView

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets a minimum row height for the entire control, in Twips.
Public Property Get RowHeightMin() As Long
Const ProcName As String = "RowHeightMin"
On Error GoTo Err

RowHeightMin = mRowHeightMin

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets a minimum row height for the entire control, in Twips.
Public Property Let RowHeightMin(ByVal Value As Long)
Const ProcName As String = "RowHeightMin"
On Error GoTo Err

mRowHeightMin = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingRowHeightMin, mRowHeightMin

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns True if the specified row is visible.
Public Property Get RowIsVisible( _
                    ByVal index As Long) As Boolean
Const ProcName As String = "RowIsVisible"
On Error GoTo Err

If index < mFixedRows Then
    RowIsVisible = True
ElseIf index < mTopRow Then
    RowIsVisible = False
ElseIf index < mBottomRow Then
    RowIsVisible = True
ElseIf index > mBottomRow Then
    RowIsVisible = False
ElseIf mColNextTop <= UserControl.ScaleHeight Then
    RowIsVisible = True
Else
    RowIsVisible = False
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns the distance, in Twips, between the upper left corner of the control and the upper left corner of a specified row.
Public Property Get RowPos( _
                    ByVal index As Long) As Long
Const ProcName As String = "RowPos"
On Error GoTo Err

ensureCellVisible index, mLeftCol
RowPos = getCellTop(index, mLeftCol)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Sets the position in order of the specified row (ie changes the row number).
Public Property Let RowPosition( _
                    ByVal index As Long, _
                    ByVal Value As Long)
Const ProcName As String = "RowPosition"
On Error GoTo Err

MoveRow index, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the total number of columns or rows in the Hierarchical FlexGrid.
Public Property Get Rows() As Long
Const ProcName As String = "Rows"
On Error GoTo Err

Rows = mRows

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the total number of columns or rows in the Hierarchical FlexGrid
Public Property Let Rows(ByVal Value As Long)
Const ProcName As String = "Rows"
On Error GoTo Err

Dim i As Long

AssertArgument Value >= 0, "Rows must be >= 0"
AssertArgument mFixedRows <= 0 Or Value > mFixedRows, "Rows must be at least one greater than FixedRows"

enableDrawing False
clearView

If Value > mRows Then
    For i = mRows To Value - 1
        addRow
    Next
Else
    For i = Value To mRows - 1
        removeLastRow
    Next
End If

mRows = Value

paintView
enableDrawing True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the starting or ending row or column for a range of cells. Not available at design time.
Public Property Get RowSel() As Long
Const ProcName As String = "RowSel"
On Error GoTo Err

RowSel = mRowSel

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Determines the starting or ending row or column for a range of cells. Not available at design time.
Public Property Let RowSel(ByVal Value As Long)
Const ProcName As String = "RowSel"
On Error GoTo Err

AssertArgument Value >= 0 And Value <= mRows - 1, "Value out of bounds"

SelectCells mRow, mCol, Value, mColSel

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the row sizing mode.
Public Property Get RowSizingMode() As RowSizingSettings
Const ProcName As String = "RowSizingMode"
On Error GoTo Err

RowSizingMode = mRowSizingMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the row sizing mode.
Public Property Let RowSizingMode(ByVal Value As RowSizingSettings)
Const ProcName As String = "RowSizingMode"
On Error GoTo Err

mRowSizingMode = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingRowSizingMode, mRowSizingMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
'Const ProcName As String = "ScaleHeight"
'On Error GoTo Err
'
'TransPanel.ScaleHeight() = New_ScaleHeight
'PropertyChanged "ScaleHeight"
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

'
'
'Public Property Get ScaleHeight() As Single
'Const ProcName As String = "ScaleHeight"
'On Error GoTo Err
'
'ScaleHeight = TransPanel.ScaleHeight
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

'Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
'Const ProcName As String = "ScaleLeft"
'On Error GoTo Err
'
'TransPanel.ScaleLeft() = New_ScaleLeft
'PropertyChanged "ScaleLeft"
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

'
'
'Public Property Get ScaleLeft() As Single
'Const ProcName As String = "ScaleLeft"
'On Error GoTo Err
'
'ScaleLeft = TransPanel.ScaleLeft
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

'Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
'Const ProcName As String = "ScaleMode"
'On Error GoTo Err
'
'TransPanel.ScaleMode() = New_ScaleMode
'PropertyChanged "ScaleMode"
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

'
'
'Public Property Get ScaleMode() As Integer
'Const ProcName As String = "ScaleMode"
'On Error GoTo Err
'
'ScaleMode = TransPanel.ScaleMode
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

'Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
'Const ProcName As String = "ScaleTop"
'On Error GoTo Err
'
'TransPanel.ScaleTop() = New_ScaleTop
'PropertyChanged "ScaleTop"
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

'
'
'Public Property Get ScaleTop() As Single
'Const ProcName As String = "ScaleTop"
'On Error GoTo Err
'
'ScaleTop = TransPanel.ScaleTop
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

'Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
'Const ProcName As String = "ScaleWidth"
'On Error GoTo Err
'
'TransPanel.ScaleWidth() = New_ScaleWidth
'PropertyChanged "ScaleWidth"
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

'
'
'Public Property Get ScaleWidth() As Single
'Const ProcName As String = "ScaleWidth"
'On Error GoTo Err
'
'ScaleWidth = TransPanel.ScaleWidth
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

' Returns or sets whether the Hierarchical FlexGrid has horizontal or vertical scroll bars.
Public Property Get ScrollBars() As ScrollBarsSettings
Const ProcName As String = "ScrollBars"
On Error GoTo Err

ScrollBars = mScrollBars

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets whether the Hierarchical FlexGrid has horizontal or vertical scroll bars.
Public Property Let ScrollBars(ByVal Value As ScrollBarsSettings)
Const ProcName As String = "ScrollBars"
On Error GoTo Err

If Value <> mScrollBars Then
    mScrollBars = Value
    If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingScrollBars, mScrollBars
    repaintView
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'' Returns or sets whether the Hierarchical FlexGrid should scroll its contents while the user moves the scroll box along the scroll bars.
'Public Property Get ScrollTrack() As Boolean
'End Property
'
'' Returns or sets whether the Hierarchical FlexGrid should scroll its contents while the user moves the scroll box along the scroll bars.
'Public Property Let ScrollTrack(ByVal Value As Boolean)
'End Property

' Returns or sets whether a Hierarchical FlexGrid should allow regular cell selection, selection by rows, or selection by columns.
Public Property Get SelectionMode() As SelectionModeSettings
Const ProcName As String = "SelectionMode"
On Error GoTo Err

SelectionMode = mSelectionMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets whether a Hierarchical FlexGrid should allow regular cell selection, selection by rows, or selection by columns.
Public Property Let SelectionMode(ByVal Value As SelectionModeSettings)
Const ProcName As String = "SelectionMode"
On Error GoTo Err

mSelectionMode = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingSelectionMode, mSelectionMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' An action-type Public Property erty that sorts selected rows according to specified criteria. Not available at design time; write-only at run time.
Public Property Let Sort(ByVal Value As Long)
Const ProcName As String = "Sort"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the text content of a cell or range of cells.
Public Property Get Text() As String
Attribute Text.VB_UserMemId = -517
Const ProcName As String = "Text"
On Error GoTo Err

Text = getCell(mRow, mCol).Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the text content of a cell or range of cells.
Public Property Let Text(ByVal Value As String)
Const ProcName As String = "Text"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    setCellValue mRow, mCol, Value
    If Not mEditingCell Then paintCell mRow, mCol
Else
    sel = getSelection
    
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            setCellValue i, j, Value
            paintCell i, j
        Next
    Next
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the text content of an arbitrary cell (single subscript).
Public Property Get TextArray( _
                    ByVal index As Long) As String
Const ProcName As String = "TextArray"
On Error GoTo Err

TextArray = TextMatrix(Int(index / mCols), index Mod mCols)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the text content of an arbitrary cell (single subscript).
Public Property Let TextArray( _
                    ByVal index As Long, _
                    ByVal Value As String)
Const ProcName As String = "TextArray"
On Error GoTo Err

TextMatrix(Int(index / mCols), index Mod mCols) = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the text content of an arbitrary cell (row/column subscripts).
Public Property Get TextMatrix( _
                    ByVal pRow As Long, _
                    ByVal pCol As Long) As String
Const ProcName As String = "TextMatrix"
On Error GoTo Err

TextMatrix = getCell(pRow, pCol).Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the text content of an arbitrary cell (row/column subscripts).
Public Property Let TextMatrix( _
                    ByVal pRow As Long, _
                    ByVal pCol As Long, _
                    ByVal Value As String)
Const ProcName As String = "TextMatrix"
On Error GoTo Err

setCellValue pRow, pCol, Value
paintCell pRow, pCol

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets 3-D effects for displaying text.
Public Property Get TextStyle() As TextStyleSettings
Const ProcName As String = "TextStyle"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets 3-D effects for displaying text.
Public Property Let TextStyle(ByVal Value As TextStyleSettings)
Const ProcName As String = "TextStyle"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets 3-D effects for displaying text.
Public Property Get TextStyleFixed() As TextStyleSettings
Const ProcName As String = "TextStyleFixed"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets 3-D effects for displaying text.
Public Property Let TextStyleFixed(ByVal Value As TextStyleSettings)
Const ProcName As String = "TextStyleFixed"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

If mTheme Is Value Then Exit Property
Set mTheme = Value
If mTheme Is Nothing Then Exit Property

enableDrawing False

Appearance = mTheme.Appearance
BorderStyle = mTheme.BorderStyle
BackColorBkg = mTheme.GridRowBackColorEven
BackColor = mTheme.BackColor
BackColorFixed = mTheme.GridBackColorFixed
Font = mTheme.BaseFont
If Not mTheme.GridFont Is Nothing Then Font = mTheme.GridFont
FontFixed = mTheme.BaseFont
If Not mTheme.GridFontFixed Is Nothing Then FontFixed = mTheme.GridFontFixed
ForeColor = mTheme.GridForeColor
ForeColorFixed = mTheme.GridForeColorFixed
GridColor = mTheme.GridLineColor
GridColorFixed = mTheme.GridLineColorFixed
RowBackColorEven = mTheme.GridRowBackColorEven
RowBackColorOdd = mTheme.GridRowBackColorOdd
RowForeColorEven = mTheme.GridRowForeColorEven
RowForeColorOdd = mTheme.GridRowForeColorOdd

enableDrawing True

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

' Returns or sets the uppermost row displayed in the Hierarchical FlexGrid. Not available at design time.
Public Property Get TopRow() As Long
Const ProcName As String = "TopRow"
On Error GoTo Err

TopRow = mTopRow

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets the uppermost row displayed in the Hierarchical FlexGrid. Not available at design time.
Public Property Let TopRow(ByVal Value As Long)
Const ProcName As String = "TopRow"
On Error GoTo Err

ScrollToRow Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'Public Property Get TransparencyColor() As OLE_COLOR
'Const ProcName As String = "TransparencyColor"
'On Error GoTo Err
'
'TransparencyColor = TransPanel.TransparencyColor
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property
'
'Public Property Let TransparencyColor( _
'                ByVal Value As OLE_COLOR)
'Const ProcName As String = "TransparencyColor"
'On Error GoTo Err
'
'TransPanel.TransparencyColor = Value
'
'Exit Property
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Property

' Returns or sets whether text within a cell should be allowed to wrap to multiple lines.
Public Property Get WordWrap() As Boolean
Const ProcName As String = "WordWrap"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

' Returns or sets whether text within a cell should be allowed to wrap to multiple lines.
Public Property Let WordWrap(ByVal Value As Boolean)
Const ProcName As String = "WordWrap"
On Error GoTo Err

Assert False, "Property not implemented", ErrorCodes.ErrUnsupportedOperationException

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

' Adds a new row of values to a Hierarchical FlexGrid control at run time.
Public Sub AddItem( _
                    ByVal Item As String, _
                    Optional ByVal pIndex As Long = -1)
Const ProcName As String = "AddItem"
On Error GoTo Err

enableDrawing False
clearView

insertARow pIndex

If pIndex = -1 Then pIndex = mRows - 1

Dim values() As String
values = Split(Item, vbTab)

Dim i As Long
Do While i <= UBound(values) And i < mCols
    getCell(pIndex, i).Value = values(i)
    i = i + 1
Loop

paintView
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function AllocateBorders() As Long
Const ProcName As String = "AllocateBorders"
On Error GoTo Err

AllocateBorders = mNextBordersIndex
mNextBordersIndex = mNextBordersIndex + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function AllocateCell() As Long
Const ProcName As String = "AllocateCell"
On Error GoTo Err

AllocateCell = mNextCellIndex
mNextCellIndex = mNextCellIndex + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub BeginCellEdit(ByVal pRow As Long, ByVal pColumn As Long)
Const ProcName As String = "BeginCellEdit"
On Error GoTo Err
Assert Not mEditingCell, "Already editing a cell"

mEditingCell = True
mSavedRow = mRow
mRow = pRow
mSavedCol = mCol
mCol = pColumn

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

' Clears the contents of the TWGrid, excluding column headers. Any individual cell formatting is removed.
Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

Dim i As Long
Dim j As Long

enableDrawing False

For i = mFixedRows To mRows - 1
    For j = 0 To mCols - 1
        getCell(i, j).Clear
        paintCell i, j
    Next
Next

ScrollToCell 0, 0
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

' Clears the contents of the TWGrid, excluding column headers, but retains cell formatting.
Public Sub ClearContents()
Const ProcName As String = "ClearContents"
On Error GoTo Err

Dim i As Long
Dim j As Long

enableDrawing False

For i = mFixedRows To mRows - 1
    For j = 0 To mCols - 1
        TextMatrix(i, j) = ""
    Next
Next

ScrollToCell 0, 0
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

' Clears information about the order and name of columns displayed.
Public Sub ClearStructure()
Const ProcName As String = "ClearStructure"
On Error GoTo Err

enableDrawing False

clearView

Initialise
ScrollToCell 0, 0

enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'Public Sub DrawLine( _
'        ByVal x1 As Long, _
'        ByVal y1 As Long, _
'        ByVal x2 As Long, _
'        ByVal y2 As Long)
'Const ProcName As String = "DrawLine"
'On Error GoTo Err
'
'TransPanel.DrawLine x1, y1, x2, y2
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

Public Sub EndCellEdit()
Const ProcName As String = "EndCellEdit"
On Error GoTo Err

Assert mEditingCell, "Not currently editing a cell"

mEditingCell = False

paintCell mRow, mCol

mRow = mSavedRow
mCol = mSavedCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ExtendSelection( _
                ByVal pRow As Long, _
                ByVal pCol As Long)
Const ProcName As String = "ExtendSelection"
On Error GoTo Err

hideSelection
mRowSel = pRow
mColSel = pCol
showSelection

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetBottomBorderLabel( _
                ByVal index As Long) As Control
Const ProcName As String = "GetBottomBorderLabel"
On Error GoTo Err

If index > BottomBorder.UBound Then
    Load BottomBorder(index)
End If
Set GetBottomBorderLabel = BottomBorder(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetCellAlignment( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As AlignmentSettings
Const ProcName As String = "GetCellAlignment"
On Error GoTo Err

With getCell(pRow, pCol)
    If .IsSetAlign Then
        GetCellAlignment = .Align
    ElseIf isCellFixed(pRow, pCol) Then
        GetCellAlignment = mColTable(pCol).FixedAlign
    Else
        GetCellAlignment = mColTable(pCol).Align
    End If
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetCellBackcolor( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Long
Const ProcName As String = "GetCellBackcolor"
On Error GoTo Err

With getCell(pRow, pCol)
    If .BackColor = 0 Then
        If isCellFixed(.Row, .Col) Then
            GetCellBackcolor = mBackColorFixed
        ElseIf pRow Mod 2 = 0 Then
            If mRowBackColorEven <> 0 Then
                GetCellBackcolor = mRowBackColorEven
            Else
                GetCellBackcolor = mBackColor
            End If
        Else
            If mRowBackColorOdd <> 0 Then
                GetCellBackcolor = mRowBackColorOdd
            Else
                GetCellBackcolor = mBackColor
            End If
        End If
    Else
        GetCellBackcolor = .BackColor
    End If
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetCellForecolor( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Long
Const ProcName As String = "GetCellForecolor"
On Error GoTo Err

With getCell(pRow, pCol)
    If .ForeColor = 0 Then
        If isCellFixed(.Row, .Col) Then
            GetCellForecolor = mForeColorFixed
        ElseIf pRow Mod 2 = 0 Then
            If mRowForeColorEven <> 0 Then
                GetCellForecolor = mRowForeColorEven
            Else
                GetCellForecolor = mForeColor
            End If
        Else
            If mRowForeColorOdd <> 0 Then
                GetCellForecolor = mRowForeColorOdd
            Else
                GetCellForecolor = mForeColor
            End If
        End If
    Else
        GetCellForecolor = .ForeColor
    End If
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetCellLabel( _
                ByVal index As Long) As Control
Const ProcName As String = "GetCellLabel"
On Error GoTo Err

If index > Cell.UBound Then
    Load Cell(index)
End If
Set GetCellLabel = Cell(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetColFromX( _
                ByVal X As Long) As Long
Const ProcName As String = "GetColFromX"
On Error GoTo Err

Dim i As Long

GetColFromX = -1

For i = 0 To mFixedCols - 1
    If getCellLeft(mTopRow, i) > X Then
        GetColFromX = i - 1
        Exit For
    End If
Next

If GetColFromX <> -1 Then Exit Function

'For i = mFixedCols To mRightCol + 1
For i = mLeftCol To mRightCol + 1
    If getCellLeft(mTopRow, i) > X Then
        If i = mLeftCol Then
            GetColFromX = mFixedCols - 1
        Else
            GetColFromX = i - 1
        End If
        Exit For
    End If
Next

If GetColFromX = -1 Then GetColFromX = mRightCol + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Friend Function GetDefaultCellFont( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As StdFont
Const ProcName As String = "GetDefaultCellFont"
On Error GoTo Err

If isCellFixed(pRow, pCol) Then
    Set GetDefaultCellFont = mDefaultFixedFont
Else
    Set GetDefaultCellFont = mDefaultCellFont
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetLeftBorderLabel( _
                ByVal index As Long) As Control
Const ProcName As String = "GetLeftBorderLabel"
On Error GoTo Err

If index > LeftBorder.UBound Then
    Load LeftBorder(index)
End If
Set GetLeftBorderLabel = LeftBorder(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetRightBorderLabel( _
                ByVal index As Long) As Control
Const ProcName As String = "GetRightBorderLabel"
On Error GoTo Err

If index > RightBorder.UBound Then
    Load RightBorder(index)
End If
Set GetRightBorderLabel = RightBorder(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetTextHeight( _
                ByVal Font As StdFont, _
                ByVal IsFixed As Boolean) As Long
Const ProcName As String = "GetTextHeight"
On Error GoTo Err

If Not Font Is Nothing Then
    Set FontPicture.Font = Font
    GetTextHeight = FontPicture.textHeight(SampleCellText)
ElseIf IsFixed Then
    GetTextHeight = mDefaultFixedTextHeight
Else
    GetTextHeight = mDefaultCellTextHeight
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetTopBorderLabel( _
                ByVal index As Long) As Control
Const ProcName As String = "GetTopBorderLabel"
On Error GoTo Err

If index > TopBorder.UBound Then
    Load TopBorder(index)
End If
Set GetTopBorderLabel = TopBorder(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetRowFromY( _
                ByVal Y As Long) As Long
Const ProcName As String = "GetRowFromY"
On Error GoTo Err

Dim i As Long

GetRowFromY = -1

For i = 0 To mFixedRows - 1
    If getCellTop(i, mLeftCol) > Y Then
        GetRowFromY = i - 1
        Exit For
    End If
Next

If GetRowFromY <> -1 Then Exit Function

'For i = mFixedRows To mBottomRow + 1
For i = mTopRow To mBottomRow + 1
    If getCellTop(i, mLeftCol) > Y Then
        If i = mTopRow Then
            GetRowFromY = mFixedRows - 1
        Else
            GetRowFromY = i - 1
        End If
        Exit For
    End If
Next

If GetRowFromY = -1 Then GetRowFromY = mBottomRow + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Friend Function GetValueLabel( _
                ByVal index As Long) As Control
Const ProcName As String = "GetValueLabel"
On Error GoTo Err

If index > ValueLabel.UBound Then
    Load ValueLabel(index)
End If
Set GetValueLabel = ValueLabel(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub HideHottrack()
Const ProcName As String = "HideHottrack"
On Error GoTo Err

HottrackLabel.Visible = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub InvertCellColors()
Const ProcName As String = "InvertCellColors"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mFillStyle = TwGridFillSingle Or mEditingCell Then
    With getCell(mRow, mCol)
        .Invert
        If Not mEditingCell Then .Paint
    End With
Else
    sel = getSelection
    
    For i = sel.RowMin To sel.RowMax
        For j = sel.ColMin To sel.ColMax
            With getCell(i, j)
                .Invert
                .Paint
            End With
        Next
    Next
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub InsertRow(Optional ByVal pIndex As Long = -1)
Const ProcName As String = "InsertRow"
On Error GoTo Err

enableDrawing False
clearView

insertARow pIndex

paintView
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub LoadFromConfig( _
                ByVal config As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

Set mConfig = config
If mConfig Is Nothing Then Exit Sub

If mConfig.GetSetting(ConfigSettingSelectionMode) <> "" Then mSelectionMode = mConfig.GetSetting(ConfigSettingSelectionMode)
If mConfig.GetSetting(ConfigSettingScrollBars) <> "" Then mScrollBars = mConfig.GetSetting(ConfigSettingScrollBars)
If mConfig.GetSetting(ConfigSettingRowSizingMode) <> "" Then mRowSizingMode = mConfig.GetSetting(ConfigSettingRowSizingMode)
If mConfig.GetSetting(ConfigSettingRowHeightMin) <> "" Then mRowHeightMin = mConfig.GetSetting(ConfigSettingRowHeightMin)
If mConfig.GetSetting(ConfigSettingRowBackColorOdd) <> "" Then mRowBackColorOdd = mConfig.GetSetting(ConfigSettingRowBackColorOdd)
If mConfig.GetSetting(ConfigSettingRowBackColorEven) <> "" Then mRowBackColorEven = mConfig.GetSetting(ConfigSettingRowBackColorEven)
If mConfig.GetSetting(ConfigSettingHighlight) <> "" Then mHighlight = mConfig.GetSetting(ConfigSettingHighlight)
If mConfig.GetSetting(ConfigSettingGridLines) <> "" Then mGridLines = mConfig.GetSetting(ConfigSettingGridLines)
If mConfig.GetSetting(ConfigSettingGridLinesFixed) <> "" Then mGridLinesFixed = mConfig.GetSetting(ConfigSettingGridLinesFixed)
If mConfig.GetSetting(ConfigSettingGridLineWidth) <> "" Then mGridLineWidth = mConfig.GetSetting(ConfigSettingGridLineWidth)
If mConfig.GetSetting(ConfigSettingGridColorFixed) <> "" Then mGridColorFixed = mConfig.GetSetting(ConfigSettingGridColorFixed)
If mConfig.GetSetting(ConfigSettingGridColor) <> "" Then mGridColor = mConfig.GetSetting(ConfigSettingGridColor)
If mConfig.GetSetting(ConfigSettingForeColorFixed) <> "" Then mForeColorFixed = mConfig.GetSetting(ConfigSettingForeColorFixed)
If mConfig.GetSetting(ConfigSettingForeColor) <> "" Then mForeColor = mConfig.GetSetting(ConfigSettingForeColor)
If Not mConfig.GetConfigurationSection(ConfigSectionFont) Is Nothing Then
    Set mDefaultCellFont = loadFontFromSettings(mConfig.GetConfigurationSection(ConfigSectionFont))
    setDefaultCellFont
End If
If Not mConfig.GetConfigurationSection(ConfigSectionFontFixed) Is Nothing Then
    Set mDefaultFixedFont = loadFontFromSettings(mConfig.GetConfigurationSection(ConfigSectionFontFixed))
    setDefaultFixedFont
End If
If mConfig.GetSetting(ConfigSettingBorderStyle) <> "" Then mBorderStyle = mConfig.GetSetting(ConfigSettingBorderStyle)
If mConfig.GetSetting(ConfigSettingBackColorFixed) <> "" Then mBackColorFixed = mConfig.GetSetting(ConfigSettingBackColorFixed)
If mConfig.GetSetting(ConfigSettingBackColorBkg) <> "" Then mBackColorBkg = mConfig.GetSetting(ConfigSettingBackColorBkg)
If mConfig.GetSetting(ConfigSettingBackColor) <> "" Then mBackColor = mConfig.GetSetting(ConfigSettingBackColor)
If mConfig.GetSetting(ConfigSettingAllowUserResizing) <> "" Then mAllowUserResizing = mConfig.GetSetting(ConfigSettingAllowUserResizing)
If mConfig.GetSetting(ConfigSettingAllowUserReordering) <> "" Then mAllowUserReordering = mConfig.GetSetting(ConfigSettingAllowUserReordering)

If mConfig.GetConfigurationSection(ConfigSectionColumns) Is Nothing Then
    mConfig.AddConfigurationSection ConfigSectionColumns
    storeColumnSettings
Else
    loadColumnSettings
End If

repaintView

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub MoveColumn( _
                ByVal pFromCol As Long, _
                ByVal pToCol As Long)
Const ProcName As String = "MoveColumn"
On Error GoTo Err

Dim tempCol As ColTableEntry
Dim tempCell As GridCell

Dim startcol As Long
Dim endCol As Long
Dim stepBy As Long
Dim i As Long
Dim j As Long

If pFromCol < FixedCols Then pFromCol = FixedCols
If pToCol < FixedCols Then pToCol = FixedCols
If pFromCol >= Cols Then pFromCol = Cols - 1
If pToCol >= Cols Then pToCol = Cols - 1

If pFromCol = pToCol Then Exit Sub

enableDrawing False
clearView

tempCol = mColTable(pFromCol)

If pFromCol < pToCol Then
    startcol = pFromCol + 1
    endCol = pToCol
    stepBy = 1
    For i = startcol To endCol Step stepBy
        mColTable(i - stepBy) = mColTable(i)
        If Not mConfig Is Nothing Then
            mConfig.GetConfigurationSection(ConfigSectionColumns).GetConfigurationSection( _
                                                                    ConfigSectionColumn & "(" & mColTable(i).Id & ")").MoveDown
        End If
    Next
Else
    startcol = pFromCol - 1
    endCol = pToCol
    stepBy = -1
    For i = startcol To endCol Step stepBy
        mColTable(i - stepBy) = mColTable(i)
        If Not mConfig Is Nothing Then
            mConfig.GetConfigurationSection(ConfigSectionColumns).GetConfigurationSection( _
                                                                    ConfigSectionColumn & "(" & mColTable(i).Id & ")").MoveUp
        End If
    Next
End If

mColTable(pToCol) = tempCol

For i = 0 To mRows - 1
    Set tempCell = mRowTable(i).Cells(pFromCol)
    For j = startcol To endCol Step stepBy
        Set mRowTable(i).Cells(j - stepBy) = mRowTable(i).Cells(j)
        mRowTable(i).Cells(j - stepBy).Col = j - stepBy
    Next
    Set mRowTable(i).Cells(pToCol) = tempCell
    mRowTable(i).Cells(pToCol).Col = pToCol
Next

paintView
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

' Forces a complete repaint of a form or control.
Public Sub MoveRow( _
                ByVal pFromRow As Long, _
                ByVal pToRow As Long)
Const ProcName As String = "moveRow"
On Error GoTo Err

Dim lTempRow As RowTableEntry
Dim lStartRow As Long
Dim lEndRow As Long
Dim lStepBy As Long
Dim i As Long
Dim j As Long

If pFromRow < FixedRows Then pFromRow = FixedRows
If pToRow < FixedRows Then pToRow = FixedRows
If pFromRow >= Rows Then pFromRow = Rows - 1
If pToRow >= Rows Then pToRow = Rows - 1

If pFromRow = pToRow Then Exit Sub

enableDrawing False
clearView

If pFromRow < pToRow Then
    lStartRow = pFromRow + 1
    lEndRow = pToRow
    lStepBy = 1
Else
    lStartRow = pFromRow - 1
    lEndRow = pToRow
    lStepBy = -1
End If

lTempRow = mRowTable(pFromRow)

For i = lStartRow To lEndRow Step lStepBy
    mRowTable(i - lStepBy) = mRowTable(i)
    For j = 0 To mCols - 1
        mRowTable(i).Cells(j).Row = i - lStepBy
    Next
Next

mRowTable(pToRow) = lTempRow
For j = 0 To mCols - 1
    mRowTable(pToRow).Cells(j).Row = pToRow
Next

paintView
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Refresh()
Const ProcName As String = "Refresh"
On Error GoTo Err

UserControl.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveItem( _
                ByVal pRow As Long)
Const ProcName As String = "RemoveItem"
On Error GoTo Err

Dim i As Long
Dim j As Long
Dim rte As RowTableEntry

clearView

For i = 0 To mCols - 1
    removeCell pRow, i
Next

If pRow < mRows - 1 Then
    CopyMemory VarPtr(mRowTable(pRow)), VarPtr(mRowTable(pRow + 1)), Len(rte) * (mRows - 1 - pRow)
End If
mRows = mRows - 1

' now need to adjust the cellTable entries for all cells in rows following the removed
' row by decrementing the row number
For i = pRow To mRows - 1
    For j = 0 To mCols - 1
        getCell(i, j).Row = getCell(i, j).Row - 1
    Next
Next

paintView

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''The Underscore following "Scale" is necessary because it
''is a Reserved Word in VBA.
'
'
'Public Sub Scale_(Optional x1 As Variant, Optional y1 As Variant, Optional x2 As Variant, Optional y2 As Variant)
'Const ProcName As String = "Scale_"
'On Error GoTo Err
'
'TransPanel.Scale_ x1, y1, x2, y2
'
'Exit Sub
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Sub

'
'
'Public Function ScaleX(ByVal Width As Single, ByVal FromScale As Variant, ByVal ToScale As Variant) As Single
'Const ProcName As String = "ScaleX"
'On Error GoTo Err
'
'ScaleX = TransPanel.ScaleX(Width, FromScale, ToScale)
'
'Exit Function
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Function

'
'
'Public Function ScaleY(ByVal Height As Single, ByVal FromScale As Variant, ByVal ToScale As Variant) As Single
'Const ProcName As String = "ScaleY"
'On Error GoTo Err
'
'ScaleY = TransPanel.ScaleY(Height, FromScale, ToScale)
'
'Exit Function
'
'Err:
'gHandleUnexpectedError ProcName, ModuleName
'End Function

Public Sub ScrollToCell( _
                ByVal pRow As Long, _
                ByVal pCol As Long)
Const ProcName As String = "ScrollToCell"
On Error GoTo Err

enableDrawing False
clearView

If pCol < mFixedCols Then
    mLeftCol = mFixedCols
ElseIf pCol < HScroll.Max Then
    mLeftCol = pCol
Else
    mLeftCol = HScroll.Max
End If
HScroll.Value = mLeftCol

If pRow < mFixedRows Then
    mTopRow = mFixedRows
ElseIf pRow < VScroll.Max Then
    mTopRow = pRow
Else
    mTopRow = VScroll.Max
End If
VScroll.Value = mTopRow

paintView
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScrollToCol( _
                ByVal pCol As Long)
Const ProcName As String = "ScrollToCol"
On Error GoTo Err

enableDrawing False
clearView
If pCol < mFixedCols Then
    mLeftCol = mFixedCols
ElseIf pCol < HScroll.Max Then
    mLeftCol = pCol
Else
    mLeftCol = HScroll.Max
End If
HScroll.Value = mLeftCol
paintView
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ScrollToRow( _
                ByVal pRow As Long)
Const ProcName As String = "ScrollToRow"
On Error GoTo Err

enableDrawing False
clearView
If pRow < mFixedRows Then
    mTopRow = mFixedRows
ElseIf pRow < VScroll.Max Then
    mTopRow = pRow
Else
    mTopRow = VScroll.Max
End If
VScroll.Value = mTopRow
paintView
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SelectAll()
Const ProcName As String = "SelectAll"
On Error GoTo Err

SelectCells mFixedRows, mFixedCols, mRows - 1, mCols - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SelectCell( _
                ByVal pRow As Long, _
                ByVal pCol As Long)
Const ProcName As String = "SelectCell"
On Error GoTo Err

SelectCells pRow, pCol, pRow, pCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SelectCells( _
                ByVal pRow1 As Long, _
                ByVal pCol1 As Long, _
                ByVal pRow2 As Long, _
                ByVal pCol2 As Long)
Const ProcName As String = "SelectCells"
On Error GoTo Err

hideSelection
mRow = pRow1
mRowSel = pRow2
mCol = pCol1
mColSel = pCol2
showSelection

RaiseEvent SelectionChanged(mRow, mCol, mRowSel, mColSel)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SelectCol( _
                ByVal pCol As Long)
Const ProcName As String = "SelectCol"
On Error GoTo Err

SelectCells mFixedRows, pCol, mRows - 1, pCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SelectRow( _
                ByVal pRow As Long)
Const ProcName As String = "SelectRow"
On Error GoTo Err

SelectCells pRow, mFixedCols, pRow, mCols - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ShowHottrack( _
                ByVal pLeft As Long, _
                ByVal pTop As Long, _
                ByVal Width As Long)
Const ProcName As String = "ShowHottrack"
On Error GoTo Err

HottrackLabel.Move pLeft, pTop - HottrackLabel.Height, Width
HottrackLabel.ZOrder 0
HottrackLabel.Visible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addCell( _
                ByVal pRow As Long, _
                ByVal pCol As Long)
Const ProcName As String = "addCell"
On Error GoTo Err

Set mRowTable(pRow).Cells(pCol) = New GridCell
mRowTable(pRow).Cells(pCol).Initialise Me, pRow, pCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub addCol( _
                ByVal Id As String)
Const ProcName As String = "addCol"
On Error GoTo Err

Dim i As Long

If mCols > UBound(mColTable) Then
    ReDim Preserve mColTable(2 * (UBound(mColTable) + 1) - 1) As ColTableEntry
End If

mColTable(mCols).Width = -1
mColTable(mCols).Align = 0
mColTable(mCols).Id = IIf(Id <> "", Id, GenerateGUIDString)
If Not mConfig Is Nothing Then storeColSettings mCols

If mRows > 0 Then
    For i = 0 To mRows - 1
        ReDim Preserve mRowTable(i).Cells(mCols) As GridCell
        addCell i, mCols
    Next
End If

mCols = mCols + 1

calcHScroll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addRow()
Const ProcName As String = "addRow"
On Error GoTo Err

Dim i As Long

If mRows > UBound(mRowTable) Then
    ReDim Preserve mRowTable(2 * (UBound(mRowTable) + 1) - 1) As RowTableEntry
End If

mRowTable(mRows).Height = -1

If mCols > 0 Then
    ReDim mRowTable(mRows).Cells(mCols - 1) As GridCell
    For i = 0 To mCols - 1
        addCell mRows, i
    Next
End If

mRows = mRows + 1

calcVScroll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub adjustForHScroll()
Const ProcName As String = "adjustForHScroll"
On Error GoTo Err

Dim rowTop As Long

rowTop = mColNextTop - RowHeight(mBottomRow) - gMaximumLongs(mGridLineWidthTwipsY, mGridLineWidthTwipsYFixed)
Do While rowTop >= HScroll.Top
    unmapRow mBottomRow
    mBottomRow = mBottomRow - 1
    rowTop = rowTop - RowHeight(mBottomRow) - gMaximumLongs(mGridLineWidthTwipsY, mGridLineWidthTwipsYFixed)
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub adjustForVScroll()
Const ProcName As String = "adjustForVScroll"
On Error GoTo Err

Dim colLeft As Long

colLeft = mRowNextLeft - ColWidth(mRightCol) - gMaximumLongs(mGridLineWidthTwipsX, mGridLineWidthTwipsXFixed)
Do While colLeft >= VScroll.Left
    unmapCol mRightCol
    mRightCol = mRightCol - 1
    colLeft = colLeft - ColWidth(mRightCol) - gMaximumLongs(mGridLineWidthTwipsX, mGridLineWidthTwipsXFixed)
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function alignToPixelX( _
                ByVal Value As Long) As Long
Const ProcName As String = "alignToPixelX"
On Error GoTo Err

alignToPixelX = Int((Value + Screen.TwipsPerPixelX - 1) / Screen.TwipsPerPixelX) * Screen.TwipsPerPixelX

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function alignToPixelY( _
                ByVal Value As Long) As Long
Const ProcName As String = "alignToPixelY"
On Error GoTo Err

alignToPixelY = Int((Value + Screen.TwipsPerPixelY - 1) / Screen.TwipsPerPixelY) * Screen.TwipsPerPixelY

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function availableScrollSpaceH() As Long
Const ProcName As String = "availableScrollSpaceH"
On Error GoTo Err

availableScrollSpaceH = availableSpaceH - mFixedColsWidth

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function availableScrollSpaceV() As Long
Const ProcName As String = "availableScrollSpaceV"
On Error GoTo Err

availableScrollSpaceV = availableSpaceV - mFixedRowsHeight

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function availableSpaceH() As Long
Const ProcName As String = "availableSpaceH"
On Error GoTo Err

If VScroll.Visible Then
    availableSpaceH = UserControl.ScaleWidth - VScroll.Width
Else
    availableSpaceH = UserControl.ScaleWidth
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function availableSpaceV() As Long
Const ProcName As String = "availableSpaceV"
On Error GoTo Err

If HScroll.Visible Then
    availableSpaceV = UserControl.ScaleHeight - HScroll.Height
Else
    availableSpaceV = UserControl.ScaleHeight
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub calcFixedColsWidth()
Const ProcName As String = "calcFixedColsWidth"
On Error GoTo Err

Dim i As Long
mFixedColsWidth = 0
For i = 0 To mFixedCols - 1
    mFixedColsWidth = mFixedColsWidth + ColWidth(i)
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub calcFixedRowsHeight()
Const ProcName As String = "calcFixedRowsHeight"
On Error GoTo Err

Dim i As Long

mFixedRowsHeight = 0
For i = 0 To mFixedRows - 1
    mFixedRowsHeight = mFixedRowsHeight + RowHeight(i)
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function calcGridLineWidthTwipsX(ByVal pGridLines As GridLineSettings, ByVal pGridLineWidth As Long) As Long
Select Case pGridLines
Case TwGridGridNone
    calcGridLineWidthTwipsX = 0
Case TwGridGridFlat
    calcGridLineWidthTwipsX = pGridLineWidth * Screen.TwipsPerPixelX
Case TwGridGridInset
    calcGridLineWidthTwipsX = 0
Case TwGridGridRaised
    calcGridLineWidthTwipsX = 0
End Select
End Function

Private Function calcGridLineWidthTwipsY(ByVal pGridLines As GridLineSettings, ByVal pGridLineWidth As Long) As Long
Select Case pGridLines
Case TwGridGridNone
    calcGridLineWidthTwipsY = 0
Case TwGridGridFlat
    calcGridLineWidthTwipsY = pGridLineWidth * Screen.TwipsPerPixelY
Case TwGridGridInset
    calcGridLineWidthTwipsY = 0
Case TwGridGridRaised
    calcGridLineWidthTwipsY = 0
End Select
End Function

Private Sub calcHScroll()
Const ProcName As String = "calcHScroll"
On Error GoTo Err

If Not mHScrollActive Then Exit Sub

HScroll.Min = mFixedCols
HScroll.Max = mCols

Dim availSpace As Long
availSpace = availableScrollSpaceH

Dim i As Long
For i = mCols - 1 To mFixedCols Step -1
    Dim Width As Long
    Width = Width + ColWidth(i) + mGridLineWidthTwipsX
    If Width > availSpace Then Exit For
    HScroll.Max = HScroll.Max - 1
Next

If HScroll.Max = mCols Then HScroll.Max = mCols - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub calcVScroll()
Const ProcName As String = "calcVScroll"
On Error GoTo Err

If Not mVScrollActive Then Exit Sub

VScroll.Min = mFixedRows
VScroll.Max = mRows

Dim availSpace As Long
availSpace = availableScrollSpaceV

Dim i As Long
For i = mRows - 1 To mFixedRows Step -1
    Dim Height As Long
    Height = Height + RowHeight(i) + mGridLineWidthTwipsY
    If Height > availSpace Then Exit For
    VScroll.Max = VScroll.Max - 1
Next

If VScroll.Max = mRows Then VScroll.Max = mRows - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub captureMouse()
Const ProcName As String = "captureMouse"
On Error GoTo Err

If Not mGotMouse Then
    SetCapture UserControl.hWnd
    mGotMouse = True
    Debug.Print "Capture mouse"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub checkCellHighlighting()
Const ProcName As String = "checkCellHighlighting"
On Error GoTo Err

Dim lCell As GridCell

If mCurrMouseRow < 0 Or mCurrMouseCol < 0 Then
    clearHottracking
    Exit Sub
End If

If Not isCellFixed(mCurrMouseRow, mCurrMouseCol) Then
    clearHottracking
    Exit Sub
End If

If isCellColHeader(mCurrMouseRow) Then
    If mouseInCellHorizResizeZone(mCurrMouseRow, mCurrMouseCol - 1, mCurrMouseX) Then
        Set lCell = getCell(mCurrMouseRow, mCurrMouseCol - 1)
    Else
        Set lCell = getCell(mCurrMouseRow, mCurrMouseCol)
    End If
ElseIf isCellRowHeader(mCurrMouseCol) Then
    If mouseInCellVertResizeZone(mCurrMouseRow - 1, mCurrMouseCol, mCurrMouseY) Then
        Set lCell = getCell(mCurrMouseRow - 1, mCurrMouseCol)
    Else
        Set lCell = getCell(mCurrMouseRow, mCurrMouseCol)
    End If
End If

If Not lCell Is mHighlightedCell Then
    clearHottracking
    If lCell.IsMapped Then
        lCell.HottrackOn
        Set mHighlightedCell = lCell
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearCellBorders( _
                ByVal cellIndex As Long)
Const ProcName As String = "clearCellBorders"
On Error GoTo Err

LeftBorder(cellIndex).Visible = False
RightBorder(cellIndex).Visible = False
TopBorder(cellIndex).Visible = False
BottomBorder(cellIndex).Visible = False

Err:

Exit Sub
End Sub

Private Sub clearFocusRect()
Const ProcName As String = "clearFocusRect"
On Error GoTo Err

FocusBox.Visible = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearHottracking()
Const ProcName As String = "clearHottracking"
On Error GoTo Err

If Not mHighlightedCell Is Nothing Then
    mHighlightedCell.HottrackOff
    Set mHighlightedCell = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearSelection()
Const ProcName As String = "clearSelection"
On Error GoTo Err

mRow = -1
mRowSel = -1
mCol = -1
mColSel = -1
RaiseEvent SelectionChanged(mRow, mCol, mRowSel, mColSel)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearView()
Const ProcName As String = "clearView"
On Error GoTo Err

Dim i As Long
Dim lCell As GridCell

If mMappedCells.Count = 0 Then Exit Sub

For Each lCell In mMappedCells
    unmapCell lCell.Col, lCell.Row
Next

Set mMappedCells = New Collection

mNextCellIndex = 1
mNextBordersIndex = 1

GridPicture.Visible = False
VScroll.Visible = False
mVScrollActive = False
HScroll.Visible = False
mHScrollActive = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub colMoverMove( _
                ByVal X As Long)
Const ProcName As String = "colMoverMove"
On Error GoTo Err

Dim Cancel As Boolean

If Not mCurrMouseCol < mLeftCol Then
    RaiseEvent ColMoving(mColMover.startcol, mCurrMouseCol, Cancel)
    If Not Cancel Then
        mColMover.moveTo mCurrMouseCol, _
                        getCellLeft(mTopRow, mCurrMouseCol), _
                        X
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createColMover( _
                ByVal pCell As GridCell, _
                ByVal X As Long)
Const ProcName As String = "createColMover"
On Error GoTo Err

Dim myRect As GDI_RECT

GetWindowRect UserControl.hWnd, myRect
Set mColMover = New ColumnMover
mColMover.Initialise pCell, _
                    myRect.Top * Screen.TwipsPerPixelY + _
                        pCell.Top + _
                        IIf(mBorderStyle = BorderStyleSettings.BorderStyleSingle, Screen.TwipsPerPixelY, 0), _
                    myRect.Left * Screen.TwipsPerPixelX + _
                        pCell.Left + _
                        IIf(mBorderStyle = BorderStyleSettings.BorderStyleSingle, Screen.TwipsPerPixelX, 0), _
                    GetColFromX(X), _
                    X, _
                    ColResizeLine, _
                    getFixedRowsHeight

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createColResizer( _
                ByVal pCol As Long)
Const ProcName As String = "createColResizer"
On Error GoTo Err

Set mColResizer = New ColumnResizer
mColResizer.Initialise pCol, _
                    ColPos(pCol), _
                    ColWidth(pCol), _
                    UserControl.ScaleHeight, _
                    ColResizeLine

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createRowMover( _
                ByVal pCell As GridCell, _
                ByVal Y As Long)
Const ProcName As String = "createRowMover"
On Error GoTo Err

Dim myRect As GDI_RECT

GetWindowRect UserControl.hWnd, myRect
Set mRowMover = New RowMover
mRowMover.Initialise pCell, _
                    (myRect.Top + 1) * Screen.TwipsPerPixelY + _
                        pCell.Top + _
                    IIf(mBorderStyle = BorderStyleSettings.BorderStyleSingle, Screen.TwipsPerPixelY, 0), _
                    myRect.Left * Screen.TwipsPerPixelX + _
                        pCell.Left + _
                        IIf(mBorderStyle = BorderStyleSettings.BorderStyleSingle, Screen.TwipsPerPixelX, 0), _
                    GetRowFromY(Y), _
                    Y, _
                    RowResizeLine, _
                    getFixedColsWidth

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub createRowResizer( _
                ByVal pRow As Long)
Const ProcName As String = "createRowResizer"
On Error GoTo Err

Debug.Print "Create row resizer"
Set mRowResizer = New RowResizer
mRowResizer.Initialise pRow, _
                    RowPos(pRow), _
                    RowHeight(pRow), _
                    UserControl.ScaleHeight, _
                    RowResizeLine

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deleteColMover()
Const ProcName As String = "deleteColMover"
On Error GoTo Err

Set mColMover = Nothing
mInFocus = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deleteColResizer()
Const ProcName As String = "deleteColResizer"
On Error GoTo Err

    If Not mColResizer Is Nothing Then
        mColResizer.endResize
        Set mColResizer = Nothing
    End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deleteRowMover()
Const ProcName As String = "deleteRowMover"
On Error GoTo Err

Set mRowMover = Nothing
mInFocus = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deleteRowResizer()
Const ProcName As String = "deleteRowResizer"
On Error GoTo Err

If Not mRowResizer Is Nothing Then
    Debug.Print "Delete row resizer"
    mRowResizer.endResize
    Set mRowResizer = Nothing
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub enableDrawing( _
                ByVal enable As Boolean, _
                Optional ByVal force As Boolean)
Const ProcName As String = "enableDrawing"
On Error GoTo Err

Static sCount As Long

If enable Then
    If sCount = 0 Then
        ' nowt to do
    ElseIf force Then
        sCount = 0
        LockWindowUpdate 0
    Else
        sCount = sCount - 1
        If sCount = 0 Then LockWindowUpdate 0
    End If
Else
    If sCount = 0 Then
        LockWindowUpdate GridPicture.hWnd
    End If
    sCount = sCount + 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ensureCellVisible( _
                ByVal pRow As Long, _
                ByVal pCol As Long)
Const ProcName As String = "ensureCellVisible"
On Error GoTo Err

Dim rowVis As Boolean
Dim colVis As Boolean

rowVis = RowIsVisible(pRow)
colVis = ColIsVisible(pCol)

If rowVis And colVis Then
    ' nothing to do
ElseIf rowVis Then
    ScrollToCol pCol
ElseIf colVis Then
    ScrollToRow pRow
Else
    ScrollToCell pRow, pCol
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getCell( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As GridCell
Const ProcName As String = "getCell"
On Error GoTo Err

If pRow >= mRows Or pCol >= mCols Then Exit Function
Set getCell = mRowTable(pRow).Cells(pCol)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellFontBold( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Boolean
Const ProcName As String = "getCellFontBold"
On Error GoTo Err

getCellFontBold = getCell(pRow, pCol).Font.Bold

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellFontItalic( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Boolean
Const ProcName As String = "getCellFontItalic"
On Error GoTo Err

getCellFontItalic = getCell(pRow, pCol).Font.Italic

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellFontName( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As String
Const ProcName As String = "getCellFontName"
On Error GoTo Err

getCellFontName = getCell(pRow, pCol).Font.Name

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellFontSize( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Single
Const ProcName As String = "getCellFontSize"
On Error GoTo Err

getCellFontSize = getCell(pRow, pCol).Font.Size

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellFontStrikethrough( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Boolean
Const ProcName As String = "getCellFontStrikethrough"
On Error GoTo Err

getCellFontStrikethrough = getCell(pRow, pCol).Font.Strikethrough

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellFontUnderline( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Boolean
Const ProcName As String = "getCellFontUnderline"
On Error GoTo Err

getCellFontUnderline = getCell(pRow, pCol).Font.Underline

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellHeight( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Long
Const ProcName As String = "getCellHeight"
On Error GoTo Err

getCellHeight = getCell(pRow, pCol).Height

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellLeft( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Long
Const ProcName As String = "getCellLeft"
On Error GoTo Err

If pCol >= mCols Then
    getCellLeft = mRowNextLeft
    Exit Function
End If

getCellLeft = getCell(pRow, pCol).Left

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellTop( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Long
Const ProcName As String = "getCellTop"
On Error GoTo Err

If pRow >= mRows Then
    getCellTop = mColNextTop
    Exit Function
End If

getCellTop = getCell(pRow, pCol).Top

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellValue( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As String
Const ProcName As String = "getCellValue"
On Error GoTo Err

getCellValue = getCell(pRow, pCol).Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getCellWidth( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Long
Const ProcName As String = "getCellWidth"
On Error GoTo Err

getCellWidth = getCell(pRow, pCol).Width

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getDefaultColWidth( _
                ByVal pCol As Long) As Long
Const ProcName As String = "getDefaultColWidth"
On Error GoTo Err

getDefaultColWidth = IIf(isCellRowHeader(pCol), mDefaultFixedWidth, mDefaultCellWidth)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getDefaultRowHeight( _
                ByVal pRow As Long) As Long
Const ProcName As String = "getDefaultRowHeight"
On Error GoTo Err

If isCellColHeader(pRow) Then
    getDefaultRowHeight = mDefaultFixedTextHeight + 2 * gTextPaddingTwips
ElseIf mDefaultRowHeight <> 0 Then
    getDefaultRowHeight = mDefaultRowHeight
ElseIf mDefaultFixedTextHeight > mDefaultCellTextHeight Then
    getDefaultRowHeight = mDefaultFixedTextHeight + 2 * gTextPaddingTwips
Else
    getDefaultRowHeight = mDefaultCellTextHeight + 2 * gTextPaddingTwips
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFixedColsWidth() As Long
Const ProcName As String = "getFixedColsWidth"
On Error GoTo Err

Dim i As Long

For i = 0 To mFixedCols - 1
    getFixedColsWidth = getFixedColsWidth + getCellWidth(0, i)
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFixedRowsHeight() As Long
Const ProcName As String = "getFixedRowsHeight"
On Error GoTo Err

Dim i As Long

For i = 0 To mFixedRows - 1
    getFixedRowsHeight = getFixedRowsHeight + getCellHeight(i, 0)
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getMouseCell( _
                ByVal X As Long, _
                ByVal Y As Long) As GridCell
Const ProcName As String = "getMouseCell"
On Error GoTo Err

Set getMouseCell = getCell(GetRowFromY(Y), GetColFromX(X))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getSelection() As SelectionSpecifier
Const ProcName As String = "getSelection"
On Error GoTo Err

If mColSel > mCol Then
    getSelection.ColMin = mCol
    getSelection.ColMax = mColSel
Else
    getSelection.ColMin = mColSel
    getSelection.ColMax = mCol
End If

If mRowSel > mRow Then
    getSelection.RowMin = mRow
    getSelection.RowMax = mRowSel
Else
    getSelection.RowMin = mRowSel
    getSelection.RowMax = mRow
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getUserControlFont() As StdFont
Const ProcName As String = "getUserControlFont"
On Error GoTo Err

Dim aFont As New StdFont
aFont.Bold = UserControl.Ambient.Font.Bold
aFont.Italic = UserControl.Ambient.Font.Italic
aFont.Name = UserControl.Ambient.Font.Name
aFont.Size = UserControl.Ambient.Font.Size
aFont.Strikethrough = UserControl.Ambient.Font.Strikethrough
aFont.Underline = UserControl.Ambient.Font.Underline
Set getUserControlFont = aFont

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub hideSelection()
Const ProcName As String = "hideSelection"
On Error GoTo Err

Dim sel As SelectionSpecifier
Dim i As Long
Dim j As Long

If mRow < 0 Or mCol < 0 Then Exit Sub

clearFocusRect

sel = getSelection
For i = sel.RowMin To sel.RowMax
    For j = sel.ColMin To sel.ColMax
        paintCell i, j
    Next
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Initialise()
Const ProcName As String = "Initialise"
On Error GoTo Err

enableDrawing False

mNextCellIndex = 1  ' we don't use Cell(0) so we can use index 0 to mean 'none'
mNextBordersIndex = 1
ReDim mRowTable(0) As RowTableEntry
ReDim mColTable(0) As ColTableEntry

Set mMappedCells = New Collection

mTopRow = 0
mBottomRow = 0
mLeftCol = 0
mRightCol = 0

clearSelection

mRows = 0
mCols = 0

mFixedRows = 0
mFixedCols = 0

mDefaultRowHeight = 0
mRowHeightMin = 0

mAppearance = AppearanceSettings.AppearanceFlat
mBorderStyle = BorderStyleSettings.BorderStyleSingle
mBackColorBkg = vbApplicationWorkspace
setUserControlAppearance

On Error Resume Next
setInitialFonts
On Error GoTo Err

mGridLines = TwGridGridFlat
mGridLinesFixed = TwGridGridRaised
mGridColor = SystemColorConstants.vbGrayText
mGridColorFixed = SystemColorConstants.vbGrayText
GridPicture.BackColor = mGridColorFixed
mGridLineWidth = 1

mForeColor = SystemColorConstants.vbWindowText
mForeColorFixed = SystemColorConstants.vbButtonText
mBackColor = SystemColorConstants.vbWindowBackground
mBackColorFixed = SystemColorConstants.vbButtonFace
mForeColorSel = SystemColorConstants.vbHighlightText
mBackColorSel = SystemColorConstants.vbHighlight

mHighlight = TwGridHighlightAlways
mFocusRect = TwGridFocusLight
mSelectionMode = TwGridSelectionFree
mAllowBigSelection = True

mAllowUserResizing = TwGridResizeBoth
mRowSizingMode = TwGridRowSizeIndividual

mScrollBars = TwGridScrollBarBoth
VScroll.LargeChange = ScrollLargeChange
VScroll.SmallChange = ScrollSmallChange
VScroll.Min = 0
VScroll.Max = 1
HScroll.Min = 0
HScroll.Max = 1

mCurrMouseRow = -1
mCurrMouseCol = -1
mCurrMouseX = -1
mCurrMouseY = -1

enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub insertARow(ByVal pIndex As Long)
Const ProcName As String = "insertARow"
On Error GoTo Err

addRow

If pIndex <> -1 Then
    MoveRow mRows - 1, pIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function isCellColHeader( _
                ByVal pRow As Long) As Boolean
Const ProcName As String = "isCellColHeader"
On Error GoTo Err

If pRow < mFixedRows Then isCellColHeader = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isCellCommonHeader( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Boolean
Const ProcName As String = "isCellCommonHeader"
On Error GoTo Err

If pRow < mFixedRows And _
    pCol < mFixedCols _
Then
    isCellCommonHeader = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isCellFixed( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Boolean
Const ProcName As String = "isCellFixed"
On Error GoTo Err

If pRow < mFixedRows Or _
    pCol < mFixedCols _
Then
    isCellFixed = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isCellRowHeader( _
                ByVal pCol As Long) As Boolean
Const ProcName As String = "isCellRowHeader"
On Error GoTo Err

If pCol < mFixedCols Then isCellRowHeader = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isMouseDown() As Boolean
Const ProcName As String = "isMouseDown"
On Error GoTo Err

isMouseDown = mLeftMouseDown Or mRightMouseDown Or mMiddleMouseDown

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub loadColumnSettings()
Const ProcName As String = "loadColumnSettings"
On Error GoTo Err

Dim cs As ConfigurationSection
Dim i As Long

For Each cs In mConfig.GetConfigurationSection(ConfigSectionColumns)
    If i > mCols - 1 Then
        addCol cs.InstanceQualifier
    Else
        mColTable(i).Id = cs.InstanceQualifier
    End If
    mColTable(i).Align = cs.GetSetting(ConfigSettingColumnAlignment)
    mColTable(i).FixedAlign = cs.GetSetting(ConfigSettingColumnFixedAlignment)
    mColTable(i).Width = cs.GetSetting(ConfigSettingColumnWidth)
    i = i + 1
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function loadFontFromSettings( _
                ByVal cs As ConfigurationSection) As StdFont
Const ProcName As String = "loadFontFromSettings"
On Error GoTo Err

Dim aFont As New StdFont
aFont.Bold = cs.GetSetting(ConfigSettingFontBold)
aFont.Name = cs.GetSetting(ConfigSettingFontName)
aFont.Italic = cs.GetSetting(ConfigSettingFontItalic)
aFont.Size = cs.GetSetting(ConfigSettingFontSize)
aFont.Strikethrough = cs.GetSetting(ConfigSettingFontStrikethrough)
aFont.Underline = cs.GetSetting(ConfigSettingFontUnderline)
Set loadFontFromSettings = aFont

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub mapCell( _
                ByVal pCol As Long, _
                ByVal pRow As Long, _
                ByVal pLeft As Long, _
                ByVal pTop As Long, _
                ByVal pWidth As Long, _
                ByVal pHeight As Long)
Const ProcName As String = "mapCell"
On Error GoTo Err

Dim lCell As GridCell
Set lCell = getCell(pRow, pCol)

Dim lFixed As Boolean
lFixed = isCellFixed(pRow, pCol)

lCell.Map lFixed, _
            mGridLines, _
            mGridLinesFixed, _
            pLeft, _
            pTop, _
            pWidth, _
            pHeight
mMappedCells.Add lCell, CStr(ObjPtr(lCell))
lCell.Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mapCol( _
                ByVal pCol As Long)
Const ProcName As String = "mapCol"
On Error GoTo Err

Dim Width As Long

mColNextTop = 0

Width = alignToPixelX(ColWidth(pCol))

mapColCells pCol, 0, mFixedRows - 1, Width
mapColCells pCol, mTopRow, mRows - 1, Width

mRightCol = pCol
mRowNextLeft = mRowNextLeft + Width '+ gMaximumLongs(mGridLineWidthTwipsXFixed, mGridLineWidthTwipsX)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mapColCells( _
                ByVal pCol As Long, _
                ByVal pFromRow As Long, _
                ByVal pToRow As Long, _
                ByVal pWidth As Long)
Const ProcName As String = "mapColCells"
On Error GoTo Err

Dim lRow As Long
Dim lHeight As Long

lRow = pFromRow
Do While mColNextTop < UserControl.ScaleHeight And lRow <= pToRow
    lHeight = alignToPixelY(RowHeight(lRow))
    mapCell pCol, _
            lRow, _
            mRowNextLeft, _
            mColNextTop, _
            pWidth, _
            lHeight
    mBottomRow = lRow
    mColNextTop = mColNextTop + lHeight '+ gMaximumLongs(mGridLineWidthTwipsYFixed, mGridLineWidthTwipsY)
    lRow = lRow + 1
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub MouseDown( _
                ByVal Button As Integer, _
                ByVal Shift As Integer, _
                ByVal X As Single, _
                ByVal Y As Single)
Const ProcName As String = "MouseDown"
On Error GoTo Err

Dim lCell As GridCell

setModifierKeyState Shift
setMouseButtonState Button

'captureMouse

If mouseOutOfBounds(X, Y) Then Exit Sub
    
Set lCell = setMousePosition(X, Y)

If moveRowResizer(Y) Then Exit Sub
If moveColResizer(X) Then Exit Sub

If Not lCell Is Nothing Then
    With lCell
        If isCellColHeader(.Row) And mColResizer Is Nothing Then
            If (mAllowUserReordering And TwGridReorderColumns) Then
                createColMover lCell, X
            End If
        End If
        If isCellRowHeader(.Col) And mRowResizer Is Nothing Then
            If (mAllowUserReordering And TwGridReorderRows) Then
                createRowMover lCell, Y
            End If
        End If
    End With
End If

RaiseEvent MouseDown(Button, _
                    Shift, _
                    UserControl.ScaleX(X, UserControl.ScaleMode, vbContainerPosition), _
                    UserControl.ScaleY(Y, UserControl.ScaleMode, vbContainerPosition))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function mouseInCellHorizResizeZone( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal X As Long) As Boolean
Const ProcName As String = "mouseInCellHorizResizeZone"
On Error GoTo Err

If pCol < 0 Then Exit Function

mouseInCellHorizResizeZone = getCell(pRow, pCol).IsXInCellHorizResizeZone(X)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function mouseInCellVertResizeZone( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal Y As Long) As Boolean
Const ProcName As String = "mouseInCellVertResizeZone"
On Error GoTo Err

If pRow < 0 Then Exit Function

mouseInCellVertResizeZone = getCell(pRow, pCol).IsYInCellVertResizeZone(Y)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function mouseInScrollzone( _
                ByVal X As Single, _
                ByVal Y As Single) As Boolean
If Not (mHScrollActive Or mVScrollActive) Then Exit Function
If mHScrollActive Then
    If Y >= HScroll.Top - HScroll.Height Then mouseInScrollzone = True: Exit Function
End If
If mVScrollActive Then
    If X >= VScroll.Left - VScroll.Width Then mouseInScrollzone = True
End If
End Function

Private Sub MouseMove( _
                ByVal hWnd As Long, _
                ByVal Button As Integer, _
                ByVal Shift As Integer, _
                ByVal X As Single, _
                ByVal Y As Single)
Const ProcName As String = "MouseMove"
On Error GoTo Err

Dim lCell As GridCell

setModifierKeyState Shift
setMouseButtonState Button

If mMouseTracker Is Nothing Then
    Set mMouseTracker = New MouseTracker
    mMouseTracker.Initialise hWnd
    mMouseTracker.TrackLeave
End If

If isMouseDown Then captureMouse

If mouseOutOfBounds(X, Y) Then Exit Sub

If mPopupScrollbars And mouseInScrollzone(X, Y) Then showScrollbars
    
Set lCell = setMousePosition(X, Y)

If moveColResizer(X) Then Exit Sub
If moveRowResizer(Y) Then Exit Sub

If moveColMover(X) Then Exit Sub
If moveRowMover(Y) Then Exit Sub

If lCell Is Nothing Then
    deleteRowResizer
    deleteColResizer
    
    clearHottracking
    RaiseEvent MouseMove(Button, _
                        Shift, _
                        UserControl.ScaleX(X, UserControl.ScaleMode, vbContainerPosition), _
                        UserControl.ScaleY(Y, UserControl.ScaleMode, vbContainerPosition))
    Exit Sub
End If

With lCell
    checkCellHighlighting
    
    If Not isCellRowHeader(.Col) Then
        deleteRowResizer
    Else
        If Not mRowResizer Is Nothing Then
            If Not mouseInCellVertResizeZone(mRowResizer.Row, .Col, Y) Then deleteRowResizer
        End If
        
        If mRowResizer Is Nothing Then
            If (mAllowUserResizing And TwGridResizeRows) Then
                If mouseInCellVertResizeZone(.Row - 1, .Col, Y) Then
                    createRowResizer .Row - 1
                ElseIf mouseInCellVertResizeZone(.Row, .Col, Y) Then
                    createRowResizer .Row
                End If
            End If
        End If
    End If
    
    If Not isCellColHeader(.Row) Then
        deleteColResizer
    Else
        If Not mColResizer Is Nothing Then
            If Not mouseInCellHorizResizeZone(.Row, mColResizer.Col, X) Then deleteColResizer
        End If
    
        If mColResizer Is Nothing Then
            If (mAllowUserResizing And TwGridResizeColumns) Then
                If mouseInCellHorizResizeZone(.Row, .Col - 1, X) Then
                    createColResizer .Col - 1
                ElseIf mouseInCellHorizResizeZone(.Row, .Col, X) Then
                    createColResizer .Col
                End If
            End If
        End If
    End If
End With

RaiseEvent MouseMove(Button, _
                    Shift, _
                    UserControl.ScaleX(X, UserControl.ScaleMode, vbContainerPosition), _
                    UserControl.ScaleY(Y, UserControl.ScaleMode, vbContainerPosition))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function mouseOutOfBounds( _
                ByVal X As Long, _
                ByVal Y As Long) As Boolean
Const ProcName As String = "mouseOutOfBounds"
On Error GoTo Err

If X < 0 Or _
    Y < 0 Or _
    X >= IIf(VScroll.Visible, VScroll.Left, UserControl.Width) Or _
    Y >= IIf(HScroll.Visible, HScroll.Top, UserControl.Height) _
Then
    If (Not mLeftMouseDown) And (Not mRightMouseDown) Then
        deleteRowResizer
        deleteColResizer
        deleteRowMover
        deleteColMover
        clearHottracking
        'releaseMouse
        Screen.MousePointer = vbDefault
        mouseOutOfBounds = True
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub MouseUp( _
                ByVal Button As Integer, _
                ByVal Shift As Integer, _
                ByVal X As Single, _
                ByVal Y As Single)
Const ProcName As String = "MouseUp"
On Error GoTo Err

Dim lCell As GridCell

setModifierKeyState Shift
setMouseButtonState Button

releaseMouse

clearHottracking

If mouseOutOfBounds(X, Y) Then Exit Sub
    
Set lCell = setMousePosition(X, Y)

If resizedRow Then Exit Sub
If resizedCol Then Exit Sub

hideSelection

If transposeCols Then Exit Sub
If transposeRows Then Exit Sub

If Not lCell Is Nothing Then
    With lCell
        If mShiftDown Then
            setExtendedSelection .Row, .Col
        Else
            setSelection .Row, .Col
        End If
        ensureCellVisible .Row, .Col
    End With
End If

'showSelection

RaiseEvent MouseUp(Button, _
                    Shift, _
                    UserControl.ScaleX(X, UserControl.ScaleMode, vbContainerPosition), _
                    UserControl.ScaleY(Y, UserControl.ScaleMode, vbContainerPosition))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
                    
End Sub

Private Function moveColMover( _
                ByVal X As Long) As Boolean
Const ProcName As String = "moveColMover"
On Error GoTo Err

If Not mColMover Is Nothing Then
    colMoverMove X
    moveColMover = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function moveColResizer( _
                ByVal X As Long) As Boolean
Const ProcName As String = "moveColResizer"
On Error GoTo Err

If Not mColResizer Is Nothing Then
    If mLeftMouseDown Then
        mColResizer.moveTo X
        moveColResizer = True
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function moveRowMover( _
                ByVal Y As Long) As Boolean
Const ProcName As String = "moveRowMover"
On Error GoTo Err

If Not mRowMover Is Nothing Then
    RowMoverMove Y
    moveRowMover = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function moveRowResizer( _
                ByVal Y As Long) As Boolean
Const ProcName As String = "moveRowResizer"
On Error GoTo Err

If Not mRowResizer Is Nothing Then
    If mLeftMouseDown Then
        mRowResizer.moveTo Y
        moveRowResizer = True
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needHScrollBar() As Boolean
Const ProcName As String = "needHScrollBar"
On Error GoTo Err

If (mScrollBars And TwGridScrollBarHorizontal) = 0 Then Exit Function
If mRowNextLeft > availableSpaceH Or _
    mLeftCol > mFixedCols Or _
    mRightCol < mCols - 1 Then needHScrollBar = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function needVScrollBar() As Boolean
Const ProcName As String = "needVScrollBar"
On Error GoTo Err

If (mScrollBars And TwGridScrollBarVertical) = 0 Then Exit Function
If mColNextTop > availableSpaceV Or _
    mTopRow > mFixedRows Or _
    mBottomRow < mRows - 1 Then needVScrollBar = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub pageDown()
Const ProcName As String = "pageDown"
On Error GoTo Err

ScrollToRow mBottomRow

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub pageLeft()
Const ProcName As String = "pageLeft"
On Error GoTo Err

Dim Width As Long
Dim i As Long

For i = mLeftCol - 1 To mFixedCols Step -1
    Width = Width + ColWidth(i)
    If Width > UserControl.ScaleWidth - mFixedColsWidth Then Exit For
Next
i = i + 1
If i < mLeftCol Then
    ScrollToCol i
Else
    ScrollToCol i - 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub pageUp()
Const ProcName As String = "pageUp"
On Error GoTo Err

Dim Height As Long
Dim i As Long

For i = mTopRow - 1 To mFixedRows Step -1
    Height = Height + RowHeight(i)
    If Height > UserControl.ScaleHeight - mFixedRowsHeight Then Exit For
Next
i = i + 1
If i < mTopRow Then
    ScrollToRow i
Else
    ScrollToRow i - 1
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub pageRight()
Const ProcName As String = "pageRight"
On Error GoTo Err

ScrollToCol mRightCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub paintCell( _
                ByVal pRow As Long, _
                ByVal pCol As Long)
Const ProcName As String = "paintCell"
On Error GoTo Err

getCell(pRow, pCol).Paint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub paintGrid()
Const ProcName As String = "paintGrid"
On Error GoTo Err

GridPicture.BackColor = mGridColor
GridPicture.Height = mColNextTop
GridPicture.Width = mRowNextLeft

' now paint rectangles the size of the fixed rows and columns to
' provide the fixed grid

GridPicture.FillColor = mGridColorFixed
GridPicture.FillStyle = vbFSSolid

'GridPicture.Line (0, 0)-(mRowNextLeft - Screen.TwipsPerPixelX, _
'                        getCellTop(mTopRow, mLeftCol) - mGridLineWidthTwipsX - Screen.TwipsPerPixelY), _
'                mGridColorFixed, _
'                B

GridPicture.Line (0, 0)-(mRowNextLeft, getCellTop(mTopRow, mLeftCol) - Screen.TwipsPerPixelY), mGridColorFixed, B

'GridPicture.Line (0, 0)-(getCellLeft(mTopRow, mLeftCol) - mGridLineWidthTwipsY - Screen.TwipsPerPixelX, _
'                        mColNextTop - Screen.TwipsPerPixelY), _
'                mGridColorFixed, _
'                B

GridPicture.Line (0, 0)-(getCellLeft(mTopRow, mLeftCol) - Screen.TwipsPerPixelX, mColNextTop), mGridColorFixed, B

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub paintView()
Const ProcName As String = "paintView"
On Error GoTo Err

Dim lCol As Long

If Not mRedraw Then Exit Sub

If mRows = 0 Or mCols = 0 Then Exit Sub

If mMappedCells.Count <> 0 Then clearView

mRowNextLeft = 0
mBottomRow = 0
mRightCol = 0

Do While mRowNextLeft < UserControl.ScaleWidth And lCol < mFixedCols
    mapCol lCol
    lCol = lCol + 1
Loop

lCol = mLeftCol
Do While mRowNextLeft < UserControl.ScaleWidth And lCol < mCols
    mapCol lCol
    lCol = lCol + 1
Loop

paintGrid
setScrollBars

setMousePosition mCurrMouseX, mCurrMouseY

GridPicture.Visible = True

ColResizeLine.y2 = GridPicture.Height
RowResizeLine.x2 = GridPicture.Width

FocusBox.ZOrder 0
showSelection

Set mHighlightedCell = Nothing
checkCellHighlighting

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub releaseMouse()
Const ProcName As String = "releaseMouse"
On Error GoTo Err

If mGotMouse Then
    ReleaseCapture
    mGotMouse = False
    Debug.Print "Release mouse"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function removeCell( _
                ByVal pRow As Long, _
                ByVal pCol As Long) As Long
Const ProcName As String = "removeCell"
On Error GoTo Err

On Error Resume Next
mMappedCells.Remove CStr(ObjPtr(getCell(pRow, pCol)))
On Error GoTo Err

getCell(pRow, pCol).Finish
Set mRowTable(pRow).Cells(pCol) = Nothing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub removeLastCol()
Const ProcName As String = "removeLastCol"
On Error GoTo Err

Dim i As Long

If mRows > 0 Then
    For i = 0 To mRows - 1
        removeCell i, mCols - 1
    Next
End If

If Not mConfig Is Nothing Then mConfig.GetConfigurationSection(ConfigSectionColumns).RemoveConfigurationSection ConfigSectionColumn & "(" & mColTable(mCols - 1).Id & ")"

mColTable(mCols - 1).Align = 0
mColTable(mCols - 1).FixedAlign = 0
mColTable(mCols - 1).Id = ""
mColTable(mCols - 1).Width = 0

mCols = mCols - 1
If mCol >= mCols Then mCol = mCols - 1
If mColSel >= mCols Then mColSel = mCols - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeLastRow()
Const ProcName As String = "removeLastRow"
On Error GoTo Err

Dim i As Long

If mCols > 0 Then
    For i = 0 To mCols - 1
        removeCell mRows - 1, i
    Next
End If
mRows = mRows - 1
If mRow >= mRows Then mRow = mRows - 1
If mRowSel >= mRows Then mRowSel = mRows - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub repaintView()
Const ProcName As String = "repaintView"
On Error GoTo Err

enableDrawing False
clearView
paintView
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resize()
Const ProcName As String = "resize"
On Error GoTo Err

Static prevHeight As Long
Static prevWidth As Long

If UserControl.Height = 0 Or _
    UserControl.Width = 0 _
Then
    Exit Sub
End If

If UserControl.Height = prevHeight And _
    UserControl.Width = prevWidth _
Then
    Exit Sub
End If

prevHeight = UserControl.Height
prevWidth = UserControl.Width

GridPicture.Width = UserControl.ScaleWidth
GridPicture.Height = UserControl.ScaleHeight

If (UserControl.ScaleHeight - HScroll.Height) < 0 Then
    HScroll.Top = 0
Else
    HScroll.Top = UserControl.ScaleHeight - HScroll.Height
End If
HScroll.Left = 0
HScroll.Width = UserControl.ScaleWidth

VScroll.Top = 0
If (UserControl.ScaleWidth - VScroll.Width) < 0 Then
    VScroll.Left = 0
Else
    VScroll.Left = UserControl.ScaleWidth - VScroll.Width
End If
VScroll.Height = UserControl.ScaleHeight

'FillerPicture.Width = VScroll.Width
'FillerPicture.Height = HScroll.Height
'FillerPicture.Left = VScroll.Left
'FillerPicture.Top = HScroll.Top

'TransPanel.Left = 0
'TransPanel.Top = 0
'TransPanel.Width = UserControl.Width
'TransPanel.Height = UserControl.Height

repaintView

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function resizedCol() As Boolean
Const ProcName As String = "resizedCol"
On Error GoTo Err

If mColResizer Is Nothing Then Exit Function

ColWidth(mColResizer.Col) = ColWidth(mColResizer.Col) + mColResizer.endResize
deleteColResizer

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function resizedRow() As Boolean
Const ProcName As String = "resizedRow"
On Error GoTo Err

If mRowResizer Is Nothing Then Exit Function

If mRowSizingMode = TwGridRowSizeIndividual Or mRowResizer.Row < mFixedRows Then
    RowHeight(mRowResizer.Row) = RowHeight(mRowResizer.Row) + mRowResizer.endResize
Else
    mDefaultRowHeight = RowHeight(mRowResizer.Row) + mRowResizer.endResize
    If mDefaultRowHeight < mRowHeightMin Then mDefaultRowHeight = mRowHeightMin
    calcFixedRowsHeight
    calcVScroll
    repaintView
End If
deleteRowResizer
resizedRow = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub RowMoverMove( _
                ByVal Y As Long)
Const ProcName As String = "RowMoverMove"
On Error GoTo Err

Dim Cancel As Boolean

If Not mCurrMouseRow < mTopRow Then
    RaiseEvent RowMoving(mRowMover.StartRow, mCurrMouseRow, Cancel)
    If Not Cancel Then
        mRowMover.moveTo mCurrMouseRow, _
                        getCellTop(mCurrMouseRow, mLeftCol), _
                        Y
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub setCellAlignment( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal pAlign As AlignmentSettings)
Const ProcName As String = "setCellAlignment"
On Error GoTo Err

getCell(pRow, pCol).Align = pAlign

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCellBackColor( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal Value As Long)
Const ProcName As String = "setCellBackColor"
On Error GoTo Err

getCell(pRow, pCol).BackColor = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCellFontBold( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal Value As Boolean)
Const ProcName As String = "setCellFontBold"
On Error GoTo Err

getCell(pRow, pCol).Bold = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCellFontItalic( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal Value As Boolean)
Const ProcName As String = "setCellFontItalic"
On Error GoTo Err

getCell(pRow, pCol).Italic = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCellFontName( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal Value As String)
Const ProcName As String = "setCellFontName"
On Error GoTo Err

getCell(pRow, pCol).FontName = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCellFontSize( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal Value As Single)
Const ProcName As String = "setCellFontSize"
On Error GoTo Err

getCell(pRow, pCol).FontSize = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCellFontStrikethrough( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal Value As Boolean)
Const ProcName As String = "setCellFontStrikethrough"
On Error GoTo Err

getCell(pRow, pCol).Strikethrough = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCellFontUnderline( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal Value As Boolean)
Const ProcName As String = "setCellFontUnderline"
On Error GoTo Err

getCell(pRow, pCol).Underline = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCellForeColor( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal Value As Long)
Const ProcName As String = "setCellForeColor"
On Error GoTo Err

getCell(pRow, pCol).ForeColor = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setCellValue( _
                ByVal pRow As Long, _
                ByVal pCol As Long, _
                ByVal Value As String)
Const ProcName As String = "setCellValue"
On Error GoTo Err

getCell(pRow, pCol).Value = Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setDefaultCellFont()
Const ProcName As String = "setDefaultCellFont"
On Error GoTo Err

Set FontPicture.Font = mDefaultCellFont
mDefaultCellTextWidth = FontPicture.textWidth(SampleCellText)
mDefaultCellWidth = mDefaultCellTextWidth + 2 * gTextPaddingTwips
mDefaultCellTextHeight = FontPicture.textHeight(SampleCellText)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setDefaultFixedFont()
Const ProcName As String = "setDefaultFixedFont"
On Error GoTo Err

Set FontPicture.Font = mDefaultFixedFont
mDefaultFixedTextWidth = FontPicture.textWidth(SampleFixedText)
mDefaultFixedWidth = mDefaultFixedTextWidth + 2 * gTextPaddingTwips
mDefaultFixedTextHeight = FontPicture.textHeight(SampleFixedText)

calcFixedColsWidth
calcFixedRowsHeight

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setExtendedSelection( _
                ByVal pRow As Long, _
                ByVal pCol As Long)
Const ProcName As String = "setExtendedSelection"
On Error GoTo Err

If isCellCommonHeader(pRow, pCol) Then
    If mAllowBigSelection Then
        SelectAll
    ElseIf mSelectionMode = TwGridSelectionFree Then
        ExtendSelection mFixedRows, mFixedCols
    End If
ElseIf isCellColHeader(pRow) Then
    If mSelectionMode = TwGridSelectionByColumn Then
        ExtendSelection Rows - 1, pCol
    ElseIf mSelectionMode = TwGridSelectionByRow Then
        ' do nothing
    ElseIf mAllowBigSelection Then
        ExtendSelection Rows - 1, pCol
    Else
        ExtendSelection mFixedRows, pCol
    End If
ElseIf isCellRowHeader(pCol) Then
    If mSelectionMode = TwGridSelectionByColumn Then
        ' do nothing
    ElseIf mSelectionMode = TwGridSelectionByRow Then
        ExtendSelection pRow, mCols - 1
    ElseIf mAllowBigSelection Then
        ExtendSelection pRow, Cols - 1
    Else
        ExtendSelection pRow, mFixedCols
    End If
Else
    ExtendSelection pRow, pCol
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setInitialFonts()
Const ProcName As String = "setInitialFonts"
On Error GoTo Err

Set mDefaultCellFont = getUserControlFont
setDefaultCellFont

Set mDefaultFixedFont = getUserControlFont
setDefaultFixedFont

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setModifierKeyState( _
                ByVal Shift As Integer)
Const ProcName As String = "setModifierKeyState"
On Error GoTo Err

mShiftDown = (Shift And vbShiftMask)
mControlDown = (Shift And vbCtrlMask)
mAltDown = (Shift And vbAltMask)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setMouseButtonState( _
                ByVal buttons As Long)
Const ProcName As String = "setMouseButtonState"
On Error GoTo Err

mLeftMouseDown = buttons And vbLeftButton
mRightMouseDown = buttons And vbRightButton
mMiddleMouseDown = buttons And vbMiddleButton

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setMousePosition( _
                ByVal X As Long, _
                ByVal Y As Long) As GridCell
Const ProcName As String = "setMousePosition"
On Error GoTo Err

If X < 0 Or Y < 0 Then Exit Function

mCurrMouseX = X
mCurrMouseY = Y

Set setMousePosition = getMouseCell(X, Y)
If setMousePosition Is Nothing Then
    mCurrMouseRow = -1
    mCurrMouseCol = -1
Else
    mCurrMouseRow = setMousePosition.Row
    mCurrMouseCol = setMousePosition.Col
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setScrollBars()
Const ProcName As String = "setScrollBars"
On Error GoTo Err

HScroll.Visible = False
mHScrollActive = False

VScroll.Visible = False
mVScrollActive = False

'FillerPicture.Visible = False

If needHScrollBar Then
    mHScrollActive = True
    If Not mPopupScrollbars Then HScroll.Visible = True
    'HScroll.Enabled = True
    HScroll.Width = UserControl.Width
    If needVScrollBar Then
        mVScrollActive = True
        If Not mPopupScrollbars Then VScroll.Visible = True
        'VScroll.Enabled = True
        VScroll.Height = UserControl.Height - HScroll.Height
        HScroll.Width = UserControl.Width - VScroll.Width
        'If Not mPopupScrollbars Then FillerPicture.Visible = True
    End If
ElseIf needVScrollBar Then
    mVScrollActive = True
    If Not mPopupScrollbars Then VScroll.Visible = True
    'VScroll.Enabled = True
    VScroll.Height = UserControl.Height
    If needHScrollBar Then
        mHScrollActive = True
        If Not mPopupScrollbars Then HScroll.Visible = True
        'HScroll.Enabled = True
        VScroll.Height = UserControl.Height - HScroll.Height
        HScroll.Width = UserControl.Width - VScroll.Width
        'If Not mPopupScrollbars Then FillerPicture.Visible = True
    End If
End If

If mHScrollActive Then
    If Not mPopupScrollbars Then adjustForHScroll
    calcHScroll
End If

If mVScrollActive Then
    If Not mPopupScrollbars Then adjustForVScroll
    calcVScroll
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setSelection( _
                ByVal pRow As Long, _
                ByVal pCol As Long)
Const ProcName As String = "setSelection"
On Error GoTo Err

If isCellCommonHeader(pRow, pCol) Then
    If mAllowBigSelection Then SelectAll
ElseIf isCellColHeader(pRow) Then
    If mAllowBigSelection Or mSelectionMode = TwGridSelectionByColumn Then SelectCol pCol
ElseIf isCellRowHeader(pCol) Then
    If mAllowBigSelection Or mSelectionMode = TwGridSelectionByRow Then SelectRow pRow
Else
    If mSelectionMode = TwGridSelectionByColumn Then
        SelectCol pCol
    ElseIf mSelectionMode = TwGridSelectionByRow Then
        SelectRow pRow
    Else
        SelectCell pRow, pCol
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setUserControlAppearance()
Const ProcName As String = "setUserControlAppearance"
On Error GoTo Err

enableDrawing False
clearView

' these three properties have to all be set together, in this order, otherwise
' they produce odd interactions
UserControl.Appearance = mAppearance
UserControl.BorderStyle = mBorderStyle
UserControl.BackColor = mBackColorBkg

paintView
enableDrawing True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showFocusRect( _
                ByVal pCell As GridCell)
Const ProcName As String = "showFocusRect"
On Error GoTo Err

If mFocusRect = TwGridFocusNone Then Exit Sub
If Not pCell.IsMapped Then Exit Sub

pCell.Paint

Dim lGridlineWidthX As Long
lGridlineWidthX = IIf(pCell.IsFixed, mGridLineWidthTwipsXFixed, mGridLineWidthTwipsX)

Dim lGridlineWidthY As Long
lGridlineWidthY = IIf(pCell.IsFixed, mGridLineWidthTwipsYFixed, mGridLineWidthTwipsY)

Dim lLeft As Long
Dim lTop As Long
Dim lWidth As Long
Dim lHeight As Long

Select Case mFocusRect
Case TwGridFocusNone
    Exit Sub
Case TwGridFocusLight
    FocusBox.BorderWidth = 1
    FocusBox.BorderStyle = VBRUN.BorderStyleConstants.vbBSSolid
    lLeft = pCell.Left
    lWidth = pCell.Width - lGridlineWidthX
    lTop = pCell.Top
    lHeight = pCell.Height - lGridlineWidthY
Case TwGridFocusHeavy
    FocusBox.BorderWidth = 2
    FocusBox.BorderStyle = VBRUN.BorderStyleConstants.vbBSSolid
    lLeft = pCell.Left + 1 * Screen.TwipsPerPixelX
    lWidth = pCell.Width - 1 * Screen.TwipsPerPixelX - lGridlineWidthX
    lTop = pCell.Top + 1 * Screen.TwipsPerPixelY
    lHeight = pCell.Height - 1 * Screen.TwipsPerPixelY - lGridlineWidthY
Case TwGridFocusBroken
    FocusBox.BorderWidth = 1
    FocusBox.BorderStyle = VBRUN.BorderStyleConstants.vbBSDot
    lLeft = pCell.Left
    lWidth = pCell.Width - lGridlineWidthX
    lTop = pCell.Top
    lHeight = pCell.Height - lGridlineWidthY
Case Else
    Exit Sub
End Select

FocusBox.BorderColor = mFocusRectColor
FocusBox.Move lLeft, lTop, lWidth, lHeight
FocusBox.Visible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showScrollbars()
Const ProcName As String = "showScrollbars"
On Error GoTo Err

If Not mScrollbarHideTLI Is Nothing Then mScrollbarHideTLI.Cancel

Set mScrollbarHideTLI = GetGlobalTimerList.Add(Empty, 2)

If mHScrollActive Then HScroll.Visible = True
If mVScrollActive Then VScroll.Visible = True
'If mHScrollActive And mVScrollActive Then FillerPicture.Visible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showSelection()
Const ProcName As String = "showSelection"
On Error GoTo Err

clearFocusRect

If mRow < 0 Or mCol < 0 Then Exit Sub

If mHighlight = TwGridHighlightNever Then Exit Sub
If mHighlight = TwGridHighlightWithFocus And Not mInFocus Then Exit Sub

Dim sel As SelectionSpecifier
sel = getSelection

Dim i As Long
For i = sel.RowMin To sel.RowMax
    Dim j As Long
    For j = sel.ColMin To sel.ColMax
        getCell(i, j).PaintSelected mForeColorSel, mBackColorSel
    Next
Next

showFocusRect getCell(mRow, mCol)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeColumnSettings()
Const ProcName As String = "storeColumnSettings"
On Error GoTo Err

Dim i As Long

For i = 0 To mCols - 1
    storeColSettings i
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeColSettings(ByVal index As Long)
Const ProcName As String = "storeColSettings"
On Error GoTo Err

Dim cs As ConfigurationSection

Set cs = mConfig.AddConfigurationSection(ConfigSectionColumns).AddConfigurationSection( _
                                                                ConfigSectionColumn & _
                                                                "(" & mColTable(index).Id & ")")
cs.SetSetting ConfigSettingColumnWidth, mColTable(index).Width
cs.SetSetting ConfigSettingColumnAlignment, mColTable(index).Align
cs.SetSetting ConfigSettingColumnFixedAlignment, mColTable(index).FixedAlign

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeDefaultCellFontSettings()
Const ProcName As String = "storeDefaultCellFontSettings"
On Error GoTo Err

storeFontSettings mConfig.AddConfigurationSection(ConfigSectionFont), mDefaultCellFont

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeDefaultFixedFontSettings()
Const ProcName As String = "storeDefaultFixedFontSettings"
On Error GoTo Err

storeFontSettings mConfig.AddConfigurationSection(ConfigSectionFontFixed), mDefaultFixedFont

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeFontSettings( _
                ByVal cs As ConfigurationSection, _
                ByVal theFont As StdFont)
Const ProcName As String = "storeFontSettings"
On Error GoTo Err

cs.SetSetting ConfigSettingFontBold, theFont.Bold
cs.SetSetting ConfigSettingFontName, theFont.Name
cs.SetSetting ConfigSettingFontItalic, theFont.Italic
cs.SetSetting ConfigSettingFontSize, theFont.Size
cs.SetSetting ConfigSettingFontStrikethrough, theFont.Strikethrough
cs.SetSetting ConfigSettingFontUnderline, theFont.Underline

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function transposeCols() As Boolean
Const ProcName As String = "transposeCols"
On Error GoTo Err

Dim fromCol As Long
Dim toCol As Long

If mColMover Is Nothing Then Exit Function

fromCol = mColMover.startcol
toCol = mColMover.EndMove
If toCol = -1 Then
ElseIf toCol < fromCol Then
    MoveColumn fromCol, toCol
    clearSelection
    transposeCols = True
    RaiseEvent ColMoved(fromCol, toCol)
ElseIf toCol > fromCol Then
    MoveColumn fromCol, toCol
    clearSelection
    transposeCols = True
    RaiseEvent ColMoved(fromCol, toCol)
Else
    transposeCols = True
End If
deleteColMover

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function transposeRows() As Boolean
Const ProcName As String = "transposeRows"
On Error GoTo Err

Dim lFromRow As Long
Dim lToRow As Long

If mRowMover Is Nothing Then Exit Function

lFromRow = mRowMover.StartRow
lToRow = mRowMover.EndMove
If lToRow = -1 Then
ElseIf lToRow < lFromRow Then
    MoveRow lFromRow, lToRow
    clearSelection
    transposeRows = True
    RaiseEvent RowMoved(lFromRow, lToRow)
ElseIf lToRow > lFromRow + 1 Then
    MoveRow lFromRow, lToRow - 1
    clearSelection
    transposeRows = True
    RaiseEvent RowMoved(lFromRow, lToRow - 1)
Else
    transposeRows = True
End If
deleteRowMover

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub unmapCell( _
                ByVal pCol As Long, _
                ByVal pRow As Long)
Const ProcName As String = "unmapCell"
On Error GoTo Err

getCell(pRow, pCol).UnMap
mMappedCells.Remove CStr(ObjPtr(getCell(pRow, pCol)))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unmapCol( _
                ByVal pCol As Long)
Const ProcName As String = "unmapCol"
On Error GoTo Err

If pCol > mRightCol Then Exit Sub

unmapColCells pCol, 0, mFixedRows - 1
unmapColCells pCol, mTopRow, mBottomRow

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unmapColCells( _
                ByVal pCol As Long, _
                ByVal pFromRow As Long, _
                ByVal pToRow As Long)
Const ProcName As String = "unmapColCells"
On Error GoTo Err

Dim lRow As Long

For lRow = pFromRow To pToRow
    unmapCell pCol, lRow
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unmapRow( _
                ByVal pRow As Long)
Const ProcName As String = "unmapRow"
On Error GoTo Err

If pRow > mBottomRow Then Exit Sub

unmapRowCells pRow, 0, mFixedCols - 1
unmapRowCells pRow, mLeftCol, mRightCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub unmapRowCells( _
                ByVal pRow As Long, _
                ByVal pFromCol As Long, _
                ByVal pToCol As Long)
Const ProcName As String = "unmapRowCells"
On Error GoTo Err

Dim lCol As Long

For lCol = pFromCol To pToCol
    unmapCell lCol, pRow
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


