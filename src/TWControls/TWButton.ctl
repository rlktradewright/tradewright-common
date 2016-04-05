VERSION 5.00
Begin VB.UserControl TWButton 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3135
   DefaultCancel   =   -1  'True
   ScaleHeight     =   1800
   ScaleWidth      =   3135
   Begin VB.CommandButton Button 
      Height          =   1095
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "TWButton"
Attribute VB_GlobalNameSpace = False
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

Implements IDeferredAction
Implements ISubclassable
Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event Click()
Attribute Click.VB_UserMemId = -600
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseEnter()
Event MouseLeave()
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

'@================================================================================
' Enums
'@================================================================================

Private Enum DeferredActions
    DeferredActionFireMouseEnter
    DeferredActionFireMouseLeave
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "TWButton"

Private Const RoundedRectHeight                     As Long = 9
Private Const RoundedRectWidth                      As Long = 9

Private Const TwipsPerPoint                         As Long = 28

'@================================================================================
' Member variables
'@================================================================================

Private mForeColor                                  As OLE_COLOR

Private mPrevWindowProcAddress                      As Long

Private WithEvents mFont                            As StdFont
Attribute mFont.VB_VarHelpID = -1
Private mFontHandle                                 As Long

Private WithEvents mDisabledFont                    As StdFont
Attribute mDisabledFont.VB_VarHelpID = -1
Private mDisabledFontHandle                         As Long

Private WithEvents mMouseOverFont                   As StdFont
Attribute mMouseOverFont.VB_VarHelpID = -1
Private mMouseOverFontHandle                        As Long

Private WithEvents mPushedFont                      As StdFont
Attribute mPushedFont.VB_VarHelpID = -1
Private mPushedFontHandle                           As Long

Private mBackColor                                  As Long
Private mhBackBrush                                 As Long

Private mDisabledBackColor                          As OLE_COLOR
Private mhDisabledBrush                             As Long
Private mDisabledForeColor                          As OLE_COLOR

Private mMouseOverBackColor                         As OLE_COLOR
Private mhMouseOverBrush                            As Long
Private mMouseOverForeColor                         As OLE_COLOR

Private mPushedBackColor                            As OLE_COLOR
Private mhPushedBrush                               As Long
Private mPushedForeColor                            As OLE_COLOR

Private mGotMouse                                   As Boolean
Private mPushed                                     As Boolean
Private mMouseDown                                  As Boolean
Private mTrackingMouse                              As Boolean

Private mDefaultBorderColor                         As Long
Private mhDefaultBorderPen                          As Long

Private mNonDefaultBorderColor                      As Long
Private mhNonDefaultBorderPen                       As Long

Private mFocusedBorderColor                         As Long
Private mhFocusedBorderPen                          As Long

Private mClientRect                                 As GDI_RECT

Private mInhibitClickEvent                          As Boolean

Private mTheme                                      As ITheme

Private mNoDraw                                     As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
Debug.Print "AccessKeyPress: " & KeyAscii
If KeyAscii = 13 Or KeyAscii = 27 Then RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
If PropertyName = "DisplayAsDefault" Then
    debugPrint True, "DisplayAsDefault: " & CStr(UserControl.Ambient.DisplayAsDefault)
    InvalidateRect Button.hWnd, mClientRect, 1
    debugPrint False, "DisplayAsDefault"
End If
End Sub

Private Sub UserControl_EnterFocus()
If Enabled Then Button.SetFocus
End Sub

Private Sub UserControl_Initialize()
Set mFont = New StdFont
Set UserControl.Font = mFont
End Sub

Private Sub UserControl_InitProperties()
ForeColor = vbButtonText
BackColor = vbButtonFace
'MouseOverBackColor = vbButtonFace
'MouseOverForeColor = vbButtonText
'PushedBackColor = vb3DHighlight
'PushedForeColor = vbButtonText
'DisabledBackColor = vbInactiveBorder
'DisabledForeColor = vbButtonText
DefaultBorderColor = &HF0FF00
NonDefaultBorderColor = vbBlack
FocusedBorderColor = &HF8FFA8
Set Font = UserControl.Ambient.Font
Set UserControl.Font = mFont
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
DefaultBorderColor = PropBag.ReadProperty("DefaultBorderColor", &HEEEECF)
DisabledBackColor = PropBag.ReadProperty("DisabledBackColor", vbInactiveBorder)
Set DisabledFont = PropBag.ReadProperty("DisabledFont", Nothing)
DisabledForeColor = PropBag.ReadProperty("DisabledForeColor", vbBlack)
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
FocusedBorderColor = PropBag.ReadProperty("FocusedBorderColor", &HF8FFA8)
Set Font = PropBag.ReadProperty("Font", UserControl.Ambient.Font)
ForeColor = PropBag.ReadProperty("ForeColor", vbButtonText)
MouseOverBackColor = PropBag.ReadProperty("MouseOverBackColor", vbButtonFace)
Set MouseOverFont = PropBag.ReadProperty("MouseOverFont", Nothing)
MouseOverForeColor = PropBag.ReadProperty("MouseOverforeColor", vbBlack)
NonDefaultBorderColor = PropBag.ReadProperty("NonDefaultBorderColor", vbBlack)
PushedBackColor = PropBag.ReadProperty("PushedBackColor", vb3DHighlight)
Set PushedFont = PropBag.ReadProperty("PushedFont", Nothing)
PushedForeColor = PropBag.ReadProperty("PushedForeColor", vbBlack)

Button.Appearance = PropBag.ReadProperty("Appearance", 1)
Button.Cancel = PropBag.ReadProperty("Cancel", False)
Button.Caption = PropBag.ReadProperty("Caption", "")
Button.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
Button.MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
Button.MousePointer = PropBag.ReadProperty("MousePointer", 0)
Button.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
Button.ToolTipText = PropBag.ReadProperty("ToolTipText", "")

RightToLeft = PropBag.ReadProperty("RightToLeft", False)

StartSubclassing Me
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

resize

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub UserControl_Terminate()
DeleteObject mhBackBrush
DeleteObject mhMouseOverBrush
DeleteObject mhPushedBrush
DeleteObject mhDefaultBorderPen
DeleteObject mhNonDefaultBorderPen
trackMouse Button.hWnd, pCancel:=True
If mPrevWindowProcAddress <> 0 Then StopSubclassing Button.hWnd
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Call PropBag.WriteProperty("Appearance", Button.Appearance, 1)
Call PropBag.WriteProperty("BackColor", mBackColor, vbButtonFace)
Call PropBag.WriteProperty("Cancel", Button.Cancel, False)
Call PropBag.WriteProperty("Caption", Button.Caption, "")
Call PropBag.WriteProperty("CausesValidation", Button.CausesValidation, True)
Call PropBag.WriteProperty("DefaultBorderColor", mDefaultBorderColor, &HEEEECF)
Call PropBag.WriteProperty("DisabledBackColor", mDisabledBackColor, vbInactiveBorder)
Call PropBag.WriteProperty("DisabledFont", mDisabledFont, Nothing)
Call PropBag.WriteProperty("DisabledForeColor", mDisabledForeColor, vbBlack)
Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
Call PropBag.WriteProperty("FocusedBorderColor", mFocusedBorderColor, &HF8FFA8)
Call PropBag.WriteProperty("Font", Button.Font, Ambient.Font)
Call PropBag.WriteProperty("ForeColor", mForeColor, vbButtonText)
Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
Call PropBag.WriteProperty("MouseOverBackColor", mMouseOverBackColor, vbButtonFace)
Call PropBag.WriteProperty("MouseOverFont", mMouseOverFont, Nothing)
Call PropBag.WriteProperty("MouseOverForeColor", mMouseOverForeColor, vbBlack)
Call PropBag.WriteProperty("MousePointer", Button.MousePointer, 0)
Call PropBag.WriteProperty("NonDefaultBorderColor", mNonDefaultBorderColor, vbBlack)
Call PropBag.WriteProperty("OLEDropMode", Button.OLEDropMode, 0)
Call PropBag.WriteProperty("PushedBackColor", mPushedBackColor, vb3DHighlight)
Call PropBag.WriteProperty("PushedFont", mPushedFont, Nothing)
Call PropBag.WriteProperty("PushedForeColor", mPushedForeColor, vbBlack)
Call PropBag.WriteProperty("RightToLeft", Button.RightToLeft, False)
Call PropBag.WriteProperty("ToolTipText", Button.ToolTipText, "")
Call PropBag.WriteProperty("UseMaskColor", Button.UseMaskColor, False)
End Sub

'@================================================================================
' DeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal Data As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

If CLng(Data) = DeferredActions.DeferredActionFireMouseEnter Then
    RaiseEvent MouseEnter
ElseIf CLng(Data) = DeferredActions.DeferredActionFireMouseLeave Then
    RaiseEvent MouseLeave
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' ISubclassable Interface Members
'@================================================================================

Private Function ISubclassable_HandleWindowMessage(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const ProcName As String = "ISubclassable_HandleWindowMessage"
On Error GoTo Err

Dim lhDc As Long

If uMsg = WM_ERASEBKGND Then
    debugPrint True, "WM_ERASEBKGND"
    paintBackground hWnd, wParam, SendMessage(hWnd, BM_GETSTATE, 0, 0)
    ISubclassable_HandleWindowMessage = 1
    debugPrint False, "WM_ERASEBKGND"

ElseIf uMsg = WM_KILLFOCUS Then
    debugPrint True, "WM_KILLFOCUS"
    
    LockWindowUpdate hWnd
    ISubclassable_HandleWindowMessage = CallWindowProc(mPrevWindowProcAddress, hWnd, uMsg, wParam, lParam)
    LockWindowUpdate 0
    
    InvalidateRect hWnd, mClientRect, 0
    debugPrint False, "WM_KILLFOCUS"

ElseIf uMsg = WM_LBUTTONDOWN Then
    debugPrint True, "WM_LBUTTONDOWN"
    mMouseDown = True
    LockWindowUpdate hWnd
    ISubclassable_HandleWindowMessage = CallWindowProc(mPrevWindowProcAddress, hWnd, uMsg, wParam, lParam)
    LockWindowUpdate 0
    debugPrint False, "WM_LBUTTONDOWN"

ElseIf uMsg = WM_LBUTTONUP Then
    debugPrint True, "WM_LBUTTONUP"
    mMouseDown = False
    LockWindowUpdate hWnd
    ISubclassable_HandleWindowMessage = CallWindowProc(mPrevWindowProcAddress, hWnd, uMsg, wParam, lParam)
    LockWindowUpdate 0
    'releaseMouse
    debugPrint False, "WM_LBUTTONUP"

ElseIf uMsg = WM_MOUSELEAVE Then
    debugPrint True, "WM_MOUSELEAVE"
    mTrackingMouse = False
    mGotMouse = False
    DeferAction Me, DeferredActions.DeferredActionFireMouseLeave
    InvalidateRect hWnd, mClientRect, 1
    debugPrint False, "WM_MOUSELEAVE"

ElseIf uMsg = WM_MOUSEMOVE Then
    debugPrint True, "WM_MouseMove"
    If Not mTrackingMouse Then
        mTrackingMouse = True
        trackMouse hWnd
        'captureMouse hwnd
        DeferAction Me, DeferredActions.DeferredActionFireMouseEnter
    End If
    
    mInhibitClickEvent = True
    If isCursorInButton(hWnd) Then
        mInhibitClickEvent = False
        If Not mGotMouse Then
            mGotMouse = True
            On Error Resume Next
            Debug.Print "Cancel=" & UserControl.Extender.Cancel
            InvalidateRect hWnd, mClientRect, 1
        End If
    ElseIf mGotMouse Then
        mGotMouse = False
        InvalidateRect hWnd, mClientRect, 1
    End If

    debugPrint False, "WM_MouseMove"

ElseIf uMsg = WM_PAINT Then
    debugPrint True, "WM_PAINT"
    Dim lPaint As PAINTSTRUCT
    lhDc = BeginPaint(hWnd, lPaint)
    
    Dim lState As Long
    lState = SendMessage(hWnd, BM_GETSTATE, 0, 0)
    
    If lPaint.fErase Then paintBackground hWnd, lhDc, lState
    
    paintForegroundRect chooseForeColor(hWnd, lState), chooseFont(hWnd, lState), lhDc, mClientRect, lState
    EndPaint hWnd, lPaint
    ISubclassable_HandleWindowMessage = 0
    debugPrint False, "WM_PAINT"

ElseIf uMsg = WM_SETFOCUS Then
    debugPrint True, "WM_SETFOCUS"
    
    LockWindowUpdate hWnd
    ISubclassable_HandleWindowMessage = CallWindowProc(mPrevWindowProcAddress, hWnd, uMsg, wParam, lParam)
    LockWindowUpdate 0
    
    InvalidateRect hWnd, mClientRect, 0
    debugPrint False, "WM_SETFOCUS"
Else
    ISubclassable_HandleWindowMessage = CallWindowProc(mPrevWindowProcAddress, hWnd, uMsg, wParam, lParam)
End If


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get ISubclassable_hWnd() As Long
ISubclassable_hWnd = Button.hWnd
End Property

Private Property Let ISubclassable_PrevWindowProcAddress(ByVal RHS As Long)
mPrevWindowProcAddress = RHS
End Property

Private Property Get ISubclassable_PrevWindowProcAddress() As Long
ISubclassable_PrevWindowProcAddress = mPrevWindowProcAddress
End Property

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

Private Sub Button_Click()
debugPrint True, "Button_Click"
If Not mInhibitClickEvent Then RaiseEvent Click
debugPrint False, "Button_Click"
End Sub

Private Sub Button_KeyDown(KeyCode As Integer, Shift As Integer)
debugPrint True, "Button_KeyDown"
RaiseEvent KeyDown(KeyCode, Shift)
debugPrint False, "Button_KeyDown"
End Sub

Private Sub Button_KeyPress(KeyAscii As Integer)
debugPrint True, "Button_KeyPress"
RaiseEvent KeyPress(KeyAscii)
debugPrint False, "Button_KeyPress"
End Sub

Private Sub Button_KeyUp(KeyCode As Integer, Shift As Integer)
debugPrint True, "Button_KeyUp"
RaiseEvent KeyUp(KeyCode, Shift)
debugPrint False, "Button_KeyUp"
End Sub

Private Sub Button_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
debugPrint True, "Button_MouseDown"
RaiseEvent MouseDown(Button, Shift, X, Y)
debugPrint False, "Button_MouseDown"
End Sub

Private Sub Button_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Button_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
debugPrint True, "Button_MouseUp"
RaiseEvent MouseUp(Button, Shift, X, Y)
debugPrint False, "Button_MouseUp"
End Sub

Private Sub Button_OLECompleteDrag(Effect As Long)
RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub Button_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Button_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub Button_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub Button_OLESetData(Data As DataObject, DataFormat As Integer)
RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub Button_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

'@================================================================================
' mDisabledFont Event Handlers
'@================================================================================

Private Sub mDisabledFont_FontChanged(ByVal PropertyName As String)
If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 1
End Sub

'@================================================================================
' mFont Event Handlers
'@================================================================================

Private Sub mFont_FontChanged(ByVal PropertyName As String)
Set UserControl.Font = mFont
If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 1
End Sub

'@================================================================================
' mMouseOverFont Event Handlers
'@================================================================================

Private Sub mMouseOverFont_FontChanged(ByVal PropertyName As String)
If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 1
End Sub

'@================================================================================
' mPushedFont Event Handlers
'@================================================================================

Private Sub mPushedFont_FontChanged(ByVal PropertyName As String)
If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 1
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Appearance() As Integer
Attribute Appearance.VB_UserMemId = -520
Appearance = Button.Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
Button.Appearance() = New_Appearance
PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
BackColor = Button.BackColor
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
setBrushProperty "BackColor", mBackColor, Value, mhBackBrush
End Property



Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = -518
Caption = Button.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
Button.Caption() = New_Caption
PropertyChanged "Caption"
End Property



Public Property Get CausesValidation() As Boolean
CausesValidation = Button.CausesValidation
End Property

Public Property Let CausesValidation(ByVal New_CausesValidation As Boolean)
Button.CausesValidation() = New_CausesValidation
PropertyChanged "CausesValidation"
End Property

Public Property Get DefaultBorderColor() As OLE_COLOR
DefaultBorderColor = mDefaultBorderColor
End Property

Public Property Let DefaultBorderColor(ByVal Value As OLE_COLOR)
setPenProperty "DefaultBorderColor", mDefaultBorderColor, Value, mhDefaultBorderPen
End Property

Public Property Get DisabledBackColor() As OLE_COLOR
DisabledBackColor = mDisabledBackColor
End Property

Public Property Let DisabledBackColor(ByVal Value As OLE_COLOR)
setBrushProperty "DisabledBackColor", mDisabledBackColor, Value, mhDisabledBrush
End Property

Public Property Get DisabledFont() As Font
Set DisabledFont = mDisabledFont
End Property

Public Property Set DisabledFont(ByVal Value As Font)
Const ProcName As String = "DisabledFont"
On Error GoTo Err

If fontsEqual(Value, mDisabledFont) And mDisabledFontHandle <> 0 Then Exit Property
If mDisabledFontHandle <> 0 Then If DeleteObject(mDisabledFontHandle) = 0 Then HandleWin32Error
mDisabledFontHandle = 0

Set mDisabledFont = Value
If Not mDisabledFont Is Nothing Then
    With mDisabledFont
       .Bold = Value.Bold
       .Italic = Value.Italic
       .Name = Value.Name
       .Size = Value.Size
       .Strikethrough = Value.Strikethrough
       .Underline = Value.Underline
       .Weight = Value.Weight
    End With
    mDisabledFontHandle = createFontHandle(mDisabledFont)
End If

PropertyChanged "DisabledFont"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get DisabledForeColor() As OLE_COLOR
DisabledForeColor = mDisabledForeColor
End Property

Public Property Let DisabledForeColor(ByVal Value As OLE_COLOR)
mDisabledForeColor = Value

If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 0
End Property

'


Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
UserControl.Enabled() = New_Enabled
PropertyChanged "Enabled"
InvalidateRect Button.hWnd, mClientRect, 1
End Property

Public Property Get FocusedBorderColor() As OLE_COLOR
FocusedBorderColor = mFocusedBorderColor
End Property

Public Property Let FocusedBorderColor(ByVal Value As OLE_COLOR)
setPenProperty "FocusedBorderColor", mFocusedBorderColor, Value, mhFocusedBorderPen
End Property



Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
Set Font = mFont
End Property

Public Property Set Font(ByVal Value As Font)
Const ProcName As String = "Font"
On Error GoTo Err

AssertArgument Not Value Is Nothing
If fontsEqual(Value, mFont) And mFontHandle <> 0 Then Exit Property
If mFontHandle <> 0 Then If DeleteObject(mFontHandle) = 0 Then HandleWin32Error

Set mFont = Value
With mFont
   .Bold = Value.Bold
   .Italic = Value.Italic
   .Name = Value.Name
   .Size = Value.Size
   .Strikethrough = Value.Strikethrough
   .Underline = Value.Underline
   .Weight = Value.Weight
End With
mFontHandle = createFontHandle(mFont)
Set Button.Font = mFont
PropertyChanged "Font"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property


'MemberInfo=10,0,0,0
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
ForeColor = mForeColor
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
mForeColor = Value
PropertyChanged "ForeColor"
If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 0
End Property



Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
hWnd = UserControl.hWnd ' Button.hwnd
End Property



Public Property Get MouseIcon() As Picture
Set MouseIcon = Button.MouseIcon
End Property

Public Property Let MouseIcon(ByVal New_MouseIcon As Picture)
Set Button.MouseIcon = New_MouseIcon
PropertyChanged "MouseIcon"
End Property

Public Property Get MouseOverBackColor() As OLE_COLOR
MouseOverBackColor = mMouseOverBackColor
End Property

Public Property Let MouseOverBackColor(ByVal Value As OLE_COLOR)
setBrushProperty "MouseOverBackColor", mMouseOverBackColor, Value, mhMouseOverBrush
End Property

Public Property Get MouseOverFont() As Font
Set MouseOverFont = mMouseOverFont
End Property

Public Property Set MouseOverFont(ByVal Value As Font)
Const ProcName As String = "MouseOverFont"
On Error GoTo Err

If fontsEqual(Value, mMouseOverFont) And mMouseOverFontHandle <> 0 Then Exit Property
If mMouseOverFontHandle <> 0 Then If DeleteObject(mMouseOverFontHandle) = 0 Then HandleWin32Error
mMouseOverFontHandle = 0

Set mMouseOverFont = Value
If Not mMouseOverFont Is Nothing Then
    With mMouseOverFont
       .Bold = Value.Bold
       .Italic = Value.Italic
       .Name = Value.Name
       .Size = Value.Size
       .Strikethrough = Value.Strikethrough
       .Underline = Value.Underline
       .Weight = Value.Weight
    End With
    mMouseOverFontHandle = createFontHandle(mMouseOverFont)
End If

PropertyChanged "MouseOverFont"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MouseOverForeColor() As OLE_COLOR
MouseOverForeColor = mMouseOverForeColor
End Property

Public Property Let MouseOverForeColor(ByVal Value As OLE_COLOR)
mMouseOverForeColor = Value

If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 0
End Property



Public Property Get MousePointer() As Integer
MousePointer = Button.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
Button.MousePointer() = New_MousePointer
PropertyChanged "MousePointer"
End Property

Public Property Get NonDefaultBorderColor() As OLE_COLOR
NonDefaultBorderColor = mNonDefaultBorderColor
End Property

Public Property Let NonDefaultBorderColor(ByVal Value As OLE_COLOR)
setPenProperty "NonDefaultBorderColor", mNonDefaultBorderColor, Value, mhNonDefaultBorderPen
End Property



Public Property Get OLEDropMode() As Integer
OLEDropMode = Button.OLEDropMode
End Property

Public Property Let OLEDropMode(ByVal New_OLEDropMode As Integer)
Button.OLEDropMode() = New_OLEDropMode
PropertyChanged "OLEDropMode"
End Property



Public Sub OLEDrag()
Button.OLEDrag
End Sub

Public Property Get PushedBackColor() As OLE_COLOR
PushedBackColor = mPushedBackColor
End Property

Public Property Let PushedBackColor(ByVal Value As OLE_COLOR)
setBrushProperty "PushedBackColor", mPushedBackColor, Value, mhPushedBrush
End Property

Public Property Get PushedFont() As Font
Set PushedFont = mPushedFont
End Property

Public Property Set PushedFont(ByVal Value As Font)
Const ProcName As String = "PushedFont"
On Error GoTo Err

If fontsEqual(Value, mPushedFont) And mPushedFontHandle <> 0 Then Exit Property
If mPushedFontHandle <> 0 Then If DeleteObject(mPushedFontHandle) = 0 Then HandleWin32Error
mPushedFontHandle = 0

Set mPushedFont = Value
If Not mPushedFont Is Nothing Then
    With mPushedFont
       .Bold = Value.Bold
       .Italic = Value.Italic
       .Name = Value.Name
       .Size = Value.Size
       .Strikethrough = Value.Strikethrough
       .Underline = Value.Underline
       .Weight = Value.Weight
    End With
    mPushedFontHandle = createFontHandle(mPushedFont)
End If

PropertyChanged "PushedFont"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PushedForeColor() As OLE_COLOR
PushedForeColor = mPushedForeColor
End Property

Public Property Let PushedForeColor(ByVal Value As OLE_COLOR)
mPushedForeColor = Value
PropertyChanged "PushedForeColor"

If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 0
End Property



Public Sub Refresh()
Attribute Refresh.VB_UserMemId = -550
Button.Refresh
End Sub


'MemberInfo=0,0,0,False
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_UserMemId = -611
RightToLeft = Button.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
Button.RightToLeft = New_RightToLeft
PropertyChanged "RightToLeft"
End Property



Public Property Get Style() As Integer
Style = Button.Style
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

mNoDraw = True

BackColor = mTheme.ButtonBackColor
Set Font = mTheme.BaseFont
DefaultBorderColor = mTheme.DefaultBorderColor

DisabledBackColor = mTheme.DisabledBackColor
Set DisabledFont = mTheme.DisabledFont
DisabledForeColor = mTheme.DisabledForeColor

FocusedBorderColor = mTheme.FocusBorderColor
If Not mTheme.ButtonFont Is Nothing Then Set Font = mTheme.ButtonFont
ForeColor = mTheme.ButtonForeColor

MouseOverBackColor = mTheme.MouseOverBackColor
Set MouseOverFont = mTheme.MouseOverFont
MouseOverForeColor = mTheme.MouseOverForeColor

NonDefaultBorderColor = mTheme.NonDefaultBorderColor

PushedBackColor = mTheme.PushedBackColor
Set PushedFont = mTheme.PushedFont
PushedForeColor = mTheme.PushedForeColor

mNoDraw = False
InvalidateRect Button.hWnd, mClientRect, 1

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



Public Property Get ToolTipText() As String
ToolTipText = Button.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
Button.ToolTipText() = New_ToolTipText
PropertyChanged "ToolTipText"
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub captureMouse(ByVal pHwnd As Long)
Const ProcName As String = "captureMouse"
On Error GoTo Err

debugPrint True, "Capture mouse"
SetCapture pHwnd
debugPrint False, "Capture mouse"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function chooseFont(ByVal pHwnd As Long, ByVal pButtonState As Long) As Long
Const ProcName As String = "chooseFont"
On Error GoTo Err

If Not UserControl.Enabled Then
    chooseFont = IIf(mDisabledFontHandle = 0, mFontHandle, mDisabledFontHandle)
ElseIf (pButtonState And BST_PUSHED) And isCursorInButton(pHwnd) Then
    chooseFont = IIf(mPushedFontHandle = 0, mFontHandle, mPushedFontHandle)
ElseIf mGotMouse Then
    chooseFont = IIf(mMouseOverFontHandle = 0, mFontHandle, mMouseOverFontHandle)
Else
    chooseFont = mFontHandle
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function chooseForeColor(ByVal pHwnd As Long, ByVal pButtonState As Long) As Long
Const ProcName As String = "chooseForeColor"
On Error GoTo Err

If Not UserControl.Enabled Then
    chooseForeColor = IIf(mDisabledForeColor = 0, mForeColor, mDisabledForeColor)
ElseIf (pButtonState And BST_PUSHED) And isCursorInButton(pHwnd) Then
    chooseForeColor = IIf(mPushedForeColor = 0, mForeColor, mPushedForeColor)
ElseIf mGotMouse Then
    chooseForeColor = IIf(mMouseOverForeColor = 0, mForeColor, mMouseOverForeColor)
Else
    chooseForeColor = mForeColor
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createFontHandle(ByVal pFont As StdFont) As Long
Const ProcName As String = "createFontHandle"
On Error GoTo Err

createFontHandle = CreateFont(pFont.Size * TwipsPerPoint / Screen.TwipsPerPixelY, _
                            0, _
                            0, _
                            0, _
                            IIf(pFont.Bold, FW_BOLD, FW_NORMAL), _
                            IIf(pFont.Italic, 1, 0), _
                            IIf(pFont.Underline, 1, 0), _
                            IIf(pFont.Strikethrough, 1, 0), _
                            DEFAULT_CHARSET, _
                            OUT_DEFAULT_PRECIS, _
                            CLIP_DEFAULT_PRECIS, _
                            DEFAULT_QUALITY, _
                            DEFAULT_PITCH Or FF_DONTCARE, _
                            StrPtr(pFont.Name))
If createFontHandle = 0 Then HandleWin32Error


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createSolidPenAndInvalidate(ByVal pPrevPen As Long, ByVal pColor As Long) As Long
If pPrevPen <> 0 Then DeleteObject pPrevPen
createSolidPenAndInvalidate = CreatePen(PS_SOLID, 1, NormalizeColor(pColor))
If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 1
End Function

Private Sub debugPrint(ByVal pEnter As Boolean, pMsg As String)
Static sIndent As Long
If Not pEnter Then sIndent = sIndent - 1
Debug.Print String(sIndent * 4, " ") & IIf(pEnter, "Enter: ", "Exit: ") & pMsg
If pEnter Then sIndent = sIndent + 1
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
If pFont1.Size <> pFont2.Size Then Exit Function
If pFont1.Strikethrough <> pFont2.Strikethrough Then Exit Function
If pFont1.Underline <> pFont2.Underline Then Exit Function
If pFont1.Weight <> pFont2.Weight Then Exit Function

fontsEqual = True
End Function

Private Function isCursorInButton(ByVal pHwnd As Long) As Boolean
Const ProcName As String = "isCursorInButton"
On Error GoTo Err

Dim lPt As GDI_POINT
If GetCursorPos(lPt) = 0 Then Exit Function
If ScreenToClient(pHwnd, lPt) = 0 Then Exit Function

isCursorInButton = (PtInRect(mClientRect, lPt.X, lPt.Y) <> 0)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub paintBackground(ByVal pHwnd As Long, ByVal phdc As Long, ByVal pButtonState As Long)
Const ProcName As String = "paintBackground"
On Error GoTo Err

debugPrint True, "Paint background: DisplayAsDefault=" & CStr(UserControl.Ambient.DisplayAsDefault)

If Not UserControl.Enabled Then
    SelectObject phdc, GetStockObject(DC_PEN)
    SetDCPenColor phdc, mDisabledBackColor
ElseIf (pButtonState And BST_FOCUS) Then
    SelectObject phdc, mhFocusedBorderPen
ElseIf UserControl.Ambient.DisplayAsDefault Then
    SelectObject phdc, mhDefaultBorderPen
Else
    SelectObject phdc, mhNonDefaultBorderPen
End If

If Not UserControl.Enabled Then
    SelectObject phdc, mhDisabledBrush
ElseIf (pButtonState And BST_PUSHED) And isCursorInButton(pHwnd) Then
    SelectObject phdc, mhPushedBrush
ElseIf mGotMouse Then
    SelectObject phdc, mhMouseOverBrush
Else
    SelectObject phdc, mhBackBrush
End If

If RoundRect(phdc, mClientRect.Left, mClientRect.Top, mClientRect.Right - 1, mClientRect.Bottom - 1, RoundedRectWidth, RoundedRectHeight) = 0 Then HandleWin32Error
debugPrint False, "Paint background"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub paintForegroundRect(ByVal pForeColor As Long, ByVal pFontHandle As Long, ByVal phdc As Long, ByRef pClientRect As GDI_RECT, ByVal pButtonState As Long)
Const ProcName As String = "paintForegroundRect"
On Error GoTo Err

debugPrint True, "Paint foreground"

If Len(Button.Caption) <> 0 Then
    Dim lPrevFontHandle As Long
    lPrevFontHandle = setFont(phdc, pFontHandle)
    
    Dim lClipRect As GDI_RECT
    lClipRect = pClientRect
    InflateRect lClipRect, -2, -2
    
    Dim lOriginalHeight As Long
    lOriginalHeight = lClipRect.Bottom - lClipRect.Top
    Dim lOriginalWIdth As Long
    lOriginalWIdth = lClipRect.Right - lClipRect.Left
    
    If SetTextColor(phdc, pForeColor) = CLR_INVALID Then HandleWin32Error
    SetBkMode phdc, TRANSPARENT
    
    SetTextAlign phdc, TA_LEFT + TA_TOP + TA_NOUPDATECP
    
    Dim lTextRect As GDI_RECT
    lTextRect = lClipRect
    If DrawText(phdc, _
                StrPtr(Button.Caption), _
                Len(Button.Caption), _
                lTextRect, _
                DT_CALCRECT + DT_CENTER + DT_WORDBREAK) = 0 Then HandleWin32Error
    
    Dim lAdjustPositionY As Long
    lAdjustPositionY = (lOriginalHeight - (lTextRect.Bottom - lTextRect.Top)) / 2
    lTextRect.Top = lTextRect.Top + lAdjustPositionY
    lTextRect.Bottom = lTextRect.Bottom + lAdjustPositionY
    
    Dim lAdjustPositionX As Long
    lAdjustPositionX = (lOriginalWIdth - (lTextRect.Right - lTextRect.Left)) / 2
    lTextRect.Left = lTextRect.Left + lAdjustPositionX
    lTextRect.Right = lTextRect.Right + lAdjustPositionX
    
    Dim lRegionHandle As Long
    lRegionHandle = CreateRectRgn(lClipRect.Left, _
                            lClipRect.Top, _
                            lClipRect.Right, _
                            lClipRect.Bottom)
    If lRegionHandle = 0 Then HandleWin32Error
    
    If SelectClipRgn(phdc, lRegionHandle) = 0 Then HandleWin32Error
    
    If DeleteObject(lRegionHandle) = 0 Then HandleWin32Error
    
    If User32.DrawText(phdc, _
                        StrPtr(Button.Caption), _
                        Len(Button.Caption), _
                        lTextRect, _
                        DT_CENTER + DT_WORDBREAK) = 0 Then HandleWin32Error
    
    releaseFont phdc, lPrevFontHandle
End If

'If (pButtonState And BST_FOCUS) Then
'    SelectClipRgn phdc, 0
'    Dim lFocusRect As GDI_RECT
'    lFocusRect = pClientRect
'    InflateRect lFocusRect, -4, -4
'    DrawFocusRect phdc, lFocusRect
'End If

debugPrint False, "Paint foreground"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub releaseFont(ByVal phdc As Long, ByVal pPrevFontHandle As Long)
Const ProcName As String = "releaseFont"
On Error GoTo Err

If pPrevFontHandle = 0 Then Exit Sub

If SelectObject(phdc, pPrevFontHandle) = 0 Then HandleWin32Error

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub releaseMouse()
Const ProcName As String = "releaseMouse"
On Error GoTo Err

debugPrint True, "Release mouse"
ReleaseCapture
debugPrint False, "Release mouse"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub resize()
Const ProcName As String = "resize"
On Error GoTo Err

Button.Move 0, 0, UserControl.Width, UserControl.Height

If GetClientRect(UserControl.hWnd, mClientRect) = 0 Then HandleWin32Error

Dim lWindowRect As GDI_RECT
If GetWindowRect(UserControl.hWnd, lWindowRect) = 0 Then HandleWin32Error

Dim lRgn As Long
lRgn = CreateRoundRectRgn(0, 0, lWindowRect.Right - lWindowRect.Left, lWindowRect.Bottom - lWindowRect.Top, RoundedRectWidth, RoundedRectHeight)
If lRgn = 0 Then HandleWin32Error

If SetWindowRgn(UserControl.hWnd, lRgn, 1) = 0 Then HandleWin32Error

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setBrushProperty( _
                ByVal pPropertyName As String, _
                ByRef pCurrColor As Long, _
                ByRef pNewColor As Long, _
                ByRef pBrushHandle As Long)
If pNewColor = pCurrColor And pBrushHandle <> 0 Then Exit Sub
pCurrColor = pNewColor
PropertyChanged pPropertyName

If pBrushHandle <> 0 Then DeleteObject pBrushHandle
pBrushHandle = CreateSolidBrush(NormalizeColor(pNewColor))

If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 1
End Sub

Private Function setFont(ByVal phdc As Long, ByVal pFontHandle As Long) As Long
Const ProcName As String = "setFont"
On Error GoTo Err

Assert pFontHandle <> 0, "Font handle not set"

setFont = SelectObject(phdc, pFontHandle)
If setFont = 0 Then HandleWin32Error

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setPenProperty( _
                ByVal pPropertyName As String, _
                ByRef pCurrColor As Long, _
                ByRef pNewColor As Long, _
                ByRef pPenHandle As Long)
If pNewColor = pCurrColor And pPenHandle <> 0 Then Exit Sub
pCurrColor = pNewColor
If pPenHandle <> 0 Then DeleteObject pPenHandle
pPenHandle = CreatePen(PS_SOLID, 1, NormalizeColor(pNewColor))
PropertyChanged pPropertyName

If Not mNoDraw Then InvalidateRect Button.hWnd, mClientRect, 0
End Sub

Private Sub trackMouse(ByVal pHwnd As Long, Optional ByVal pCancel As Boolean)
Const ProcName As String = "trackMouse"
On Error GoTo Err

Dim tm As TRACKMOUSEEVENTSTRUCT
tm.cbSize = Len(tm)
tm.hwndTrack = pHwnd
tm.dwFlags = TME_LEAVE
If pCancel Then tm.dwFlags = tm.dwFlags Or TME_CANCEL
If TrackMouseEvent(tm) = 0 Then HandleWin32Error

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
