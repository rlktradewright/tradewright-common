VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.UserControl TWModelessMsgBox 
   BackColor       =   &H80000005&
   ClientHeight    =   1260
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4875
   DefaultCancel   =   -1  'True
   ScaleHeight     =   1260
   ScaleWidth      =   4875
   Begin TWControls40.TWButton Button3 
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWButton Button2 
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin TWControls40.TWButton Button1 
      Height          =   315
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TWModelessMsgBox.ctx":0000
            Key             =   "Critical"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TWModelessMsgBox.ctx":0452
            Key             =   "Question"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TWModelessMsgBox.ctx":08A4
            Key             =   "Exclamation"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "TWModelessMsgBox.ctx":0CF6
            Key             =   "Information"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox IconPicture 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "TWModelessMsgBox"
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

Implements IThemeable

'@================================================================================
' Events
'@================================================================================

Event Result(ByVal Value As MsgBoxResults)
                
'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "TWModelessMsgBox"

'@================================================================================
' Member variables
'@================================================================================

Private mNonClientMetrics As NONCLIENTMETRICSW
Private mFILLER(999) As Long
Private mCancellable As Boolean
Private mIsToolWindow As Boolean

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Initialize()
Const ProcName As String = "UserControl_Initialize"
On Error GoTo Err

mNonClientMetrics.cbSize = Len(mNonClientMetrics)
If SystemParametersInfo(SPI_GETNONCLIENTMETRICS, Len(mNonClientMetrics), VarPtr(mNonClientMetrics), 0) = 0 Then
    ' Windows version prior to 0x0600 don't have the last field in NONCLIENTMETRICSW
    mNonClientMetrics.cbSize = Len(mNonClientMetrics) - 4
    If SystemParametersInfo(SPI_GETNONCLIENTMETRICS, Len(mNonClientMetrics) - 4, VarPtr(mNonClientMetrics), 0) = 0 Then HandleWin32Error
End If

mIsToolWindow = GetWindowLong(hWnd, GWL_EXSTYLE) And WS_EX_TOOLWINDOW

setFont

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub UserControl_Show()
Const ProcName As String = "UserControl_Show"
On Error GoTo Err

If Not UserControl.Ambient.UserMode Then
    Initialise "This is the modeless message box control, which may be " & vbCrLf & _
                "used to display a message to the user in a context where " & vbCrLf & _
                "the MsgBox function would be inappropriate", _
                vbOKOnly + vbInformation

End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
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

Private Sub Button1_Click()
Const ProcName As String = "Button1_Click"
On Error GoTo Err

processClick Button1.Caption

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Button2_Click()
Const ProcName As String = "Button2_Click"
On Error GoTo Err

processClick Button2.Caption

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Button3_Click()
Const ProcName As String = "Button3_Click"
On Error GoTo Err

processClick Button3.Caption

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
If mTheme Is Nothing Then Exit Property

UserControl.BackColor = mTheme.TextBackColor
gApplyTheme mTheme, UserControl.Controls

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Initialise( _
                ByVal prompt As String, _
                ByVal buttons As MsgBoxStyles, _
                Optional ByVal title As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

Dim failPoint As String
failPoint = 50

Clear

failPoint = 100

UserControl.BackColor = UserControl.Ambient.BackColor
Text.BackColor = UserControl.Ambient.BackColor
IconPicture.BackColor = UserControl.Ambient.BackColor

failPoint = 200

setPrompt prompt

failPoint = 300

Dim numbuttons As Long
numbuttons = setButtons(buttons)

failPoint = 400

setIcons buttons

failPoint = 500

setOptions buttons

failPoint = 600

sizeControl numbuttons, title

failPoint = 700

If UserControl.Ambient.UserMode Then sizeContainer

failPoint = 800

If UserControl.Ambient.UserMode Then setWindowCaption title

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, pFailpoint:=failPoint
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

IconPicture.Picture = Nothing
IconPicture.Visible = False
Button1.Default = True
Button1.Cancel = False
Button2.Cancel = False
Button3.Cancel = False
Text.Alignment = vbLeftJustify
mCancellable = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function createMessageFont() As Long
Const ProcName As String = "createMessageFont"
On Error GoTo Err

Dim lf As LOGFONTW
lf = mNonClientMetrics.lfMessageFont
createMessageFont = CreateFont(lf.lfHeight, _
                            lf.lfWidth, _
                            lf.lfEscapement, _
                            lf.lfOrientation, _
                            lf.lfWeight, _
                            lf.lfItalic, _
                            lf.lfUnderline, _
                            lf.lfStrikeOut, _
                            lf.lfCharSet, _
                            lf.lfOutPrecision, _
                            lf.lfClipPrecision, _
                            lf.lfQuality, _
                            lf.lfPitchAndFamily, _
                            StrPtr(StrConv(lf.lfFaceName, vbUnicode)))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createTitleFont() As Long
Const ProcName As String = "createTitleFont"
On Error GoTo Err

Dim lf As LOGFONTW
If mIsToolWindow Then
    lf = mNonClientMetrics.lfSmCaptionFont
Else
    lf = mNonClientMetrics.lfCaptionFont
End If

createTitleFont = CreateFont(lf.lfHeight, _
                            lf.lfWidth, _
                            lf.lfEscapement, _
                            lf.lfOrientation, _
                            lf.lfWeight, _
                            lf.lfItalic, _
                            lf.lfUnderline, _
                            lf.lfStrikeOut, _
                            lf.lfCharSet, _
                            lf.lfOutPrecision, _
                            lf.lfClipPrecision, _
                            lf.lfQuality, _
                            lf.lfPitchAndFamily, _
                            StrPtr(StrConv(lf.lfFaceName, vbUnicode)))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
                                        
End Function

Private Sub getTextWidthAndHeight( _
                ByVal prompt As String, _
                ByRef textWidth As Long, _
                ByRef textHeight As Long)
Const ProcName As String = "getTextWidthAndHeight"
On Error GoTo Err

Dim dc As Long
dc = GetDC(Text.hWnd)

Dim newFont As Long
newFont = createMessageFont

Dim prevFont As Long
prevFont = SelectObject(dc, newFont)

Dim tm As TEXTMETRICW
GetTextMetrics dc, tm

Dim lines() As String
lines = Split(prompt, vbCrLf)

Dim i As Long
For i = 0 To UBound(lines)
    Dim line As String
    line = lines(i)
    
    Dim textDimensions As Long
    textDimensions = GetTabbedTextExtent(dc, _
                                        StrPtr(line), _
                                        Len(line), _
                                        0, _
                                        0)
                                        
    Dim lineWidth As Long
    lineWidth = textDimensions And &HFFFF&
    If lineWidth > textWidth Then textWidth = lineWidth
Next

SelectObject dc, prevFont
DeleteObject newFont

textWidth = textWidth * Screen.TwipsPerPixelX

textHeight = Screen.TwipsPerPixelY * ((UBound(lines) + 1) * tm.tmHeight)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getTitleBarWidth( _
                ByVal title As String) As Long
Const ProcName As String = "getTitleBarWidth"
On Error GoTo Err

Dim hWnd As Long
hWnd = UserControl.ContainerHwnd

Dim dc As Long
dc = getDCEx(hWnd, 0, DCX_WINDOW)

Dim newFont As Long
newFont = createTitleFont

Dim prevFont As Long
prevFont = SelectObject(dc, newFont)
                                        
Dim titleSize As GDI_SIZE
GetTextExtentPoint32 dc, StrPtr(title), Len(title), titleSize

SelectObject dc, prevFont
DeleteObject newFont

getTitleBarWidth = titleSize.cx

If mCancellable Then
    ' the title bar will have a control menu icon and a close button
   getTitleBarWidth = getTitleBarWidth + 2 * (IIf(mIsToolWindow, GetSystemMetrics(SM_CXSMSIZE), GetSystemMetrics(SM_CXSIZE))) + 2
End If
getTitleBarWidth = getTitleBarWidth * Screen.TwipsPerPixelX

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub processClick( _
                ByVal Value As String)
Const ProcName As String = "processClick"
On Error GoTo Err

If Value = "&Abort" Then
    RaiseEvent Result(MsgBoxAbort)
ElseIf Value = "Cancel" Then
    RaiseEvent Result(MsgBoxCancel)
ElseIf Value = "&Ignore" Then
    RaiseEvent Result(MsgBoxIgnore)
ElseIf Value = "&No" Then
    RaiseEvent Result(MsgBoxNo)
ElseIf Value = "OK" Then
    RaiseEvent Result(MsgBoxOK)
ElseIf Value = "&Retry" Then
    RaiseEvent Result(MsgBoxRetry)
ElseIf Value = "&Yes" Then
    RaiseEvent Result(MsgBoxYes)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setButtons( _
                ByVal buttons As MsgBoxStyles) As Long
Const ProcName As String = "setButtons"
On Error GoTo Err

Button1.Caption = "OK"
Button1.Visible = True
Button1.Cancel = True
mCancellable = True

Dim numbuttons As Long
numbuttons = 1

If (buttons And MsgBoxOKCancel) = MsgBoxOKCancel Then
    Button1.Caption = "OK"
    Button1.Visible = True
    Button2.Caption = "Cancel"
    Button2.Visible = True
    Button2.Cancel = True
    mCancellable = True
    Button3.Visible = False
    numbuttons = 2
End If
If (buttons And MsgBoxAbortRetryIgnore) = MsgBoxAbortRetryIgnore Then
    Button1.Caption = "&Abort"
    Button1.Cancel = False
    mCancellable = False
    Button1.Visible = True
    Button2.Caption = "&Retry"
    Button2.Visible = True
    Button3.Caption = "&Ignore"
    Button3.Visible = True
    numbuttons = 3
End If
If (buttons And MsgBoxYesNoCancel) = MsgBoxYesNoCancel Then
    Button1.Caption = "&Yes"
    Button1.Visible = True
    Button2.Caption = "&No"
    Button2.Visible = True
    Button3.Caption = "Cancel"
    Button3.Visible = True
    Button3.Cancel = True
    mCancellable = True
    numbuttons = 3
End If
If (buttons And MsgBoxYesNo) = MsgBoxYesNo Then
    Button1.Caption = "&Yes"
    Button1.Visible = True
    Button1.Cancel = False
    mCancellable = False
    Button2.Caption = "&No"
    Button2.Visible = True
    Button3.Visible = False
    numbuttons = 2
End If
If (buttons And MsgBoxRetryCancel) = MsgBoxRetryCancel Then
    Button1.Caption = "&Retry"
    Button1.Visible = True
    Button2.Caption = "Cancel"
    Button2.Visible = True
    Button2.Cancel = True
    mCancellable = True
    Button3.Visible = False
    numbuttons = 2
End If
If (buttons And MsgBoxDefaultButton1) = MsgBoxDefaultButton1 Then
    If numbuttons >= 1 Then Button1.Default = True
End If
If (buttons And MsgBoxDefaultButton2) = MsgBoxDefaultButton2 Then
    If numbuttons >= 2 Then Button2.Default = True
End If
If (buttons And MsgBoxDefaultButton3) = MsgBoxDefaultButton3 Then
    If numbuttons >= 3 Then Button3.Default = True
End If
If (buttons And MsgBoxDefaultButton4) = MsgBoxDefaultButton4 Then

End If

setButtons = numbuttons

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setFont()
Const ProcName As String = "setFont"
On Error GoTo Err

Dim lFont As New StdFont
With mNonClientMetrics.lfMessageFont
    lFont.Bold = .lfWeight >= FW_BOLD
    lFont.Charset = .lfCharSet
    lFont.Italic = .lfItalic
    lFont.Name = StrConv(.lfFaceName, vbUnicode)
    lFont.Size = (-72 * .lfHeight) / (1440 / Screen.TwipsPerPixelY)
    lFont.Strikethrough = .lfStrikeOut
    lFont.Underline = .lfUnderline
    lFont.Weight = .lfWeight
End With

Set Text.Font = lFont
Set UserControl.Font = lFont

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setIcons( _
                ByVal buttons As MsgBoxStyles)
Const ProcName As String = "setIcons"
On Error GoTo Err

If (buttons And MsgBoxCritical) = MsgBoxCritical Then
    IconPicture.Picture = ImageList1.ListImages.Item("Critical").Picture
End If
If (buttons And MsgBoxQuestion) = MsgBoxQuestion Then
    IconPicture.Picture = ImageList1.ListImages.Item("Question").Picture
End If
If (buttons And MsgBoxExclamation) = MsgBoxExclamation Then
    IconPicture.Picture = ImageList1.ListImages.Item("Exclamation").Picture
End If
If (buttons And MsgBoxInformation) = MsgBoxInformation Then
    IconPicture.Picture = ImageList1.ListImages.Item("Information").Picture
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setOptions( _
                ByVal buttons As MsgBoxStyles)
Const ProcName As String = "setOptions"
On Error GoTo Err

If (buttons And MsgBoxApplicationModal) = MsgBoxApplicationModal Then

End If
If (buttons And MsgBoxSystemModal) = MsgBoxSystemModal Then

End If
If (buttons And MsgBoxMsgBoxHelpButton) = MsgBoxMsgBoxHelpButton Then

End If
If (buttons And MsgBoxMsgBoxRight) = MsgBoxMsgBoxRight Then
    Text.Alignment = vbRightJustify
End If
If (buttons And MsgBoxMsgBoxRtlReading) = MsgBoxMsgBoxRtlReading Then

End If
If (buttons And MsgBoxMsgBoxSetForeground) = MsgBoxMsgBoxSetForeground Then

End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setPrompt( _
                ByVal prompt As String)
Const ProcName As String = "setPrompt"
On Error GoTo Err

If prompt = "" Then
    Text.Width = 3375
    Text.Height = 615
Else
    Dim textWidth As Long
    Dim textHeight As Long
    getTextWidthAndHeight prompt, textWidth, textHeight
    
    textWidth = textWidth + 2 * Screen.TwipsPerPixelX
    If textWidth > 0.75 * Screen.Width Then
        Text.Width = 0.75 * Screen.Width
    Else
        Text.Width = textWidth
    End If

    If textHeight > 0.75 * Screen.Height Then
        Text.Height = 0.75 * Screen.Height
    Else
        Text.Height = textHeight
    End If
End If

Text.Text = prompt

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setWindowCaption( _
                ByVal title As String)
Const ProcName As String = "setWindowCaption"
On Error GoTo Err

SetWindowText UserControl.ContainerHwnd, StrPtr(title)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub sizeContainer()
Const ProcName As String = "sizeContainer"
On Error GoTo Err

Dim hWnd As Long
hWnd = UserControl.ContainerHwnd

Dim frameThicknessX As Long
frameThicknessX = GetSystemMetrics(SM_CXDLGFRAME)

Dim frameThicknessY As Long
frameThicknessY = GetSystemMetrics(SM_CYDLGFRAME)

Dim captionHeight As Long
If mIsToolWindow Then
    captionHeight = GetSystemMetrics(SM_CYSMCAPTION)
Else
    captionHeight = GetSystemMetrics(SM_CYCAPTION)
End If

Dim myWidth As Long
myWidth = UserControl.Width / Screen.TwipsPerPixelX

Dim myHeight As Long
myHeight = UserControl.Height / Screen.TwipsPerPixelY

Dim windowWidth As Long
windowWidth = myWidth + 2 * frameThicknessX

Dim windowHeight As Long
windowHeight = myHeight + 2 * frameThicknessY + captionHeight

If Not mCancellable Then
    SetWindowLong hWnd, GWL_STYLE, GetWindowLong(hWnd, GWL_STYLE) And Not WS_SYSMENU
End If

SetWindowPos hWnd, _
            HWND_NOTOPMOST, _
            (Screen.Width / Screen.TwipsPerPixelX - windowWidth) / 2, _
            (Screen.Height / Screen.TwipsPerPixelY - windowHeight) / 2, _
            windowWidth, _
            windowHeight, _
            SWP_SHOWWINDOW Or IIf(Not mCancellable, SWP_FRAMECHANGED, 0)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub sizeControl( _
                ByVal numbuttons As Long, _
                ByVal title As String)
Const ProcName As String = "sizeControl"
On Error GoTo Err

If IconPicture.Picture.Handle = 0 Then
    Text.Left = IconPicture.Left
Else
    IconPicture.Visible = True
    Text.Left = IconPicture.Left + IconPicture.Width + 120
End If

Dim Width As Long
Width = Text.Left + Text.Width + 120

Dim Height As Single
Height = Text.Top + Text.Height + 120
If Height < IconPicture.Top + IconPicture.Height + 120 Then Height = IconPicture.Top + IconPicture.Height + 120
Button1.Top = Height
Button2.Top = Height
Button3.Top = Height
Height = Height + 240 + Button1.Height

Dim buttonsWidth As Long
buttonsWidth = numbuttons * Button1.Width + (numbuttons - 1) * 105

If Width < buttonsWidth + 240 Then Width = buttonsWidth + 240

If UserControl.Ambient.UserMode Then
    Dim titleBarWidth As Long
    titleBarWidth = getTitleBarWidth(title)
    If titleBarWidth > Width Then Width = titleBarWidth
End If

Button1.Left = (Width - buttonsWidth) / 2
Button2.Left = Button1.Left + 1200
Button3.Left = Button1.Left + 2400

If UserControl.Width <> Width Then UserControl.Width = Width
If UserControl.Height <> Height Then UserControl.Height = Height

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

