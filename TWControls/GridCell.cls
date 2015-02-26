VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GridCell"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
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

Private Enum CellFlags
    CellFlagAlign = 1
    CellFlagBold = 2
    CellFlagItalic = 4
    CellFlagStrikethrough = &H10&
    CellFlagUnderline = &H20&
    CellFlagInvertColors = &H40&
    CellFlagAlignmentChanged = &H80&
    CellFlagColorChanged = &H100&
    CellFlagFontChanged = &H200&
    CellFlagValueChanged = &H400&
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GridCell"

Private Const ResizeZoneSizePixels                  As Long = 3
'@================================================================================
' Member variables
'@================================================================================

Private mGrid                                       As TWGrid

Private mRow                                        As Long
Private mCol                                        As Long

Private mFlags                                      As CellFlags
Private mForeColor                                  As Long
Private mBackColor                                  As Long

Private mFontName                                   As String
Private mFontSize                                   As Currency ' NB this is the correct type for a StdFont's size

Private mValue                                      As String
Private mAlign                                      As AlignmentSettings

Private mCellLabelIndex                             As Long
Private mBorderLabelsIndex                          As Long

Private mLeft                                       As Long
Private mTop                                        As Long
Private mHeight                                     As Long
Private mWidth                                      As Long

Private mIsMapped                                   As Boolean
Private mIsFixed                                    As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mCellLabelIndex = -1
mBorderLabelsIndex = -1
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Align( _
                ByVal Value As AlignmentSettings)
Const ProcName As String = "Align"
On Error GoTo Err

mAlign = Value
setFlag CellFlagAlign Or CellFlagAlignmentChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Align() As AlignmentSettings
Const ProcName As String = "Align"
On Error GoTo Err

Align = mAlign

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let BackColor( _
                ByVal Value As Long)
Const ProcName As String = "BackColor"
On Error GoTo Err

mBackColor = Value
setFlag CellFlagColorChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get BackColor() As Long
Const ProcName As String = "BackColor"
On Error GoTo Err

BackColor = mBackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Bold( _
                ByVal Value As Boolean)
Const ProcName As String = "Bold"
On Error GoTo Err

If Value Then
    setFlag CellFlagBold Or CellFlagFontChanged
Else
    clearFlag CellFlagBold
    setFlag CellFlagFontChanged
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Bold() As Boolean
Const ProcName As String = "Bold"
On Error GoTo Err

Bold = isFlagSet(CellFlagBold)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Col( _
                ByVal Value As Long)
Const ProcName As String = "Col"
On Error GoTo Err

mCol = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Col() As Long
Const ProcName As String = "Col"
On Error GoTo Err

Col = mCol

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Font() As StdFont
Const ProcName As String = "Font"
On Error GoTo Err

Set Font = getFont

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let FontName( _
                ByVal Value As String)
Const ProcName As String = "FontName"
On Error GoTo Err

mFontName = Value
setFlag CellFlagFontChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get FontName() As String
Const ProcName As String = "FontName"
On Error GoTo Err

FontName = mFontName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let FontSize( _
                ByVal Value As Currency)
Const ProcName As String = "FontSize"
On Error GoTo Err

mFontSize = Value
setFlag CellFlagFontChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let ForeColor( _
                ByVal Value As Long)
Const ProcName As String = "ForeColor"
On Error GoTo Err

mForeColor = Value
setFlag CellFlagColorChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ForeColor() As Long
Const ProcName As String = "ForeColor"
On Error GoTo Err

ForeColor = mForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Height( _
                ByVal Value As Long)
Const ProcName As String = "Height"
On Error GoTo Err

mHeight = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Height() As Long
Const ProcName As String = "Height"
On Error GoTo Err

Height = mHeight

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get IsFixed() As Boolean
IsFixed = mIsFixed
End Property

Friend Property Get IsMapped() As Boolean
Const ProcName As String = "IsMapped"
On Error GoTo Err

IsMapped = mIsMapped

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get IsSetAlign() As Boolean
Const ProcName As String = "IsSetAlign"
On Error GoTo Err

IsSetAlign = isFlagSet(CellFlagAlign)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Italic( _
                ByVal Value As Boolean)
Const ProcName As String = "Italic"
On Error GoTo Err

If Value Then
    setFlag CellFlagItalic Or CellFlagFontChanged
Else
    clearFlag CellFlagItalic
    setFlag CellFlagFontChanged
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Italic() As Boolean
Const ProcName As String = "italic"
On Error GoTo Err

Italic = isFlagSet(CellFlagItalic)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Left( _
                ByVal Value As Long)
Const ProcName As String = "Left"
On Error GoTo Err

mLeft = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Left() As Long
Const ProcName As String = "Left"
On Error GoTo Err

Left = mLeft

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Row( _
                ByVal Value As Long)
Const ProcName As String = "Row"
On Error GoTo Err

mRow = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Row() As Long
Const ProcName As String = "Row"
On Error GoTo Err

Row = mRow

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Strikethrough( _
                ByVal Value As Boolean)
Const ProcName As String = "Strikethrough"
On Error GoTo Err

If Value Then
    setFlag CellFlagStrikethrough Or CellFlagFontChanged
Else
    clearFlag CellFlagStrikethrough
    setFlag CellFlagFontChanged
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Strikethrough() As Boolean
Const ProcName As String = "Strikethrough"
On Error GoTo Err

Strikethrough = isFlagSet(CellFlagStrikethrough)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Top( _
                ByVal Value As Long)
Const ProcName As String = "top"
On Error GoTo Err

mTop = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Top() As Long
Const ProcName As String = "top"
On Error GoTo Err

Top = mTop

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Underline( _
                ByVal Value As Boolean)
Const ProcName As String = "underline"
On Error GoTo Err

If Value Then
    setFlag CellFlagUnderline Or CellFlagFontChanged
Else
    clearFlag CellFlagUnderline
    setFlag CellFlagFontChanged
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Underline() As Boolean
Const ProcName As String = "underline"
On Error GoTo Err

Underline = isFlagSet(CellFlagUnderline)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Value( _
                ByVal Value As String)
Const ProcName As String = "Value"
On Error GoTo Err

mValue = Value
setFlag CellFlagValueChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Value() As String
Const ProcName As String = "Value"
On Error GoTo Err

Value = mValue

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ValueLabel() As Label
Const ProcName As String = "ValueLabel"
On Error GoTo Err

Set ValueLabel = mGrid.GetValueLabel(mCellLabelIndex)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Width( _
                ByVal Value As Long)
Const ProcName As String = "Width"
On Error GoTo Err

mWidth = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Width() As Long
Const ProcName As String = "Width"
On Error GoTo Err

Width = mWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

Value = ""

BackColor = 0
ForeColor = 0

Dim defaultFont As StdFont
Set defaultFont = mGrid.GetDefaultCellFont(mRow, mCol)
Bold = defaultFont.Bold
Italic = defaultFont.Italic
FontName = defaultFont.Name
FontSize = defaultFont.Size
Strikethrough = defaultFont.Strikethrough
Underline = defaultFont.Underline

clearFlag CellFlagAlign
setFlag CellFlagAlignmentChanged

clearFlag CellFlagInvertColors

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

Set mGrid = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub HottrackOff()
Const ProcName As String = "HottrackOff"
On Error GoTo Err

Invert
Paint
If Not mGrid Is Nothing Then mGrid.HideHottrack

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub HottrackOn()
Const ProcName As String = "HottrackOn"
On Error GoTo Err

Invert
showColorAttributes
mGrid.ShowHottrack Left, _
                    Top + Height, _
                    Width

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pGrid As TWGrid, _
                ByVal pRow As Long, _
                ByVal pCol As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mGrid = pGrid
mRow = pRow
mCol = pCol

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Invert()
Const ProcName As String = "Invert"
On Error GoTo Err

flipFlag CellFlagInvertColors
setFlag CellFlagColorChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function IsXInCellHorizResizeZone( _
                ByVal X As Long) As Boolean
Const ProcName As String = "IsXInCellHorizResizeZone"
On Error GoTo Err

If X >= mLeft + mWidth - (ResizeZoneSizePixels * Screen.TwipsPerPixelX) And _
    X <= mLeft + mWidth + (ResizeZoneSizePixels * Screen.TwipsPerPixelX) _
Then
    IsXInCellHorizResizeZone = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function IsYInCellVertResizeZone( _
                ByVal Y As Long) As Boolean
Const ProcName As String = "IsYInCellVertResizeZone"
On Error GoTo Err

If Y >= mTop + mHeight - (ResizeZoneSizePixels * Screen.TwipsPerPixelY) And _
    Y <= mTop + mHeight + (ResizeZoneSizePixels * Screen.TwipsPerPixelY) _
Then
    IsYInCellVertResizeZone = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Map( _
                ByVal pIsFixed As Boolean, _
                ByVal pGridLines As GridLineSettings, _
                ByVal pGridLinesFixed As GridLineSettings, _
                ByVal pLeft As Long, _
                ByVal pTop As Long, _
                ByVal pWidth As Long, _
                ByVal pHeight As Long)
Const ProcName As String = "Map"
On Error GoTo Err

mIsFixed = pIsFixed
mLeft = pLeft
mTop = pTop
mWidth = pWidth
mHeight = pHeight

Dim lValueLeft As Long
lValueLeft = mLeft + gTextPaddingTwips

Dim lEffectiveWidth As Long
Dim lValueWidth As Long

If mIsFixed Then
    lEffectiveWidth = mWidth - mGrid.GridLineWidthTwipsXFixed
    lValueWidth = lEffectiveWidth - 2 * gTextPaddingTwips
    If pGridLinesFixed = TwGridGridInset Or pGridLinesFixed = TwGridGridRaised Then
        lValueWidth = lValueWidth - 2 * g3dBorderWidthTwipsX
        lValueLeft = lValueLeft + g3dBorderWidthTwipsX
    End If
Else
    lEffectiveWidth = mWidth - mGrid.GridLineWidthTwipsX
    lValueWidth = lEffectiveWidth - 2 * gTextPaddingTwips
    If pGridLines = TwGridGridInset Or pGridLines = TwGridGridRaised Then
        lValueWidth = lValueWidth - 2 * g3dBorderWidthTwipsX
        lValueLeft = lValueLeft + g3dBorderWidthTwipsX
    End If
End If

If lEffectiveWidth < 0 Then lEffectiveWidth = 0
If lValueWidth < 0 Then lValueWidth = 0

Dim lEffectiveHeight As Long
If mIsFixed Then
    lEffectiveHeight = mHeight - mGrid.GridLineWidthTwipsYFixed
Else
    lEffectiveHeight = mHeight - mGrid.GridLineWidthTwipsY
End If

If lEffectiveHeight < 0 Then lEffectiveHeight = 0

mCellLabelIndex = mGrid.AllocateCell
Dim cellLabel As Label
Set cellLabel = mGrid.GetCellLabel(mCellLabelIndex)

cellLabel.Move mLeft, mTop, lEffectiveWidth, lEffectiveHeight
cellLabel.ZOrder 0

Dim ValueLabel As Label
Set ValueLabel = mGrid.GetValueLabel(mCellLabelIndex)
ValueLabel.WordWrap = mIsFixed
ValueLabel.ZOrder 0
ValueLabel.Width = lValueWidth

Dim textHeight As Long
textHeight = mGrid.GetTextHeight(getFont, mIsFixed)
If textHeight < lEffectiveHeight Then
    ValueLabel.Height = textHeight
Else
'    If CellLabel.Height - 2 * gTextPaddingTwips > 0 Then
'        ValueLabel.Height = CellLabel.Height - 2 * gTextPaddingTwips
'    Else
'        ValueLabel.Height = 0
'    End If
    ValueLabel.Height = lEffectiveHeight
End If
ValueLabel.Left = lValueLeft

showBorders IIf(mIsFixed, pGridLinesFixed, pGridLines)

mIsMapped = True
setFlag CellFlagAlignmentChanged Or CellFlagColorChanged Or CellFlagFontChanged Or CellFlagValueChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Paint()
Const ProcName As String = "Paint"
On Error GoTo Err

If mCellLabelIndex > 0 Then
    showAttributes
    showValue
    mGrid.GetCellLabel(mCellLabelIndex).Visible = True
    mGrid.GetValueLabel(mCellLabelIndex).Visible = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub PaintSelected( _
                ByVal pForeColorSel As Long, _
                ByVal pBackColorSel As Long)
Dim lLabel As Label

Const ProcName As String = "PaintSelected"
On Error GoTo Err

If mCellLabelIndex > 0 Then
    Set lLabel = mGrid.GetCellLabel(mCellLabelIndex)
    lLabel.BackColor = pBackColorSel
    
    Set lLabel = mGrid.GetValueLabel(mCellLabelIndex)
    lLabel.BackColor = pBackColorSel
    lLabel.ForeColor = pForeColorSel
End If
setFlag CellFlagColorChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub UnMap()
Dim cellLabel As Label
Dim ValueLabel As Label
Dim borderlabel As Label

Const ProcName As String = "UnMap"
On Error GoTo Err

mIsMapped = False

If mCellLabelIndex > 0 Then
    Set cellLabel = mGrid.GetCellLabel(mCellLabelIndex)
    cellLabel.Visible = False

    Set ValueLabel = mGrid.GetValueLabel(mCellLabelIndex)
    ValueLabel.Visible = False
End If

If mBorderLabelsIndex > 0 Then
    Set borderlabel = mGrid.GetTopBorderLabel(mBorderLabelsIndex)
    borderlabel.Visible = False
    
    Set borderlabel = mGrid.GetLeftBorderLabel(mBorderLabelsIndex)
    borderlabel.Visible = False
    
    Set borderlabel = mGrid.GetBottomBorderLabel(mBorderLabelsIndex)
    borderlabel.Visible = False
    
    Set borderlabel = mGrid.GetRightBorderLabel(mBorderLabelsIndex)
    borderlabel.Visible = False
End If

mCellLabelIndex = -1
mBorderLabelsIndex = -1

mLeft = 0
mTop = 0
mHeight = 0
mWidth = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clearFlag( _
                ByVal flag As CellFlags)
Const ProcName As String = "clearFlag"
On Error GoTo Err

mFlags = mFlags And (Not flag)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function cloneFont( _
                ByVal Font As StdFont) As StdFont
Dim lFont As New StdFont
Const ProcName As String = "cloneFont"
On Error GoTo Err

With lFont
    .Bold = Font.Bold
    .Italic = Font.Italic
    .Name = Font.Name
    .Size = Font.Size
    .Strikethrough = Font.Strikethrough
    .Underline = Font.Underline
    .Weight = Font.Weight
End With
Set cloneFont = lFont

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function convertAlignment( _
                ByVal pAlign As AlignmentSettings, _
                ByVal Value As String) As AlignmentConstants
Const ProcName As String = "convertAlignment"
On Error GoTo Err

Select Case pAlign
Case TwGridAlignLeftTop
    convertAlignment = vbLeftJustify
Case TwGridAlignLeftCenter
    convertAlignment = vbLeftJustify
Case TwGridAlignLeftBottom
    convertAlignment = vbLeftJustify
Case TwGridAlignCenterTop
    convertAlignment = vbCenter
Case TwGridAlignCenterCenter
    convertAlignment = vbCenter
Case TwGridAlignCenterBottom
    convertAlignment = vbCenter
Case TwGridAlignRightTop
    convertAlignment = vbRightJustify
Case TwGridAlignRightCenter
    convertAlignment = vbRightJustify
Case TwGridAlignRightBottom
    convertAlignment = vbRightJustify
Case TwGridAlignGeneral
    If IsNumeric(Value) Then
        convertAlignment = vbRightJustify
    Else
        convertAlignment = vbLeftJustify
    End If
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Sub flipFlag( _
                ByVal flag As CellFlags)
Const ProcName As String = "flipFlag"
On Error GoTo Err

mFlags = mFlags Xor flag

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getCustomFont() As StdFont
Dim aFont As StdFont
Const ProcName As String = "getCustomFont"
On Error GoTo Err

Set aFont = cloneFont(mGrid.GetDefaultCellFont(mRow, mCol))
If mFontName <> "" Then aFont.Name = mFontName
If isFlagSet(CellFlagBold) Then aFont.Bold = True
If isFlagSet(CellFlagItalic) Then aFont.Italic = True
If mFontSize > 0 Then aFont.Size = mFontSize
If isFlagSet(CellFlagStrikethrough) Then aFont.Strikethrough = True
If isFlagSet(CellFlagUnderline) Then aFont.Underline = True
Set getCustomFont = aFont

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFont() As StdFont
Const ProcName As String = "getFont"
On Error GoTo Err

If useDefaultFont Then
    Set getFont = mGrid.GetDefaultCellFont(mRow, mCol)
Else
    Set getFont = getCustomFont
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function invertColor(ByVal pColor As Long) As Long
invertColor = NormalizeColor(pColor) Xor &HFFFFFF
End Function

Private Function isFlagSet( _
                ByVal flag As CellFlags) As Boolean
Const ProcName As String = "isFlagSet"
On Error GoTo Err

isFlagSet = CBool((mFlags And flag) <> 0)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub setFlag( _
                ByVal flag As CellFlags)
Const ProcName As String = "setFlag"
On Error GoTo Err

mFlags = mFlags Or flag

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showAlignment()
Const ProcName As String = "showAlignment"
On Error GoTo Err

If Not isFlagSet(CellFlagAlignmentChanged) Then Exit Sub

Dim ValueLabel As Label
Set ValueLabel = mGrid.GetValueLabel(mCellLabelIndex)

Dim lRowHeight As Long
lRowHeight = mGrid.RowHeight(mRow)

Dim lTop As Long
Select Case mGrid.GetCellAlignment(mRow, mCol)
Case TwGridAlignLeftTop, TwGridAlignCenterTop, TwGridAlignRightTop
    lTop = (lRowHeight - ValueLabel.Height) / 2
    If lTop >= gTextPaddingTwips Then
        lTop = gTextPaddingTwips
    ElseIf lTop < 0 Then
        lTop = 0
    End If
Case TwGridAlignLeftCenter, TwGridAlignCenterCenter, TwGridAlignRightCenter, TwGridAlignGeneral
    lTop = (lRowHeight - ValueLabel.Height) / 2
Case TwGridAlignLeftBottom, TwGridAlignCenterBottom, TwGridAlignRightBottom
    lTop = (lRowHeight - ValueLabel.Height) / 2
    If lTop >= gTextPaddingTwips Then
        lTop = lRowHeight - ValueLabel.Height - gTextPaddingTwips
    ElseIf lTop < 0 Then
        lTop = lRowHeight - ValueLabel.Height
    End If
End Select

ValueLabel.Alignment = convertAlignment(mGrid.GetCellAlignment(mRow, mCol), ValueLabel.Caption)
ValueLabel.Top = mTop + lTop

clearFlag CellFlagAlignmentChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showAttributes()

Const ProcName As String = "showAttributes"
On Error GoTo Err

showFont

showAlignment

showColorAttributes

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Sub showBorders(ByVal pGridLines As GridLineSettings)
Const ProcName As String = "showBorders"
On Error GoTo Err

If pGridLines = TwGridGridFlat Then Exit Sub
If pGridLines = TwGridGridNone Then Exit Sub

Dim lRaised As Boolean
lRaised = (pGridLines = TwGridGridRaised)

Dim cellLabel As Label
Set cellLabel = mGrid.GetCellLabel(mCellLabelIndex)
With cellLabel
    .Width = mWidth
    .Height = mHeight
    
    mBorderLabelsIndex = mGrid.AllocateBorders
    
    Dim borderlabel As Label
    If lRaised Then
        Set borderlabel = mGrid.GetTopBorderLabel(mBorderLabelsIndex)
    Else
        Set borderlabel = mGrid.GetBottomBorderLabel(mBorderLabelsIndex)
    End If
    borderlabel.Move .Left, .Top, .Width, g3dBorderWidthTwipsY
    borderlabel.Visible = True
    borderlabel.ZOrder 0
    
    If lRaised Then
        Set borderlabel = mGrid.GetLeftBorderLabel(mBorderLabelsIndex)
    Else
        Set borderlabel = mGrid.GetRightBorderLabel(mBorderLabelsIndex)
    End If
    borderlabel.Move .Left, .Top, g3dBorderWidthTwipsX, .Height
    borderlabel.Visible = True
    borderlabel.ZOrder 0
    
    If lRaised Then
        Set borderlabel = mGrid.GetBottomBorderLabel(mBorderLabelsIndex)
    Else
        Set borderlabel = mGrid.GetTopBorderLabel(mBorderLabelsIndex)
    End If
    borderlabel.Move .Left, .Top + .Height - g3dBorderWidthTwipsY, .Width
    borderlabel.Visible = True
    borderlabel.ZOrder 0
    
    If lRaised Then
        Set borderlabel = mGrid.GetRightBorderLabel(mBorderLabelsIndex)
    Else
        Set borderlabel = mGrid.GetLeftBorderLabel(mBorderLabelsIndex)
    End If
    borderlabel.Move .Left + .Width - g3dBorderWidthTwipsX, .Top, borderlabel.Width, .Height
    borderlabel.Visible = True
    borderlabel.ZOrder 0
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showColorAttributes()
Const ProcName As String = "showColorAttributes"
On Error GoTo Err

If Not isFlagSet(CellFlagColorChanged) Then Exit Sub

Dim lBackColor As Long
lBackColor = mGrid.GetCellBackcolor(mRow, mCol)

Dim lForeColor As Long
lForeColor = mGrid.GetCellForecolor(mRow, mCol)

If isFlagSet(CellFlagInvertColors) Then
'    If lForeColor = SystemColorConstants.vbWindowText Then
'        lForeColor = SystemColorConstants.vbInactiveTitleBarText
'    ElseIf lForeColor = SystemColorConstants.vbInactiveTitleBarText Then
'        lForeColor = SystemColorConstants.vbWindowText
'    ElseIf lForeColor = SystemColorConstants.vbButtonText Then
'        lForeColor = SystemColorConstants.vbHighlight
'    ElseIf lForeColor > 0 Then
'        lForeColor = IIf((lForeColor Xor &HFFFFFF) = 0, 1, lForeColor Xor &HFFFFFF)
'    End If
    lForeColor = invertColor(lForeColor)
    
'    If lBackColor = SystemColorConstants.vbWindowBackground Then
'        lBackColor = SystemColorConstants.vbHighlight
'    ElseIf lBackColor = SystemColorConstants.vbHighlight Then
'        lBackColor = SystemColorConstants.vbWindowBackground
'    ElseIf lBackColor = SystemColorConstants.vbButtonFace Then
'        lBackColor = SystemColorConstants.vb3DHighlight
'    ElseIf lBackColor > 0 Then
'        lBackColor = IIf((lBackColor Xor &HFFFFFF) = 0, 1, lBackColor Xor &HFFFFFF)
'    End If
    lBackColor = invertColor(lBackColor)
End If

mGrid.GetCellLabel(mCellLabelIndex).BackColor = lBackColor
mGrid.GetValueLabel(mCellLabelIndex).BackColor = lBackColor
mGrid.GetValueLabel(mCellLabelIndex).ForeColor = lForeColor

clearFlag CellFlagColorChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showFont()
Dim ValueLabel As Label

Const ProcName As String = "showFont"
On Error GoTo Err

If Not isFlagSet(CellFlagFontChanged) Then Exit Sub

Set ValueLabel = mGrid.GetValueLabel(mCellLabelIndex)

Set ValueLabel.Font = getFont

clearFlag CellFlagFontChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub showValue()
Dim ValueLabel As Label

Const ProcName As String = "showValue"
On Error GoTo Err

If Not isFlagSet(CellFlagValueChanged) Then Exit Sub

Set ValueLabel = mGrid.GetValueLabel(mCellLabelIndex)
ValueLabel.Caption = mValue

clearFlag CellFlagValueChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function useDefaultFont() As Boolean
Dim defaultFont As StdFont
Const ProcName As String = "useDefaultFont"
On Error GoTo Err

Set defaultFont = mGrid.GetDefaultCellFont(mRow, mCol)

If mFontName <> "" And mFontName <> defaultFont.Name Then Exit Function
If Bold <> defaultFont.Bold Then Exit Function
If Italic <> defaultFont.Italic Then Exit Function
If mFontSize > 0 And mFontSize <> defaultFont.Size Then Exit Function
If Strikethrough <> defaultFont.Strikethrough Then Exit Function
If Underline <> defaultFont.Underline Then Exit Function

useDefaultFont = True
'useDefaultFont = Not (mFontName <> "" Or mFontSize > 0 Or isFlagSet(CellFlagBold Or CellFlagItalic Or CellFlagStrikethrough Or CellFlagUnderline))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



