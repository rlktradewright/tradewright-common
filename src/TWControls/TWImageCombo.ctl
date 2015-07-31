VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.UserControl TWImageCombo 
   ClientHeight    =   1470
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2625
   ScaleHeight     =   1470
   ScaleWidth      =   2625
   ToolboxBitmap   =   "TWImageCombo.ctx":0000
   Begin MSComctlLib.ImageCombo Combo1 
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
   End
End
Attribute VB_Name = "TWImageCombo"
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

Event Change()
Event Click()
Event DropDown()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event OLECompleteDrag(Effect As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "TWCombo"

'@================================================================================
' Member variables
'@================================================================================

Private mListWidth                          As Long
Private mAppearance                         As AppearanceSettings

Private mTheme                                      As ITheme

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_EnterFocus()
If Enabled Then Combo1.SetFocus
End Sub

Private Sub UserControl_Initialize()
mAppearance = cc3D
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_ReadProperties"
On Error GoTo Err

mAppearance = PropBag.ReadProperty("Appearance", AppearanceConstants.cc3D)
Combo1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
Combo1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000008)
UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
Set Combo1.Font = PropBag.ReadProperty("Font", Ambient.Font)
Combo1.CausesValidation = PropBag.ReadProperty("CausesValidation", True)
Combo1.Locked = PropBag.ReadProperty("Locked", False)
Combo1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
Combo1.OLEDropMode = PropBag.ReadProperty("OLEDropMode", 0)
Combo1.OLEDragMode = PropBag.ReadProperty("OLEDragMode", 0)
Combo1.Text = PropBag.ReadProperty("Text", "Combo1")
Combo1.Indentation = PropBag.ReadProperty("Indentation", 0)
ListWidth = PropBag.ReadProperty("ListWidth", 0)

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

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_WriteProperties"
On Error GoTo Err

Call PropBag.WriteProperty("Appearance", mAppearance, AppearanceConstants.cc3D)
Call PropBag.WriteProperty("BackColor", Combo1.BackColor, &H80000005)
Call PropBag.WriteProperty("ForeColor", Combo1.ForeColor, &H80000008)
Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
Call PropBag.WriteProperty("Font", Combo1.Font, Ambient.Font)
Call PropBag.WriteProperty("CausesValidation", Combo1.CausesValidation, True)
Call PropBag.WriteProperty("Locked", Combo1.Locked, False)
Call PropBag.WriteProperty("MousePointer", Combo1.MousePointer, 0)
Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
Call PropBag.WriteProperty("OLEDropMode", Combo1.OLEDropMode, 0)
Call PropBag.WriteProperty("OLEDragMode", Combo1.OLEDragMode, 0)
Call PropBag.WriteProperty("Text", Combo1.Text, "Combo1")
Call PropBag.WriteProperty("Indentation", Combo1.Indentation, 0)
Call PropBag.WriteProperty("ListWidth", ListWidth, 0)
    
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
' Combo1 Event Handlers
'@================================================================================

Private Sub Combo1_Change()
Const ProcName As String = "Combo1_Change"
On Error GoTo Err

Static prevValue As String
If Combo1.Text <> "" And Combo1.Text = prevValue Then Exit Sub

If Not Combo1.SelectedItem Is Nothing Then
    If Combo1.Text = Combo1.SelectedItem.Text Then Exit Sub
End If

Dim l As Long
l = Len(Combo1.Text)
If Combo1.Text <> "" Then
    Dim selItem As Long
    selItem = findSubStringIndex(Combo1.Text)
    If selItem <> 0 Then
        Combo1.ComboItems(selItem).Selected = True
        Combo1.SelStart = l
        Combo1.SelLength = Len(Combo1.Text) - l
    End If
End If
RaiseEvent Change
prevValue = Combo1.Text

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Combo1_Click()
Const ProcName As String = "Combo1_Click"
On Error GoTo Err

    RaiseEvent Click

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Combo1_DropDown()
    RaiseEvent DropDown
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
Const ProcName As String = "Combo1_KeyDown"
On Error GoTo Err

Dim i As Long
Dim posn As Long
If KeyCode = vbKeyUp Then
    posn = 0
    If Combo1.SelectedItem Is Nothing Then
        For i = 1 To Combo1.ComboItems.Count
            If StrComp(Combo1.Text, Combo1.ComboItems(i).Text, vbTextCompare) = 0 Then
                Combo1.ComboItems(i).Selected = True
                Exit Sub
            End If
            If StrComp(Combo1.ComboItems(i).Text, Combo1.Text, vbTextCompare) < 0 Then posn = i
        Next
        If posn = Combo1.ComboItems.Count Then
            KeyCode = 0
        Else
            posn = posn + 1
        End If
        Combo1.ComboItems(posn).Selected = True
    End If
ElseIf KeyCode = vbKeyDown Then
    posn = Combo1.ComboItems.Count
    If Combo1.SelectedItem Is Nothing Then
        For i = Combo1.ComboItems.Count To 1 Step -1
            If StrComp(Combo1.Text, Combo1.ComboItems(i).Text, vbTextCompare) = 0 Then
                Combo1.ComboItems(i).Selected = True
                Exit Sub
            End If
            If StrComp(Combo1.Text, Combo1.ComboItems(i).Text, vbTextCompare) < 0 Then posn = i
        Next
        If posn = 1 Then
            KeyCode = 0
        Else
            posn = posn - 1
        End If
        Combo1.ComboItems(posn).Selected = True
    End If
End If
RaiseEvent KeyDown(KeyCode, Shift)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
Const ProcName As String = "Combo1_KeyPress"
On Error GoTo Err

If KeyAscii = vbKeyBack Then
    If Combo1.SelStart <> 0 And Combo1.SelLength <> 0 Then
        Dim l As Long
        l = Combo1.SelLength
        Combo1.SelStart = Combo1.SelStart - 1
        Combo1.SelLength = l + 1
    End If
End If
RaiseEvent KeyPress(KeyAscii)

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub Combo1_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub Combo1_OLESetData(Data As MSComctlLib.DataObject, DataFormat As Integer)
    RaiseEvent OLESetData(Data, DataFormat)
End Sub

Private Sub Combo1_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)
End Sub

Private Sub Combo1_OLEDragOver(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, X, Y, State)
End Sub

Private Sub Combo1_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, X, Y)
End Sub

Private Sub Combo1_OLECompleteDrag(Effect As Long)
    RaiseEvent OLECompleteDrag(Effect)
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
Const ProcName As String = "Combo1_Validate"
On Error GoTo Err

If Combo1.Text = "" Then Exit Sub
If Combo1.SelectedItem Is Nothing Then
    Dim i As Long
    For i = 1 To Combo1.ComboItems.Count
        If Combo1.Text = Combo1.ComboItems(i).Text Then
            Combo1.ComboItems(i).Selected = True
            Exit Sub
        End If
    Next
    Cancel = True
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Appearance(ByVal Value As AppearanceSettings)
Const ProcName As String = "Appearance"
On Error GoTo Err

mAppearance = Value
PropertyChanged "Appearance"
resize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Appearance() As AppearanceSettings
Appearance = mAppearance
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_UserMemId = -501
Const ProcName As String = "BackColor"
On Error GoTo Err

BackColor = Combo1.BackColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let BackColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "BackColor"
On Error GoTo Err

Combo1.BackColor() = Value
PropertyChanged "BackColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CausesValidation() As Boolean
Const ProcName As String = "CausesValidation"
On Error GoTo Err

    CausesValidation = Combo1.CausesValidation

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let CausesValidation(ByVal Value As Boolean)
Const ProcName As String = "CausesValidation"
On Error GoTo Err

    Combo1.CausesValidation() = Value
    PropertyChanged "CausesValidation"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ComboItems() As IComboItems
Const ProcName As String = "ComboItems"
On Error GoTo Err

    Set ComboItems = Combo1.ComboItems

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Set ComboItems(ByVal Value As IComboItems)
Const ProcName As String = "ComboItems"
On Error GoTo Err

    Set Combo1.ComboItems = Value
    PropertyChanged "ComboItems"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_UserMemId = -514
Const ProcName As String = "Enabled"
On Error GoTo Err

Enabled = UserControl.Enabled

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Enabled(ByVal Value As Boolean)
Const ProcName As String = "Enabled"
On Error GoTo Err

UserControl.Enabled() = Value
PropertyChanged "Enabled"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Font() As Font
Attribute Font.VB_UserMemId = -512
Const ProcName As String = "Font"
On Error GoTo Err

    Set Font = Combo1.Font

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Set Font(ByVal Value As Font)
Const ProcName As String = "Font"
On Error GoTo Err

    Set Combo1.Font = Value
    PropertyChanged "Font"
    resize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_UserMemId = -513
Const ProcName As String = "ForeColor"
On Error GoTo Err

    ForeColor = Combo1.ForeColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ForeColor(ByVal Value As OLE_COLOR)
Const ProcName As String = "ForeColor"
On Error GoTo Err

    Combo1.ForeColor() = Value
    PropertyChanged "ForeColor"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Function GetFirstVisible() As IComboItem
Const ProcName As String = "GetFirstVisible"
On Error GoTo Err

    GetFirstVisible = Combo1.GetFirstVisible()

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Property Get hWnd() As Long
Attribute hWnd.VB_UserMemId = -515
Const ProcName As String = "hWnd"
On Error GoTo Err

    hWnd = UserControl.hWnd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ImageList() As Object
Const ProcName As String = "ImageList"
On Error GoTo Err

    Set ImageList = Combo1.ImageList

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Set ImageList(ByVal Value As Object)
Const ProcName As String = "ImageList"
On Error GoTo Err

    Set Combo1.ImageList = Value
    PropertyChanged "ImageList"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Indentation() As Integer
Const ProcName As String = "Indentation"
On Error GoTo Err

    Indentation = Combo1.Indentation

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Indentation(ByVal Value As Integer)
Const ProcName As String = "Indentation"
On Error GoTo Err

    Combo1.Indentation() = Value
    PropertyChanged "Indentation"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
Const ProcName As String = "Locked"
On Error GoTo Err

    Locked = Combo1.Locked

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Locked(ByVal Value As Boolean)
Const ProcName As String = "Locked"
On Error GoTo Err

    Combo1.Locked() = Value
    PropertyChanged "Locked"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ListWidth() As Long
Const ProcName As String = "ListWidth"
On Error GoTo Err

ListWidth = mListWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ListWidth(ByVal Value As Long)
Const ProcName As String = "ListWidth"
On Error GoTo Err

mListWidth = Value
SendMessage Combo1.hWnd, CB_SETDROPPEDWIDTH, Value / Screen.TwipsPerPixelX, 0
PropertyChanged "ListWidth"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MouseIcon() As Picture
Const ProcName As String = "MouseIcon"
On Error GoTo Err

    Set MouseIcon = Combo1.MouseIcon

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Set MouseIcon(ByVal Value As Picture)
Const ProcName As String = "MouseIcon"
On Error GoTo Err

    Set Combo1.MouseIcon = Value
    PropertyChanged "MouseIcon"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get MousePointer() As Integer
Const ProcName As String = "MousePointer"
On Error GoTo Err

    MousePointer = Combo1.MousePointer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let MousePointer(ByVal Value As Integer)
Const ProcName As String = "MousePointer"
On Error GoTo Err

    Combo1.MousePointer() = Value
    PropertyChanged "MousePointer"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Sub OLEDrag()
Const ProcName As String = "OLEDrag"
On Error GoTo Err

    Combo1.OLEDrag

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Property Get OLEDragMode() As Integer
Const ProcName As String = "OLEDragMode"
On Error GoTo Err

    OLEDragMode = Combo1.OLEDragMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let OLEDragMode(ByVal Value As Integer)
Const ProcName As String = "OLEDragMode"
On Error GoTo Err

    Combo1.OLEDragMode() = Value
    PropertyChanged "OLEDragMode"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get OLEDropMode() As Integer
Const ProcName As String = "OLEDropMode"
On Error GoTo Err

    OLEDropMode = Combo1.OLEDropMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let OLEDropMode(ByVal Value As Integer)
Const ProcName As String = "OLEDropMode"
On Error GoTo Err

    Combo1.OLEDropMode() = Value
    PropertyChanged "OLEDropMode"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SelectedItem(ByVal Value As ComboItem)
Const ProcName As String = "SelectedItem"
On Error GoTo Err

Set Combo1.SelectedItem = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Set SelectedItem(ByVal Value As ComboItem)
Const ProcName As String = "SelectedItem"
On Error GoTo Err

Set Combo1.SelectedItem = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SelectedItem() As ComboItem
Const ProcName As String = "SelectedItem"
On Error GoTo Err

Set SelectedItem = Combo1.SelectedItem

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SelLength() As Long
Const ProcName As String = "SelLength"
On Error GoTo Err

    SelLength = Combo1.SelLength

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SelLength(ByVal Value As Long)
Const ProcName As String = "SelLength"
On Error GoTo Err

    Combo1.SelLength() = Value
    PropertyChanged "SelLength"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SelStart() As Long
Const ProcName As String = "SelStart"
On Error GoTo Err

    SelStart = Combo1.SelStart

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SelStart(ByVal Value As Long)
Const ProcName As String = "SelStart"
On Error GoTo Err

    Combo1.SelStart() = Value
    PropertyChanged "SelStart"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get SelText() As String
Const ProcName As String = "SelText"
On Error GoTo Err

    SelText = Combo1.SelText

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let SelText(ByVal Value As String)
Const ProcName As String = "SelText"
On Error GoTo Err

    Combo1.SelText() = Value
    PropertyChanged "SelText"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
Attribute Text.VB_MemberFlags = "34"
Const ProcName As String = "Text"
On Error GoTo Err

    Text = Combo1.Text

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Text(ByVal Value As String)
Const ProcName As String = "Text"
On Error GoTo Err

    Combo1.Text() = Value
    PropertyChanged "Text"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Theme() As ITheme
Set Theme = mTheme
End Property

Public Property Let Theme(ByVal Value As ITheme)
Const ProcName As String = "Theme"
On Error GoTo Err

Set mTheme = Value
Appearance = mTheme.Appearance
BackColor = mTheme.TextBackColor
ForeColor = mTheme.TextForeColor
If Not mTheme.ComboFont Is Nothing Then Set Font = mTheme.ComboFont

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



Public Property Get ToolTipText() As String
Const ProcName As String = "ToolTipText"
On Error GoTo Err

    ToolTipText = Combo1.ToolTipText

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ToolTipText(ByVal Value As String)
Const ProcName As String = "ToolTipText"
On Error GoTo Err

    Combo1.ToolTipText() = Value
    PropertyChanged "ToolTipText"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



Public Property Get WhatsThisHelpID() As Long
Const ProcName As String = "WhatsThisHelpID"
On Error GoTo Err

    WhatsThisHelpID = Combo1.WhatsThisHelpID

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let WhatsThisHelpID(ByVal Value As Long)
Const ProcName As String = "WhatsThisHelpID"
On Error GoTo Err

    Combo1.WhatsThisHelpID() = Value
    PropertyChanged "WhatsThisHelpID"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================



Public Sub Refresh()
Const ProcName As String = "Refresh"
On Error GoTo Err

    Combo1.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function findSubStringIndex( _
                ByRef str As String) As Long
Const ProcName As String = "findSubStringIndex"
On Error GoTo Err

Dim i As Long
For i = 1 To Combo1.ComboItems.Count
    If Len(Combo1.ComboItems(i).Text) >= Len(str) Then
        If StrComp(str, Left$(Combo1.ComboItems(i).Text, Len(str)), vbTextCompare) = 0 Then
            findSubStringIndex = i
            Exit Function
        End If
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub resize()
Const ProcName As String = "resize"
On Error GoTo Err

If mAppearance = cc3D Then
    Combo1.Move 0, 0, UserControl.Width
    UserControl.Height = Combo1.Height
Else
    Combo1.Move -2 * Screen.TwipsPerPixelX, -2 * Screen.TwipsPerPixelY, UserControl.Width + 4 * Screen.TwipsPerPixelX
    UserControl.Height = Combo1.Height - 4 * Screen.TwipsPerPixelY
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
