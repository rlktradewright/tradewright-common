VERSION 5.00
Begin VB.UserControl TransparentPanel 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   4725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7350
   ControlContainer=   -1  'True
   ScaleHeight     =   4725
   ScaleWidth      =   7350
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4080
      ScaleHeight     =   495
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "TransparentPanel"
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

'@================================================================================
' Events
'@================================================================================

Event Click()

Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)

Event KeyPress(KeyAscii As Integer)

Event KeyUp(KeyCode As Integer, Shift As Integer)

Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Event Paint()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "TransparentPanel"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, _
                    Shift, _
                    UserControl.ScaleX(X, UserControl.ScaleMode, vbContainerPosition), _
                    UserControl.ScaleY(Y, UserControl.ScaleMode, vbContainerPosition))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, _
                    Shift, _
                    UserControl.ScaleX(X, UserControl.ScaleMode, vbContainerPosition), _
                    UserControl.ScaleY(Y, UserControl.ScaleMode, vbContainerPosition))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, _
                    Shift, _
                    UserControl.ScaleX(X, UserControl.ScaleMode, vbContainerPosition), _
                    UserControl.ScaleY(Y, UserControl.ScaleMode, vbContainerPosition))
End Sub

Private Sub UserControl_Paint()
Const ProcName As String = "UserControl_Paint"
On Error GoTo Err

Debug.Print "TransPanel:Paint"
RaiseEvent Paint
paintIt

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next
End Sub

Private Sub UserControl_Resize()
Const ProcName As String = "UserControl_Resize"
On Error GoTo Err

Picture1.Width = UserControl.Width
Picture1.Height = UserControl.Height

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Const ProcName As String = "UserControl_WriteProperties"
On Error GoTo Err

PropBag.WriteProperty "ScaleWidth", Picture1.ScaleWidth, UserControl.Width
PropBag.WriteProperty "ScaleTop", Picture1.ScaleTop, 0
PropBag.WriteProperty "ScaleMode", Picture1.ScaleMode, 1
PropBag.WriteProperty "ScaleLeft", Picture1.ScaleLeft, 0
PropBag.WriteProperty "ScaleHeight", Picture1.ScaleHeight, UserControl.Height

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Control Event Handlers
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get hDC() As Long
Const ProcName As String = "hDC"
On Error GoTo Err

hDC = UserControl.hDC

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get TransparencyColor() As OLE_COLOR
Const ProcName As String = "TransparencyColor"
On Error GoTo Err

TransparencyColor = UserControl.MaskColor

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let TransparencyColor( _
                ByVal Value As OLE_COLOR)
Const ProcName As String = "TransparencyColor"
On Error GoTo Err

UserControl.MaskColor = Value
UserControl.BackColor = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

' this method is hidden at present
Public Sub DrawLine( _
        ByVal x1 As Long, _
        ByVal y1 As Long, _
        ByVal x2 As Long, _
        ByVal y2 As Long)
Attribute DrawLine.VB_MemberFlags = "40"
Const ProcName As String = "DrawLine"
On Error GoTo Err

UserControl.DrawMode = vbXorPen
UserControl.DrawWidth = 2
UserControl.Line (x1, y1)-(x2, y2), vbRed
paintIt

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Refresh()
Const ProcName As String = "Refresh"
On Error GoTo Err

paintIt

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'The Underscore following "Scale" is necessary because it
'is a Reserved Word in VBA.


Public Sub Scale_(Optional x1 As Variant, Optional y1 As Variant, Optional x2 As Variant, Optional y2 As Variant)
Const ProcName As String = "Scale_"
On Error GoTo Err

Picture1.Scale (x1, y1)-(x2, y2)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."
Const ProcName As String = "ScaleHeight"
On Error GoTo Err

Picture1.ScaleHeight() = New_ScaleHeight
PropertyChanged "ScaleHeight"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



Public Property Get ScaleHeight() As Single
Const ProcName As String = "ScaleHeight"
On Error GoTo Err

ScaleHeight = Picture1.ScaleHeight

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."
Const ProcName As String = "ScaleLeft"
On Error GoTo Err

Picture1.ScaleLeft() = New_ScaleLeft
PropertyChanged "ScaleLeft"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



Public Property Get ScaleLeft() As Single
Const ProcName As String = "ScaleLeft"
On Error GoTo Err

ScaleLeft = Picture1.ScaleLeft

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ScaleMode(ByVal New_ScaleMode As Integer)
Attribute ScaleMode.VB_Description = "Returns/sets a value indicating measurement units for object coordinates when using graphics methods or positioning controls."
Const ProcName As String = "ScaleMode"
On Error GoTo Err

Picture1.ScaleMode() = New_ScaleMode
PropertyChanged "ScaleMode"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



Public Property Get ScaleMode() As Integer
Const ProcName As String = "ScaleMode"
On Error GoTo Err

ScaleMode = Picture1.ScaleMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."
Const ProcName As String = "ScaleTop"
On Error GoTo Err

Picture1.ScaleTop() = New_ScaleTop
PropertyChanged "ScaleTop"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



Public Property Get ScaleTop() As Single
Const ProcName As String = "ScaleTop"
On Error GoTo Err

ScaleTop = Picture1.ScaleTop

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."
Const ProcName As String = "ScaleWidth"
On Error GoTo Err

Picture1.ScaleWidth() = New_ScaleWidth
PropertyChanged "ScaleWidth"

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



Public Property Get ScaleWidth() As Single
Const ProcName As String = "ScaleWidth"
On Error GoTo Err

ScaleWidth = Picture1.ScaleWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property



Public Function ScaleX(ByVal Width As Single, ByVal FromScale As Variant, ByVal ToScale As Variant) As Single
Attribute ScaleX.VB_Description = "Converts the value for the width of a Form, PictureBox, or Printer from one unit of measure to another."
Const ProcName As String = "ScaleX"
On Error GoTo Err

ScaleX = Picture1.ScaleX(Width, FromScale, ToScale)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



Public Function ScaleY(ByVal Height As Single, ByVal FromScale As Variant, ByVal ToScale As Variant) As Single
Const ProcName As String = "ScaleY"
On Error GoTo Err

ScaleY = Picture1.ScaleY(Height, FromScale, ToScale)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub paintIt()
Const ProcName As String = "paintIt"
On Error GoTo Err

Set UserControl.MaskPicture = UserControl.Image
UserControl.Refresh

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

