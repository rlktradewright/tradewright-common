VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HatchedBrush"
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
Implements INativeBrush

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

Private Const ModuleName                            As String = "HatchedBrush"

Private Const ConfigSettingColor                    As String = "&Color"
Private Const ConfigSettingStyle                    As String = "&Style"

'@================================================================================
' Member variables
'@================================================================================

Private mColor As Long
Private mStyle As HatchStyles

Private mHandle As Long

Private mConfig As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mStyle = HatchStyles.HatchHorizontal
End Sub

Private Sub Class_Terminate()
deleteBrushObject
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
' INativeBrush Interface Members
'@================================================================================

Private Property Get INativeBrush_Handle() As Long
Const ProcName As String = "INativeBrush_Handle"
On Error GoTo Err

If mHandle = 0 Then createTheBrush
INativeBrush_Handle = mHandle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Private Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"
On Error GoTo Err

deleteBrushObject

AssertArgument gIsValidColor(Value), "Not a valid color"

mColor = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Color() As Long
Color = mColor
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
mConfig.SetSetting ConfigSettingColor, mColor
mConfig.SetSetting ConfigSettingStyle, mStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Private Property Let Style(ByVal Value As HatchStyles)
Const ProcName As String = "Style"
On Error GoTo Err

Select Case Value
Case HatchHorizontal
Case HatchVertical
Case HatchUpwardDiagonal
Case HatchDownwardDiagonal
Case HatchCross
Case HatchDiagonalCross
Case Else
    AssertArgument False, "Not a valid HatchStyle"
End Select

deleteBrushObject

mStyle = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Style() As HatchStyles
Style = mStyle
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pColor As Long, _
                ByVal pStyle As HatchStyles)
Const ProcName As String = "Initialise"
On Error GoTo Err

Color = pColor
Style = pStyle

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
mColor = mConfig.GetSetting(ConfigSettingColor, "0")
mStyle = mConfig.GetSetting(ConfigSettingStyle, CStr(HatchStyles.HatchCross))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveFromConfig()
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

Private Sub createTheBrush()
Const ProcName As String = "createTheBrush"
Dim failpoint As String
On Error GoTo Err

Dim lColor As Long

lColor = gNormalizeColor(mColor)

mHandle = CreateHatchBrush(&H80000000 Xor mStyle, lColor)
If mHandle = 0 Then gHandleWin32Error GetLastError

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deleteBrushObject()
Const ProcName As String = "deleteBrushObject"
On Error GoTo Err

If mHandle <> 0 Then
    If DeleteObject(mHandle) = 0 Then gHandleWin32Error GetLastError
    mHandle = 0
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub


