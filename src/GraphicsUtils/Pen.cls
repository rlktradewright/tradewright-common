VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pen"
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

Private Const ModuleName                            As String = "Pen"

Private Const ConfigSettingColor                    As String = "&Color"
Private Const ConfigSettingWidth                    As String = "&Width"
Private Const ConfigSettingLineStyle                As String = "&LineStyle"
Private Const ConfigSettingHatchStyle               As String = "&HatchStyle"
Private Const ConfigSettingCustomDashPattern        As String = "&CustomDashPattern"
Private Const ConfigSettingEndcaps                  As String = "&EndCaps"
Private Const ConfigSettingJoinMode                 As String = "&JoinMode"

'@================================================================================
' Member variables
'@================================================================================

Private mColor As Long
Private mWidth As Double
Private mLineStyle As LineStyles
Private mHatchStyle As HatchStyles
Private mHandle As Long
Private mCustomDashPattern() As Single
Private mIsPixelPen As Boolean
Private mEndcaps As Endcaps
Private mJoinMode As JoinModes

Private mConfig As ConfigurationSection

' only used for a logical pen
Private WithEvents mGraphics As Graphics
Attribute mGraphics.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mEndcaps = EndcapRound
mJoinMode = JoinMiter
End Sub

Private Sub Class_Terminate()
deletePenObject
Debug.Print "Pen terminated"
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mGraphics Event Handlers
'@================================================================================

Private Sub mGraphics_BoundaryChanged()
deletePenObject
End Sub

Private Sub mGraphics_Resize()
deletePenObject
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Let Color(ByVal Value As Long)
Const ProcName As String = "Color"
On Error GoTo Err

AssertArgument gIsValidColor(Value), "Not a valid color"

deletePenObject
mColor = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingColor, mColor

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

mConfig.SetSetting ConfigSettingColor, mColor
mConfig.SetSetting ConfigSettingWidth, mWidth
mConfig.SetSetting ConfigSettingLineStyle, mLineStyle
mConfig.SetSetting ConfigSettingHatchStyle, mHatchStyle
mConfig.SetSetting ConfigSettingEndcaps, mEndcaps
mConfig.SetSetting ConfigSettingJoinMode, mJoinMode
mConfig.SetSetting ConfigSettingCustomDashPattern, gSingleArrayToString(mCustomDashPattern)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Let CustomDashPattern(ByRef Value() As Single)
Const ProcName As String = "CustomDashPattern"
On Error GoTo Err

On Error Resume Next
Dim size As Long
size = UBound(Value) + 1
On Error GoTo Err

AssertArgument size <= 16, "The array length must not exceed 16"

deletePenObject
mCustomDashPattern = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingCustomDashPattern, gSingleArrayToString(mCustomDashPattern)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get CustomDashPattern() As Single()
CustomDashPattern = mCustomDashPattern
End Property

Public Property Let Endcaps(ByVal Value As Endcaps)
Const ProcName As String = "Endcaps"
On Error GoTo Err

Select Case Value
Case EndCapFlat
Case EndcapRound
Case EndcapSquare
Case Else
    AssertArgument False, "Value must be a member of the Endcaps enum"
End Select

deletePenObject
mEndcaps = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingEndcaps, mEndcaps

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Endcaps() As Endcaps
Endcaps = mEndcaps
End Property

Friend Property Get Handle() As Long

Const ProcName As String = "Handle"
On Error GoTo Err

If mHandle = 0 Then createThePen

Handle = mHandle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Let HatchStyle(ByVal Value As HatchStyles)
Const ProcName As String = "HatchStyle"
On Error GoTo Err

Select Case Value
Case HatchNone
Case HatchHorizontal
Case HatchVertical
Case HatchUpwardDiagonal
Case HatchDownwardDiagonal
Case HatchCross
Case HatchDiagonalCross
Case Else
    AssertArgument False, "Not a valid HatchStyle"
End Select

deletePenObject

mHatchStyle = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingHatchStyle, mHatchStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get HatchStyle() As HatchStyles
HatchStyle = mHatchStyle
End Property

Public Property Get IsPixelPen() As Boolean
IsPixelPen = mIsPixelPen
End Property

Public Property Let JoinMode(ByVal Value As JoinModes)
Const ProcName As String = "JoinMode"
On Error GoTo Err

Select Case Value
Case JoinBevel
Case JoinMiter
Case JoinRound
Case Else
    AssertArgument False, "Value must be a member of the JoinModes enum"
End Select

deletePenObject
mJoinMode = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingJoinMode, mJoinMode

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get JoinMode() As JoinModes
JoinMode = mJoinMode
End Property

Public Property Let LineStyle(ByVal Value As LineStyles)
Const ProcName As String = "LineStyle"
On Error GoTo Err

Select Case Value
Case LineSolid
Case LineDash
Case LineDot
Case LineDashDot
Case LineDashDotDot
Case LineInvisible
Case LineInsideSolid
Case LineAlternate
Case LineCustom
Case Else
    AssertArgument False, "Not a valid LineStyle"
End Select

deletePenObject
mLineStyle = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingLineStyle, mLineStyle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get LineStyle() As LineStyles
LineStyle = mLineStyle
End Property

Public Property Let Width(ByVal Value As Double)
Const ProcName As String = "Width"
On Error GoTo Err

AssertArgument Value > 0, "Widfh must be greater than zero"

deletePenObject
mWidth = Value
If Not mConfig Is Nothing Then mConfig.SetSetting ConfigSettingWidth, mWidth

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Width() As Double
Width = mWidth
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pColor As Long, _
                ByVal pWidth As Double, _
                ByVal pLineStyle As LineStyles, _
                ByVal pHatchStyle As HatchStyles, _
                ByVal pIsPixelPen As Boolean, _
                ByVal pGraphics As Graphics)
Const ProcName As String = "Initialise"
On Error GoTo Err

Color = pColor
Width = pWidth
LineStyle = pLineStyle
HatchStyle = pHatchStyle
mIsPixelPen = pIsPixelPen

AssertArgument mIsPixelPen Or Not pGraphics Is Nothing, "pGrapics cannot be nothing for a logical pen"

Set mGraphics = pGraphics

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
mWidth = mConfig.GetSetting(ConfigSettingWidth, 1)
mLineStyle = mConfig.GetSetting(ConfigSettingLineStyle, CStr(LineSolid))
mHatchStyle = mConfig.GetSetting(ConfigSettingHatchStyle, CStr(HatchStyles.HatchNone))
mEndcaps = mConfig.GetSetting(ConfigSettingEndcaps, CStr(EndCapFlat))
mJoinMode = mConfig.GetSetting(ConfigSettingJoinMode, CStr(JoinMiter))
mCustomDashPattern = gSingleArrayFromString(mConfig.GetSetting(ConfigSettingCustomDashPattern))
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

Private Sub createThePen()
Const ProcName As String = "createThePen"
Dim failpoint As String
On Error GoTo Err

Dim lPenHandle As Long
Dim lSettings As LOGBRUSH
Dim lPenStyle As Long
Dim lWidth As Long
Dim lCustomDashPatternLength As Long

lSettings.lbColor = gNormalizeColor(mColor)

lPenStyle = mLineStyle
lPenStyle = lPenStyle Or mEndcaps
lPenStyle = lPenStyle Or mJoinMode

If mHatchStyle = HatchNone Then
    lSettings.lbStyle = BS_SOLID
Else
    lSettings.lbStyle = BS_HATCHED
    lSettings.lbHatch = mHatchStyle Xor &H80000000
End If

If Not mIsPixelPen Then
    lPenStyle = lPenStyle Or PS_GEOMETRIC
    lWidth = mGraphics.ConvertLogicalToGdiX(mWidth)
ElseIf mWidth = 1 And ((lPenStyle And PS_INSIDEFRAME) <> PS_INSIDEFRAME) Then
    lPenStyle = lPenStyle Or PS_COSMETIC
    lWidth = 1
Else
    lPenStyle = lPenStyle Or PS_GEOMETRIC
    lWidth = mWidth * LogicalUnitsPerPixel
End If


If (lPenStyle And PS_USERSTYLE) = PS_USERSTYLE Then
    On Error Resume Next
    lCustomDashPatternLength = UBound(mCustomDashPattern) + 1
    On Error GoTo Err
    Assert lCustomDashPatternLength <> 0, "No dash style has been specified"
        
    ReDim lCustomDashPattern(lCustomDashPatternLength - 1) As Long
    
    Dim i As Long
    For i = 0 To UBound(mCustomDashPattern)
        If (lPenStyle And PS_GEOMETRIC) <> PS_GEOMETRIC Then
            lCustomDashPattern(i) = mCustomDashPattern(i)
        ElseIf mIsPixelPen Then
            lCustomDashPattern(i) = mCustomDashPattern(i) * LogicalUnitsPerPixel
        Else
            lCustomDashPattern(i) = mGraphics.ConvertLogicalToGdiX(mCustomDashPattern(i))
        End If
    Next
    
    mHandle = ExtCreatePen(lPenStyle, lWidth, lSettings, lCustomDashPatternLength, VarPtr(lCustomDashPattern(0)))
Else
    mHandle = ExtCreatePen(lPenStyle, lWidth, lSettings, 0, 0)
End If

If mHandle = 0 Then
    gHandleWin32Error GetLastError
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deletePenObject()
Const ProcName As String = "deletePenObject"
On Error GoTo Err

If mHandle <> 0 Then
    If DeleteObject(mHandle) = 0 Then gHandleWin32Error GetLastError
    mHandle = 0
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

