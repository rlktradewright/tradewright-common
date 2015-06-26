VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Size"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

Private Const ModuleName                        As String = "Size"

Private Const ConfigSettingScaleUnitHeight      As String = "&ScaleUnitHeight"
Private Const ConfigSettingScaleUnitWidth       As String = "&ScaleUnitWidth"
Private Const ConfigSettingHeight               As String = "&Height"
Private Const ConfigSettingWidth                As String = "&Width"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mWidth As Double
Private mHeight As Double

Private mScaleUnitWidth As ScaleUnits
Private mScaleUnitHeight As ScaleUnits

Private mConfig As ConfigurationSection

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' xxxx Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

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


mConfig.SetSetting ConfigSettingWidth, mWidth
mConfig.SetSetting ConfigSettingHeight, mHeight
mConfig.SetSetting ConfigSettingScaleUnitWidth, mScaleUnitWidth
mConfig.SetSetting ConfigSettingScaleUnitHeight, mScaleUnitHeight

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get ScaleUnitHeight() As ScaleUnits
ScaleUnitHeight = mScaleUnitHeight
End Property

Public Property Get ScaleUnitWidth() As ScaleUnits
ScaleUnitWidth = mScaleUnitWidth
End Property

Public Property Get Width() As Double
Width = mWidth
End Property

Public Property Get Height() As Double
Height = mHeight
End Property

'================================================================================
' Methods
'================================================================================

Public Function Equals(ByVal aDim As size) As Boolean
Const ProcName As String = "Equals"
Dim failpoint As String
On Error GoTo Err

If aDim Is Nothing Then Exit Function
If aDim Is Me Then
    Equals = True
    Exit Function
End If
With aDim
    If .ScaleUnitWidth <> mScaleUnitWidth Then Exit Function
    If .ScaleUnitHeight <> mScaleUnitHeight Then Exit Function
    If .Width <> mWidth Then Exit Function
    If .Height <> mHeight Then Exit Function
End With
Equals = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pWidth As Double, _
                ByVal pHeight As Double, _
                ByVal pScaleUnitWidth As ScaleUnits, _
                ByVal pScaleUnitHeight As ScaleUnits)
mWidth = pWidth
mHeight = pHeight
mScaleUnitWidth = pScaleUnitWidth
mScaleUnitHeight = pScaleUnitHeight
End Sub

Friend Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)

Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
On Error GoTo Err

Set mConfig = pConfig
mHeight = mConfig.GetSetting(ConfigSettingHeight, "0")
mWidth = mConfig.GetSetting(ConfigSettingWidth, "0")
mScaleUnitHeight = mConfig.GetSetting(ConfigSettingScaleUnitHeight, "0")
mScaleUnitWidth = mConfig.GetSetting(ConfigSettingScaleUnitWidth, "0")

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

Public Function WidthLogical( _
                ByVal pGraphics As Graphics) As Double
Const ProcName As String = "WidthLogical"
Dim failpoint As String
On Error GoTo Err

Select Case mScaleUnitWidth
Case ScaleUnitLogical
    WidthLogical = mWidth
Case ScaleUnitRelative
    WidthLogical = pGraphics.ConvertRelativeToLogicalX(mWidth)
Case ScaleUnitCm
    WidthLogical = pGraphics.ConvertDistanceToLogicalX(mWidth)
Case ScaleUnitPixels
    WidthLogical = pGraphics.ConvertPixelsToLogicalX(mWidth)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function HeightLogical( _
                ByVal pGraphics As Graphics) As Double
Const ProcName As String = "HeightLogical"
Dim failpoint As String
On Error GoTo Err

Select Case mScaleUnitHeight
Case ScaleUnitLogical
    HeightLogical = mHeight
Case ScaleUnitRelative
    HeightLogical = pGraphics.ConvertRelativeToLogicalY(mHeight)
Case ScaleUnitCm
    HeightLogical = pGraphics.ConvertDistanceToLogicalY(mHeight)
Case ScaleUnitPixels
    HeightLogical = pGraphics.ConvertPixelsToLogicalY(mHeight)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'================================================================================
' Helper Functions
'================================================================================





