Attribute VB_Name = "GBrush"
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

Private Const ModuleName                            As String = "GBrush"

Private Const ConfigSettingClassName                As String = "&ClassName"

'@================================================================================
' Member variables
'@================================================================================

Private mSolidBrushes                               As New Collection

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateBrush( _
                Optional ByVal pColor As Long = vbBlack) As SolidBrush
Const ProcName As String = "gCreateBrush"
On Error GoTo Err

Set gCreateBrush = New SolidBrush
gCreateBrush.Initialise pColor

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateHatchedBrush( _
                Optional ByVal pColor As Long = vbBlack, _
                Optional ByVal pStyle As HatchStyles = HatchHorizontal) As HatchedBrush
Const ProcName As String = "gCreateHatchedBrush"
On Error GoTo Err

Set gCreateHatchedBrush = New HatchedBrush
gCreateHatchedBrush.Initialise pColor, pStyle

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateRadialGradientBrush( _
                ByRef pGradientOrigin As TPoint, _
                ByRef pCentre As TPoint, _
                ByVal pRadiusX As Double, _
                ByVal pRadiusY As Double, _
                ByVal pPad As Boolean, _
                ByRef pColors() As Long) As RadialGradientBrush
Const ProcName As String = "gCreateRadialGradientBrush"
On Error GoTo Err

Set gCreateRadialGradientBrush = New RadialGradientBrush
gCreateRadialGradientBrush.Initialise pGradientOrigin, pCentre, pRadiusX, pRadiusY, pPad, pColors
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateTextureBrush( _
                ByVal pBitmap As Bitmap) As TextureBrush
Const ProcName As String = "gCreateTextureBrush"
On Error GoTo Err

Set gCreateTextureBrush = New TextureBrush
gCreateTextureBrush.Initialise pBitmap

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateTiledGradientBrush( _
                ByRef pBrushBoundary As TRectangle, _
                ByVal pDirection As LinearGradientDirections, _
                ByVal pPad As Boolean, _
                ByVal pTileMode As TileModes, _
                ByRef pColors() As Long, _
                ByRef pIntensities() As Double, _
                ByRef pPositions() As Double) As TiledGradientBrush
Const ProcName As String = "gCreateTiledGradientBrush"
On Error GoTo Err

Set gCreateTiledGradientBrush = New TiledGradientBrush
gCreateTiledGradientBrush.Initialise pBrushBoundary, pDirection, pPad, pTileMode, pColors, pIntensities, pPositions
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetBrush( _
                Optional ByVal pColor As Long = vbBlack) As SolidBrush
Const ProcName As String = "gGetBrush"
On Error GoTo Err

On Error Resume Next
Set gGetBrush = mSolidBrushes(CStr(pColor))
On Error GoTo Err

If gGetBrush Is Nothing Then
    Set gGetBrush = gCreateBrush(pColor)
    mSolidBrushes.Add gGetBrush, CStr(pColor)
End If
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gLoadBrushFromConfig( _
                ByVal pConfig As ConfigurationSection) As IBrush
Const ProcName As String = "gLoadBrushFromConfig"
On Error GoTo Err

Set gLoadBrushFromConfig = CreateObject(pConfig.GetSetting(ConfigSettingClassName))

If TypeOf gLoadBrushFromConfig Is HatchedBrush Then
    Dim lHatchedBrush As HatchedBrush
    Set lHatchedBrush = gLoadBrushFromConfig
    lHatchedBrush.LoadFromConfig pConfig
ElseIf TypeOf gLoadBrushFromConfig Is RadialGradientBrush Then
    Dim lRGBrush As RadialGradientBrush
    Set lRGBrush = gLoadBrushFromConfig
    lRGBrush.LoadFromConfig pConfig
ElseIf TypeOf gLoadBrushFromConfig Is SolidBrush Then
    Dim lSolidBrush As SolidBrush
    Set lSolidBrush = gLoadBrushFromConfig
    lSolidBrush.LoadFromConfig pConfig
ElseIf TypeOf gLoadBrushFromConfig Is TextureBrush Then
    Dim lTextureBrush As TextureBrush
    Set lTextureBrush = gLoadBrushFromConfig
    lTextureBrush.LoadFromConfig pConfig
ElseIf TypeOf gLoadBrushFromConfig Is TiledGradientBrush Then
    Dim lTGBrush As TiledGradientBrush
    Set lTGBrush = gLoadBrushFromConfig
    lTGBrush.LoadFromConfig pConfig
Else
    Assert False, "Invalid brush class name setting"
End If
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gSetBrushClassInConfig( _
                ByVal pBrush As IBrush, _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "gSetBrushClassInConfig"
On Error GoTo Err

pConfig.SetSetting ConfigSettingClassName, TypeName(pBrush)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




