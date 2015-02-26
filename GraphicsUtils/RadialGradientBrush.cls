VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RadialGradientBrush"
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
Implements ISyntheticBrush

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

Private Const ModuleName                            As String = "RadialGradientBrush"

Private Const ConfigSettingCentre                   As String = "&Centre"
Private Const ConfigSettingColors                   As String = "&Colors"
Private Const ConfigSettingPad                      As String = "&Pad"
Private Const ConfigSettingGradientOrigin           As String = "&GradientOrigin"
Private Const ConfigSettingRadiusX                  As String = "&RadiusX"
Private Const ConfigSettingRadiusY                  As String = "&RadiusY"

'@================================================================================
' Member variables
'@================================================================================

Private mCentre                             As TPoint
Private mRadiusX                            As Double
Private mRadiusY                            As Double

Private mBrushRect                          As GDI_RECT

Private mColors()                           As Long

Private mVertices()                         As TRIVERTEX
Private mVerticesIndex                      As Long
Private mMeshTriangles()                    As GRADIENT_TRIANGLE
Private mMeshTrianglesIndex                  As Long

Private mFillBoundary                       As GDI_RECT
Private mPad                                As Boolean

Private mGraphics                           As Graphics

Private mGradientOrigin                     As TPoint

Private mConfig                             As ConfigurationSection

Private mPadBrush                           As IBrush

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

ReDim mColors(1) As Long
mColors(0) = vbRed
mColors(1) = &HFFFF00

mCentre.X = 0.5
mCentre.Y = 0.5

mGradientOrigin.X = 0.5
mGradientOrigin.Y = 0.5

mRadiusX = 0.5
mRadiusY = 0.5

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
' ISyntheticBrush Interface Members
'@================================================================================

Private Sub ISyntheticBrush_Fill( _
                pRect As GDI_RECT, _
                ByVal pClipRegionHandle As Long, _
                ByVal pGraphics As Graphics)
Const ProcName As String = "ISyntheticBrush_Fill"
On Error GoTo Err

Set mGraphics = pGraphics

fillRect pRect

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Colors() As Long()
Colors = mColors
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
mConfig.SetSetting ConfigSettingCentre, TPointToShortString(mCentre)
mConfig.SetSetting ConfigSettingColors, gLongArrayToString(mColors)
mConfig.SetSetting ConfigSettingPad, CStr(mPad)
mConfig.SetSetting ConfigSettingGradientOrigin, TPointToShortString(mGradientOrigin)
mConfig.SetSetting ConfigSettingRadiusX, CStr(mRadiusX)
mConfig.SetSetting ConfigSettingRadiusY, CStr(mRadiusY)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get GradientOrigin() As TPoint
GradientOrigin = mGradientOrigin
End Property

Public Property Get Pad() As Boolean
Pad = mPad
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByRef pGradientOrigin As TPoint, _
                ByRef pCentre As TPoint, _
                ByVal pRadiusX As Double, _
                ByVal pRadiusY As Double, _
                ByVal pPad As Boolean, _
                ByRef pColors() As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

mGradientOrigin = pGradientOrigin
mCentre = pCentre
mRadiusX = pRadiusX
mRadiusY = pRadiusY
mPad = pPad
setColors pColors
If mPad Then Set mPadBrush = gCreateBrush(mColors(UBound(mColors)))
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
mCentre = TPointFromShortString(mConfig.GetSetting(ConfigSettingCentre))
mColors = gLongArrayFromString(mConfig.GetSetting(ConfigSettingColors))
mPad = CBool(mConfig.GetSetting(ConfigSettingPad))
mGradientOrigin = TPointFromShortString(mConfig.GetSetting(ConfigSettingGradientOrigin))
mRadiusX = CDbl(mConfig.GetSetting(ConfigSettingRadiusX))
mRadiusY = CDbl(mConfig.GetSetting(ConfigSettingRadiusY))

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

Private Sub addMeshTriangle( _
                ByVal pVertexIndex1 As Long, _
                ByVal pVertexIndex2 As Long, _
                ByVal pVertexIndex3 As Long)
If mMeshTrianglesIndex > UBound(mMeshTriangles) Then ReDim Preserve mMeshTriangles(2 * (UBound(mMeshTriangles) + 1) - 1) As GRADIENT_TRIANGLE

mMeshTriangles(mMeshTrianglesIndex).Vertex1 = pVertexIndex1
mMeshTriangles(mMeshTrianglesIndex).Vertex2 = pVertexIndex2
mMeshTriangles(mMeshTrianglesIndex).Vertex3 = pVertexIndex3

mMeshTrianglesIndex = mMeshTrianglesIndex + 1
End Sub

Private Sub addVertex( _
                ByVal pX As Long, _
                ByVal pY As Long, _
                ByVal pColor As Long)
                
Const ProcName As String = "addVertex"
On Error GoTo Err

If mVerticesIndex > UBound(mVertices) Then ReDim Preserve mVertices(2 * (UBound(mVertices) + 1) - 1) As TRIVERTEX
mVertices(mVerticesIndex).X = pX
mVertices(mVerticesIndex).Y = pY
mVertices(mVerticesIndex).Red = pColor And &HFF&
mVertices(mVerticesIndex).Green = (pColor And &HFF00&) / &H100&
mVertices(mVerticesIndex).Blue = (pColor And &HFF0000) / &H10000
mVertices(mVerticesIndex).Alpha = 0

mVerticesIndex = mVerticesIndex + 1
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub calculateGradientFillParameters()
Const ProcName As String = "calculateGradientFillParameters"
On Error GoTo Err

If gLogger.IsLoggable(LogLevelHighDetail) Then
    Static et As ElapsedTimer
    If et Is Nothing Then Set et = New ElapsedTimer
    et.StartTiming
End If

If BeginPath(mGraphics.hDC) = 0 Then gHandleWin32Error GetLastError
mGraphics.UsePenHandle ghPathPen
If Ellipse(mGraphics.hDC, mBrushRect.Left, mBrushRect.Top, mBrushRect.Right, mBrushRect.Bottom) = 0 Then gHandleWin32Error GetLastError
mGraphics.ReleasePen
EndPath mGraphics.hDC
 
If FlattenPath(mGraphics.hDC) = 0 Then gHandleWin32Error GetLastError

Dim lNumberOfPoints As Long
ReDim lPoints(0) As GDI_POINT
ReDim ltypes(0) As Byte

lNumberOfPoints = GetPath(mGraphics.hDC, lPoints(0), ltypes(0), 0)
If lNumberOfPoints = 0 Then gHandleWin32Error GetLastError

ReDim lPoints(lNumberOfPoints - 1) As GDI_POINT
ReDim ltypes(lNumberOfPoints - 1) As Byte
If GetPath(mGraphics.hDC, lPoints(0), ltypes(0), lNumberOfPoints) = 0 Then gHandleWin32Error GetLastError

addVertex mFillBoundary.Left + (mFillBoundary.Right - mFillBoundary.Left) * mGradientOrigin.X, _
        mFillBoundary.Bottom - (mFillBoundary.Bottom - mFillBoundary.Top) * mGradientOrigin.Y, _
        mColors(0)
        
Dim i As Long
Dim lVertexNumber
For i = 0 To UBound(lPoints)
    If ltypes(i) = PT_MOVETO Or ltypes(i) = PT_LINETO Then
        lVertexNumber = lVertexNumber + 1
        addVertex lPoints(i).X, lPoints(i).Y, mColors(1)
        If lVertexNumber > 1 Then addMeshTriangle 0, lVertexNumber, lVertexNumber - 1
    End If
Next

addMeshTriangle 0, 1, lVertexNumber

AbortPath mGraphics.hDC

If gLogger.IsLoggable(LogLevelHighDetail) Then
    gLogger.Log "Microseconds", ProcName, ModuleName, LogLevelHighDetail, CStr(Int(et.ElapsedTimeMicroseconds))
End If
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Clear()
ReDim mVertices(3) As TRIVERTEX
mVerticesIndex = 0
ReDim mMeshTriangles(3) As GRADIENT_TRIANGLE
mMeshTrianglesIndex = 0
End Sub

Private Sub doFill()
Const ProcName As String = "doFill"
On Error GoTo Err

If gLogger.IsLoggable(LogLevelHighDetail) Then
    Static et As ElapsedTimer
    If et Is Nothing Then Set et = New ElapsedTimer
    et.StartTiming
End If

If mPad Then mGraphics.FillRectangleGdi mPadBrush, mFillBoundary

If gLogger.IsLoggable(LogLevelHighDetail) Then
    gLogger.Log "Microseconds: Rectangle", ProcName, ModuleName, LogLevelHighDetail, CStr(Int(et.ElapsedTimeMicroseconds))
    et.StartTiming
End If

GradientFill mGraphics.hDC, _
            mVertices(0), _
            mVerticesIndex, _
            VarPtr(mMeshTriangles(0)), _
            mMeshTrianglesIndex, _
            GRADIENT_FILL_TRIANGLE

If gLogger.IsLoggable(LogLevelHighDetail) Then
    gLogger.Log "Microseconds: GradientFill", ProcName, ModuleName, LogLevelHighDetail, CStr(Int(et.ElapsedTimeMicroseconds))
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fillRect( _
                ByRef pRect As GDI_RECT)
Const ProcName As String = "Fill"
On Error GoTo Err

If IsRectEmpty(pRect) Then Exit Sub

If EqualRect(pRect, mFillBoundary) Then
    'Debug.Print "RadialGradientBrush: fill boundary unchanged"
Else
    'Debug.Print "RadialGradientBrush: fill boundary changed"
    Clear
    mFillBoundary = pRect
    setBrushRect
    calculateGradientFillParameters
End If

doFill

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setBrushRect()
Dim lCentre As GDI_POINT

lCentre.X = mFillBoundary.Left + (mFillBoundary.Right - mFillBoundary.Left) * mCentre.X
lCentre.Y = mFillBoundary.Bottom - (mFillBoundary.Bottom - mFillBoundary.Top) * mCentre.Y

mBrushRect.Left = lCentre.X - mRadiusX * (mFillBoundary.Right - mFillBoundary.Left)
mBrushRect.Right = lCentre.X + mRadiusX * (mFillBoundary.Right - mFillBoundary.Left)
mBrushRect.Top = lCentre.Y - mRadiusY * (mFillBoundary.Bottom - mFillBoundary.Top)
mBrushRect.Bottom = lCentre.Y + mRadiusY * (mFillBoundary.Bottom - mFillBoundary.Top)
End Sub

Private Sub setColors(ByRef pColors() As Long)
Const ProcName As String = "setColors"
On Error GoTo Err

On Error Resume Next
Dim l As Long
l = -1
l = UBound(pColors)
On Error GoTo Err

AssertArgument l >= 1, "pColors must have at least two elements"

Dim i As Long
For i = 0 To UBound(pColors)
    AssertArgument gIsValidColor(pColors(i)), "pColor(" & i & ") is not a valid color"
Next

mColors = pColors

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



