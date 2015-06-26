VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Point"
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

Private Const ModuleName                As String = "Point"

Private Const ConfigSettingCoordSystemX As String = "&CoordSystemX"
Private Const ConfigSettingCoordSystemY As String = "&CoordSystemY"
Private Const ConfigSettingX            As String = "&X"
Private Const ConfigSettingY            As String = "&Y"
Private Const ConfigSectionOffset       As String = "Offset"

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mX As Double
Private mY As Double
Private mCoordSystemX As CoordinateSystems
Private mCoordSystemY As CoordinateSystems
Private mOffset As size

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


mConfig.SetSetting ConfigSettingX, mX
mConfig.SetSetting ConfigSettingY, mY
mConfig.SetSetting ConfigSettingCoordSystemX, mCoordSystemX
mConfig.SetSetting ConfigSettingCoordSystemY, mCoordSystemY
If Not mOffset Is Nothing Then mOffset.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionOffset)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Public Property Get CoordinateSystemX() As CoordinateSystems
CoordinateSystemX = mCoordSystemX
End Property

Public Property Get CoordinateSystemY() As CoordinateSystems
CoordinateSystemY = mCoordSystemY
End Property

Public Property Get Offset() As size
Set Offset = mOffset
End Property

Public Property Get X() As Double
X = mX
End Property

Public Property Get Y() As Double
Y = mY
End Property

'================================================================================
' Methods
'================================================================================

Public Function ChangeCoordinateSystem( _
                ByVal pCoordSystemX As CoordinateSystems, _
                ByVal pCoordSystemY As CoordinateSystems, _
                ByVal pGraphics As Graphics) As Point
Dim newX As Double
Dim newY As Double

Const ProcName As String = "ChangeCoordinateSystem"
On Error GoTo Err

If pCoordSystemX = Me.CoordinateSystemX And pCoordSystemY = Me.CoordinateSystemY Then
    Set ChangeCoordinateSystem = Me
    Exit Function
End If

newX = gTransformCoordX(Me.X, Me.CoordinateSystemX, pCoordSystemX, pGraphics)
newY = gTransformCoordY(Me.Y, Me.CoordinateSystemY, pCoordSystemY, pGraphics)

Set ChangeCoordinateSystem = gNewPoint(newX, newY, pCoordSystemX, pCoordSystemY)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Equals(ByVal aPoint As Point) As Boolean
Const ProcName As String = "Equals"
Dim failpoint As String
On Error GoTo Err

If aPoint Is Nothing Then Exit Function
If aPoint Is Me Then
    Equals = True
    Exit Function
End If
With aPoint
    If .CoordinateSystemX <> mCoordSystemX Then Exit Function
    If .CoordinateSystemY <> mCoordSystemY Then Exit Function
    If .X <> mX Then Exit Function
    If .Y <> mY Then Exit Function
End With
Equals = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal X As Double, _
                ByVal Y As Double, _
                ByVal coordSystemX As CoordinateSystems, _
                ByVal coordSystemY As CoordinateSystems, _
                ByVal pOffset As size)
mX = X
mY = Y
mCoordSystemX = coordSystemX
mCoordSystemY = coordSystemY
Set mOffset = pOffset
End Sub

Friend Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)

Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
On Error GoTo Err

Set mConfig = pConfig
mX = mConfig.GetSetting(ConfigSettingX, "0")
mY = mConfig.GetSetting(ConfigSettingY, "0")
mCoordSystemX = mConfig.GetSetting(ConfigSettingCoordSystemX, "0")
mCoordSystemY = mConfig.GetSetting(ConfigSettingCoordSystemY, "0")
Set mOffset = gLoadSizeFromConfig(mConfig.AddConfigurationSection(ConfigSectionOffset))
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

Public Function ToTPoint( _
                ByVal pGraphics As Graphics) As TPoint
Const ProcName As String = "ToTPoint"
On Error GoTo Err

ToTPoint.X = XLogical(pGraphics)
ToTPoint.Y = YLogical(pGraphics)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function XLogical( _
                ByVal pGraphics As Graphics) As Double
Const ProcName As String = "XLogical"
Dim failpoint As String
On Error GoTo Err

Select Case Me.CoordinateSystemX
Case CoordsLogical
    XLogical = Me.X
Case CoordsRelative
    XLogical = pGraphics.ConvertRelativeToLogicalX(Me.X) + pGraphics.Boundary.Left
Case CoordsDistance
    XLogical = pGraphics.ConvertDistanceToLogicalX(Me.X) + pGraphics.Boundary.Left
Case CoordsCounterDistance
    XLogical = pGraphics.ConvertDistanceToLogicalX(pGraphics.WidthCm - Me.X) + pGraphics.Boundary.Left
Case CoordsPixels
    XLogical = pGraphics.ConvertPixelsToLogicalX(Me.X) + pGraphics.Boundary.Left
Case CoordsCounterPixels
    XLogical = pGraphics.ConvertPixelsToLogicalX(pGraphics.WidthPixels - Me.X) + pGraphics.Boundary.Left
End Select

If Not Me.Offset Is Nothing Then XLogical = XLogical + Me.Offset.WidthLogical(pGraphics)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Property Get YLogical( _
                ByVal pGraphics As Graphics) As Double
Const ProcName As String = "YLogical"
Dim failpoint As String
On Error GoTo Err

Select Case Me.CoordinateSystemY
Case CoordsLogical
    YLogical = Me.Y
Case CoordsRelative
    YLogical = pGraphics.ConvertRelativeToLogicalY(Me.Y) + pGraphics.Boundary.Bottom
Case CoordsDistance
    YLogical = pGraphics.ConvertDistanceToLogicalY(Me.Y) + pGraphics.Boundary.Bottom
Case CoordsCounterDistance
    YLogical = pGraphics.ConvertDistanceToLogicalY(pGraphics.HeightCm - Me.Y) + pGraphics.Boundary.Bottom
Case CoordsPixels
    YLogical = pGraphics.ConvertPixelsToLogicalY(Me.Y) + pGraphics.Boundary.Left
Case CoordsCounterPixels
    YLogical = pGraphics.ConvertPixelsToLogicalY(pGraphics.HeightPixels - Me.Y) + pGraphics.Boundary.Bottom
End Select

If Not Me.Offset Is Nothing Then YLogical = YLogical + Me.Offset.HeightLogical(pGraphics)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'================================================================================
' Helper Functions
'================================================================================




