VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

Event PositionChanged()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "Sun"

'@================================================================================
' Member variables
'@================================================================================

Private mSun                                        As Ellipse
Private WithEvents mSunTimer                        As IntervalTimer
Attribute mSunTimer.VB_VarHelpID = -1

Private mOrbitCentre                                As Point
Private mOrbitSize                                  As Size
Private mStartAnomaly                               As Double

Private mStartTime                                  As Date

Private mGraphics                                   As Graphics

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mSunTimer Event Handlers
'@================================================================================

Private Sub mSunTimer_TimerExpired(ev As TWUtilities40.TimerExpiredEventData)
calcPosition mStartAnomaly - ((2.5 * Pi / 180#) * ((GetTimestamp - mStartTime) * 86400#))
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get OrbitSize() As Size
Set OrbitSize = mOrbitSize
End Property

Public Property Get Position() As Point
Set Position = mSun.Position
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Finish()
mSunTimer.StopTimer
End Sub

Public Sub Initialise( _
                ByVal pSun As Ellipse, _
                ByVal pOrbitCentre As Point, _
                ByVal pOrbitSize As Size, _
                ByVal pStartingAnomaly As Double, _
                ByVal pGraphics As Graphics)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mGraphics = pGraphics

Set mSun = pSun
mSun.Brush = CreateBrush(vbWhite)
mSun.Layer = LayerBackground + 1
mSun.Size = NewSize(1#, 1#, ScaleUnitCm, ScaleUnitCm)

Set mOrbitCentre = pOrbitCentre
Set mOrbitSize = pOrbitSize
mStartAnomaly = pStartingAnomaly

mStartTime = Now
calcPosition mStartAnomaly

Set mSunTimer = CreateIntervalTimer(100, ExpiryTimeUnitMilliseconds, 100)
mSunTimer.StartTimer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub calcPosition(ByVal pAnomaly As Double)
Const ProcName As String = "calcPosition"
On Error GoTo Err

mSun.Position = NewPoint( _
    mOrbitCentre.XLogical(mGraphics) + mOrbitSize.WidthLogical(mGraphics) / 2# * Cos(pAnomaly), _
    mOrbitCentre.YLogical(mGraphics) + mOrbitSize.HeightLogical(mGraphics) / 2# * Sin(pAnomaly))

RaiseEvent PositionChanged
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


