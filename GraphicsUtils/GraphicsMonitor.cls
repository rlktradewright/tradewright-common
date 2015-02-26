VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicsMonitor"
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

Private Const ModuleName                            As String = "GraphicsMonitor"

'@================================================================================
' Member variables
'@================================================================================

Private mGraphics                                   As Graphics
Private mPrevGaugeX                                 As Double
Private mPrevGaugeY                                 As Double
Private mPrevBoundary                               As TRectangle

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

Friend Sub Initialise( _
                ByVal pGraphics As Graphics)
Set mGraphics = pGraphics
mPrevGaugeX = mGraphics.GaugeX
mPrevGaugeY = mGraphics.GaugeY
mPrevBoundary = mGraphics.Boundary
End Sub

Public Function GetGraphicsChanges() As GraphicsChangeTypes
GetGraphicsChanges = mGraphics.GetGraphicsChanges(mPrevGaugeX, mPrevGaugeY, mPrevBoundary)
mPrevGaugeX = mGraphics.GaugeX
mPrevGaugeY = mGraphics.GaugeY
mPrevBoundary = mGraphics.Boundary
End Function

'@================================================================================
' Helper Functions
'@================================================================================




