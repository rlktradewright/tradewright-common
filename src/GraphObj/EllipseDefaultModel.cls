VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EllipseDefaultModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements IEllipseModel

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

Private Const ModuleName                            As String = "EllipseDefaultModel"

'@================================================================================
' Member variables
'@================================================================================

Private mPosition                                   As Point
Private mSize                                       As Size
Private mOrientation                                As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
gRemoveAllChangeListeners Me
End Sub

'@================================================================================
' IEllipseModel Interface Members
'@================================================================================

Private Sub IEllipseModel_AddChangeListener(ByVal pListener As IChangeListener)
gAddChangeListener Me, pListener
End Sub

Private Sub IEllipseModel_ClearOrientation()
gFireChange Me, EllipseOrientationCleared
End Sub

Private Sub IEllipseModel_ClearPosition()
gFireChange Me, EllipsePositionCleared
End Sub

Private Sub IEllipseModel_ClearSize()
gFireChange Me, EllipseSizeCleared
End Sub

Private Property Let IEllipseModel_Orientation(ByVal Value As Double)
mOrientation = Value
gFireChange Me, EllipseOrientationChanged
End Property

Private Property Get IEllipseModel_Orientation() As Double
IEllipseModel_Orientation = mOrientation
End Property

Private Property Let IEllipseModel_Position(ByVal Value As Point)
Const ProcName As String = "IEllipseModel_Position"
On Error GoTo Err

If Value Is Nothing Then
    If Not mPosition Is Nothing Then
        Set mPosition = Nothing
        gFireChange Me, EllipsePositionChanged
    End If
    Exit Property
ElseIf Not mPosition Is Nothing Then
    If Value.Equals(mPosition) Then
        If mPosition.CoordinateSystemX = CoordsLogical And mPosition.CoordinateSystemY = CoordsLogical Then
            Exit Property
        End If
    End If
End If

Set mPosition = Value

gFireChange Me, EllipsePositionChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IEllipseModel_Position() As Point
Set IEllipseModel_Position = mPosition
End Property

Private Sub IEllipseModel_RemoveChangeListener(ByVal pListener As IChangeListener)
gRemoveChangeListener Me, pListener
End Sub

Private Property Let IEllipseModel_Size(ByVal Value As Size)
Set mSize = Value
gFireChange Me, EllipseSizeChanged
End Property

Private Property Get IEllipseModel_Size() As Size
Set IEllipseModel_Size = mSize
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================






