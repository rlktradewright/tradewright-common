VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RectangleDefaultModel"
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

Implements IRectangleModel

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

Private Const ModuleName                            As String = "RectangleDefaultModel"

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
' IRectangleModel Interface Members
'@================================================================================

Private Sub IRectangleModel_AddChangeListener(ByVal pListener As IChangeListener)
gAddChangeListener Me, pListener
End Sub

Private Sub IRectangleModel_ClearOrientation()
gFireChange Me, RectangleOrientationCleared
End Sub

Private Sub IRectangleModel_ClearPosition()
gFireChange Me, RectanglePositionCleared
End Sub

Private Sub IRectangleModel_ClearSize()
gFireChange Me, RectangleSizeCleared
End Sub

Private Property Let IRectangleModel_Orientation(ByVal Value As Double)
mOrientation = Value
gFireChange Me, RectangleOrientationChanged
End Property

Private Property Get IRectangleModel_Orientation() As Double
IRectangleModel_Orientation = mOrientation
End Property

Private Property Let IRectangleModel_Position(ByVal Value As Point)
Const ProcName As String = "IRectangleModel_Position"
On Error GoTo Err

If Value Is Nothing Then
    If Not mPosition Is Nothing Then
        Set mPosition = Nothing
        gFireChange Me, RectanglePositionChanged
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

gFireChange Me, RectanglePositionChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Property Get IRectangleModel_Position() As Point
Set IRectangleModel_Position = mPosition
End Property

Private Sub IRectangleModel_RemoveChangeListener(ByVal pListener As IChangeListener)
gRemoveChangeListener Me, pListener
End Sub

Private Property Let IRectangleModel_Size(ByVal Value As Size)
Set mSize = Value
gFireChange Me, RectangleSizeChanged
End Property

Private Property Get IRectangleModel_Size() As Size
Set IRectangleModel_Size = mSize
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





