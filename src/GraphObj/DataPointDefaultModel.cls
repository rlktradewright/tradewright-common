VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataPointDefaultModel"
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

' TODO: replace DataPoint with the related graphic object's class name
Implements IDataPointModel

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

' TODO: replace DataPoint with the related graphic object's class name
Private Const ModuleName                            As String = "DataPointDefaultModel"

'@================================================================================
' Member variables
'@================================================================================

Private mPosition                                   As Point
Private mSize                                       As Size
Private mOrientation                                As Double
Private mValue                                      As Double

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IDataPointModel Interface Members
'@================================================================================

Private Sub IDataPointModel_AddChangeListener(ByVal pListener As IChangeListener)
gAddChangeListener Me, pListener
End Sub

Private Sub IDataPointModel_ClearOrientation()
gFireChange Me, DataPointOrientationCleared
End Sub

Private Sub IDataPointModel_ClearPosition()
gFireChange Me, DataPointPositionCleared
End Sub

Private Sub IDataPointModel_ClearSize()
gFireChange Me, DataPointSizeCleared
End Sub

Private Property Let IDataPointModel_Orientation(ByVal Value As Double)
mOrientation = Value
gFireChange Me, DataPointOrientationChanged
End Property

Private Property Get IDataPointModel_Orientation() As Double
IDataPointModel_Orientation = mOrientation
End Property

Private Property Let IDataPointModel_Position(ByVal Value As Point)
Const ProcName As String = "IDataPointModel_Position"
On Error GoTo Err

If Value Is Nothing Then
    If Not mPosition Is Nothing Then
        Set mPosition = Nothing
        gFireChange Me, DataPointPositionChanged
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

gFireChange Me, DataPointPositionChanged

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Private Property Get IDataPointModel_Position() As Point
Set IDataPointModel_Position = mPosition
End Property

Private Sub IDataPointModel_RemoveChangeListener(ByVal pListener As IChangeListener)
gRemoveChangeListener Me, pListener
End Sub

Private Property Let IDataPointModel_Size(ByVal Value As Size)
Set mSize = Value
gFireChange Me, DataPointSizeChanged
End Property

Private Property Get IDataPointModel_Size() As Size
Set IDataPointModel_Size = mSize
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






Public Property Let IDataPointModel_Value(ByVal Value As Double)
mValue = Value
gFireChange Me, DataPointValueChanged
End Property

Public Property Get IDataPointModel_Value() As Double
IDataPointModel_Value = mValue
End Property

