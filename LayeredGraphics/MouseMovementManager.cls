VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MouseMovementManager"
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

Private Const ModuleName                            As String = "MouseMovementManager"

'@================================================================================
' Member variables
'@================================================================================

Private mModel                                      As LayeredGraphicsModel
Private mCurrentObj                                 As IGraphicObject

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

Friend Sub Initialise(ByVal pModel As LayeredGraphicsModel)
Set mModel = pModel
End Sub

Friend Sub MouseMove(pButton As Integer, pShift As Integer, ByVal pX As Single, ByVal pY As Single)
Dim targetObj As IGraphicObject
Const ProcName As String = "MouseMove"
On Error GoTo Err

Set targetObj = mModel.GetFirstVisibleObjectAtPoint(pX, pY)
If mCurrentObj Is Nothing Then
    Set mCurrentObj = targetObj
    If Not mCurrentObj Is Nothing Then
        mCurrentObj.MouseEnter pButton, pShift, pX, pY
    End If
ElseIf Not targetObj Is mCurrentObj Then
    mCurrentObj.MouseLeave pButton, pShift, pX, pY
    Set mCurrentObj = targetObj
    If Not mCurrentObj Is Nothing Then mCurrentObj.MouseEnter pButton, pShift, pX, pY
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




