VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphObjSeriesRemovalTask"
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

Implements ITask

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


Private Const ModuleName                    As String = "GraphObjSeriesRemovalTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                        As TaskContext

Private mGraphObjSeries                     As IGraphicObjectSeries
Private mModel                              As LayeredGraphicsModel

Private mRemoveObjectsFromModel             As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "GraphObjSeriesRemovalTask terminated"
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub ITask_cancel()
mTaskContext.Finish Empty, True
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

If mGraphObjSeries.Count <> 0 Then

    ' we remove the latest object first, since later objects may have references to earlier
    ' objects, so that removing the earliest first would mean they can't actually be
    ' terminated
    If mRemoveObjectsFromModel Then mModel.RemoveGraphicObjectEx mGraphObjSeries.Item(mGraphObjSeries.Count), True, False
    
    mGraphObjSeries.Remove mGraphObjSeries.Count
Else
    mGraphObjSeries.Finish
    Set mGraphObjSeries = Nothing
    Set mModel = Nothing
    
    mTaskContext.Finish Empty, False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
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

Friend Sub Initialise( _
                ByVal pGraphObjSeries As IGraphicObjectSeries, _
                ByVal pModel As LayeredGraphicsModel, _
                ByVal pRemoveObjectsFromModel As Boolean)
Set mGraphObjSeries = pGraphObjSeries
If pRemoveObjectsFromModel Then
    Set mModel = pModel
    mRemoveObjectsFromModel = True
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




