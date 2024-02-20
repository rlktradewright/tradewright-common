VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryObjectsTask"
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

Private Const ModuleName                    As String = "QueryObjectsTask"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mQueryExecutor           As QueryExecutor
Attribute mQueryExecutor.VB_VarHelpID = -1

Private mRs                                 As ADODB.Recordset
Attribute mRs.VB_VarHelpID = -1

Private mConnectionString                   As String
Private mSQL                                As String

Private mTaskContext                        As TaskContext

Private mFactory                            As DataObjectFactory

Private mDataObjs                           As DataObjects

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "QueryObjectsTask terminated"
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub ITask_Cancel()
Const ProcName As String = "ITask_Cancel"
Dim failpoint As String
On Error GoTo Err

Cancel

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub ITask_run()

Const ProcName As String = "ITask_run"
Dim failpoint As String
On Error GoTo Err

Debug.Print ModuleName & ":ITask_Run"

If mTaskContext.CancelPending Then
    Cancel
    Exit Sub
End If

If mRs Is Nothing Then
    failpoint = "010"
    Set mQueryExecutor = New QueryExecutor
    mQueryExecutor.Query mConnectionString, mSQL
    mTaskContext.Suspend -1
ElseIf mRs.EOF Then
    failpoint = "020"
    mRs.Close
    Set mRs = Nothing
    mTaskContext.Finish mDataObjs, False
Else
    failpoint = "030"
    mDataObjs.Add gBuildDataObjectFromRecordset(mFactory, mRs)
    failpoint = "040"
    mRs.MoveNext
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName, failpoint
End Sub

Private Property Let ITask_TaskContext( _
                ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mQueryExecutor Event Handlers
'@================================================================================

Private Sub mQueryExecutor_Error(ev As TWUtilities40.ErrorEventData)
Const ProcName As String = "mQueryExecutor_Error"
On Error GoTo Err

mTaskContext.Error ev.errorCode, ev.ErrorMessage, ev.ErrorSource

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub mQueryExecutor_RecordsetRetrieved(ByVal rs As ADODB.Recordset)
Const ProcName As String = "mQueryExecutor_RecordsetRetrieved"
On Error GoTo Err

Set mDataObjs = New DataObjects
Set mRs = rs
Set mQueryExecutor = Nothing
mTaskContext.Continue

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal factory As DataObjectFactory, _
                ByVal connectionString As String, _
                ByVal sql As String)
Set mFactory = factory
mConnectionString = connectionString
mSQL = sql
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Cancel()
Const ProcName As String = "Cancel"
Dim failpoint As String
On Error GoTo Err

If Not mRs Is Nothing Then
    If Not gIsStateSet(mRs.State, adStateOpen) Then mRs.Cancel
End If
If Not mQueryExecutor Is Nothing Then mQueryExecutor.Cancel
mTaskContext.Finish Empty, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


