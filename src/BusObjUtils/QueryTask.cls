VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryTask"
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

Private Const ModuleName                    As String = "QueryTask"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mQueryExecutor           As QueryExecutor
Attribute mQueryExecutor.VB_VarHelpID = -1

Private WithEvents mRs                      As ADODB.Recordset
Attribute mRs.VB_VarHelpID = -1

Private mConnectionString                   As String
Private mSQL                                As String

Private mTaskContext                        As TaskContext

Private mFieldNames()                       As String

Private mSpecifiers                         As FieldSpecifiers

Private mSummsBuilder                       As DOSummariesBuilder

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "QueryTask terminated"
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

If mTaskContext.CancelPending Then
    Cancel
    Exit Sub
End If
    
If mRs Is Nothing Then
    Set mQueryExecutor = New QueryExecutor
    mQueryExecutor.Query mConnectionString, mSQL
    mTaskContext.Suspend -1
ElseIf mRs.EOF Then
    mRs.Close
    Set mRs = Nothing

    mTaskContext.Finish mSummsBuilder.DataObjectSummaries, False
Else
    mSummsBuilder.Add gBuildSummaryFromRecordset(mRs, mFieldNames, mSpecifiers)
    mRs.MoveNext
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
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

createSummsBuilder
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
                ByVal connectionString As String, _
                ByVal sql As String, _
                ByRef FieldNames() As String, _
                ByVal specifiers As FieldSpecifiers)
mConnectionString = connectionString
mSQL = sql
mFieldNames = FieldNames
Set mSpecifiers = specifiers
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

Private Sub createSummsBuilder()
Dim i As Long
Dim spec As FieldSpecifier
Const ProcName As String = "createSummsBuilder"
On Error GoTo Err

Set mSummsBuilder = New DOSummariesBuilder
For i = 0 To UBound(mFieldNames)
    spec = mSpecifiers(mFieldNames(i))
    mSummsBuilder.AddFieldDetails spec
Next

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub
