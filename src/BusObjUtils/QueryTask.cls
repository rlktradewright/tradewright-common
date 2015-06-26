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

Private WithEvents mConnector               As Connector
Attribute mConnector.VB_VarHelpID = -1
Private WithEvents mConnection              As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
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
    
If mConnection Is Nothing Then
    Debug.Print "QueryTask: Connect to DB"
    Set mConnector = New Connector
    mConnector.Initialise mConnectionString, adUseClient, adAsyncConnect
    mConnector.Connect
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
                ByVal value As TaskContext)
Set mTaskContext = value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'@================================================================================
' mConnection Event Handlers
'@================================================================================

Private Sub mConnection_ExecuteComplete( _
                ByVal RecordsAffected As Long, _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pCommand As ADODB.Command, _
                ByVal pRecordset As ADODB.Recordset, _
                ByVal pConnection As ADODB.Connection)

Const ProcName As String = "mConnection_ExecuteComplete"
Dim failpoint As String
On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    If pError.Number = 3712 Then
        ' this error is "Operation has been cancelled by the user"
        ' This should no longer happen but we'll make this check just in case
        gLogger.Log "Error 3712 returned from ADO: operation cancelled by user: retrying", ProcName, ModuleName
        mConnector.Connect
        Exit Sub
    End If
    
    Dim lErrMsg As String
    lErrMsg = gGenerateErrorMessage(pError) & gGenerateConnectionErrorMessages(pConnection)

    gLogger.Log lErrMsg, ProcName, ModuleName
    
    mTaskContext.Error pError.Number, lErrMsg, ProjectName & "." & ModuleName & ":" & ProcName
    If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
Else
    ' note that for some reason the very first ADO command for some time
    ' results in mRs_FetchComplete occurring before mConnection_ExecuteComplete,
    ' but subsequently it is always the other way around! So it is possible that
    ' mRs may no longer exist at this point
    If Not mRs Is Nothing Then Set mRs.ActiveConnection = Nothing
    Debug.Print "QueryTask: close connection"
    If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mConnector Event Handlers
'@================================================================================

Private Sub mConnector_Connected(ByVal pConnection As ADODB.Connection)
Const ProcName As String = "mConnector_Connected"
Dim failpoint As String
On Error GoTo Err

Set mConnection = pConnection
    Set mRs = New ADODB.Recordset
    Set mRs.ActiveConnection = mConnection
    mRs.Properties("Initial Fetch Size") = 0
    mRs.Open mSQL, _
            , _
            adOpenStatic, _
            adLockBatchOptimistic, _
            ExecuteOptionEnum.adAsyncFetchNonBlocking + _
            ExecuteOptionEnum.adAsyncExecute

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mConnector_ConnectFailed(ByVal errorCode As Long, ByVal errorDesc As String)
Const ProcName As String = "mConnector_ConnectFailed"
Dim failpoint As String
On Error GoTo Err

mTaskContext.Error errorCode, errorDesc

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' mRs Event Handlers
'@================================================================================

Private Sub mRs_FetchComplete( _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)

Const ProcName As String = "mRs_FetchComplete"
Dim failpoint As String

On Error GoTo Err

If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    If pError.Number = 3712 Then
        ' this error is "Operation has been cancelled by the user"
        ' This should no longer happen but we'll make this check just in case
        gLogger.Log "Error 3712 returned from ADO: operation cancelled by user: retrying", ProcName, ModuleName
        mConnector.Connect
        Exit Sub
    End If
    
    Dim lErrMsg As String
    lErrMsg = gGenerateErrorMessage(pError)

    gLogger.Log lErrMsg, ProcName, ModuleName
    
    mTaskContext.Error pError.Number, lErrMsg, ProjectName & "." & ModuleName & ":" & ProcName
    If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    
    Set mRs = Nothing
    Exit Sub
End If

createSummsBuilder

mTaskContext.Continue

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, pFailpoint:=failpoint
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
ElseIf Not mConnection Is Nothing Then
    If gIsStateSet(mConnection.State, adStateOpen) Then
        mConnection.Close
    Else
        mConnection.Cancel
    End If
End If
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
