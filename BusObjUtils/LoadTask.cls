VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LoadTask"
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

Implements Task

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

Private Const ModuleName                    As String = "LoadTask"

'@================================================================================
' Member variables
'@================================================================================

Private mFactory                            As DataObjectFactory

Private WithEvents mConnector               As Connector
Attribute mConnector.VB_VarHelpID = -1
Private WithEvents mConnection              As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
Private WithEvents mRs                      As ADODB.Recordset
Attribute mRs.VB_VarHelpID = -1

Private mConnectionString                   As String
Private mSQL                                As String

Private mTaskContext                        As TaskContext

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
Debug.Print "LoadTask terminated"
End Sub

'@================================================================================
' Task Interface Members
'@================================================================================

Private Sub Task_Cancel()
Cancel
End Sub

Private Sub Task_run()

Const ProcName As String = "Task_run"
Dim failpoint As String
On Error GoTo Err

If mTaskContext.CancelPending Then
    Cancel
    Exit Sub
End If

If mConnection Is Nothing Then
    Debug.Print "LoadTask: Connect to DB"
    Set mConnector = New Connector
    mConnector.Initialise mConnectionString, adUseClient, adAsyncConnect
    mConnector.Connect
    mTaskContext.Suspend -1
Else
    If Not mRs.EOF Then
        If mRs.RecordCount = 1 Then
            mTaskContext.Finish mFactory.MakeNewFromRecordset(mRs), False
        Else
            mTaskContext.Error ErrorCodes.ErrIllegalStateException, _
                                "More than 1 such data object exists", _
                                ProjectName & "." & ModuleName & ":" & "mConnection_FetchComplete"
        End If
    Else
        mTaskContext.Finish Empty, False
    End If
    
    Set mRs = Nothing
    
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let Task_TaskContext( _
                ByVal value As TaskContext)
Set mTaskContext = value
End Property

Private Property Get task_TaskName() As String
task_TaskName = mTaskContext.Name
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
    Dim lErrMsg As String
    lErrMsg = gGenerateErrorMessage(pError) & gGenerateConnectionErrorMessages(pConnection)

    gLogger.Log "Fetch error " & pError.Number & vbCrLf & lErrMsg, _
                ProcName, ModuleName, LogLevelNormal
    
    If pError.Number = 3712 Then
        ' 'Operation has been cancelled by user' - retry
        mConnector.Connect
    Else
        mTaskContext.Error pError.Number, lErrMsg, ProjectName & "." & ModuleName & ":" & ProcName
        If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    End If
Else
    If Not mRs Is Nothing Then Set mRs.ActiveConnection = Nothing
    Debug.Print "LoadTask: close connection"
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
mTaskContext.Error errorCode, errorDesc
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
    Dim lErrMsg As String
    lErrMsg = gGenerateErrorMessage(pError)

    gLogger.Log "Fetch error " & pError.Number & vbCrLf & lErrMsg, _
                ProcName, ModuleName, LogLevelNormal
    
    If pError.Number = 3712 Then
        ' 'Operation has been cancelled by user' - retry
        mConnector.Connect
    Else
        mTaskContext.Error pError.Number, lErrMsg, ProjectName & "." & ModuleName & ":" & ProcName
        If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    End If
    Set mRs = Nothing
    Exit Sub
End If

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
If Err.Number = 3712 Then Exit Sub
gHandleUnexpectedError ProcName, ModuleName
End Sub
