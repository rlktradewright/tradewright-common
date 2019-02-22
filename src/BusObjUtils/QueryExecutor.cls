VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "QueryExecutor"
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

'@================================================================================
' Events
'@================================================================================

Event Error(ByRef ev As ErrorEventData)
Event RecordsetRetrieved(ByVal rs As Recordset)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "QueryExecutor"

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

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

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
    Static sRetryCount As Long
    handleError pError, pConnection, sRetryCount
Else
    ' note that for some reason the very first ADO command for some time
    ' results in mRs_FetchComplete occurring before mConnection_ExecuteComplete,
    ' but subsequently it is always the other way around! So it is possible that
    ' mRs may no longer exist at this point
    If Not mRs Is Nothing Then Set mRs.ActiveConnection = Nothing
    
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

Debug.Print ModuleName & ":mConnector_Connected"
Set mConnection = pConnection
Set mRs = New ADODB.Recordset
mRs.cursorlocation = adUseClient
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

Dim ev As ErrorEventData
ev.Source = Me
ev.errorCode = errorCode
ev.ErrorMessage = errorDesc
ev.ErrorSource = ProjectName & "." & ModuleName & ":" & ProcName
RaiseEvent Error(ev)

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
    Static sRetryCount As Long
    handleError pError, mConnection, sRetryCount
Else
    ' don't do the following here because it sometimes gives a failure saying 'operation
    ' not allowed when running asynchronously'
    'Set mRs.ActiveConnection = Nothing
    'Debug.Print "QueryObjectsTask: close connection"
    'If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close
    
    RaiseEvent RecordsetRetrieved(mRs)
End If

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

Friend Sub Cancel()
Const ProcName As String = "Cancel"
Dim failpoint As String
On Error GoTo Err

If mConnection Is Nothing Then
ElseIf gIsStateSet(mConnection.State, adStateOpen) Then
    mConnection.Close
Else
    mConnection.Cancel
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Query( _
                ByVal connectionString As String, _
                ByVal sql As String)
mConnectionString = connectionString
mSQL = sql

Set mConnector = New Connector
mConnector.Initialise mConnectionString, adUseClient, adAsyncConnect
mConnector.Connect

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub handleError( _
                ByVal pError As ADODB.Error, _
                ByVal pConnection As ADODB.Connection, _
                ByRef pRetryCount As Long)
Const ProcName As String = "handleError"
On Error GoTo Err

If pError.Number = ADODB.ErrorValueEnum.adErrOperationCancelled Then
    ' error 3712: this error is "Operation has been cancelled by the user"
    ' This should no longer happen but we'll make this check just in case
    gLogger.Log "Error 3712 returned from ADO: operation cancelled by user: retrying", ProcName, ModuleName
    mConnector.Connect
    Exit Sub
End If

If pError.Number = SqlQueryTimeoutError Then
    pRetryCount = pRetryCount + 1
    If pRetryCount <= 5 Then
        gLogger.Log "Query timeout: retrying", ProcName, ModuleName
        mConnector.Connect
        Exit Sub
    End If
End If

Dim lErrMsg As String
lErrMsg = gGenerateErrorMessage(pError) & gGenerateConnectionErrorMessages(pConnection)

gLogger.Log lErrMsg, ProcName, ModuleName

If gIsStateSet(mConnection.State, adStateOpen) Then mConnection.Close

Dim ev As ErrorEventData
ev.Source = Me
ev.errorCode = pError.Number
ev.ErrorMessage = lErrMsg
ev.ErrorSource = ProjectName & "." & ModuleName & ":" & ProcName
RaiseEvent Error(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


