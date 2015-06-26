VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connector"
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

Event Connected( _
                ByVal pConnection As Connection)

Event ConnectFailed( _
                ByVal errorCode As Long, _
                ByVal errorDesc As String)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "Connector"

'@================================================================================
' Member variables
'@================================================================================

Private mConnectionString                   As String

Private WithEvents mConnection              As ADODB.Connection
Attribute mConnection.VB_VarHelpID = -1
Private mConnected                          As Boolean

' need the following because sometimes when running async under heavy Load,
' ADO appears not to notify the connection completion event, so we never fetch
' the bars. So need to retry the connection if we don't get the event within
' a certain time.
Private WithEvents mConnectionTLI           As TimerListItem
Attribute mConnectionTLI.VB_VarHelpID = -1

Private mCursorlocation                     As CursorLocationEnum
Private mOptions                            As ConnectOptionEnum

Private mCancelled                          As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
If Not mConnection Is Nothing Then
    If gIsStateSet(mConnection.State, adStateOpen) Then
        mConnection.Close
    Else
        mConnection.Cancel
    End If
End If
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@===============================================================================
' mConnection Event Handlers
'@===============================================================================

Private Sub mConnection_ConnectComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
Const ProcName As String = "mConnection_ConnectComplete"
Dim failpoint As String
On Error GoTo Err

If mCancelled Then Exit Sub

If Not mConnectionTLI Is Nothing Then mConnectionTLI.Cancel
If adStatus = EventStatusEnum.adStatusErrorsOccurred Then
    gLogger.Log "Connector: connection error occurred" & vbCrLf & _
                                gGenerateErrorMessage(pError) & _
                                gGenerateConnectionErrorMessages(pConnection), ProcName, ModuleName, LogLevelSevere
    RaiseEvent ConnectFailed(pError.Number, _
                            ModuleName & ": connection error " & vbCrLf & _
                            gGenerateErrorMessage(pError) & _
                            gGenerateConnectionErrorMessages(pConnection))
Else
    mConnected = True
    RaiseEvent Connected(mConnection)
End If


Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mConnection_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
Const ProcName As String = "mConnection_InfoMessage"
On Error GoTo Err

gLogger.Log "Connector: InfoMessage event received" & vbCrLf & _
                            gGenerateErrorMessage(pError) & _
                            gGenerateConnectionErrorMessages(mConnection), ProcName, ModuleName, LogLevelDetail

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@===============================================================================
' mConnectionTLI Event Handlers
'@===============================================================================

Private Sub mConnectionTLI_StateChange(ev As StateChangeEventData)
Const ProcName As String = "mConnectionTLI_StateChange"
On Error GoTo Err

If mConnected Or gIsStateSet(mConnection.State, adStateOpen) Then Exit Sub
If ev.State <> TimerListItemStates.TimerListItemStateExpired Then Exit Sub

gLogger.Log "connect timed out: cancelling connection", ProcName, ModuleName, LogLevelNormal

' note that mConnection.Cancel causes mConnection_ConnectComplete to be called
' immediately (ie synchronously)
mCancelled = True
mConnection.Cancel
mCancelled = False

Set mConnection = Nothing

gLogger.Log "reconnecting", ProcName, ModuleName, LogLevelNormal
connectToDB

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Connection() As ADODB.Connection
Const ProcName As String = "Connection"
Dim failpoint As String
On Error GoTo Err

Assert mConnected, "Connection not yet established or connection failed"

Set Connection = mConnection

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Connect()
Const ProcName As String = "Connect"
Dim failpoint As String
On Error GoTo Err

connectToDB

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal connectionString As String, _
                ByVal cursorlocation As CursorLocationEnum, _
                ByVal Options As ConnectOptionEnum)
mConnectionString = connectionString
mCursorlocation = cursorlocation
mOptions = Options
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub connectToDB()
Const ProcName As String = "connectToDB"
Dim failpoint As String
On Error GoTo Err

If Not mConnection Is Nothing Then
    If gIsStateSet(mConnection.State, adStateOpen) Then
        mConnection.Close
    Else
        mConnection.Cancel
    End If
    Set mConnection = Nothing
End If

Set mConnection = New ADODB.Connection

mConnection.cursorlocation = mCursorlocation
mConnection.Open mConnectionString, , , mOptions

If mOptions = adAsyncConnect Then Set mConnectionTLI = GetGlobalTimerList.Add(Empty, ConnectCompletionTimeoutMillisecs, ExpiryTimeUnitMilliseconds)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



