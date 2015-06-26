VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FormattingLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class is used to provide an easy means of logging information with
' a consistent format.
'
' Log record data is formatted as follows:
'
'   &#91;projectname.modulename:procedurename&#93; message: message-qualifier
'
' where projectname is supplied when the <code>FormattingLogger</code> object
' is created and modulename, procedurename, message and message-qualifier are
' supplied in the call to the <code>Log</code> method.
'
' A <code>FormattingLogger</code> object can only be created by a call to the
' global <code>CreateFormattingLogger</code> method.
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

Private Const ModuleName                            As String = "FormattingLogger"

'@================================================================================
' Member variables
'@================================================================================

Private mLogger                                     As Logger
Private mLogTokens(9)                               As String

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

''
' Registers an object to act as a listener for log information.
'
' @remarks
'   The listener object must implement the <code>LogListener</code>
'   interface.
' @param listener
'   The object which is to act as a listener.
'@/
Public Sub AddLogListener( _
                ByVal pListener As LogListener)
Const ProcName As String = "AddLogListener"
On Error GoTo Err

If mLogger Is Nothing Then Exit Sub
mLogger.AddLogListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pInfoType As String, _
                ByVal pProjectName As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

' note that if this FormattingLogger is create after NotifyUnhandledError
' has been called, then mLogger will be Nothing
Set mLogger = gLogManager.GetLogger(pInfoType)

mLogTokens(0) = "["
mLogTokens(1) = pProjectName
mLogTokens(2) = "."
mLogTokens(4) = ":"
mLogTokens(6) = "] "

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Indicates whether a message of a specified log level would be logged.
'
' @return
'   Returns <code>True</code> if the message would be logged, and <code>False</code>
'   otherwise.
' @param level
'   The relevant log level.
' @see
'
'@/
Public Function IsLoggable( _
                ByVal level As LogLevels) As Boolean
Const ProcName As String = "IsLoggable"
On Error GoTo Err

If mLogger Is Nothing Then Exit Function

IsLoggable = mLogger.IsLoggable(level)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Logs string data.
'
' @remarks
'   The data is logged only if the specified log level is not less than the
'   current <code>LogLevel</code> property of the <code>Logger</code>
'   object for the infotype handled by this object.
'
'   If the data is logged and the <code>logToParent</code> property of the <code>Logger</code>
'   object for the infotype handled by this object is <code>True</code>,
'   then the data is also logged by the parent logger.
'
' @param pMsg
'   The first part of the data to be logged.
' @param pProcName
'   The name of the procedure from which this method is called.
' @param pModName
'   The name of the module containing the procedure from which this method is called.
' @param pLogLevel
'   The log level of the specified Data.
' @param pMsgQualifier
'   The second part of the data to be logged.
'@/
Public Sub Log(ByVal pMsg As String, _
                ByVal pProcName As String, _
                ByVal pModName As String, _
                Optional ByVal pLogLevel As LogLevels = LogLevelNormal, _
                Optional ByVal pMsgQualifier As String = vbNullString)
Const ProcName As String = "Log"
On Error GoTo Err

If mLogger Is Nothing Then Exit Sub

If Not mLogger.IsLoggable(pLogLevel) Then Exit Sub
mLogTokens(3) = pModName
mLogTokens(5) = pProcName
mLogTokens(7) = pMsg
If Len(pMsgQualifier) <> 0 Then
    mLogTokens(8) = ": "
    mLogTokens(9) = pMsgQualifier
Else
    mLogTokens(8) = vbNullString
    mLogTokens(9) = vbNullString
End If

mLogger.Log pLogLevel, Join(mLogTokens, "")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Unregisters a listener.
'
' @param listener
'   The object which is to be unregistered.
'@/
Public Sub RemoveLogListener( _
                ByVal pListener As LogListener)
Const ProcName As String = "RemoveLogListener"
On Error GoTo Err

If mLogger Is Nothing Then Exit Sub
mLogger.RemoveLogListener pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


