VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Logger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Objects of this class are used to log information of a particular type
' (commonly known as an 'infotype').
'
' Information type names form a hierarchical namespace, components of a name
' being separated by a period character (ie '.'). The root of the tree is
' an empty string. Information type names are not case sensitive.
'
' Infotype names starting with a '$' character are reserved for use within
' the system. Applications are not able to obtain direct access to the <code>Logger</code>
' object for such infotypes.
'
' There is a single <code>Logger</code> object for each information type, which
' may be obtained by calling the global <code>GetLogger</code> method.
'
' Information logged by a <code>Logger</code> object is, by default, also logged by the
' <code>Logger</code> object for the parent information type, recursively up the
' information type namespace tree.
'
' Each <code>Logger</code> object has a <code>LogLevel</code> property which is a
' value within the <code>LogLevels</code> enum. If it is set to the value
' <code>LogLevelNull</code>, then the <code>Logger</code> object obtains its value
' from its parent <code>Logger</code> object. A value of <code>LogLevelNull</code>
' for the root <code>Logger</code> object has the same effect as <code>LogLevelNone</code>,
' ie it logs nothing.
'
' When a <code>Logger</code> object is created, its initial log level is set to
' the current default log level, as set by the global <code>DefaultLogLevel</code> property.
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


Private Const ModuleName                    As String = "Logger"

'@================================================================================
' Member variables
'@================================================================================

Private mLogLevel                           As LogLevels

Private mParent                             As Logger
Private mLogToParent                        As Boolean

Private mLogListeners                       As Listeners

Private mInfoType                           As String

Private mFinished                           As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mLogListeners = New Listeners
mLogToParent = True

mLogLevel = LogLevelUseDefault
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let InfoType( _
                ByVal Value As String)
mInfoType = Value
End Property

''
' Specifies the logging level for this <code>Logger</code> object.
'
' @param Value
'   The logging level to be used by this <code>Logger</code> object in deciding
'   whether to log.
'@/
Public Property Let LogLevel( _
                ByVal Value As LogLevels)
If mFinished Then Exit Property
mLogLevel = Value
End Property

''
' Returns the logging level for this <code>Logger</code> object.
'
' @param Value
'   The logging level currently used by this <code>Logger</code> object in deciding
'   whether to log.
'@/
Public Property Get LogLevel() As LogLevels
Const ProcName As String = "LogLevel"
On Error GoTo Err

If mFinished Then Exit Property

If mLogLevel = LogLevelNull Then
    If mParent Is Nothing Then
        LogLevel = LogLevels.LogLevelNone
    Else
        LogLevel = mParent.LogLevel
    End If
ElseIf mLogLevel = LogLevelUseDefault Then
    LogLevel = gDefaultLogLevel
Else
    LogLevel = mLogLevel
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Specifies whether this <code>Logger</code> object is to log to its parent
' <code>Logger</code> object.
'
' @param Value
'   If <code>True</code>, this <code>Logger</code> object must log to its parent
'   <code>Logger</code> object.
'@/
Public Property Let LogToParent( _
                ByVal Value As Boolean)
If mFinished Then Exit Property
mLogToParent = Value
End Property

''
' Indicates whether this <code>Logger</code> object logs to its parent
' <code>Logger</code> object.
'
' @return
'   If <code>True</code>, this <code>Logger</code> object logs to its parent
'   <code>Logger</code> object.
'@/
Public Property Get LogToParent() As Boolean
If mFinished Then Exit Property
LogToParent = mLogToParent
End Property

Friend Property Let Parent( _
                ByVal Value As Logger)
Set mParent = Value
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Registers an object to act as a listener for log information.
'
' @remarks
'   The listener object must implement the <code>LogListener</code>
'   interface.
' @param pListener
'   The object which is to act as a listener.
'@/
Public Sub AddLogListener( _
                ByVal pListener As ILogListener)
Const ProcName As String = "AddLogListener"
On Error GoTo Err

If mFinished Then Exit Sub

mLogListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Finish()
Const ProcName As String = "Finish"
On Error GoTo Err

mFinished = True

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mLogListeners.Valid Then
    sInit = True
    sSomeListeners = mLogListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ILogListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Finish
    Next
End If

mLogListeners.Clear

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
' @param pLevel
'   The relevant log level.
' @see
'
'@/
Public Function IsLoggable( _
                ByVal pLevel As LogLevels) As Boolean
Const ProcName As String = "IsLoggable"
On Error GoTo Err

If mFinished Then Exit Function

If Not gIsLogLevelPermittedForApplication(pLevel) Then Exit Function

IsLoggable = (pLevel >= LogLevel)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Logs data of a specified log level.
'
' @remarks
'   The data is logged only if the specified log level is not less than the
'   current <code>LogLevel</code> property.
'
'   If the data is logged and the <code>logToParent</code> property is <code>True</code>,
'   then the data is also logged by the parent logger.
'
'   If the data is of a type that has no string representation, then it may be ignored
'   by some <code>LogListener</code> objects. To ensure that this does not happen:
'   <ul>
'   <li>Avoid logging User Defined Types, unless you can be sure that listeners are
'   prepared to handle them.</li>
'   <li>Only log objects that implement the <code>Stringable</code> interface.</li>
'   </ul>
' @param pLevel
'   The log level of the specified data.
' @param pData
'   The data to be logged. Note that this is a variant, so any Data type
'   that can be held in a variant can be logged (but see the Remarks concerning UDTs and
'   objects).
' @param pSource
'   Information that identifies the Source of this Log Record (this information
'   could for example be a reference to an object, or something that uniquely identifies
'   an object). It need not be supplied where there is no need to distinguish between
'   log records from different sources.
'@/
Public Sub Log( _
                ByVal pLevel As LogLevels, _
                ByVal pData As Variant, _
                Optional ByVal pSource As Variant)
If mFinished Then Exit Sub

On Error Resume Next    ' to avoid problems with errors here being caught by
                        ' lower stack frames which may attempt to log them!

If IsEmpty(pData) Or IsNull(pData) Then Exit Sub
If Not IsLoggable(pLevel) Then Exit Sub
If mLogListeners.Count = 0 And Not LogToParent Then Exit Sub

Dim Logrec As New LogRecord
Logrec.Data = pData
Logrec.InfoType = mInfoType
Logrec.LogLevel = pLevel
Logrec.SequenceNumber = GLogging.gGetLoggingSequenceNum
Logrec.Timestamp = gGetTimestamp
gSetVariant Logrec.Source, pSource

If mLogListeners.Count <> 0 Then fireNotify Logrec
If LogToParent Then mParent.Logrec Logrec


End Sub

Friend Sub Logrec( _
                ByVal rec As LogRecord)
Const ProcName As String = "Logrec"
On Error GoTo Err

If IsLoggable(rec.LogLevel) Then
    fireNotify rec
    If LogToParent Then
        mParent.Logrec rec
    End If
End If

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
                ByVal pListener As ILogListener)
Const ProcName As String = "RemoveLogListener"
On Error GoTo Err

If mFinished Then Exit Sub

mLogListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireNotify( _
                ByVal Logrec As LogRecord)
Const ProcName As String = "fireNotify"
On Error GoTo Err

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mLogListeners.Valid Then
    sInit = True
    sSomeListeners = mLogListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ILogListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Notify Logrec
    Next
End If

Exit Sub

Err:
' catch and log errors occurring in the LogListener but don't reraise them.
' Since logging the error will probably invoke a different LogListener, it
' probably won't cause another error, and if it does, it will be swallowed at
' a lower level
gHandleUnexpectedError ProcName, ModuleName, , False, True
End Sub
                

