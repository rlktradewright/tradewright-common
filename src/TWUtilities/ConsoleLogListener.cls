VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConsoleLogListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' A log listener that writes log records to the console.
'
' To create an object of this class, use the global <code>CreateConsoleLogListener</code>
' method.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements LogListener

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


Private Const ModuleName                    As String = "ConsoleLogListener"

'@================================================================================
' Member variables
'@================================================================================

Private mConsole                            As Console

Private mFormatter                          As LogFormatter

Private mFinished                           As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' LogListener Interface Members
'@================================================================================

Private Sub LogListener_finish()

Const ProcName As String = "LogListener_finish"

On Error GoTo Err

mFinished = True

If mFormatter.Trailer <> "" Then mConsole.WriteLineToConsole mFormatter.Trailer
Set mConsole = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub LogListener_notify(ByVal Logrec As LogRecord)
Const ProcName As String = "LogListener_notify"

On Error GoTo Err

If mFinished Then Exit Sub
mConsole.WriteLineToConsole mFormatter.FormatRecord(Logrec)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

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
                ByVal formatter As LogFormatter, _
                ByVal timestampFormat As TimestampFormats, _
                ByVal includeInfoType As Boolean, _
                ByVal includeTimestamp As Boolean, _
                ByVal includeLogLevel As Boolean)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mConsole = gGetConsole
If formatter Is Nothing Then
    Dim blf As New BasicLogFormatter
    blf.Initialise timestampFormat, includeInfoType, includeTimestamp, includeLogLevel
    Set mFormatter = blf
Else
    Set mFormatter = formatter
End If

If mFormatter.Header <> "" Then mConsole.WriteLineToConsole mFormatter.Header

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




