VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileLogListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' A log listener that writes log records to a specified file.
'
' To create an object of this class, use the global <code>CreateFileLogListener</code>
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


Private Const ModuleName                    As String = "FileLogListener"

'@================================================================================
' Member variables
'@================================================================================

Private mFilename                           As String

Private mTs                                 As TextStream

Private mFormatter                          As LogFormatter

Private mFinished                           As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' LogListener Interface Members
'@================================================================================

Private Sub LogListener_finish()

mFinished = True

If mFormatter.Trailer <> "" Then mTs.WriteLine mFormatter.Trailer
mTs.Close
Set mTs = Nothing
End Sub

Private Sub LogListener_notify(ByVal Logrec As LogRecord)
If mFinished Then Exit Sub
mTs.WriteLine mFormatter.FormatRecord(Logrec)
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
                ByVal pFilename As String, _
                ByVal formatter As LogFormatter, _
                ByVal pOverwrite As Boolean, _
                ByVal pCreateBackup As Boolean, _
                ByVal pUnicode As Boolean, _
                ByVal timestampFormat As TimestampFormats, _
                ByVal includeInfoType As Boolean, _
                ByVal includeTimestamp As Boolean, _
                ByVal includeLogLevel As Boolean)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTs = gCreateWriteableTextFile(pFilename, pOverwrite, pCreateBackup, pUnicode, True)
If formatter Is Nothing Then
    Dim blf As New BasicLogFormatter
    blf.Initialise timestampFormat, includeInfoType, includeTimestamp, includeLogLevel
    Set mFormatter = blf
Else
    Set mFormatter = formatter
End If

If mFormatter.Header <> "" Then mTs.WriteLine mFormatter.Header

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


