VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BasicLogFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' A <code>LogFormatter</code> that provides a simple text conversion of a log record.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements ILogFormatter

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


Private Const ModuleName                    As String = "BasicLogFormatter"

'@================================================================================
' Member variables
'@================================================================================

Private mTimestampFormat As TimestampFormats
Private mIncludeInfoType As Boolean
Private mIncludeTimestamp As Boolean
Private mIncludeLogLevel As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' LogFormatter Interface Members
'@================================================================================

Private Function ILogFormatter_FormatRecord(ByVal Logrec As LogRecord) As String
Const ProcName As String = "LogFormatter_FormatRecord"
On Error GoTo Err

Dim lPrefix As String
If mIncludeTimestamp Then lPrefix = gFormatTimestamp(Logrec.Timestamp, mTimestampFormat) & " "

If mIncludeLogLevel Then lPrefix = lPrefix & gLogLevelToShortString(Logrec.LogLevel)

If mIncludeInfoType Then lPrefix = lPrefix & Logrec.InfoType & ": "

Dim lData As String
gVariantToString Logrec.Data, lData

If InStr(1, lData, vbCrLf) <> 0 Then
    Dim lSpacer As String
    lSpacer = Space(Len(lPrefix))
    lData = Replace$(lData, vbCrLf, vbCrLf & lSpacer)
End If

ILogFormatter_FormatRecord = lPrefix & lData

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get ILogFormatter_Header() As String

End Property

Private Property Get ILogFormatter_Trailer() As String

End Property

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
                ByVal timestampFormat As TimestampFormats, _
                ByVal includeInfoType As Boolean, _
                ByVal includeTimestamp As Boolean, _
                ByVal includeLogLevel As Boolean)
mTimestampFormat = timestampFormat
mIncludeInfoType = includeInfoType
mIncludeTimestamp = includeTimestamp
mIncludeLogLevel = includeLogLevel
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


