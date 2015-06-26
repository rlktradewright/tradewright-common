VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ILogFormatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This interface is implemented by classes that provide services for formatting
' <code>LogRecords</code>s.
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


Private Const ModuleName                    As String = "ILogFormatter"

'@================================================================================
' Member variables
'@================================================================================

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

''
' Returns any Header text that should precede the formatted log records.
'
' @return
'   Any Header text that should precede the formatted log records.
'@/
Public Property Get Header() As String

End Property

''
' Returns any Trailer text that should follow the formatted log records.
'
' @return
'   Any Trailer text that should follow the formatted log records.
'@/
Public Property Get Trailer() As String

End Property

'@================================================================================
' Methods
'@================================================================================

''
' Returns a string resulting from formatting the contents of a
' <code>LogRecord</code> object.
'
' @return
'   The formatted log record.
' @param Logrec
'   The <code>LogRecord</code> object to be formatted.
' @see
'
'@/
Public Function FormatRecord( _
                ByVal Logrec As LogRecord) As String
                
End Function

'@================================================================================
' Helper Functions
'@================================================================================


