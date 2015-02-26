VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogRecord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Objects of this class are the unit of logged information.
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


Private Const ModuleName                    As String = "LogRecord"

'@================================================================================
' Member variables
'@================================================================================

Private mLogLevel                           As LogLevels
Private mInfoType                           As String
Private mTimestamp                          As Date
Private mSequenceNumber                     As Long
Private mData                               As Variant
Private mSource                             As Variant

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

Friend Property Let Data( _
                ByVal Value As Variant)
gSetVariant mData, Value
End Property

''
' Returns the Data for this Log Record.
'
' @return
'   The Data for this Log Record.
'@/
Public Property Get Data() As Variant
gSetVariant Data, mData
End Property

Friend Property Let InfoType( _
                ByVal Value As String)
mInfoType = Value
End Property

''
' Returns the information type for the <code>data</code> property.
'
' @return
'   The information type for the <code>data</code> property.
'@/
Public Property Get InfoType() As String
InfoType = mInfoType
End Property

Friend Property Let LogLevel( _
                ByVal Value As LogLevels)
mLogLevel = Value
End Property

''
' Returns the log level for this Log Record.
'
' @return
'   The log level for this object.
'@/
Public Property Get LogLevel() As LogLevels
LogLevel = mLogLevel
End Property

Friend Property Let SequenceNumber( _
                ByVal Value As Long)
mSequenceNumber = Value
End Property

''
' Returns the sequence number for this Log Record.
'
' @remarks
'   Sequence numbers are allocated consecutively to log records as they are created.
'
' @return
'   The sequence number for this object.
'@/
Public Property Get SequenceNumber() As Long
SequenceNumber = mSequenceNumber
End Property

Friend Property Let Source( _
                ByVal Value As Variant)
gSetVariant mSource, Value
End Property

''
' Returns information that identifies the Source of this Log Record (this information
' could for example be a reference to an object, or something that uniquely identifies
' an object).
'
' @remarks
'   This Value may be <code>Empty</code> where there is no need to distinguish between
'   log records from different sources.
'
' @return
'   Information identifying the Source of this Log Record.
'@/
Public Property Get Source() As Variant
gSetVariant Source, mSource
End Property

Friend Property Let Timestamp( _
                ByVal Value As Date)
mTimestamp = Value
End Property

''
' Returns the Timestamp for this Log Record.
'
' @return
'   The Timestamp for this object.
'@/
Public Property Get Timestamp() As Date
Timestamp = mTimestamp
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


