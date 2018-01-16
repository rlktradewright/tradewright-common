VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LogManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Maintains information about a set of logging-related objects.
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

Private Const ModuleName                    As String = "LogManager"

'@================================================================================
' Member variables
'@================================================================================

Private mLoggers                            As SortedDictionary
Private mRootLogger                         As Logger

Private mFinished                           As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mLoggers = New SortedDictionary
mLoggers.Initialise KeyTypeString, False
Set mRootLogger = New Logger
mRootLogger.LogLevel = LogLevelAll
mRootLogger.LogToParent = False
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

'@================================================================================
' Methods
'@================================================================================

Public Function Finish()
mFinished = True

Dim lLogger As Logger
For Each lLogger In mLoggers
    lLogger.Finish
Next

mRootLogger.Finish
End Function

Public Function GetLogger( _
                ByVal pInfoType As String) As Logger
Dim templogger As Logger
Dim ParentLogger As Logger

Const ProcName As String = "GetLogger"
On Error GoTo Err

If mFinished Then Exit Function

If pInfoType = "" Then
    Set GetLogger = mRootLogger
    Exit Function
End If

gAssertArgument Left$(pInfoType, 1) <> "$", "Infotypes starting with $ are reserved for system use"

Set GetLogger = GetLoggerEx(pInfoType)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetLoggerEx( _
                ByVal InfoType As String) As Logger
Dim templogger As Logger
Dim ParentLogger As Logger

Const ProcName As String = "GetLoggerEx"

On Error GoTo Err

If InfoType = "" Then
    Set GetLoggerEx = mRootLogger
    Exit Function
End If

On Error Resume Next
Set GetLoggerEx = mLoggers.Item(InfoType)
On Error GoTo Err

If Not GetLoggerEx Is Nothing Then Exit Function

Set GetLoggerEx = New Logger
GetLoggerEx.InfoType = InfoType
mLoggers.Add GetLoggerEx, InfoType
Set templogger = GetLoggerEx

InfoType = removeElement(InfoType)
Do While InfoType <> ""
    
    Set ParentLogger = Nothing
    On Error Resume Next
    Set ParentLogger = mLoggers.Item(InfoType)
    On Error GoTo Err
    
    If Not ParentLogger Is Nothing Then
        templogger.Parent = ParentLogger
        Exit Function
    End If
    
    Set ParentLogger = New Logger
    ParentLogger.InfoType = InfoType
    mLoggers.Add ParentLogger, InfoType
    
    templogger.Parent = ParentLogger
    
    Set templogger = ParentLogger
    
    InfoType = removeElement(InfoType)

Loop

templogger.Parent = mRootLogger

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

'@================================================================================
' Helper Functions
'@================================================================================


Private Function removeElement( _
                ByVal InfoType As String) As String
Dim l As Long

Const ProcName As String = "removeElement"

On Error GoTo Err

l = InStrRev(InfoType, ".")
If l <> 0 Then
    removeElement = Left$(InfoType, l - 1)
Else
    removeElement = ""
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

