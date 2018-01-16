Attribute VB_Name = "GLogging"
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "GLogging"

'@================================================================================
' Member variables
'@================================================================================

Private mSeq                                As Long

Public gDefaultLogLevel                     As LogLevels

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

Public Property Get gErrorLogger() As Logger
Static lLogger As Logger
If lLogger Is Nothing Then Set lLogger = gLogManager.GetLogger("error")
Set gErrorLogger = lLogger
End Property

Public Property Get gLogger() As FormattingLogger
Static sLogger As FormattingLogger
If sLogger Is Nothing Then Set sLogger = gCreateFormattingLogger("twutilities.log", ProjectName)
Set gLogger = sLogger
End Property

Public Property Get gLogLogger() As Logger
Static sLogger As Logger
If sLogger Is Nothing Then Set sLogger = gLogManager.GetLogger("log")
Set gLogLogger = sLogger
End Property

Public Property Get gLogManager() As LogManager
Static lLogmanager As LogManager
If lLogmanager Is Nothing Then Set lLogmanager = New LogManager
Set gLogManager = lLogmanager
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateFormattingLogger( _
                ByVal pInfoType As String, _
                ByVal pProjectName As String) As FormattingLogger
Const ProcName As String = "gCreateFormattingLogger"
On Error GoTo Err

Set gCreateFormattingLogger = New FormattingLogger
gCreateFormattingLogger.Initialise pInfoType, pProjectName

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetLoggingSequenceNum() As Long
mSeq = mSeq + 1
gGetLoggingSequenceNum = mSeq
End Function

Public Sub gInit()
Dim lFormattingLogger As FormattingLogger
Set lFormattingLogger = gLogger

Dim lLogger As Logger
Set lLogger = gErrorLogger
End Sub

Public Function gIsLogLevelPermittedForApplication(ByVal Value As LogLevels) As Boolean
Select Case Value
Case LogLevelAll, _
        LogLevelUseDefault, _
        LogLevelNone, _
        LogLevelNull
Case Else
    gIsLogLevelPermittedForApplication = True
End Select
End Function

Public Function gLogLevelFromString( _
                ByVal Value As String) As LogLevels
Const ProcName As String = "gLogLevelFromString"
On Error GoTo Err

Select Case UCase$(Value)
Case "A", "ALL"
    gLogLevelFromString = LogLevels.LogLevelAll
Case "D", "DETAIL"
    gLogLevelFromString = LogLevels.LogLevelDetail
Case "H", "HIGH", "HIGH DETAIL", "HIGHDETAIL"
    gLogLevelFromString = LogLevels.LogLevelHighDetail
Case "I", "INFO"
    gLogLevelFromString = LogLevels.LogLevelInfo
Case "M", "MEDIUM", "MEDIUMDETAIL", "MEDIUM DETAIL"
    gLogLevelFromString = LogLevels.LogLevelMediumDetail
Case "0", "NONE"
    gLogLevelFromString = LogLevels.LogLevelNone
Case "N", "NORMAL"
    gLogLevelFromString = LogLevels.LogLevelNormal
Case "-", "NULL"
    gLogLevelFromString = LogLevels.LogLevelNull
Case "S", "SEVERE"
    gLogLevelFromString = LogLevels.LogLevelSevere
Case "W", "WARNING"
    gLogLevelFromString = LogLevels.LogLevelWarning
Case "U", "DEFAULT"
    gLogLevelFromString = LogLevels.LogLevelUseDefault
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid log level Name"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gLogLevelToShortString( _
                ByVal Value As LogLevels) As String
Select Case Value
Case LogLevels.LogLevelAll
    gLogLevelToShortString = "A "
Case LogLevels.LogLevelDetail
    gLogLevelToShortString = "D "
Case LogLevels.LogLevelHighDetail
    gLogLevelToShortString = "H "
Case LogLevels.LogLevelInfo
    gLogLevelToShortString = "I "
Case LogLevels.LogLevelMediumDetail
    gLogLevelToShortString = "M "
Case LogLevels.LogLevelNone
    gLogLevelToShortString = "0 "
Case LogLevels.LogLevelNormal
    gLogLevelToShortString = "N "
Case LogLevels.LogLevelNull
    gLogLevelToShortString = "- "
Case LogLevels.LogLevelSevere
    gLogLevelToShortString = "S "
Case LogLevels.LogLevelWarning
    gLogLevelToShortString = "W "
Case LogLevels.LogLevelUseDefault
    gLogLevelToShortString = "U "
Case Else
    gLogLevelToShortString = "? "
End Select
End Function

Public Function gLogLevelToString( _
                ByVal Value As LogLevels) As String
Select Case Value
Case LogLevels.LogLevelAll
    gLogLevelToString = "All"
Case LogLevels.LogLevelDetail
    gLogLevelToString = "Detail"
Case LogLevels.LogLevelHighDetail
    gLogLevelToString = "High detail"
Case LogLevels.LogLevelInfo
    gLogLevelToString = "Info"
Case LogLevels.LogLevelMediumDetail
    gLogLevelToString = "Medium detail"
Case LogLevels.LogLevelNone
    gLogLevelToString = "None"
Case LogLevels.LogLevelNormal
    gLogLevelToString = "Normal"
Case LogLevels.LogLevelNull
    gLogLevelToString = "Null"
Case LogLevels.LogLevelSevere
    gLogLevelToString = "Severe"
Case LogLevels.LogLevelWarning
    gLogLevelToString = "Warning"
Case LogLevels.LogLevelUseDefault
    gLogLevelToString = "Default"
Case Else
    gLogLevelToString = CStr(Value)
End Select
End Function

'@================================================================================
' Helper Functions
'@================================================================================


