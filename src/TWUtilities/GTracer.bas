Attribute VB_Name = "GTracer"
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

Private Const ModuleName                            As String = "GTracer"

Private Const TraceInfotype                         As String = "$trace"

'@================================================================================
' Member variables
'@================================================================================

Private mTracers                                    As SortedDictionary

' Don't make this Static within gBuildTraceString - my tests show that it's
' faster declared at module level
Private mTokens(10)                                 As String


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

Public Property Get gNullTracer() As Tracer
Const ProcName As String = "gNullTracer"
On Error GoTo Err

Static sTracer As Tracer

If sTracer Is Nothing Then Set sTracer = gGetTracer("")
Set gNullTracer = sTracer

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

' My tests show that using this function is more than three times faster than
' using string concatenation
Public Function gBuildTraceString( _
                ByVal pIndent As Long, _
                ByVal pAction As String, _
                ByVal pProcedureName As String, _
                ByVal pProjectName As String, _
                ByVal pModuleName As String, _
                ByVal pInfo As String) As String
Const ProcName As String = "gBuildTraceString"
On Error GoTo Err

mTokens(0) = Space(pIndent)
mTokens(1) = pAction
mTokens(2) = pProcedureName

If Len(pInfo) <> 0 Then
    mTokens(3) = ": "
Else
    mTokens(3) = ""
End If
mTokens(4) = pInfo

If Len(pProjectName) <> 0 Or Len(pModuleName) <> 0 Then
    mTokens(5) = " ("
Else
    mTokens(5) = ""
End If

mTokens(6) = pProjectName

If Len(pModuleName) <> 0 Then
    mTokens(7) = "."
Else
    mTokens(7) = ""
End If
mTokens(8) = pModuleName

If Len(pProjectName) <> 0 Or Len(pModuleName) <> 0 Then
    mTokens(9) = ")"
Else
    mTokens(9) = ""
End If

mTokens(10) = " Thread " & GetCurrentThreadId

gBuildTraceString = Join(mTokens, "")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gDisableTracing( _
                ByVal pTraceType As String)
Const ProcName As String = "gDisableTracing"
On Error GoTo Err

gLogger.Log "Disabling tracing for: " & IIf(pTraceType = "", "ALL", pTraceType), ProcName, ModuleName
gGetTracer(pTraceType).Enabled = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gEnableTracing( _
                ByVal pTraceType As String)
Const ProcName As String = "gEnableTracing"
On Error GoTo Err

gLogger.Log "Enabling tracing for: " & IIf(pTraceType = "", "ALL", pTraceType), ProcName, ModuleName
gGetTracer(pTraceType).Enabled = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gGetTracer( _
                ByVal pTraceType As String) As Tracer
Const ProcName As String = "gGetTracer"
On Error GoTo Err

pTraceType = normaliseTraceType(pTraceType)

If Not mTracers.TryItem(pTraceType, gGetTracer) Then
    Set gGetTracer = New Tracer
    gGetTracer.Initialise pTraceType
    mTracers.Add gGetTracer, pTraceType
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gInit()
Const ProcName As String = "gInit"
On Error GoTo Err

Set mTracers = New SortedDictionary

gLogManager.GetLoggerEx(TraceInfotype).LogLevel = LogLevelNormal

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function normaliseTraceType(ByVal pTraceType As String) As String
Const ProcName As String = "normaliseTraceType"
On Error GoTo Err

If pTraceType = "" Then
    normaliseTraceType = TraceInfotype
Else
    normaliseTraceType = TraceInfotype & "." & pTraceType
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
