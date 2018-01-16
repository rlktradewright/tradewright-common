VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tracer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Objects of this class are used to log tracing information (procedure entries
' and exits) to the logging subsystem.
'
' To create a <code>Tracer</code> object, call the global <code>GetTracer</code>
' method. This method has an argument that specifies the <em>trace type</em> that
' will be logged via the returned <code>Tracer</code> object.
'
' The trace type is simply
' a name consisting of components separated by a period, and it governs the log
' infotype used for logging trace items: the log infotype is the trace type prefixed
' with &quot;$trace.&quot;. Note that the trace type can be an empty string (called the
' null trace type), in which case the log infotype is &quot;$trace&quot;.
'
' To obtain trace output in the log, tracing must be enabled for the
' relevant trace type using the global <code>EnableTracing</code> method. Tracing
' may be turned off for a particular trace type by calling the global
' <code>DisableTracing</code> method.
'
' Note that enabling tracing for a given trace type does not also enable it for
' trace types that begin with the same name: for example, enabling
' &quot;mytracetype&quot; does not also enable &quot;mytracetype.anothertracetype&quot;.
' The same applies to disabling tracing.
'
' Use of a non-null trace type enables finer control of what is traced. For example,
' a dll that is used by third parties may use its own trace type so that any tracing
' it contains will not be logged simply because the calling application enables
' tracing for the null trace type.
'
' Warning: tracing imposes a considerable overhead, especially where code is structured
' into small procedures, and can generate large amounts of
' log information. It is recommended that tracing statements be conditionally compiled,
' being excluded from production versions.
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

Private Const ModuleName                            As String = "Tracer"

'@================================================================================
' Member variables
'@================================================================================

Private mEnabled                                    As Boolean
Private mLogger                                     As Logger

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

Friend Property Let Enabled(ByVal value As Boolean)
mEnabled = value
End Property

''
' Indicates whether tracing has been enabled for the trace type of this
' <code>Tracer</code> object.
' @return
'   A <code>Boolean</code> value indicating whether tracing has been enabled for the
'   trace type of this <code>Tracer</code> object.
'@/
Public Property Get Enabled() As Boolean
Enabled = mEnabled
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Writes a trace entry to the logging system recording entry to the specified procedure.
'
' @param pProcedureName
'   The name of the procedure in which the method is called.
' @param pProjectName
'   The project's name.
' @param pModuleName
'   The name of the module which contains the calling procedure.
' @param pInfo
'   Any additional information to be logged, such as argument values.
'@/
Public Sub EnterProcedure( _
                ByVal pProcedureName As String, _
                Optional ByVal pProjectName As String, _
                Optional ByVal pModuleName As String, _
                Optional ByVal pInfo As String)
Dim s As String
If Not mEnabled Then Exit Sub

gBuildTraceString "Enter: ", pProcedureName, pProjectName, pModuleName, pInfo, s
mLogger.Log LogLevelNormal, s
End Sub

''
' Writes a trace entry to the logging system recording exit from the specified procedure.
'
' @param pProcedureName
'   The name of the procedure in which the method is called.
' @param pProjectName
'   The project's name.
' @param pModuleName
'   The name of the module which contains the calling procedure.
' @param pInfo
'   Any additional information to be logged, such as argument values.
'@/
Public Sub ExitProcedure( _
                ByVal pProcedureName As String, _
                Optional ByVal pProjectName As String, _
                Optional ByVal pModuleName As String, _
                Optional ByVal pInfo As String)
Dim s As String
If Not mEnabled Then Exit Sub

gBuildTraceString "Exit:  ", pProcedureName, pProjectName, pModuleName, pInfo, s
mLogger.Log LogLevelNormal, s
End Sub

Friend Sub Initialise( _
                ByVal pLogInfoType As String)
Set mLogger = gLogManager.GetLoggerEx(pLogInfoType)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


