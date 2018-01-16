VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnhandledErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class provides a mechanism for ActiveX Dlls and ActiveX Controls to notify
' the application of unhandled errors occurring within them. These are typically errors that
' occur during event handlers, which would otherwise call the dll to be unloaded
' and the application to fail untidily.
'
' Only a single instance of this class exists: it can only be accessed via the
' global <code>UnhandledErrorHandler</code> property.
'
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

''
' Raised when an ActiveX DLL or ActiveX Control (or the application itself) notifies
' an error that it is unable to handle. The application should handle this event
' by:
'   <ol>
'   <li>taking appropriate Action (usually this means initiating a tidy closedown);</li>
'   <li><b>AND</b> setting the <code>UnhandledErrorObject</code'
'       <code>Handled</code> property to <code>True</code>.</li>
'   </ol>
'
' @remarks
'   If the application does not set the <code>Handled</code> property to <code>True</code>,
'   then the unhandled error is re-raised to ensure that the error does not
'   go unnoticed.
' @param ev
'   contains details of the unhandled error.
'
'@/
Event UnhandledError(ev As ErrorEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "UnhandledErrorHandler"

'@================================================================================
' Member variables
'@================================================================================

Private mListeners As Listeners

' indicates that the application has handled an unhandled
' event that has been notified to it. Note that this is
' only meaningful during the scope of the <code>Notify</code>
' method, and the application should set it by setting the
' <code>Handled</code> property to <code>True</code> within
' its <code>UnhandledError</code> event handler.
Private mHandled As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mListeners = New Listeners
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

''
' The application must set this property to <code>True</code> before
' exiting from the <code>UnhandledError</code> event handler. Failure
' to do this results in the unhandled error being re-raised, which
' causes in an untidy application failure: this Action is taken to
' ensure that the unhandled error does not go unnoticed.
' @param Value
'   Must be set to <code>True</code>.
' @see
'
'@/
Public Property Let Handled(ByVal Value As Boolean)
mHandled = Value
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Registers an object to act as a listener for <code>UnhandledError</code> events
' raised by this class.
'
' @remarks
'   The listener object must implement the <code>ErrorListener</code>
'   interface.
' @param listener
'   The object which is to act as a listener.
'@/
Public Sub AddErrorListener( _
                ByVal pListener As IErrorListener)
Const ProcName As String = "AddErrorListener"

On Error GoTo Err

mListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Notifies the occurrence of an unhandled error.
'
' @remarks
'   A call to this method causes details of the error to be logged and the
'   <code>UnhandledError</code> event to be raised. The application can handle
'   this event, for example by logging details and then calling the global
'   <code>EndProcess</code> method.
'
'   Note that when running your program in the IDE, your <code>Err</code> object is not shared
'   by the TWUtilities component, and it is therefore advisable to explictly pass the
'   <code>pErrorNumber</code>, <code>pErrorDesc</code> and <code>pErrorSource</code> arguments
'   to ensure that incorrect values are not assumed.
' @param pProcedureName
'   The name of the procedure in which the error was caught.
' @param pProjectName
'   The project's name.
' @param pModuleName
'   The name of the module which contains the calling procedure.
' @param pFailpoint
'   A value that indicates the region of the calling procedure where the error occurred.
' @param pErrorNumber
'   The error number. If this argument is not included in the call to this method, the current
'   value of <code>Err.Number</code> is used.
' @param pErrorDesc
'   The error description. If this argument is not included in the call to this method, the current
'   value of <code>Err.Description</code> is used.
' @param pErrorSource
'   The error Source. If this argument is not included in the call to this method, the current
'   value of <code>Err.Source</code> is used.
'@/
Public Sub Notify( _
                ByVal pProcedureName As String, _
                ByVal pModuleName As String, _
                ByVal pProjectName As String, _
                Optional ByVal pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByVal pErrorDesc As String, _
                Optional ByVal pErrorSource As String)
Const ProcName As String = "Notify"

Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)

Dim ev As ErrorEventData
ev.ErrorCode = IIf(pErrorNumber <> 0, pErrorNumber, Err.number)
ev.ErrorMessage = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
ev.ErrorSource = IIf(errSource <> "", errSource & vbCrLf, "") & _
                    IIf(pProjectName <> "", pProjectName & ".", "") & _
                    IIf(pModuleName <> "", pModuleName & ":", "") & _
                    pProcedureName & _
                    IIf(pFailpoint <> "", " At: " & pFailpoint, "")
Set ev.Source = Me

On Error GoTo Err
gErrorLogger.Log LogLevels.LogLevelSevere, "Error " & ev.ErrorCode & ": " & ev.ErrorMessage & vbCrLf & ev.ErrorSource

gTerminate

fireUnhandledError ev

If Not mHandled Then
    ' the application may not have handled the UnhandledError event, so
    ' terminate the process.
    gEndProcess 1
End If

Exit Sub
    
Err:
gEndProcess 1
End Sub

''
' Unregisters a listener.
'
' @param listener
'   The object which is to be unregistered.
'@/
Public Sub RemoveErrorListener( _
                ByVal pListener As IErrorListener)
Const ProcName As String = "RemoveErrorListener"
On Error GoTo Err

mListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireUnhandledError( _
                ByRef ev As ErrorEventData)
Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mListeners.Valid Then
    sInit = True
    sSomeListeners = mListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IErrorListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Notify ev
    Next
End If

RaiseEvent UnhandledError(ev)

End Sub


