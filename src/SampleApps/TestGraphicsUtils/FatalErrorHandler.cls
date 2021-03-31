VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FatalErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "FatalErrorHandler"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mUnhandledErrorHandler           As UnhandledErrorHandler
Attribute mUnhandledErrorHandler.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mUnhandledErrorHandler = UnhandledErrorHandler
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mUnhandledErrorHandler Event Handlers
'@================================================================================

Private Sub mUnhandledErrorHandler_UnhandledError(ev As TWUtilities30.ErrorEvent)

HandleFatalError

' Tell TWUtilities that we've now handled this unhandled error. Not actually
' needed here because HandleFatalError never returns anyway
UnhandledErrorHandler.Handled = True
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub HandleFatalError()
On Error Resume Next    ' ignore any further errors that might arise

MsgBox "A fatal error has occurred. The program will close when you click the OK button." & vbCrLf & _
        "Please email the log file located at" & vbCrLf & vbCrLf & _
        "     " & DefaultLogFileName(Command) & vbCrLf & vbCrLf & _
        "to support@tradewright.com", _
        vbCritical, _
        "Fatal error"

' At this point, we don't know what state things are in, so it's not feasible to return to
' the caller. All we can do is terminate abruptly.
'
' Note that normally one would use the End statement to terminate a VB6 program abruptly. But
' the TWUtilities component interferes with the End statement's processing and may prevent
' proper shutdown, so we use the TWUtilities component's EndProcess method instead.
'
' However if we are running in the development environment, then we call End because the
' EndProcess method kills the entire development environment as well which can have undesirable
' side effects if other components are also loaded.

If mIsInDev Then
    End
Else
    EndProcess
End If

End Sub



