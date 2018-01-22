VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IProgramErrorListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
'
' This interface is implemented by classes that override default error handling
' within other software that supports this.
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

Private Const ModuleName                            As String = "IProgramErrorListener"

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

'@================================================================================
' Methods
'@================================================================================

''
' Called by the Source object to indicate the occurrence of an unexpected program
' error.
'
' @remarks
'   The details of the error are conveyed in the <code>ev</code> parameter.
'
' @param ev
'   A <code>ErrorEventData</code> instance detailing the Source and nature
'   of the error.
'
'@/
Public Sub NotifyUnexpectedProgramError( _
                ByRef ev As ErrorEventData)
End Sub

''
' Called by the Source object to indicate the occurrence of an unhandled program
' error.
'
' @remarks
'   An unhandled program error is a program error in an event handler that, if not
'   otherwise handled, will result in VB's default error handling Action, which is to
'   display a message box with minimal information about the error: closing the message
'   box causes the program to be terminated.
'
'   The details of the error are conveyed in the <code>ev</code> parameter.
'
' @param ev
'   A <code>ErrorEventData</code> instance detailing the Source and nature
'   of the error.
'
'@/
Public Sub NotifyUnhandledProgramError( _
                ByRef ev As ErrorEventData)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


