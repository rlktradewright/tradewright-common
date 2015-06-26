VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDeferredAction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This interface is implemented by classes to enable an Action to be executed at
' some future time.
'
' Note that no means is provided to detect when a deferred Action has completed,
' or to cancel a deferred Action that has not yet Run.
'
' Deferred actions are not preserved across program execution boundaries. In other
' words, when the program exits any un-executed deferred actions are lost.
'
' To queue a deferred Action for processing, use the <code>DeferAction</code>
' global method.
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


Private Const ModuleName                    As String = "IDeferredAction"

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
' This is the entrypoint for the Action. It is called only once to initiate the
' Action. The Action should Run to completion and then exit from this method.
'
' @param Data
'   The Value that was passed to the <code>DeferAction</code> global method.
'@/
Public Sub Run( _
                ByVal Data As Variant)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================


