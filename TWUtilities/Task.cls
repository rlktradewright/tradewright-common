VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Task"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This interface is implemented by classes containing functionality that is
' potentially time consuming and processor-intensive, and that needs to be executed
' in small chunks to enable other events to be processed.
'
' It may also be implemented by classes performing processing that involves waiting
' for one or more events raised asynchronously by other components. This provides
' common mechanisms for notifying the invoker of the Progress and completion of the
' processing, and for cancelling it before completion.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

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
' Returns the Name for this task.
'
' @return
'   The task's Name.
'@/
Public Property Get TaskName() As String

End Property

''
' Sets the task's <code>TaskContext</code> object.
'
' @remarks
'   This property is set before the first call to the <code>run</code> method.
' @param Value
'   The <code>TaskContext</code> object for this task.
'@/
Public Property Let TaskContext(ByVal Value As TaskContext)

End Property

'@================================================================================
' Methods
'@================================================================================

''
' Requests that the task be cancelled.
' @remarks
'   This method is invoked when the <code>cancelTask</code> method of the task's
'   <code>TaskController</code> object is called, but only if the task is
'   currently suspended or pending. A task that is not suspended or pending should
'   check for a cancellation request by examining the <code>cancelPending</code>
'   property of its <code>TaskContext</code> object during the <code>Run</code> method.
'
' @return
'   description
' @param Name
'   description
' @see
'
'@/
Public Sub Cancel()

End Sub

''
' This is the entrypoint for the task. It is called by the Task Manager at the
' start of each 'chunk' of processing. The task itself must maintain State information
' that enables it to carry on from where it left off at the end of the previous chunk.
'
' @remarks
'   When the task is complete, it must call the <code>TaskContext</code>
'   object's <code>finish</code> method and then exit immediately from the
'   <code>run</code> method.
'@/
Public Sub Run()

End Sub

'@================================================================================
' Helper Functions
'@================================================================================
