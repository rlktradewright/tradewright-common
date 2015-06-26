VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CounterTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'================================================================================
' Description
'================================================================================
'
'

'================================================================================
' Interfaces
'================================================================================

Implements Task

'================================================================================
' Events
'================================================================================

'================================================================================
' Constants
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Types
'================================================================================

'================================================================================
' Member variables
'================================================================================

Private mIndex As Long

Private mRunCounter As Long

Private mCount As Long
Private mTaskContext As TaskContext

Private mPause As Boolean

'================================================================================
' Class Event Handlers
'================================================================================

'================================================================================
' Task Interface Members
'================================================================================

Private Sub Task_Cancel()
mTaskContext.Finish 0, True
End Sub

Private Sub Task_run()
Dim i As Long

If mTaskContext.CancelPending Then
    mTaskContext.Finish 0, True
    Exit Sub
End If

If mPause Then
    mPause = False
    mTaskContext.Suspend 10000
    Exit Sub
End If

For i = 1 To 100000: Next

mRunCounter = mRunCounter + 1
If mRunCounter = 100 Then
    mCount = mCount + 1
    mRunCounter = 0
End If

Form1.CountText(mIndex) = mCount

If mCount Mod 10 = 0 Then
    mTaskContext.Progress mCount / 10, Empty
End If

If mCount = 1000 Then
    mTaskContext.Finish 0, False
End If

End Sub

Private Property Let Task_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

'================================================================================
' XXXX Event Handlers
'================================================================================

'================================================================================
' Properties
'================================================================================

Public Property Let Index(ByVal value As Long)
mIndex = value
End Property

'================================================================================
' Methods
'================================================================================

Public Sub pause()
mPause = True
End Sub

'================================================================================
' Helper Functions
'================================================================================




Private Property Get Task_taskName() As String
Task_taskName = mTaskContext.Name
End Property
