VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FutureWaiter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IDeferredAction
Implements IStateChangeListener

'@================================================================================
' Events
'@================================================================================

Event WaitCompleted(ev As FutureWaitCompletedEventData)
Event WaitAllCompleted(ev As FutureWaitCompletedEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "FutureWaiter"

'@================================================================================
' Member variables
'@================================================================================

Private mFutures                                    As New EnumerableCollection
Private mDatas                                      As New EnumerableCollection
Private mContinuationData                           As Variant

Private mCleared                                    As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' DeferredAction Interface Members
'@================================================================================

Private Sub IDeferredAction_Run(ByVal pData As Variant)
Const ProcName As String = "IDeferredAction_Run"
On Error GoTo Err

notifyCompletion pData

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

If ev.State <> FutureStates.FutureStatePending Then notifyCompletion ev.Source

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let ContinuationData(ByVal Value As Variant)
Const ProcName As String = "ContinuationData"
On Error GoTo Err

gSetVariant mContinuationData, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ContinuationData() As Variant
Const ProcName As String = "ContinuationData"
On Error GoTo Err

gSetVariant ContinuationData, mContinuationData

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add( _
                ByVal pFuture As IFuture, _
                Optional ByVal pContinuationData As Variant)
Const ProcName As String = "Add"
On Error GoTo Err

mFutures.Add pFuture
If Not IsMissing(pContinuationData) Then mDatas.Add pContinuationData, gGetObjectKey(pFuture)
pFuture.AddStateChangeListener Me

If Not pFuture.IsPending Then
    DeferredActionManager.DeferAction Me, pFuture
    Exit Sub
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Cancel()
Const ProcName As String = "Cancel"
On Error GoTo Err

Dim lCurrentFutures As New EnumerableCollection

Dim lFuture As IFuture
For Each lFuture In mFutures
    lCurrentFutures.Add lFuture
Next

For Each lFuture In lCurrentFutures
    lFuture.Cancel
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

mCleared = True

Dim lFuture As IFuture
For Each lFuture In mFutures
    lFuture.RemoveStateChangeListener Me
Next

mFutures.Clear
mDatas.Clear
mContinuationData = Empty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Remove(ByVal pFuture As IFuture)
Const ProcName As String = "Remove"
On Error GoTo Err

removeFuture pFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub notifyCompletion(ByVal pFuture As IFuture)
Const ProcName As String = "notifyCompletion"
On Error GoTo Err

If Not mFutures.Contains(pFuture) Then Exit Sub

Dim lEv As FutureWaitCompletedEventData
Set lEv.Source = Me
Set lEv.Future = pFuture

If mDatas.Contains(gGetObjectKey(pFuture)) Then gSetVariant lEv.ContinuationData, mDatas.Item(gGetObjectKey(pFuture))

removeFuture pFuture, False

RaiseEvent WaitCompleted(lEv)

If mCleared Then
    mCleared = False
Else
    Set lEv.Future = Nothing
    lEv.ContinuationData = Empty
    If mFutures.Count = 0 Then RaiseEvent WaitAllCompleted(lEv)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub removeFuture(ByVal pFuture As IFuture, Optional ByVal pStopListening As Boolean = True)
Const ProcName As String = "removeFuture"
On Error GoTo Err

If pStopListening Then pFuture.RemoveStateChangeListener Me
If mDatas.Contains(gGetObjectKey(pFuture)) Then mDatas.Remove gGetObjectKey(pFuture)
If mFutures.Count <> 0 Then mFutures.Remove pFuture

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


