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

'''Debug.Print "FutureWaiter::IDeferredAction_Run"


Dim lEv As FutureWaitCompletedEventData

If TypeOf pData Is IFuture Then
    notifyCompletion pData
Else
    gSetVariant lEv, pData
    If lEv.Future Is Nothing Then
        '''Debug.Print "Notify all complete: " & getMyKey
        RaiseEvent WaitAllCompleted(lEv)
    Else
        '''Debug.Print "Notify completion: " & getFutureKey(lEv.Future) & "; data: " & gVariantToString(lEv.ContinuationData)
        RaiseEvent WaitCompleted(lEv)
    End If
End If

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

'''Debug.Print "FutureWaiter::IStateChangeListener_Change"
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

Dim lFutureKey As String: lFutureKey = getFutureKey(pFuture)
Dim s As String
If gFutures.Contains(lFutureKey) Then
    s = "Updated future: " & lFutureKey
Else
    s = "Added future: " & lFutureKey
    gFutures.Add pFuture, lFutureKey
End If

If IsMissing(pContinuationData) Then pContinuationData = Empty
gDatas.Add pContinuationData, lFutureKey

If Not IsEmpty(pContinuationData) Then s = s & "; data: " & gVariantToString(pContinuationData)
'''Debug.Print s

If pFuture.IsPending Then
    pFuture.AddStateChangeListener Me
Else
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

Dim en As Enumerator
Set en = gFutures.EntryEnumeratorFromKey(getMyKey)
Do While en.MoveNext
    Dim lEntry As DictionaryEntry
    Set lEntry = en.Current
    If gStartsWith(lEntry.Key, getMyKey) Then
        lCurrentFutures.Add lEntry.Data
    Else
        Exit Do
    End If
Loop

Dim lFuture As IFuture
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

Dim lEntry As DictionaryEntry
Dim en As Enumerator

Set en = gFutures.EntryEnumeratorFromKey(getMyKey)
Do While en.MoveNext
    Set lEntry = en.Current
    If lEntry.Key = getMyKey Then
        Dim lFuture As IFuture: Set lFuture = lEntry.Data
        lFuture.RemoveStateChangeListener Me
        en.Remove
    Else
        Exit Do
    End If
Loop

Set en = gDatas.EntryEnumeratorFromKey(getMyKey)
Do While en.MoveNext
    Set lEntry = en.Current
    If gStartsWith(lEntry.Key, getMyKey) Then
        en.Remove
    Else
        Exit Do
    End If
Loop

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

Private Function getFutureKey(ByVal pFuture As IFuture) As String
getFutureKey = getMyKey & "|" & gGetObjectKey(pFuture)
End Function

Private Function getMyKey() As String
Static sMyKey As String
If sMyKey = "" Then sMyKey = gGetObjectKey(Me)
getMyKey = sMyKey
End Function

Private Sub notifyCompletion(ByVal pFuture As IFuture)
Const ProcName As String = "notifyCompletion"
On Error GoTo Err

Dim lFutureKey As String: lFutureKey = getFutureKey(pFuture)
'''Debug.Print "FutureWaiter::notifyCompletion: " & lFutureKey

If Not gFutures.Contains(lFutureKey) Then
    '''Debug.Print "Future not found: " & lFutureKey
    Exit Sub
End If

Dim lEv As FutureWaitCompletedEventData
Set lEv.Source = Me
Set lEv.Future = pFuture

'''Debug.Print "  gDatas count: " & gDatas.Count
Dim lDataEntry As DictionaryEntry
Dim lDatasEn As Enumerator
Set lDatasEn = gDatas.EntryEnumerator
Do While lDatasEn.MoveNext
    Set lDataEntry = lDatasEn.Current
    '''Debug.Print "     DataEntry: key: " & lDataEntry.Key & "; data: " & gVariantToString(lDataEntry.Data)
Loop

Set lDatasEn = gDatas.EntryEnumeratorFromKey(lFutureKey)
Do While lDatasEn.MoveNext
    Set lDataEntry = lDatasEn.Current
    If lDataEntry.Key <> lFutureKey Then Exit Do
    gSetVariant lEv.ContinuationData, lDataEntry.Data
    '''Debug.Print "Defer action: notify completion: key: " & lDataEntry.Key & "; data: " & gVariantToString(lDataEntry.Data)
    DeferredActionManager.DeferAction Me, lEv
Loop

removeFuture pFuture, False

If mCleared Then
    mCleared = False
Else
    Set lEv.Future = Nothing
    lEv.ContinuationData = Empty
    Dim lFuturesEn As Enumerator: Set lFuturesEn = gFutures.EntryEnumeratorFromKey(getMyKey)
    If lFuturesEn.MoveNext Then
        Dim lFutureEntry As DictionaryEntry: Set lFutureEntry = lFuturesEn.Current
        If gStartsWith(lFutureEntry.Key, getMyKey) Then
            Exit Sub
        End If
    End If
    '''Debug.Print "Defer action: notify all completed: key: " & getMyKey
    DeferredActionManager.DeferAction Me, lEv
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub removeFuture(ByVal pFuture As IFuture, Optional ByVal pStopListening As Boolean = True)
Const ProcName As String = "removeFuture"
On Error GoTo Err

If pStopListening Then pFuture.RemoveStateChangeListener Me

Dim lFutureKey As String: lFutureKey = getFutureKey(pFuture)
Dim lDatasEn As Enumerator: Set lDatasEn = gDatas.EntryEnumeratorFromKey(lFutureKey)
Do While lDatasEn.MoveNext
    Dim lDataEntry As DictionaryEntry: Set lDataEntry = lDatasEn.Current
    If lDataEntry.Key <> lFutureKey Then Exit Do
    '''Debug.Print "Remove future: " & lFutureKey & "; data: " & gVariantToString(lDataEntry.Data)
    lDatasEn.Remove
Loop

gFutures.Remove lFutureKey

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


