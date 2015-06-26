VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeferredActionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements StateChangeListener
Implements ITimerExpiryListener

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


Private Const ModuleName                    As String = "DeferredActionManager"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITimerExpiryListener Interface Members
'@================================================================================

Private Sub ITimerExpiryListener_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "ITimerExpiryListener_TimerExpired"
On Error GoTo Err

Dim dae As DeferredActionEntry
dae = ev.Data
runDeferredAction dae

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' StateChangeListener Interface Members
'@================================================================================

Private Sub StateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "StateChangeListener_Change"
On Error GoTo Err

Dim tli As TimerListItem
Set tli = ev.Source

Select Case ev.State
Case TimerListItemStatePending
Case TimerListItemStateExpired
    Dim dae As DeferredActionEntry
    dae = tli.Data
    runDeferredAction dae
Case TimerListItemStateCancelled
End Select

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

'@================================================================================
' Methods
'@================================================================================

Friend Sub DeferAction( _
                ByVal pAction As DeferredAction, _
                Optional ByVal pData As Variant, _
                Optional ByVal pDeferTime As Variant, _
                Optional ByVal pDeferTimeUnits As ExpiryTimeUnits = ExpiryTimeUnits.ExpiryTimeUnitMilliseconds, _
                Optional ByVal pMayYield As Boolean)
Const ProcName As String = "DeferAction"
On Error GoTo Err

If IsEmpty(pDeferTime) Then pDeferTime = 0
If IsMissing(pDeferTime) Then pDeferTime = 0
If pDeferTime = 0 Then
    pDeferTime = 1
    pDeferTimeUnits = ExpiryTimeUnitMilliseconds
End If

If IsMissing(pData) Then pData = Empty

Dim dae As DeferredActionEntry

Set dae.Action = pAction
gSetVariant dae.Data, pData
If pMayYield Then
    Dim lTimer As IntervalTimer
    Set lTimer = gCreateIntervalTimer(pDeferTime, pDeferTimeUnits, , , dae)
    lTimer.AddTimerExpiryListener Me
    lTimer.StartTimer
Else
    gGetTimerList.Add(dae, pDeferTime, pDeferTimeUnits).AddStateChangeListener Me
End If

'Debug.Print "Deferred Action added: class=" & TypeName(pAction) & ": data=" & TypeName(pData)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub runDeferredAction( _
                ByRef dae As DeferredActionEntry)
Const ProcName As String = "runDeferredAction"
On Error GoTo Err

dae.Action.Run dae.Data

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub
