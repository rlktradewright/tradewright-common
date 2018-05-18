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

Implements IStateChangeListener

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
' IStateChangeListener Interface Members
'@================================================================================

Private Sub IStateChangeListener_Change(ev As StateChangeEventData)
Const ProcName As String = "IStateChangeListener_Change"
On Error GoTo Err

Dim tli As TimerListItem
Set tli = ev.Source

Select Case ev.State
Case TimerListItemStatePending
Case TimerListItemStateExpired
    Dim dae As DeferredActionEntry
    dae = tli.Data
    GDeferredActions.InitiateDeferredAction dae
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
                ByVal pAction As IDeferredAction, _
                Optional ByVal pData As Variant, _
                Optional ByVal pDeferTime As Variant, _
                Optional ByVal pDeferTimeUnits As ExpiryTimeUnits = ExpiryTimeUnits.ExpiryTimeUnitMilliseconds)
Const ProcName As String = "DeferAction"
On Error GoTo Err

If IsEmpty(pDeferTime) Then pDeferTime = 0
If IsMissing(pDeferTime) Then pDeferTime = 0

If IsMissing(pData) Then pData = Empty

Dim dae As DeferredActionEntry
Set dae.Action = pAction
gSetVariant dae.Data, pData

If pDeferTime = 0 Then
    GDeferredActions.InitiateDeferredAction dae
Else
    gGetTimerList.Add(dae, pDeferTime, pDeferTimeUnits).AddStateChangeListener Me
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

