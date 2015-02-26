Attribute VB_Name = "GTimerList"
Option Explicit

''
' Description here
'
' @remarks
' @see
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


Private Const ModuleName                    As String = "GTimerList"

'@================================================================================
' Member variables
'@================================================================================

Private mRealtimeTimerList                  As TimerList

Private mPrivateTimerLists                  As New EnumerableCollection

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

Public Function gGetSimulatedTimerList(ByVal pClock As Clock) As TimerList
Const ProcName As String = "gGetSimulatedTimerList"
On Error GoTo Err

If mPrivateTimerLists.Contains(gGetObjectKey(pClock)) Then
    Set gGetSimulatedTimerList = mPrivateTimerLists.Item(gGetObjectKey(pClock))
    Exit Function
End If

Set gGetSimulatedTimerList = New TimerList
gGetSimulatedTimerList.Initialise True, pClock

mPrivateTimerLists.Add gGetSimulatedTimerList, gGetObjectKey(pClock)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetTimerList() As TimerList
Set gGetTimerList = mRealtimeTimerList
End Function

Public Sub gInit()
Const ProcName As String = "gInit"
On Error GoTo Err

Set mRealtimeTimerList = New TimerList
mRealtimeTimerList.Initialise False, gGetClock("")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gTerm()
Const ProcName As String = "gTerm"
On Error GoTo Err

If Not mRealtimeTimerList Is Nothing Then
    mRealtimeTimerList.Clear
    Set mRealtimeTimerList = Nothing
End If

Dim tl As TimerList
For Each tl In mPrivateTimerLists
    tl.Clear
Next

mPrivateTimerLists.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


