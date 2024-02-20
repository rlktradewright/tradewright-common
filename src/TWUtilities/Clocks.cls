VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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


Private Const ModuleName                    As String = "Clocks"

'@================================================================================
' Member variables
'@================================================================================

Private mClocks                             As Collection

Private WithEvents mTimer                   As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1

Private mClocksRunning                      As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()

Set mClocks = New Collection

End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mTimer Event Handlers
'@================================================================================

Private Sub mTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mTimer_TimerExpired"
On Error GoTo Err

Static tickCount As Long

tickCount = tickCount + 1

Dim lClock As Clock
For Each lClock In mClocks
    lClock.Tick
Next

Dim lOffset As Long
lOffset = getClockOffset
If lOffset > 20 And lOffset < 980 Then
    'Debug.Print "Restart clock timer: offset=" & lOffset
    startClockTimer lOffset
End If

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, ProjectName
End Sub

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Add( _
                ByVal pClock As Clock, _
                ByVal Key As String)
Const ProcName As String = "Add"

On Error GoTo Err

mClocks.Add pClock, Key

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Item( _
                ByVal Key As String) As Clock
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mClocks(Key)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"

On Error GoTo Err

Set NewEnum = mClocks.[_NewEnum]

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub ResetClocks()
Const ProcName As String = "ResetClocks"
On Error GoTo Err

If Not mClocksRunning Then Exit Sub
startClockTimer getClockOffset

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StartClocks()
Const ProcName As String = "StartClocks"
On Error GoTo Err

If mClocksRunning Then Exit Sub
mClocksRunning = True
startClockTimer getClockOffset

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getClockOffset() As Long
Dim lTimestamp As Date
Dim millisecsSinceMidnight As Long

lTimestamp = TimestampGlobals.gGetTimestampUtc
millisecsSinceMidnight = Int((lTimestamp - Int(lTimestamp)) * 86400000)
getClockOffset = millisecsSinceMidnight - Int(millisecsSinceMidnight / 1000) * 1000
End Function

Private Sub startClockTimer(ByVal pOffset As Long)
Const ProcName As String = "startClockTimer"
On Error GoTo Err

If mTimer Is Nothing Then
    Set mTimer = New IntervalTimer
Else
    mTimer.StopTimer
End If

mTimer.Initialise 1000 - pOffset, 1000, False, Empty
mTimer.StartTimer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
