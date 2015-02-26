Attribute VB_Name = "GClock"
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


Private Const ModuleName                    As String = "GClock"

'@================================================================================
' Member variables
'@================================================================================

Private mClocks                             As Clocks
Private mLocalClock                         As Clock

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

Public Function gCreateSimulatedClock( _
                ByVal pRate As Single, _
                ByVal pTimezonename As String) As Clock
Const ProcName As String = "gCreateSimulatedClock"
On Error GoTo Err

Set gCreateSimulatedClock = New Clock
gCreateSimulatedClock.Initialise True, pRate, pTimezonename

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetClock( _
                ByVal pTimezonename As String) As Clock

Const ProcName As String = "gGetClock"

On Error GoTo Err

If pTimezonename = "" Then
    Set gGetClock = mLocalClock
Else
    On Error Resume Next
    Set gGetClock = mClocks.Item(pTimezonename)
    On Error GoTo Err
    
    If gGetClock Is Nothing Then
        Set gGetClock = New Clock
        gGetClock.Initialise False, 0, pTimezonename
        mClocks.Add gGetClock, pTimezonename
    End If
End If
mClocks.StartClocks ' start the clocks if they're not already running

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gInit()
Const ProcName As String = "gInit"
On Error GoTo Err

Dim Name As String

Set mClocks = New Clocks
Name = GTimeZone.gGetCurrentTimeZoneName
Set mLocalClock = New Clock
mLocalClock.Initialise False, 0, Name
mClocks.Add mLocalClock, Name

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gResetClocks()
Const ProcName As String = "gResetClocks"
On Error GoTo Err

mClocks.ResetClocks

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


