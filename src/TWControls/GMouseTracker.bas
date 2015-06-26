Attribute VB_Name = "GMouseTracker"
Option Explicit

''
' Description here
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

Private Const ModuleName                    As String = "GMouseTracker"

'@================================================================================
' Member variables
'@================================================================================

Private mTrackers                           As New EnumerableCollection

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

Public Sub gRegisterTracker( _
                ByVal pTracker As MouseTracker)
Const ProcName As String = "gRegisterTracker"
On Error GoTo Err

If pTracker.PreviousWindProc = 0 Then
    pTracker.PreviousWindProc = SetWindowLong(pTracker.hWnd, GWL_WNDPROC, AddressOf WindowProc)
End If

mTrackers.Add pTracker, CStr(pTracker.hWnd)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gTrackHover( _
                ByVal tracker As MouseTracker)
Dim tme As TRACKMOUSEEVENTSTRUCT

Const ProcName As String = "gTrackHover"
On Error GoTo Err

tme = queryTracking(tracker)
tme.hwndTrack = tracker.hWnd
tme.dwFlags = tme.dwFlags Or TME_HOVER
tme.dwHoverTime = HOVER_DEFAULT

If TrackMouseEvent(tme) = 0 Then
    'Debug.Print "TrackMouseEvent failed: " & err.LastDllError
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gTrackLeave( _
                ByVal tracker As MouseTracker)
Dim tme As TRACKMOUSEEVENTSTRUCT

Const ProcName As String = "gTrackLeave"
On Error GoTo Err

tme = queryTracking(tracker)
tme.hwndTrack = tracker.hWnd
tme.dwFlags = tme.dwFlags Or TME_LEAVE

If TrackMouseEvent(tme) = 0 Then
    'Debug.Print "TrackMouseEvent failed: " & err.LastDllError
Else
    'Debug.Print "Tracking mouse leave"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gDeregisterTracker( _
                ByVal pTracker As MouseTracker)
Const ProcName As String = "gDeregisterTracker"
On Error GoTo Err

If pTracker.PreviousWindProc <> 0 Then
    SetWindowLong pTracker.hWnd, GWL_WNDPROC, pTracker.PreviousWindProc
    pTracker.PreviousWindProc = 0
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const ProcName As String = "WindowProc"
On Error GoTo Err

Dim tracker As MouseTracker

Set tracker = mTrackers(CStr(hWnd))
If uMsg = WM_MOUSELEAVE Then
    'Debug.Print "WM_MOUSELEAVE"
    tracker.FireMouseLeave
    mTrackers.Remove CStr(hWnd)
    SetWindowLong tracker.hWnd, GWL_WNDPROC, tracker.PreviousWindProc
ElseIf uMsg = WM_MOUSEHOVER Then
    'Debug.Print "WM_MOUSEHOVER"
    tracker.FireMouseHover
Else
    WindowProc = CallWindowProc(tracker.PreviousWindProc, hWnd, uMsg, wParam, lParam)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function queryTracking( _
                ByVal tracker As MouseTracker) As TRACKMOUSEEVENTSTRUCT
Const ProcName As String = "queryTracking"
On Error GoTo Err

Dim tme As TRACKMOUSEEVENTSTRUCT

tme.cbSize = Len(tme)
tme.hwndTrack = tracker.hWnd
tme.dwFlags = TME_QUERY

If TrackMouseEvent(tme) = 0 Then
    'Debug.Print "TrackMouseEvent failed: " & err.LastDllError
End If

queryTracking = tme

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function



