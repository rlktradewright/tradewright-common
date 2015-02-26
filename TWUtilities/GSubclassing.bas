Attribute VB_Name = "GSubclassing"
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

Private Const ModuleName                            As String = "GSubclassing"

'@================================================================================
' Member variables
'@================================================================================

Private mControls                                   As New Collection

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

Public Sub gStartSubclassing(ByVal pControl As ISubclassable)
If pControl.PrevWindowProcAddress <> 0 Then Exit Sub

mControls.Add ObjPtr(pControl), CStr(pControl.hwnd)
pControl.PrevWindowProcAddress = SetWindowLong(pControl.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Public Sub gStopSubclassing(ByVal hwnd As Long)
Dim lControl As ISubclassable
Set lControl = getControlFromHwnd(hwnd)

mControls.Remove CStr(hwnd)

If lControl.PrevWindowProcAddress = 0 Then Exit Sub
SetWindowLong hwnd, GWL_WNDPROC, lControl.PrevWindowProcAddress
lControl.PrevWindowProcAddress = 0

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getControlFromHwnd(hwnd) As ISubclassable
Const ProcName As String = "getControlFromHwnd"
On Error GoTo Err

Dim lControlAddress As Long
lControlAddress = mControls.Item(CStr(hwnd))

Dim lControl As ISubclassable
CopyMemory VarPtr(lControl), VarPtr(lControlAddress), 4

Set getControlFromHwnd = lControl

ZeroMemory VarPtr(lControl), 4

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function WindowProc( _
                ByVal hwnd As Long, _
                ByVal uMsg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long
Const ProcName As String = "WindowProc"
On Error GoTo Err

WindowProc = getControlFromHwnd(hwnd).HandleWindowMessage(hwnd, uMsg, wParam, lParam)

Exit Function
Err:
Dim lErrNumber As Long: lErrNumber = Err.number
Dim lErrDesc As String: lErrDesc = Err.Description
Dim lErrSource As String: lErrSource = Err.Source

gStopSubclassing hwnd
gNotifyUnhandledError ProcName, ModuleName, , lErrNumber, lErrDesc, lErrSource
End Function




