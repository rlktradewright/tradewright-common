Attribute VB_Name = "GGraphics"
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

Private Const ModuleName                            As String = "GGraphics"

'@================================================================================
' Member variables
'@================================================================================

Private mColl As New EnumerableCollection

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

Public Function gRegister(ByVal pGraphics As Graphics, ByVal phWnd As Long) As Long
Const ProcName As String = "gRegister"
On Error GoTo Err

AssertArgument Not mColl.Contains(CStr(phWnd)), "A Graphics object for this window already exists"

mColl.Add pGraphics, CStr(phWnd)
gRegister = SetWindowLong(phWnd, GWL_WNDPROC, AddressOf gWindowProc)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gUnRegister(ByVal phWnd As Long, ByVal pOrigWindProcAddress As Long)
Const ProcName As String = "gUnRegister"
On Error GoTo Err

On Error Resume Next
mColl.Remove CStr(phWnd)
If Err.Number <> 0 Then Exit Sub
On Error GoTo Err

SetWindowLong phWnd, GWL_WNDPROC, pOrigWindProcAddress

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gWindowProc( _
                ByVal hwnd As Long, _
                ByVal iMsg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long
Const ProcName As String = "gWindowProc"

On Error Resume Next

Dim lGraphics As Graphics
Set lGraphics = mColl.Item(CStr(hwnd))
If lGraphics Is Nothing Then Exit Function
gWindowProc = lGraphics.WindowProc(hwnd, iMsg, wParam, lParam)

End Function

'@================================================================================
' Helper Functions
'@================================================================================




