VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MouseTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

'@================================================================================
' Events
'@================================================================================

Event MouseHover()

Event MouseLeave()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "MouseTracker"

'@================================================================================
' Member variables
'@================================================================================

Private mHwnd                               As Long
Private mPreviousWindProc                   As Long


'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
On Error Resume Next
gDeregisterTracker Me
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get hWnd() As Long
Const ProcName As String = "hWnd"
On Error GoTo Err

hWnd = mHwnd

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let PreviousWindProc( _
                ByVal Value As Long)
Const ProcName As String = "PreviousWindProc"
On Error GoTo Err

mPreviousWindProc = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get PreviousWindProc() As Long
Const ProcName As String = "PreviousWindProc"
On Error GoTo Err

PreviousWindProc = mPreviousWindProc

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub FireMouseHover()
Const ProcName As String = "FireMouseHover"
On Error GoTo Err

RaiseEvent MouseHover

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FireMouseLeave()
Const ProcName As String = "FireMouseLeave"
On Error GoTo Err

RaiseEvent MouseLeave

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal hWnd As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

mHwnd = hWnd
gRegisterTracker Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub TrackHover()
Const ProcName As String = "TrackHover"
On Error GoTo Err

gTrackHover Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub TrackLeave()
Const ProcName As String = "TrackLeave"
On Error GoTo Err

gTrackLeave Me

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




