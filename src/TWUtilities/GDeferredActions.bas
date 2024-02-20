Attribute VB_Name = "GDeferredActions"
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


Private Const ModuleName                    As String = "GDeferredActions"

'@================================================================================
' Member variables
'@================================================================================

Private mDeferredActionManager              As New DeferredActionManager

Private mDeferredActions                    As New Collection

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

Public Property Get DeferredActionManager() As DeferredActionManager
Set DeferredActionManager = mDeferredActionManager
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub InitiateDeferredAction( _
                ByRef dae As DeferredActionEntry)
Const ProcName As String = "InitiateDeferredAction"
On Error GoTo Err

Dim index As Long: index = CLng(Rnd * &H7FFFFFFF)
mDeferredActions.Add dae, CStr(index)
'Debug.Print "GDeferredActions::InitiateDeferredAction: " & index
gPostUserMessage UserMessageExecuteDeferredAction, index, 0

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Public Sub RunDeferredAction( _
                ByVal pIndex As Long)
Const ProcName As String = "RunDeferredAction"
On Error GoTo Err

Dim lKey As String: lKey = CStr(pIndex)

Dim dae As DeferredActionEntry
dae = mDeferredActions.Item(lKey)

mDeferredActions.Remove lKey

dae.Action.Run dae.Data

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


