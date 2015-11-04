VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Listeners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type ListenerEntry
    Listener        As Object
    Next            As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "Listeners"

Private Const NullIndex                             As Long = -1

'@================================================================================
' Member variables
'@================================================================================

Private mListeners()                                As ListenerEntry
Private mListenersNextIndex                         As Long
Private mFirstFreeIndex                             As Long
Private mCount                                      As Long

Private mCurrentListeners()                         As Object
Private mCurrentListenersValid                      As Boolean
Private mCurrentListenersCount                      As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Initialise
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

Public Property Get CurrentListeners() As Object()
Const ProcName As String = "CurrentListeners"
On Error GoTo Err

gAssert mCount <> 0, "No current listeners"

If Not mCurrentListenersValid Then SetCurrentListeners

CurrentListeners = mCurrentListeners

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Add(ByVal pListener As Object)
Const ProcName As String = "Add"
On Error GoTo Err

gAssertArgument Not pListener Is Nothing, "pListener is Nothing"

If findEntry(pListener) <> NullIndex Then Exit Sub

Dim lIndex As Long
If mFirstFreeIndex <> NullIndex Then
    lIndex = mFirstFreeIndex
    mFirstFreeIndex = mListeners(mFirstFreeIndex).Next
Else
    If mListenersNextIndex > UBound(mListeners) Then ReDim Preserve mListeners(2 * (UBound(mListeners) + 1) - 1) As ListenerEntry
    lIndex = mListenersNextIndex
    mListenersNextIndex = mListenersNextIndex + 1
End If
    
Set mListeners(lIndex).Listener = pListener
mCount = mCount + 1
mCurrentListenersValid = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

Initialise

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mCount

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetListener(ByVal pIndex As Long) As Object
Const ProcName As String = "GetListener"
On Error GoTo Err

gAssertArgument pIndex > 0 And pIndex <= mCurrentListenersCount, "Index out of range"
Set GetListener = mCurrentListeners(pIndex - 1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Remove(ByVal pListener As Object)
Const ProcName As String = "Remove"
On Error GoTo Err

gAssertArgument Not pListener Is Nothing, "pListener is Nothing"

Dim i As Long: i = findEntry(pListener)

If i <> NullIndex Then
    Set mListeners(i).Listener = Nothing
    mListeners(i).Next = mFirstFreeIndex
    mFirstFreeIndex = i
    mCount = mCount - 1
    mCurrentListenersValid = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetCurrentListeners()
If mCurrentListenersValid Then Exit Sub
If mCount = 0 Then
    mCurrentListenersCount = 0
    Exit Sub
End If

If UBound(mCurrentListeners) < mCount - 1 Then ReDim mCurrentListeners(mCount - 1) As Object
mCurrentListenersCount = mCount
Dim i As Long
Dim j As Long
For i = 0 To mListenersNextIndex - 1
    If Not mListeners(i).Listener Is Nothing Then
        Set mCurrentListeners(j) = mListeners(i).Listener
        j = j + 1
    End If
Next
mCurrentListenersValid = True
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function findEntry(ByVal pListener As Object) As Long
Const ProcName As String = "findEntry"
On Error GoTo Err

findEntry = NullIndex

Dim i As Long
For i = 0 To mListenersNextIndex - 1
    If mListeners(i).Listener Is pListener Then
        findEntry = i
        Exit For
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub Initialise()
ReDim mListeners(1) As ListenerEntry
ReDim mCurrentListeners(1) As Object
mListenersNextIndex = 0
mFirstFreeIndex = NullIndex
mCurrentListenersValid = False
mCurrentListenersCount = 0
mCount = 0
End Sub

