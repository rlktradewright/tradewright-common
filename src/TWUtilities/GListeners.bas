Attribute VB_Name = "GListeners"
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

Public Type ListenerEntry
    Next                As Long
    Listener            As Object
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GListeners"

'@================================================================================
' Member variables
'@================================================================================

Private mListeners()                                As ListenerEntry
Private mNextUnusedIndex                            As Long
Private mFirstFreeIndex                             As Long

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

Public Sub AddListener( _
                ByVal pListener As Object, _
                ByVal pListHeadIndex As Long, _
                ByVal pListTailIndex As Long)
Dim lIndex As Long: lIndex = AllocateIndex
mListeners(lIndex).Next = mListeners(pListHeadIndex).Next
Set mListeners(lIndex).Listener = pListener
mListeners(pListHeadIndex).Next = lIndex
End Sub

Public Sub ClearListeners( _
                ByVal pListHeadIndex As Long, _
                ByVal pListTailIndex As Long)
Dim lIndex As Long: lIndex = mListeners(pListHeadIndex).Next
Dim lNextIndex As Long
Do While lIndex <> pListTailIndex
    lNextIndex = mListeners(lIndex).Next
    releaseIndex lIndex
    lIndex = lNextIndex
Loop
mListeners(pListHeadIndex).Next = pListTailIndex
End Sub

Public Function ContainsListener( _
                ByVal pListener As Object, _
                ByVal pListHeadIndex As Long, _
                ByVal pListTailIndex As Long) As Boolean
Dim lIndex As Long: lIndex = mListeners(pListHeadIndex).Next
Do While lIndex <> pListTailIndex
    If mListeners(lIndex).Listener Is pListener Then
        ContainsListener = True
        Exit Function
    End If
    lIndex = mListeners(lIndex).Next
Loop
End Function

Public Sub GetCurrentListeners( _
                ByRef pListeners() As Object, _
                ByVal pListHeadIndex As Long, _
                ByVal pListTailIndex As Long)
Dim lListener As Object
Dim lIndex As Long: lIndex = mListeners(pListHeadIndex).Next
Dim i As Long
Do While lIndex <> pListTailIndex
    Set pListeners(i) = mListeners(lIndex).Listener
    i = i + 1
    lIndex = mListeners(lIndex).Next
Loop
End Sub

Public Sub IntialiseListenersList( _
                ByRef pListHeadIndex As Long, _
                ByRef pListTailIndex As Long)
pListHeadIndex = AllocateIndex
pListTailIndex = AllocateIndex
mListeners(pListHeadIndex).Next = pListTailIndex
mListeners(pListTailIndex).Next = NullIndex
End Sub

Public Function RemoveListener( _
                ByVal pListener As Object, _
                ByVal pListHeadIndex As Long, _
                ByVal pListTailIndex As Long) As Boolean
Dim lPrevIndex As Long: lPrevIndex = pListHeadIndex
Dim lIndex As Long: lIndex = mListeners(lPrevIndex).Next
Do While lIndex <> pListTailIndex
    If mListeners(lIndex).Listener Is pListener Then
        mListeners(lPrevIndex).Next = mListeners(lIndex).Next
        releaseIndex lIndex
        RemoveListener = True
        Exit Function
    End If
    lPrevIndex = lIndex
    lIndex = mListeners(lIndex).Next
Loop
RemoveListener = False
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AllocateIndex() As Long
Static sInitialised As Boolean
If Not sInitialised Then
    sInitialised = True
    ReDim mListeners(511) As ListenerEntry
    mNextUnusedIndex = 0
    mFirstFreeIndex = NullIndex
End If

If mFirstFreeIndex <> NullIndex Then
    AllocateIndex = mFirstFreeIndex
    mFirstFreeIndex = mListeners(mFirstFreeIndex).Next
Else
    If mNextUnusedIndex > UBound(mListeners) Then
        ReDim Preserve mListeners(2 * (UBound(mListeners) + 1) - 1) As ListenerEntry
    End If
    AllocateIndex = mNextUnusedIndex
    mNextUnusedIndex = mNextUnusedIndex + 1
End If
End Function

Private Sub releaseIndex(ByVal pIndex As Long)
Set mListeners(pIndex).Listener = Nothing
mListeners(pIndex).Next = mFirstFreeIndex
mFirstFreeIndex = pIndex
End Sub




