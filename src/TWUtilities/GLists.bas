Attribute VB_Name = "GLists"
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

Public Type ListEntry
    Next                As Long
    Item                As Object
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GLists"

'@================================================================================
' Member variables
'@================================================================================

Private mListEntries()                              As ListEntry
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

Public Sub Add( _
                ByVal pItem As Object, _
                ByVal pListHeadIndex As Long, _
                ByVal pListTailIndex As Long)
Dim lIndex As Long: lIndex = AllocateIndex
mListEntries(lIndex).Next = mListEntries(pListHeadIndex).Next
Set mListEntries(lIndex).Item = pItem
mListEntries(pListHeadIndex).Next = lIndex
End Sub

Public Sub Clear( _
                ByVal pListHeadIndex As Long, _
                ByVal pListTailIndex As Long)
Dim lIndex As Long: lIndex = mListEntries(pListHeadIndex).Next
Dim lNextIndex As Long
Do While lIndex <> pListTailIndex
    lNextIndex = mListEntries(lIndex).Next
    releaseIndex lIndex
    lIndex = lNextIndex
Loop
mListEntries(pListHeadIndex).Next = pListTailIndex
End Sub

Public Function Contains( _
                ByVal pItem As Object, _
                ByVal pListHeadIndex As Long, _
                ByVal pListTailIndex As Long) As Boolean
Dim lIndex As Long: lIndex = mListEntries(pListHeadIndex).Next
Do While lIndex <> pListTailIndex
    If mListEntries(lIndex).Item Is pItem Then
        Contains = True
        Exit Function
    End If
    lIndex = mListEntries(lIndex).Next
Loop
End Function

Public Sub ToArray( _
                ByRef pItems() As Object, _
                ByVal pListHeadIndex As Long, _
                ByVal pListTailIndex As Long)
Dim lItem As Object
Dim lIndex As Long: lIndex = mListEntries(pListHeadIndex).Next
Dim i As Long
Do While lIndex <> pListTailIndex
    Set pItems(i) = mListEntries(lIndex).Item
    i = i + 1
    lIndex = mListEntries(lIndex).Next
Loop
End Sub

Public Sub IntialiseList( _
                ByRef pListHeadIndex As Long, _
                ByRef pListTailIndex As Long)
pListHeadIndex = AllocateIndex
pListTailIndex = AllocateIndex
mListEntries(pListHeadIndex).Next = pListTailIndex
mListEntries(pListTailIndex).Next = NullIndex
End Sub

Public Function Remove( _
                ByVal pItem As Object, _
                ByVal pListHeadIndex As Long, _
                ByVal pListTailIndex As Long) As Boolean
Dim lPrevIndex As Long: lPrevIndex = pListHeadIndex
Dim lIndex As Long: lIndex = mListEntries(lPrevIndex).Next
Do While lIndex <> pListTailIndex
    If mListEntries(lIndex).Item Is pItem Then
        mListEntries(lPrevIndex).Next = mListEntries(lIndex).Next
        releaseIndex lIndex
        Remove = True
        Exit Function
    End If
    lPrevIndex = lIndex
    lIndex = mListEntries(lIndex).Next
Loop
Remove = False
End Function

Public Sub TerminateList( _
                ByRef pListHeadIndex As Long, _
                ByRef pListTailIndex As Long)
releaseIndex pListHeadIndex
releaseIndex pListTailIndex
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AllocateIndex() As Long
Static sInitialised As Boolean
If Not sInitialised Then
    sInitialised = True
    ReDim mListEntries(511) As ListEntry
    mNextUnusedIndex = 0
    mFirstFreeIndex = NullIndex
End If

If mFirstFreeIndex <> NullIndex Then
    AllocateIndex = mFirstFreeIndex
    mFirstFreeIndex = mListEntries(mFirstFreeIndex).Next
Else
    If mNextUnusedIndex > UBound(mListEntries) Then
        gLogger.Log "Increased mListEntries size to ", "AllocateIndex", ModuleName, LogLevelDetail, CStr(UBound(mListEntries) + 1)
        ReDim Preserve mListEntries(2 * (UBound(mListEntries) + 1) - 1) As ListEntry
    End If
    AllocateIndex = mNextUnusedIndex
    mNextUnusedIndex = mNextUnusedIndex + 1
End If
End Function

Private Sub releaseIndex(ByVal pIndex As Long)
Set mListEntries(pIndex).Item = Nothing
mListEntries(pIndex).Next = mFirstFreeIndex
mFirstFreeIndex = pIndex
End Sub




