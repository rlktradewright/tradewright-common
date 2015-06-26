Attribute VB_Name = "GLinkedList"
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

Private Type ListEntry
    PrevIndex               As Long
    NextIndex               As Long
    Item                    As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GLinkedList"

Public Const NullIndex                              As Long = -1

'@================================================================================
' Member variables
'@================================================================================

Private mListEntries()                              As ListEntry
Private mListEntriesIndex                           As Long

Private mFirstFreeListEntryIndex                    As Long

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

Public Function gAllocateListEntry(ByVal pItem As Long) As Long
Const ProcName As String = "gAllocateListEntry"
If mFirstFreeListEntryIndex >= 0 Then
    gAllocateListEntry = mFirstFreeListEntryIndex
    mFirstFreeListEntryIndex = gNextListEntryIndex(mFirstFreeListEntryIndex)
Else
    If mListEntriesIndex > UBound(mListEntries) Then
        ReDim Preserve mListEntries(2 * (UBound(mListEntries) + 1) - 1) As ListEntry
        If gLogger.IsLoggable(LogLevelHighDetail) Then _
            gLogger.Log "Increased mListEntries size to", ProcName, ModuleName, LogLevelHighDetail, CStr(UBound(mListEntries) + 1)
    End If
    gAllocateListEntry = mListEntriesIndex
    mListEntriesIndex = mListEntriesIndex + 1
End If

mListEntries(gAllocateListEntry).NextIndex = NullIndex
mListEntries(gAllocateListEntry).PrevIndex = NullIndex
mListEntries(gAllocateListEntry).Item = pItem
End Function

Public Function gGetListItem(ByVal pEntryIndex As Long) As Long
gGetListItem = mListEntries(pEntryIndex).Item
End Function

Public Sub gInit()
ReDim mListEntries(511) As ListEntry
mFirstFreeListEntryIndex = NullIndex
End Sub

Public Sub gLinkListEntries( _
                ByVal pIndex As Long, _
                ByVal pNextIndex As Long)
mListEntries(pIndex).NextIndex = pNextIndex
mListEntries(pNextIndex).PrevIndex = pIndex
End Sub

Public Function gNextListEntryIndex(ByVal pIndex As Long) As Long
gNextListEntryIndex = mListEntries(pIndex).NextIndex
End Function

Public Function gPrevListEntryIndex(ByVal pIndex As Long) As Long
gPrevListEntryIndex = mListEntries(pIndex).PrevIndex
End Function

Public Sub gReleaseListEntry(ByVal pEntryIndex As Long)
gLinkListEntries gPrevListEntryIndex(pEntryIndex), gNextListEntryIndex(pEntryIndex)
SetNextlistEntryIndex pEntryIndex, mFirstFreeListEntryIndex
mFirstFreeListEntryIndex = pEntryIndex
mListEntries(pEntryIndex).Item = NullIndex
End Sub

Public Sub gReleaseList( _
                ByVal pFirstIndex As Long, _
                ByVal pLastIndex As Long)
SetNextlistEntryIndex pLastIndex, mFirstFreeListEntryIndex
mFirstFreeListEntryIndex = pFirstIndex
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub SetNextlistEntryIndex( _
                ByVal pIndex As Long, _
                ByVal pNextIndex As Long)
mListEntries(pIndex).NextIndex = pNextIndex
End Sub


