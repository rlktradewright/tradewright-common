Attribute VB_Name = "GListMembership"
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

Public Type ListMembership
    List                    As LinkedList
    EntryIndex              As Long
End Type

Private Type ListMembershipEntry
    ListMembership          As ListMembership
    NextIndex               As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GListMembership"

'@================================================================================
' Member variables
'@================================================================================

Private mListMembershipEntries()                    As ListMembershipEntry
Private mListMembershipEntriesIndex                 As Long

Private mFirstFreeListMembershipEntryIndex          As Long

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

Public Sub gAddListMembership( _
                ByVal pMembershipList As LinkedList, _
                ByVal pGraphObjIndex As Long, _
                ByVal pTargetList As LinkedList)
Dim lmIndex As Long

lmIndex = gAllocateListMembershipEntry
Set mListMembershipEntries(lmIndex).ListMembership.List = pTargetList
mListMembershipEntries(lmIndex).ListMembership.EntryIndex = pTargetList.AddItem(pGraphObjIndex)
pMembershipList.AddItem lmIndex
End Sub

Public Function gAllocateListMembershipEntry() As Long
Const ProcName As String = "gAllocateListMembershipEntry"
If mFirstFreeListMembershipEntryIndex >= 0 Then
    gAllocateListMembershipEntry = mFirstFreeListMembershipEntryIndex
    mFirstFreeListMembershipEntryIndex = nextListMembershipEntryIndex(mFirstFreeListMembershipEntryIndex)
Else
    If mListMembershipEntriesIndex > UBound(mListMembershipEntries) Then
        ReDim Preserve mListMembershipEntries(2 * (UBound(mListMembershipEntries) + 1) - 1) As ListMembershipEntry
        If gLogger.IsLoggable(LogLevelHighDetail) Then _
            gLogger.Log "Increased mListMembershipEntries size to", ProcName, ModuleName, LogLevelHighDetail, CStr(UBound(mListMembershipEntries) + 1)
    End If
    gAllocateListMembershipEntry = mListMembershipEntriesIndex
    mListMembershipEntriesIndex = mListMembershipEntriesIndex + 1
End If

mListMembershipEntries(gAllocateListMembershipEntry).NextIndex = NullIndex
End Function

Public Sub gClearListMembership( _
                ByVal pMembershipList As LinkedList)
Dim lListMembershipIndex As Long
Dim En As Enumerator

If pMembershipList.Count <> 0 Then
    Set En = pMembershipList.Enumerator
    Do While En.MoveNext
        lListMembershipIndex = CLng(En.Current)
        With mListMembershipEntries(lListMembershipIndex).ListMembership
            .List.RemoveEntry .EntryIndex
            releaseListMembershipEntry lListMembershipIndex
        End With
    Loop
End If

pMembershipList.Clear
End Sub

Public Sub gInit()
ReDim mListMembershipEntries(511) As ListMembershipEntry
mFirstFreeListMembershipEntryIndex = NullIndex
End Sub

Public Function gRemoveListMembership( _
                ByVal pMembershipList As LinkedList, _
                ByVal pTargetList As LinkedList) As Long
Dim lListMembershipIndex As Long
Dim En As Enumerator

gRemoveListMembership = NullIndex

If pMembershipList.Count <> 0 Then
    Set En = pMembershipList.Enumerator
    Do While En.MoveNext
        lListMembershipIndex = CLng(En.Current)
        With mListMembershipEntries(lListMembershipIndex).ListMembership
            If .List Is pTargetList Then
                En.Remove
                gRemoveListMembership = .EntryIndex
                releaseListMembershipEntry lListMembershipIndex
                Exit Do
            End If
        End With
    Loop
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Public Function nextListMembershipEntryIndex(ByVal pIndex As Long) As Long
nextListMembershipEntryIndex = mListMembershipEntries(pIndex).NextIndex
End Function

Private Sub releaseListMembershipEntry(ByVal pEntryIndex As Long)
SetNextlistMembershipEntryIndex pEntryIndex, mFirstFreeListMembershipEntryIndex
mFirstFreeListMembershipEntryIndex = pEntryIndex
End Sub

Private Sub SetNextlistMembershipEntryIndex( _
                ByVal pIndex As Long, _
                ByVal pNextIndex As Long)
mListMembershipEntries(pIndex).NextIndex = pNextIndex
End Sub




