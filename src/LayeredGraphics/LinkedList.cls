VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LinkedList"
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

Implements Enumerable

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

Private Const ModuleName                            As String = "LinkedList"

'@================================================================================
' Member variables
'@================================================================================

Private mFirstIndex             As Long
Private mLastIndex              As Long
Private mUpdateSequence         As Long
Private mCount                  As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Initialize
End Sub

Private Sub Class_Terminate()
gReleaseList mFirstIndex, mLastIndex
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As Enumerator
Set Enumerable_Enumerator = Enumerator
End Property

Private Function Enumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean

Dim lCookie As Long

If IsEmpty(pCookie) Then
    lCookie = mFirstIndex
Else
    lCookie = CLng(pCookie)
End If

If lCookie = mLastIndex Then
    Enumerable_NextItem = False
    Exit Function
End If
    
lCookie = gNextListEntryIndex(lCookie)
If lCookie = mLastIndex Then
    Enumerable_NextItem = False
Else
    Item = gGetListItem(lCookie)
    Enumerable_NextItem = True
End If

pCookie = lCookie
End Function

Private Sub Enumerable_Remove(pCookie As Variant)
Const ProcName As String = "Enumerable_Remove"
On Error GoTo Err

Assert Not IsEmpty(pCookie), "No current item"

Dim lEntryIndex As Long
lEntryIndex = CLng(pCookie)
If lEntryIndex = mLastIndex Then Exit Sub

pCookie = gPrevListEntryIndex(lEntryIndex)
removeAnEntry lEntryIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, cookie As Variant) As Boolean

End Function

Private Property Get Enumerable_UpdateSequence() As Long
Enumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Enumerator() As Enumerator
Set Enumerator = CreateEnumerator(Me)
End Property

Friend Property Get Count() As Long
Count = mCount
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function AddItem(ByVal pItem As Long) As Long
Dim lIndex As Long
lIndex = gAllocateListEntry(pItem)
gLinkListEntries gPrevListEntryIndex(mLastIndex), lIndex
gLinkListEntries lIndex, mLastIndex
AddItem = lIndex
mCount = mCount + 1
change
End Function

Friend Sub Clear()
If mCount = 0 Then Exit Sub
gReleaseList mFirstIndex, mLastIndex
Initialize
change
End Sub

Friend Sub RemoveEntry(ByVal pEntryIndex As Long)
If pEntryIndex = NullIndex Then Exit Sub
removeAnEntry pEntryIndex
change
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub change()
mUpdateSequence = mUpdateSequence + 1
End Sub

Private Sub Initialize()
mFirstIndex = gAllocateListEntry(-1)
mLastIndex = gAllocateListEntry(-1)
gLinkListEntries mFirstIndex, mLastIndex
mCount = 0
End Sub

Private Sub removeAnEntry(ByVal pEntryIndex As Long)
If pEntryIndex = NullIndex Then Exit Sub
gReleaseListEntry pEntryIndex
mCount = mCount - 1
End Sub



