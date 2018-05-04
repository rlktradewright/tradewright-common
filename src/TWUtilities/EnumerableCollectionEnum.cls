VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumerableCollectionEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IEnumerable

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

Private Const ModuleName                            As String = "EnumerableCollectionEnum"

'@================================================================================
' Member variables
'@================================================================================

Private mEnumColl                                   As EnumerableCollection
Private mColl                                       As Collection

Private mUpdateSequence                             As Long

Private mDontNotifyDeletions                            As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Const ProcName As String = "IEnumerable_Enumerator"
On Error GoTo Err

Set IEnumerable_Enumerator = Enumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IEnumerable_NextItem( _
                Item As Variant, _
                pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_NextItem"
On Error GoTo Err

Dim lCookie As EnumerationCookie
Dim en As IEnumVARIANT

If IsEmpty(pCookie) Then
    Set en = mColl.[_NewEnum]
    Set lCookie.Next = en
    lCookie.Current = 0
Else
    lCookie = pCookie
    Set en = lCookie.Next
End If

lCookie.Current = Abs(lCookie.Current) + 1

Dim numFetched As Long
en.Next 1, Item, VarPtr(numFetched)

IEnumerable_NextItem = (numFetched <> 0)

pCookie = lCookie

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IEnumerable_Remove(pCookie As Variant)
Const ProcName As String = "IEnumerable_Remove"
On Error GoTo Err

Dim lCookie As EnumerationCookie
lCookie = pCookie

If lCookie.Current <= 0 Then
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Current Item already removed"
End If

If mDontNotifyDeletions Then
    mEnumColl.RemoveIt mColl, CLng(lCookie.Current)
Else
    mEnumColl.Remove CLng(lCookie.Current)
End If
lCookie.Current = -(lCookie.Current - 1) ' negative indicates item has been deleted

pCookie = lCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IEnumerable_Skip(ByVal numToSkip As Long, pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_Skip"
On Error GoTo Err

Dim lCookie As EnumerationCookie
Dim en As IEnumVARIANT

If IsEmpty(pCookie) Then
    Set en = mColl.[_NewEnum]
    Set lCookie.Next = en
    lCookie.Current = 1
Else
    lCookie = pCookie
    Set en = lCookie.Next
End If

en.Skip numToSkip
lCookie.Current = Abs(lCookie.Current) + numToSkip

pCookie = lCookie

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IEnumerable_UpdateSequence() As Long
IEnumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Enumerator() As Enumerator
Const ProcName As String = "Enumerator"
On Error GoTo Err

Set Enumerator = New Enumerator
Enumerator.Initialise Me

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pEnumColl As EnumerableCollection, _
                ByVal pColl As Collection, _
                ByVal pDontNotifyDeletions As Boolean)
Set mColl = pColl
Set mEnumColl = pEnumColl
mUpdateSequence = mEnumColl.UpdateSequence
mDontNotifyDeletions = pDontNotifyDeletions
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




