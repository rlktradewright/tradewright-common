VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DictionaryEntriesEnum"
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

Private Const ModuleName                            As String = "DictionaryEntriesEnum"

'@================================================================================
' Member variables
'@================================================================================

Private mDict                                       As SortedDictionary
Private mRoot                                       As DictionaryEntry
Private mDeleteAsYouGo                              As Boolean
Private mUpdateSequence                             As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Const ProcName As String = "IEnumerable_Enumerator"
On Error GoTo Err

Set IEnumerable_Enumerator = Enumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IEnumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_NextItem"
On Error GoTo Err

Dim lCookie As EnumerationCookie
If IsEmpty(pCookie) Then
    Set lCookie.Current = Nothing
    Set lCookie.Next = Nothing
    gFindFirst lCookie, mRoot, mDeleteAsYouGo
Else
    lCookie = pCookie
    gFindNext lCookie, mDeleteAsYouGo
End If

If lCookie.Current Is Nothing Then
    Item = Empty
Else
    Set Item = lCookie.Current
End If

pCookie = lCookie
If Not IsEmpty(Item) Then IEnumerable_NextItem = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IEnumerable_Remove(pCookie As Variant)
Const ProcName As String = "IEnumerable_Remove"
On Error GoTo Err

Dim lCookie As EnumerationCookie
lCookie = pCookie
mDict.RemoveCurrent lCookie
pCookie = lCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IEnumerable_Skip(ByVal numToSkip As Long, pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_Skip"
Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Skip not supported"
End Function

Private Property Get IEnumerable_UpdateSequence() As Long
Const ProcName As String = "IEnumerable_UpdateSequence"
On Error GoTo Err

IEnumerable_UpdateSequence = IIf(mDeleteAsYouGo, mUpdateSequence, mDict.UpdateSequence)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
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
                ByVal pDict As SortedDictionary, Optional ByVal pDeleteAsYouGo As Boolean)
Set mDict = pDict
Set mRoot = pDict.Root
mDeleteAsYouGo = pDeleteAsYouGo
mUpdateSequence = mDict.UpdateSequence
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




