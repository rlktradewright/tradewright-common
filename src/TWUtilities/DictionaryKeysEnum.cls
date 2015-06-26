VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DictionaryKeysEnum"
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

Private Const ModuleName                            As String = "DictionaryKeysEnum"

'@================================================================================
' Member variables
'@================================================================================

Private mSc As SortedDictionary

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Const ProcName As String = "IEnumerable_Enumerator"

On Error GoTo Err

Set IEnumerable_Enumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IEnumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean
Dim lCookie As EnumerationCookie
Dim lEntry As DictionaryEntry

Const ProcName As String = "IEnumerable_NextItem"

On Error GoTo Err

If IsEmpty(pCookie) Then
    Set lCookie.Current = Nothing
    Set lCookie.Next = Nothing
    mSc.FindFirst lCookie
Else
    lCookie = pCookie
    mSc.FindNext lCookie
End If

If lCookie.Current Is Nothing Then
    Item = Empty
Else
    Set lEntry = lCookie.Current
    gSetVariant Item, lEntry.Key
End If

pCookie = lCookie
If Not IsEmpty(Item) Then IEnumerable_NextItem = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IEnumerable_Remove(pCookie As Variant)
Dim lCookie As EnumerationCookie
Const ProcName As String = "IEnumerable_Remove"

On Error GoTo Err

lCookie = pCookie
mSc.RemoveCurrent lCookie
pCookie = lCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IEnumerable_Skip(ByVal numToSkip As Long, pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_Skip"

On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Skip not supported"

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IEnumerable_UpdateSequence() As Long
IEnumerable_UpdateSequence = mSc.UpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal sc As SortedDictionary)
Set mSc = sc
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getEnumerator() As Enumerator
Const ProcName As String = "getEnumerator"

On Error GoTo Err

Set getEnumerator = New Enumerator
getEnumerator.Initialise Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


