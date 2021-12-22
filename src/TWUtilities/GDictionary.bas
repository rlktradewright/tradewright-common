Attribute VB_Name = "GDictionary"
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

Private Const ModuleName                            As String = "GDictionary"

'@================================================================================
' Member variables
'@================================================================================

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

Public Function gCompare( _
                ByVal value1 As Variant, _
                ByVal value2 As Variant, _
                ByVal pKeyType As DictionaryKeyTypes) As Long
Const ProcName As String = "gCompare"
On Error GoTo Err

Select Case pKeyType
Case KeyTypeInteger, KeyTypeFloat, KeyTypeDate
    If value1 = value2 Then
        gCompare = 0
    ElseIf value1 > value2 Then
        gCompare = 1
    Else
        gCompare = -1
    End If
Case KeyTypeString
    gCompare = StrComp(value1, value2, vbTextCompare)
Case KeyTypeCaseSensitiveString
    gCompare = StrComp(value1, value2, vbBinaryCompare)
Case KeyTypeComparable
    Dim obj1 As IComparable
    Dim obj2 As IComparable
    Set obj1 = value1
    Set obj2 = value2
    gCompare = obj1.CompareTo(obj2)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gFindEntry( _
                ByVal pKey As Variant, _
                ByVal pRoot As DictionaryEntry, _
                ByVal pKeyType As DictionaryKeyTypes, _
                ByRef pEntry As DictionaryEntry) As Long
Const ProcName As String = "gFindEntry"
On Error GoTo Err

Dim currentEntry As DictionaryEntry
Set currentEntry = pRoot
Set pEntry = currentEntry

Do While Not currentEntry Is Nothing
    Dim cmp As Long
    cmp = gCompare(pKey, currentEntry.Key, pKeyType)
    If cmp = 0 Then
        Do While Not currentEntry.Left Is Nothing
            If gCompare(currentEntry.Left.Key, currentEntry.Key, pKeyType) = 0 Then
                Set currentEntry = currentEntry.Left
            Else
                Exit Do
            End If
        Loop
        Set pEntry = currentEntry
        gFindEntry = 0
        Exit Function
    ElseIf cmp < 0 Then
        Set pEntry = currentEntry
        gFindEntry = 1
        Set currentEntry = currentEntry.Left
    Else
        Set pEntry = currentEntry
        gFindEntry = -1
        Set currentEntry = currentEntry.Right
    End If
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gFindFirst( _
                ByRef pCookie As EnumerationCookie, _
                ByVal pRoot As DictionaryEntry, _
                ByVal pDeleteAsYouGo As Boolean)
Const ProcName As String = "gFindFirst"
On Error GoTo Err

Dim first As DictionaryEntry
Set first = gFirstEntry(pRoot)
Set pCookie.Current = first
Set pCookie.Next = gSuccessor(first, pDeleteAsYouGo)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gFindFirstFromKey( _
                ByVal pKey As Variant, _
                ByVal pKeyType As DictionaryKeyTypes, _
                ByRef pCookie As EnumerationCookie, _
                ByVal pRoot As DictionaryEntry, _
                ByVal pDeleteAsYouGo As Boolean)
Const ProcName As String = "gFindFirst"
On Error GoTo Err

Set pCookie.Current = Nothing
Set pCookie.Next = Nothing

Dim lInitialEntry As DictionaryEntry
If gFindEntry(pKey, pRoot, pKeyType, lInitialEntry) >= 0 Then
    Set pCookie.Current = lInitialEntry
    Set pCookie.Next = gSuccessor(lInitialEntry, pDeleteAsYouGo)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gFindNext( _
                ByRef pCookie As EnumerationCookie, _
                ByVal pDeleteAsYouGo As Boolean)
Const ProcName As String = "gFindNext"
On Error GoTo Err

Set pCookie.Current = pCookie.Next
Set pCookie.Next = gSuccessor(pCookie.Current, pDeleteAsYouGo)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gFirstEntry(ByVal pRoot As DictionaryEntry) As DictionaryEntry
Const ProcName As String = "gFirstEntry"
On Error GoTo Err

Dim e As DictionaryEntry
Set e = pRoot
If Not e Is Nothing Then
    Do While Not e.Left Is Nothing
        Set e = e.Left
    Loop
End If
Set gFirstEntry = e

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gPredecessor( _
                ByVal pCurrent As DictionaryEntry) As DictionaryEntry
Const ProcName As String = "gPredecessor"
On Error GoTo Err

Dim e As DictionaryEntry

If pCurrent Is Nothing Then

ElseIf Not pCurrent.Left Is Nothing Then
    Set e = pCurrent.Left
    Do While Not e.Left Is Nothing
        Set e = e.Right
    Loop
    Set gPredecessor = e
Else
    Set e = pCurrent.Parent
    
    Dim ch As DictionaryEntry
    Set ch = pCurrent
    
    Do While Not e Is Nothing
        If Not ch Is e.Right Then Exit Do
        Set ch = e
        Set e = e.Parent
    Loop
    Set gPredecessor = e
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gSuccessor( _
                ByVal pCurrent As DictionaryEntry, _
                ByVal pDeleteAsYouGo As Boolean) As DictionaryEntry
Const ProcName As String = "gSuccessor"
On Error GoTo Err

Dim e As DictionaryEntry

If pCurrent Is Nothing Then

ElseIf Not pCurrent.Right Is Nothing Then
    Set e = pCurrent.Right
    Do While Not e.Left Is Nothing
        Set e = e.Left
    Loop
    Set gSuccessor = e
Else
    Set e = pCurrent.Parent
    
    Dim ch As DictionaryEntry
    Set ch = pCurrent
    
    Do While Not e Is Nothing
        If Not ch Is e.Right Then
            If pDeleteAsYouGo Then e.Left = Nothing
            Exit Do
        End If
        If pDeleteAsYouGo Then e.Right = Nothing
        Set ch = e
        Set e = e.Parent
    Loop
    Set gSuccessor = e
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




