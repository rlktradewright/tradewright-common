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

Sub gFindFirst( _
                ByRef pCookie As EnumerationCookie, _
                ByVal pRoot As DictionaryEntry)
Const ProcName As String = "gFindFirst"
On Error GoTo Err

Dim first As DictionaryEntry
Set first = gFirstEntry(pRoot)
Set pCookie.Current = first
Set pCookie.Next = gSuccessor(first)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gFindNext( _
                ByRef pCookie As EnumerationCookie)
Const ProcName As String = "gFindNext"
On Error GoTo Err

Set pCookie.Current = pCookie.Next
Set pCookie.Next = gSuccessor(pCookie.Current)

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

Public Function gSuccessor( _
                ByVal pCurrent As DictionaryEntry) As DictionaryEntry
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
        If Not ch Is e.Right Then Exit Do
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




