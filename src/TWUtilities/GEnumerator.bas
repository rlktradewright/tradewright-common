Attribute VB_Name = "GEnumerator"
Option Explicit

''
' Description here
'
' @remarks
' @see
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


Private Const ModuleName                    As String = "GEnumerator"

Private Const VariantLength                 As Long = 16

'@================================================================================
' Member variables
'@================================================================================

Public gRedirected                          As Boolean

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

Public Function GetNext( _
                ByVal this As Object, _
                ByVal numElementsRequested As Long, _
                ByRef items As Variant, _
                ByVal lpNumElementsFetched As Long) As Long
Dim i As Long
Dim lEnumerator As Enumerator
Dim anItem As Variant
Dim emptyItem As Variant
Dim numElementsFetched As Long

GetNext = S_FALSE

On Error GoTo Err

Set lEnumerator = this
For i = 1 To numElementsRequested
    If lEnumerator.GetNext(anItem) Then
        ' move the item into place in the caller's buffer
        CopyMemory VarPtr(items) + (i - 1) * VariantLength, VarPtr(anItem), VariantLength
        numElementsFetched = numElementsFetched + 1
        
        ' now turn anItem into an empty variant, so that any objects or strings
        ' it may contain don't get released on exit from the function
        CopyMemory VarPtr(anItem), VarPtr(emptyItem), VariantLength
    Else
        If lpNumElementsFetched <> 0 Then
            CopyMemory lpNumElementsFetched, VarPtr(numElementsFetched), 4
        End If
        Exit Function
    End If
Next

GetNext = S_OK
Exit Function
                
Err:

numElementsFetched = 0

' undo the variants we've passed back
For i = i To 1 Step -1
    CopyMemory VarPtr(anItem), VarPtr(items) + (i - 1) * VariantLength, VariantLength
    anItem = Empty ' release any objects or strings
    CopyMemory VarPtr(items) + (i - 1) * VariantLength, VarPtr(emptyItem), VariantLength
Next

If lpNumElementsFetched <> 0 Then
    CopyMemory lpNumElementsFetched, VarPtr(0&), 4
End If

GetNext = Err.number Or &H80000000 ' convert error code to an HRESULT

End Function

Public Function Skip( _
                ByVal this As Object, _
                ByVal numElementsToSkip As Long) As Long
Dim lEnumerator As Enumerator

On Error GoTo Err

Set lEnumerator = this

If lEnumerator.Skip(numElementsToSkip) Then
    Skip = S_OK
Else
    Skip = S_FALSE
End If
Exit Function

Err:
Skip = Err.number Or &H80000000 ' convert error code to an HRESULT
End Function

'@================================================================================
' Helper Functions
'@================================================================================


