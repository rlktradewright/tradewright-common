VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' A <code>StringBuilder</code> object provides a potentially more efficient means of
' manipulating strings than the standard Visual Basic string-handling functions and
' the concatenation (<code>&amp;</code>) operator.
'
' Where a large number of strings must be concatenated to produce a single string, the
' Visual Basic string concatenation <code>&amp;</code> operator can be inefficient
' because every use of it involves the two strings on either side being copied to
' form a new string.
'
' The <code>StringBuilder</code> class avoids excess copying by maintaining a buffer
' into which each appended string is copied - thus when building a string from many
' parts, only each part to be added is copied into the buffer. Copying of what has
' already been assembled is only necessary when the buffer needs to be extended, and
' this can be minimised by careful setting of the <code>StringBuilder</code> object's
' Capacity when it is created.
'
' Note that <code>StringBuilder</code> has overheads of its own that are greater than
' those of the <code>&amp;</code> operator for small numbers of concatenations. It is not
' possible to give a hard and fast rule as to when <code>StringBuilder</code> becomes
' more efficient than <code>&amp;</code>, but as a rule of thumb, anything above
' 100 concatenation operations is best done with a <code>StringBuilder</code>.
'
' A <code>StringBuilder</code> object cannot be created with the <code>New</code>
' operator: use the global <code>CreateStringBuilder</code> method.
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


Private Const ModuleName                    As String = "StringBuilder"

'@================================================================================
' Member variables
'@================================================================================


Private mBuffer                             As String
Private mBufferIndex                        As Long     ' the index of the next
                                                        ' free position in mBuffer

Private mMaxCapacity                        As Long

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

''
' Returns the current Capacity of the <code>StringBuilder</code> object. This is the maximum
' size of string that the object can hold without needing to extend the buffer.
'
' @return
'   The current Capacity of this <code>StringBuilder</code> object.
'@/
Public Property Get Capacity() As Long
Capacity = Len(mBuffer)
End Property

''
' Returns the character at the specified position (the first character is at position 1).
'
' @return
'   The character at the specified position.
' @param position
'   The position for which the character is required.
'@/
Public Property Get Character( _
                ByVal position As Long) As String
Const ProcName As String = "Character"
On Error GoTo Err

gAssertArgument position > 0 Or position <= Length, "position must be > 0 and <= current Length"

Character = Mid$(mBuffer, position, 1)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the Length in characters of the built string.
'
' @return
'   The Length in characters of the built string.
'@/
Public Property Get Length() As Long
Length = mBufferIndex / 2
End Property
                
''
' Returns the maximum Capacity of the <code>StringBuilder</code> object. This is the Length
' in characters of the longest string that can be built with this object.
'
' @return
'   The maximum Capacity of the <code>StringBuilder</code> object.
'@/
Public Property Get MaxCapacity() As Long
MaxCapacity = mMaxCapacity
End Property
                

'@================================================================================
' Methods
'@================================================================================

''
' Appends the specified string to the string under construction.
'
' @param Value
'   The string to be appended.
'@/
Public Sub Append( _
                ByVal Value As String)
Const ProcName As String = "Append"
On Error GoTo Err

EnsureCapacity Length + Len(Value)

CopyMemory bufferElementAddress(mBufferIndex), StrPtr(Value), LenB(Value)
mBufferIndex = mBufferIndex + LenB(Value)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Appends the specified string, together with following carriage return and new line characters,
' to the string under construction.
'
' @remarks
'   This method is significantly more efficient than calling the <code>append</code> method
'   separately for the string and the line terminator.
' @param Value
'   The string to be appended.
'@/
Public Sub AppendLine( _
                ByVal Value As String)
Const ProcName As String = "AppendLine"
On Error GoTo Err

EnsureCapacity Length + Len(Value) + 2

CopyMemory bufferElementAddress(mBufferIndex), StrPtr(Value), LenB(Value)
mBufferIndex = mBufferIndex + LenB(Value)
CopyMemory bufferElementAddress(mBufferIndex), StrPtr(vbCrLf), LenB(vbCrLf)
mBufferIndex = mBufferIndex + LenB(vbCrLf)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Clears any data that has been previously added.
'
'@/
Public Sub Clear()
mBufferIndex = 0
End Sub

''
' Copies the string under construction, or a substring of it, to (a subarray of) a
' byte array.
'
' @param destination
'   The byte array into which the string is to be copied.
' @param startPosition
'   The 1-based index of the first character in the string under construction to be
'   copied.
' @param destinationIndex
'   The zero-based index of the array element where the copied string is to start.
' @param Length
'   The Length of the substring to be copied. If omitted or zero, the string is copied
'   from the starting point to the end.
'@/
Public Sub CopyToBytes( _
                Destination() As Byte, _
                Optional ByVal startPosition As Long = 1, _
                Optional ByVal destinationIndex As Long, _
                Optional ByVal Length As Long)
Const ProcName As String = "CopyToBytes"
On Error GoTo Err

gAssertArgument startPosition > 0, "startPosition must be > 0"
gAssertArgument destinationIndex >= 0, "destinationIndex must be >= 0"
gAssertArgument Length >= 0, "Length must be >= 0"
gAssertArgument startPosition + Length - 1 <= Me.Length, "Request exceeds Length of Source"


If Length = 0 Then Length = Me.Length - startPosition + 1

gAssertArgument 2 * (destinationIndex + Length) - 1 <= UBound(Destination), "Request exceeds Length of destination"

CopyMemory VarPtr(Destination(2 * destinationIndex)), _
            bufferElementAddress(2 * (startPosition - 1)), _
            2 * Length

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Adjusts the size of the internal buffer so that a string of the specified Length
' can he held without further need to adjust the buffer size.
'
' @remarks
'   If the buffer is already larger than the required Capacity, no change is made (the
'   buffer size is never reduced).
' @return
'   The actual Capacity (after any adjustment if necessary).
' @param reqdCapacity
'   The required Capacity in characters.
'@/
Public Function EnsureCapacity( _
                ByVal reqdCapacity As Long) As Long
Const ProcName As String = "EnsureCapacity"
On Error GoTo Err

gAssertArgument reqdCapacity >= 0, "Capacity cannot be negative"
gAssertArgument reqdCapacity <= MaxCapacity, "Max Capacity would be exceeded"

If reqdCapacity < Capacity Then
    EnsureCapacity = Capacity
    Exit Function
End If

Dim newCapacity As Long
newCapacity = Capacity
Do While reqdCapacity > newCapacity
    If 2 * newCapacity <= MaxCapacity Then
        newCapacity = 2 * newCapacity
    Else
        newCapacity = MaxCapacity
    End If
Loop

If newCapacity <> Capacity Then
    mBuffer = mBuffer & Space$(newCapacity - Capacity)
End If

EnsureCapacity = Capacity

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
                
End Function

Friend Sub Initialise( _
                Optional ByVal initialCapacity As Long = 128, _
                Optional ByVal MaxCapacity As Long = &H7FFFFFFF, _
                Optional ByVal initialvalue As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

gAssertArgument initialCapacity >= 1 And MaxCapacity >= 1, "initialCapacity and MaxCapacity must be > 0 and <= &HFFFFFFF"

gAssertArgument initialCapacity <= MaxCapacity, "initialCapacity cannot be greater than MaxCapacity"
gAssertArgument Len(initialvalue) <= MaxCapacity, "Initial Value exceeds max Capacity"

mBuffer = Space$(initialCapacity)
mMaxCapacity = MaxCapacity

If Not IsEmpty(initialvalue) Then
    Append initialvalue
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Inserts the specified string at the specified position in the string under
' construction.
'
' @param Value
'   The string to be inserted.
' @param position
'   The position at which the string is to be inserted. The first position is 1.
'   If the value is greater than the Length of the string under construction, the
'   string is appended.
'
'@/
Public Sub Insert( _
                ByVal Value As String, _
                ByVal position As Long)
Const ProcName As String = "Insert"
On Error GoTo Err

gAssertArgument position > 0, "position must be > 0"

If Value = "" Then Exit Sub

EnsureCapacity Length + Len(Value)

If position > (Length + 1) Then position = Length + 1

If position < (Length + 1) Then
    CopyMemory bufferElementAddress(2 * (position + Len(Value) - 1)), _
                bufferElementAddress(2 * (position - 1)), _
                2 * (Length - position + 1)
End If

CopyMemory bufferElementAddress(2 * (position - 1)), StrPtr(Value), 2 * Len(Value)
mBufferIndex = mBufferIndex + 2 * Len(Value)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Places the specified string at the specified position in the string under
' construction, overwriting its previous contents.
'
' @param Value
'   The string to be included.
' @param position
'   The position at which the string is to be included. The first position is 1.
'
'@/
Public Sub Overwrite( _
                ByVal Value As String, _
                ByVal position As Long)
Const ProcName As String = "Overwrite"
On Error GoTo Err

gAssertArgument position > 0 And position <= Length + 1, "position invalid"

If Value = "" Then Exit Sub

EnsureCapacity position + Len(Value) - 1

CopyMemory bufferElementAddress(2 * (position - 1)), StrPtr(Value), 2 * Len(Value)
mBufferIndex = mBufferIndex + gMax(0, 2 * (position + Len(Value) - 1 - Me.Length))

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Removes the specified substring from the string under construction.
'
' @param StartIndex
'   The starting position of the string to be removed. The first character is
'   at position 1.
' @param lengthToRemove
'   The Length of the string to be removed.
'
'@/
Public Sub Remove( _
                ByVal StartIndex As Long, _
                ByVal lengthToRemove As Long)
Const ProcName As String = "Remove"
On Error GoTo Err

gAssertArgument StartIndex > 0 And lengthToRemove >= 0 And StartIndex + lengthToRemove <= Length + 1, "Specified segment outside buffer"

If lengthToRemove = 0 Then Exit Sub

If StartIndex <> Length Then
    CopyMemory bufferElementAddress(2 * (StartIndex - 1)), _
                bufferElementAddress(2 * (StartIndex + lengthToRemove - 1)), _
                mBufferIndex - 2 * (StartIndex + lengthToRemove - 1)
End If
mBufferIndex = mBufferIndex - 2 * lengthToRemove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Replaces one or more occurrences of the specified substring in the string
' under construction with another string.
'
' @return
'   The number of replacements made.
' @param targetString
'   The substring to be replaced.
' @param replacementString
'   The string which is to Replace the Target string.
' @param ignoreCase
'   If <code>True</code>, case is to be ignored when searching for the
'   Target string. If not specified, the search is case-sensitive.
' @param startPosition
'   The position within the string under construction which is the start of
'   the area to be searched for the Target string. If not specified, the searched
'   area starts at position 1.
' @param searchLength
'   The Length of the area of the string under construction to be searched
'   for the Target string. If not specified or negatvie, the area to be searched
'   extends from the specified position to the end of the string.
' @param maxReplacements
'   The maximum number of replacements to be made. If not specified or negative,
'   all occurrences of the Target string are replaced.
' @param reverseSearch
'   If <code>True</code>, searching starts at the end of the specified area and
'   proceeds backwards.
'@/
Public Function Replace( _
                ByVal targetString As String, _
                ByVal replacementString As String, _
                Optional ByVal ignoreCase As Boolean, _
                Optional ByVal startPosition = 1, _
                Optional ByVal searchLength As Long = -1, _
                Optional ByVal maxReplacements As Long = -1, _
                Optional ByVal reverseSearch As Boolean) As Long
Const ProcName As String = "Replace"
On Error GoTo Err

gAssertArgument startPosition + searchLength <= Length And startPosition >= 1, "Invalid startPosition or searchLength"

If maxReplacements = 0 Then Exit Function
If searchLength = 0 Then Exit Function

If maxReplacements < 0 Then maxReplacements = &H7FFFFFFF

Dim endPosition As Long
If searchLength < 0 Then
    endPosition = Length + 1
Else
    endPosition = startPosition + searchLength
End If

If Len(targetString) > (endPosition - startPosition) Then Exit Function

If ignoreCase Then targetString = UCase$(targetString)

Dim i As Long
Dim numReplacements As Long
Dim found As Boolean
If Not reverseSearch Then
    For i = startPosition To (endPosition - Len(targetString))
        found = containsStringAtPosition(i, targetString, ignoreCase)
        If found Then
            Remove i, Len(targetString)
            If i = Length Then
                Append replacementString
            Else
                Insert replacementString, i
            End If
            i = i + Len(replacementString)
            endPosition = endPosition - Len(targetString) + Len(replacementString)
            numReplacements = numReplacements + 1
            If numReplacements >= maxReplacements Then
                Replace = numReplacements
                Exit Function
            End If
        End If
    Next
Else
    For i = (endPosition - Len(targetString)) To startPosition Step -1
        found = containsStringAtPosition(i, targetString, ignoreCase)
        If found Then
            Remove i, Len(targetString)
            If i = Length Then
                Append replacementString
            Else
                Insert replacementString, i
            End If
            endPosition = endPosition - Len(targetString) + Len(replacementString)
            numReplacements = numReplacements + 1
            If numReplacements >= maxReplacements Then
                Replace = numReplacements
                Exit Function
            End If
        End If
    Next
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns a substring of the string under construction.
'
' @param startPosition
'   The 1-based index of the first character of the substring.
' @param Length
'   The Length of the substring to be returned. If omitted or zero, the Length from
'   the starting point to the end of the string under construction is assumed.
'@/
Public Function SubString( _
                Optional ByVal startPosition As Long = 1, _
                Optional ByVal Length As Long)
Const ProcName As String = "SubString"
On Error GoTo Err

gAssertArgument startPosition > 0, "startPosition must be > 0"
gAssertArgument Length >= 0, "Length must be >= 0"
gAssertArgument startPosition + Length - 1 <= Me.Length, "Request exceeds Length of Source"

If Length = 0 Then Length = Me.Length - startPosition + 1

SubString = Mid$(mBuffer, startPosition, Length)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the current value of the string under construction.
'
' @return
'   The current value of the string under construction.
'
'@/
Public Function ToString() As String
Const ProcName As String = "ToString"
On Error GoTo Err

ToString = Left$(mBuffer, Length)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the position in the string under construction of the start of the first
' word to the left of the specified position.
'
' @remarks
'   If the specified position is within a word but after its first character, the
'   position of the start of that word is returned.
'
'   If the specified position is the first character of a word, the
'   position of the start of the previous word is returned. If there is no previous word,
'   the value 1 is returned.
' @return
'   The starting position of the required word.
' @param position
'   The position from which to start looking for a word. The value must be greater than
'   zero. If it is greater than the Length of the string under construction, the start of
'   the last word in the string is found.
' @param separator
'   The character used as a separator between words. This argument must be a single
'   character, though there may be any number of occurrences of it between words.
' @see
'
'@/
Public Function WordPositionLeft( _
                ByVal position As Long, _
                Optional ByVal separator As String = " ") As Long
Const ProcName As String = "WordPositionLeft"
On Error GoTo Err

gAssertArgument position > 0, "position must be > 0"
gAssertArgument Len(separator) = 1, "separator must be a single character"

If position = 1 Or Me.Length = 0 Then
    WordPositionLeft = 1
    Exit Function
End If

If position > Me.Length Then
    position = Me.Length
End If


If Mid$(mBuffer, position, 1) = separator Then
    WordPositionLeft = previousWordPosition(position, separator)
ElseIf Mid$(mBuffer, position - 1, 1) = separator Then
    WordPositionLeft = previousWordPosition(position - 1, separator)
Else
    WordPositionLeft = previousWordPosition(position, separator)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the position in the string under construction of the start of the first
' word to the right of the specified position.
'
' @return
'   The starting position of the required word.
' @param position
'   The position from which to start looking for a word. The value must be greater than
'   zero. If it is greater than the Length of the string under construction, the value
'   returned is one greater than the Length of the string.
' @param separator
'   The character used as a separator between words. This argument must be a single
'   character, though there may be any number of occurrences of it between words.
' @see
'
'@/
Public Function WordPositionRight( _
                ByVal position As Long, _
                Optional ByVal separator As String = " ") As Long
Const ProcName As String = "WordPositionRight"
On Error GoTo Err

gAssertArgument position > 0, "position must be > 0"
gAssertArgument Len(separator) = 1, "separator must be a single character"

If position > Me.Length Then
    WordPositionRight = Me.Length + 1
    Exit Function
End If

WordPositionRight = nextWordPosition(position, separator)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function bufferElementAddress( _
                ByVal elementIndex As Long) As Long
Const ProcName As String = "bufferElementAddress"
On Error GoTo Err

bufferElementAddress = gUnsignedAdd(StrPtr(mBuffer), elementIndex)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function containsStringAtPosition( _
                ByVal position As Long, _
                ByVal targetString As String, _
                ByVal ignoreCase As Boolean) As Boolean
Const ProcName As String = "containsStringAtPosition"
On Error GoTo Err

If 2 * Len(targetString) > mBufferIndex - position Then Exit Function

If ignoreCase Then
    containsStringAtPosition = (UCase$(Mid$(mBuffer, position, Len(targetString))) = targetString)
Else
    containsStringAtPosition = Mid$(mBuffer, position, Len(targetString)) = targetString
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function nextWordPosition( _
                ByVal position As Long, _
                ByVal separator As String) As Long
Const ProcName As String = "nextWordPosition"
On Error GoTo Err

position = InStr(position, mBuffer, separator)

If position = 0 Or position > Me.Length Then
    nextWordPosition = Me.Length + 1
    Exit Function
End If

Do While Mid$(mBuffer, position, 1) = separator
    position = position + 1
    If position = Me.Length + 1 Then Exit Do
Loop

nextWordPosition = position

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function previousWordPosition( _
                ByVal position As Long, _
                ByVal separator As String) As Long
Const ProcName As String = "previousWordPosition"
On Error GoTo Err

Do While position > 1
    If Mid$(mBuffer, position, 1) <> separator Then Exit Do
    position = position - 1
Loop

If position = 1 Then
    previousWordPosition = 1
    Exit Function
End If

position = InStrRev(mBuffer, separator, position)
If position = 0 Then
    previousWordPosition = 1
Else
    previousWordPosition = position + 1
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



