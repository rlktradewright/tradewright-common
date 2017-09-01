VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortedDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Represents a collection of key/Value pairs, sorted on the key.
'
' Items in the collection are sorted in the order of the key supplied when they
' are added. Iterating through the collection returns the items in sorted order.
'
' Optionally, the collection may be specified to allow items with duplicate Keys.
' When iterating through the collection, items with duplicate Keys are returned in the
' order in which they were added.
'
' The first item with any given key may be updated.
'
' NB: you cannot create a <code>SortedDictionary</code> object with the <code>New</code>
' operator. Use the global <code>CreateSortedDictionary</code> method.
'
' <h3>Acknowledgement</h3>
' This class draws heavily on the implementation of the TreeMap class in the
' Java Collections Framework.
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IEnumerable
Implements IJSONable
Implements IStringable

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


Private Const ModuleName                    As String = "SortedDictionary"

'@================================================================================
' Member variables
'@================================================================================

Private mRoot                               As DictionaryEntry  ' Root of binary tree

Private mCount                              As Long

Private mAllowDuplicates                    As Boolean

Private mFirst                              As DictionaryEntry

Private mKeyType                            As DictionaryKeyTypes

Private mUpdateSequence                     As Long

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
    FindFirst lCookie
Else
    lCookie = pCookie
    FindNext lCookie
End If

If lCookie.Current Is Nothing Then
    Item = Empty
Else
    Set lEntry = lCookie.Current
    gSetVariant Item, lEntry.Data
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
RemoveCurrent lCookie
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
IEnumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' JSONable Interface Members
'@================================================================================

Private Function IJSONable_ToJSON() As String
Const ProcName As String = "IJSONable_ToJSON"

On Error GoTo Err

IJSONable_ToJSON = ToString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

'@================================================================================
' Stringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
Const ProcName As String = "IStringable_ToString"

On Error GoTo Err

IStringable_ToString = ToString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

''
' Returns the number of items in the collection.
'
' @return
'   The number of items in the collection.
'@/
Public Property Get Count() As Long
Count = mCount
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the entries of the collection.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get EntryEnumerator() As Enumerator
Dim ee As New DictionaryEntriesEnum
Dim en As IEnumerable
Const ProcName As String = "EntryEnumerator"

On Error GoTo Err

ee.Initialise Me
Set en = ee
Set EntryEnumerator = en.Enumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the data members of the collection.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get Enumerator() As Enumerator
Const ProcName As String = "Enumerator"

On Error GoTo Err

Set Enumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FirstItem() As Variant
Const ProcName As String = "FirstItem"
On Error GoTo Err

gAssert mCount > 0, "No items"

gSetVariant FirstItem, firstEntry.Data

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns a specific Item in the collection, identified by key.
'
' @remarks
'   If the key provided doesn’t match any existing member of the
'   collection, an error occurs.
'
'   If the collection contains more than one Item with the specified key, the
'   first such Item is returned.
' @return
'   The required Item.
' @param key
'   An expression that specifies the key of a member of the collection.
'@/
Public Property Get Item( _
                ByVal Key As Variant) As Variant
Const ProcName As String = "Item"
On Error GoTo Err

Dim requiredItem As Variant
If TryItem(Key, requiredItem) Then
    gSetVariant Item, requiredItem
Else
    Err.Raise VBErrorCodes.VbErrInvalidProcedureCall, , "Key not found"
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Updates a specific Item in the collection, identified by key.
'
' @remarks
'   If the key provided doesn’t match any existing member of the
'   collection, an error occurs.
'
'   If the collection contains more than one Item with the specified key, the
'   first such Item is updated.
' @param key
'   An expression that specifies the key of a member of the collection.
' @param Value
'   The new Value to be assigned to the specified Item.
'@/
Public Property Let Item( _
                ByVal Key As Variant, _
                ByVal Value As Variant)
Const ProcName As String = "Item"
On Error GoTo Err

Dim entry As DictionaryEntry
Dim Result As Long
Result = findEntry(Key, entry)
If Result = 0 Then
    gSetVariant entry.Data, Value
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the keys of the collection.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get KeyEnumerator() As Enumerator
Dim ke As New DictionaryKeysEnum
Dim en As IEnumerable
Const ProcName As String = "KeyEnumerator"

On Error GoTo Err

ke.Initialise Me
Set en = ke
Set KeyEnumerator = en.Enumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get UpdateSequence() As Long
UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Adds an Item to the collection.
'
' @remarks
'   An Error occurs if a specified key duplicates the key for an existing member
'   of the collection, and the collection was not specified to allow duplicates
'   when it was created.
'
' @param Item
'   The Item to be added to the collection.
' @param key
'   A key that determines the position of the Item in the collection.
'@/
Public Sub Add( _
                ByVal Item As Variant, _
                ByVal Key As Variant)
Dim Current As DictionaryEntry
Dim Parent As DictionaryEntry
Dim cmp As Long

Const ProcName As String = "Add"

On Error GoTo Err

changed

If IsEmpty(Key) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Key cannot be Empty"

Select Case mKeyType
Case KeyTypeInteger
    If Not gIsInteger(Key) Then Err.Raise ErrorCodes.ErrClassCastException
Case KeyTypeFloat
    If Not IsNumeric(Key) Then Err.Raise ErrorCodes.ErrClassCastException
Case KeyTypeString
    If Not VarType(Key) = vbString Then Err.Raise ErrorCodes.ErrClassCastException
Case KeyTypeCaseSensitiveString
    If Not VarType(Key) = vbString Then Err.Raise ErrorCodes.ErrClassCastException
Case KeyTypeDate
    If Not IsDate(Key) Then Err.Raise ErrorCodes.ErrClassCastException
Case KeyTypeComparable
    If Not TypeOf Key Is IComparable Then Err.Raise ErrorCodes.ErrClassCastException
End Select

If mRoot Is Nothing Then
    ' first entry added, so make it root
    Set mRoot = createEntry(Key, Item, Nothing)
    mCount = 1
    Exit Sub
End If

If Not mFirst Is Nothing Then
    cmp = compare(Key, mFirst.Key)
    If cmp < 0 Then Set mFirst = Nothing
End If

' find where node belongs
Set Current = mRoot

Do While Not Current Is Nothing
    Set Parent = Current
    cmp = compare(Key, Current.Key)
    If cmp = 0 And Not mAllowDuplicates Then Err.Raise VBErrorCodes.VbErrElementAlreadyExists, , "Element already exists"
    
    If cmp < 0 Then
        Set Current = Current.Left
    Else
        Set Current = Current.Right
    End If
Loop

Dim newEntry As DictionaryEntry
Set newEntry = createEntry(Key, Item, Parent)
If cmp < 0 Then
    Parent.Left = newEntry
Else
    Parent.Right = newEntry
End If
fixAfterInsertion newEntry
mCount = mCount + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Removes all items from the collection.
'@/
Public Sub Clear()
Const ProcName As String = "Clear"

On Error GoTo Err

changed
mCount = 0
Set mRoot = Nothing
Set mFirst = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Indicates whether an Item with the specified key already exists in the collection.
'
' @return
'   <code>True</code> if an Item with the specified key already exists, otherwise
'   <code>False</code>.
' @param key
'   The relevant key.
'@/
Public Function Contains(ByVal Key As Variant) As Boolean
Const ProcName As String = "Contains"
On Error GoTo Err

Dim entry As DictionaryEntry
Contains = (findEntry(Key, entry) = 0)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub FindFirst( _
                ByRef pCookie As EnumerationCookie)
Const ProcName As String = "FindFirst"

On Error GoTo Err

If mFirst Is Nothing Then
    Set mFirst = firstEntry
End If
Set pCookie.Current = mFirst
Set pCookie.Next = successor(mFirst)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FindNext( _
                ByRef pCookie As EnumerationCookie)
Const ProcName As String = "FindNext"

On Error GoTo Err

Set pCookie.Current = pCookie.Next
Set pCookie.Next = successor(pCookie.Current)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal keyType As DictionaryKeyTypes, _
                ByVal allowDuplicates As Boolean)
Dim i As Long

Const ProcName As String = "Initialise"

On Error GoTo Err

mAllowDuplicates = allowDuplicates

Select Case keyType
Case KeyTypeInteger
Case KeyTypeFloat
Case KeyTypeString
Case KeyTypeCaseSensitiveString
Case KeyTypeDate
Case KeyTypeComparable
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid Key type"
End Select
mKeyType = keyType

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"

On Error GoTo Err

Set NewEnum = getEnumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the lowest key strictly greater than the supplied key.
'
' @remarks
'
' @return
'    The lowest key higher than <code>Key</code>, or <code>Empty</code> if there
'    is no such key.
' @param key
'    The key
' @see
'
'@/
'
Public Function NextKey(ByVal Key As Variant) As Variant
Const ProcName As String = "NextKey"
On Error GoTo Err

Dim entry As DictionaryEntry
Dim Result As Integer
Result = findEntry(Key, entry)

If Result <= 0 Then
    ' either we found an entry with the supplied key, or
    ' the entry found is lower than the supplied key, so it must be the
    ' successor that we want
    Set entry = successor(entry)
    If entry Is Nothing Then
        NextKey = Empty
    Else
        gSetVariant NextKey, entry.Key
    End If
Else
    ' the entry found is higher than the supplied key, so it must be the
    ' one we want
    gSetVariant NextKey, entry.Key
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the greatest key strictly less than the supplied key.
'
' @remarks
'
' @return
'    The highest key lower than <code>Key</code>, or <code>Empty</code> if there
'    is no such key.
' @param key
'    The key
' @see
'
'@/
'
Public Function PreviousKey(ByVal Key As Variant) As Variant
Const ProcName As String = "PreviousKey"
On Error GoTo Err

Dim entry As DictionaryEntry
Dim Result As Integer
Result = findEntry(Key, entry)

If Result >= 0 Then
    ' either we found an entry with the supplied key, or
    ' the entry found is lower than the supplied key, so it must be the
    ' predecessor that we want
    Set entry = predecessor(entry)
    If entry Is Nothing Then
        PreviousKey = Empty
    Else
        gSetVariant PreviousKey, entry.Key
    End If
Else
    ' the entry found is lower than the supplied key, so it must be the
    ' one we want
    gSetVariant PreviousKey, entry.Key
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Removes an Item from the collection.
'
' @remarks
'   If the key provided doesn’t match an existing member of the
'   collection, an error occurs.
'
'   If the collection contains more than one Item with the specified key, the
'   first such Item to be added is removed.
' @return
'   Returns the Item being removed.
' @param key
'   Identifies the Item to be removed from the collection.
'@/
Public Function Remove(ByVal Key As Variant) As Variant
Const ProcName As String = "Remove"
On Error GoTo Err

Dim entry As DictionaryEntry
Dim Result As Long
Result = findEntry(Key, entry)

If Result <> 0 Then
    ' raise the same Error as the VB6 Collection object
    Err.Raise VBErrorCodes.VbErrInvalidProcedureCall, _
                , _
                "Invalid procedure call or argument"
End If

gSetVariant Remove, entry.Data
deleteEntry entry

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub RemoveCurrent( _
                ByRef pCookie As EnumerationCookie)
Const ProcName As String = "RemoveCurrent"

On Error GoTo Err

Dim lCurrentEntry As DictionaryEntry

If IsEmpty(pCookie.Current) Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Current Item already deleted"

Set lCurrentEntry = pCookie.Current
If (Not lCurrentEntry.Left Is Nothing And _
    Not lCurrentEntry.Right Is Nothing) _
Then Set pCookie.Next = pCookie.Current
deleteEntry pCookie.Current
Set pCookie.Current = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

''
' Gets a specific item in the collection, identified by key.
'
' @remarks
'   If the key provided doesn’t match any existing member of the
'   collection, <code>False</code> is returned.
'
'   If the collection contains more than one item with the specified key, the
'   first such item is returned.
' @return
'   <code>True</code> if the required item is found. Otherwise <code>False</code>.
' @param Key
'   An expression that specifies the key of a member of the collection.
' @param Item
'   The required item, if it has been found. If not found, the parameter is not changed.
'@/
Public Function TryItem( _
                ByVal Key As Variant, _
                ByRef Item As Variant) As Boolean
Const ProcName As String = "TryItem"
On Error GoTo Err

If mCount = 0 Then Exit Function

Dim entry As DictionaryEntry
Dim Result As Long
Result = findEntry(Key, entry)

If Result = 0 Then
    gSetVariant Item, entry.Data
    TryItem = True
Else
    TryItem = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub changed()
mUpdateSequence = mUpdateSequence + 1
End Sub

Private Function colorOf( _
                ByVal p As DictionaryEntry) As NodeColors
Const ProcName As String = "colorOf"

On Error GoTo Err

If p Is Nothing Then
    colorOf = BLACK
Else
    colorOf = p.Color
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function compare( _
                ByVal value1 As Variant, _
                ByVal value2 As Variant) As Long
Const ProcName As String = "compare"
On Error GoTo Err

Select Case mKeyType
Case KeyTypeInteger, KeyTypeFloat, KeyTypeDate
    If value1 = value2 Then
        compare = 0
    ElseIf value1 > value2 Then
        compare = 1
    Else
        compare = -1
    End If
Case KeyTypeString
    compare = StrComp(value1, value2, vbTextCompare)
Case KeyTypeCaseSensitiveString
    compare = StrComp(value1, value2, vbBinaryCompare)
Case KeyTypeComparable
    Dim obj1 As IComparable
    Dim obj2 As IComparable
    Set obj1 = value1
    Set obj2 = value2
    compare = obj1.CompareTo(obj2)
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function createEntry( _
                ByVal Key As Variant, _
                ByVal Data As Variant, _
                ByVal Parent As DictionaryEntry) As DictionaryEntry
Const ProcName As String = "createEntry"

On Error GoTo Err

Set createEntry = New DictionaryEntry
createEntry.Key = Key
createEntry.Data = Data
createEntry.Parent = Parent

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub deleteEntry(ByVal p As DictionaryEntry)
Const ProcName As String = "deleteEntry"

On Error GoTo Err

changed

If p Is mFirst Then Set mFirst = Nothing
mCount = mCount - 1

' If strictly internal, copy successor's element to p and then make p
' point to successor.
If (Not p.Left Is Nothing And Not p.Right Is Nothing) Then
    Dim s As DictionaryEntry: Set s = successor(p)
    p.Key = s.Key
    p.Data = s.Data
    Set p = s
End If

' Start fixup at replacement node, if it exists.
Dim replacement As DictionaryEntry
If Not p.Left Is Nothing Then
    Set replacement = p.Left
Else
    Set replacement = p.Right
End If

If (Not replacement Is Nothing) Then
    ' Link replacement to Parent
    replacement.Parent = p.Parent
    If p.Parent Is Nothing Then
        Set mRoot = replacement
    ElseIf (p Is p.Parent.Left) Then
        p.Parent.Left = replacement
    Else
        p.Parent.Right = replacement
    End If
    
    ' Null out links so they are OK to use by fixAfterDeletion.
    p.Left = Nothing
    p.Right = Nothing
    p.Parent = Nothing

    ' Fix replacement
    If (p.Color = BLACK) Then fixAfterDeletion replacement
    
ElseIf (p.Parent Is Nothing) Then   ' return if we are the only node.
    Set mRoot = Nothing
Else  '  No children. Use self as phantom replacement and unlink.
    If (p.Color = BLACK) Then fixAfterDeletion p

    If (Not p.Parent Is Nothing) Then
        If (p Is p.Parent.Left) Then
            p.Parent.Left = Nothing
        ElseIf (p Is p.Parent.Right) Then
            p.Parent.Right = Nothing
        End If
        p.Parent = Nothing
    End If
End If


Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function findEntry(ByVal Key As Variant, ByRef entry As DictionaryEntry) As Long
Const ProcName As String = "findEntry"
On Error GoTo Err

Dim currentEntry As DictionaryEntry
Set currentEntry = mRoot
Set entry = currentEntry

Do While Not currentEntry Is Nothing
    Dim cmp As Long
    cmp = compare(Key, currentEntry.Key)
    If cmp = 0 Then
        Set entry = currentEntry
        findEntry = 0
        Exit Function
    ElseIf cmp < 0 Then
        Set entry = currentEntry
        findEntry = 1
        Set currentEntry = currentEntry.Left
    Else
        Set entry = currentEntry
        findEntry = -1
        Set currentEntry = currentEntry.Right
    End If
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function firstEntry() As DictionaryEntry
Dim p As DictionaryEntry
Const ProcName As String = "firstEntry"

On Error GoTo Err

Set p = mRoot
If Not p Is Nothing Then
    Do While Not p.Left Is Nothing
        Set p = p.Left
    Loop
End If
Set firstEntry = p

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fixAfterDeletion( _
                ByVal x As DictionaryEntry)
Dim w As DictionaryEntry

Const ProcName As String = "fixAfterDeletion"

On Error GoTo Err

Do While (Not x Is mRoot) And colorOf(x) = BLACK
    If x Is leftOf(parentOf(x)) Then
        Set w = rightOf(parentOf(x))
        If colorOf(w) = Red Then
            setColor w, BLACK
            setColor parentOf(x), Red
            rotateLeft parentOf(x)
            Set w = rightOf(parentOf(x))
        End If
        
        If colorOf(leftOf(w)) = BLACK _
            And colorOf(rightOf(w)) = BLACK _
        Then
            setColor w, Red
            Set x = parentOf(x)
        Else
            If colorOf(rightOf(w)) = BLACK Then
                setColor leftOf(w), BLACK
                setColor w, Red
                rotateRight w
                w = rightOf(parentOf(x))
            End If
            setColor w, colorOf(parentOf(x))
            setColor parentOf(x), BLACK
            setColor rightOf(w), BLACK
            rotateLeft parentOf(x)
            Set x = mRoot
        End If
    Else
        Set w = leftOf(parentOf(x))
        If colorOf(w) = Red Then
            setColor w, BLACK
            setColor parentOf(x), Red
            rotateRight parentOf(x)
            Set w = leftOf(parentOf(x))
        End If
        
        If colorOf(rightOf(w)) = BLACK _
            And colorOf(leftOf(w)) = BLACK _
        Then
            setColor w, Red
            Set x = parentOf(x)
        Else
            If colorOf(leftOf(w)) = BLACK Then
                setColor leftOf(w), BLACK
                setColor w, Red
                rotateLeft w
                Set w = leftOf(parentOf(x))
            End If
            setColor w, colorOf(parentOf(x))
            setColor parentOf(x), BLACK
            setColor leftOf(w), BLACK
            rotateRight parentOf(x)
            Set x = mRoot
        End If
    End If
Loop
setColor x, BLACK

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fixAfterInsertion( _
                ByVal x As DictionaryEntry)
Dim y As DictionaryEntry

Const ProcName As String = "fixAfterInsertion"

On Error GoTo Err

Do While Not x Is mRoot
    If x.Parent.Color = BLACK Then Exit Do
    
    If parentOf(x) Is leftOf(parentOf(parentOf(x))) Then
        Set y = rightOf(parentOf(parentOf(x)))
        If colorOf(y) = Red Then

            ' uncle is Red
            setColor parentOf(x), BLACK
            setColor y, BLACK
            setColor parentOf(parentOf(x)), Red
            Set x = parentOf(parentOf(x))
        Else

            ' uncle is Black
            If x Is rightOf(parentOf(x)) Then
                Set x = parentOf(x)
                rotateLeft x
            End If
            
            setColor parentOf(x), BLACK
            setColor parentOf(parentOf(x)), Red
            rotateRight parentOf(parentOf(x))
        End If
    Else

        ' mirror image of above code
        Set y = leftOf(parentOf(parentOf(x)))
        If colorOf(y) = Red Then

            ' uncle is Red
            setColor parentOf(x), BLACK
            setColor y, BLACK
            setColor parentOf(parentOf(x)), Red
            Set x = parentOf(parentOf(x))
        Else

            ' uncle is Black
            If x Is leftOf(parentOf(x)) Then
                Set x = parentOf(x)
                rotateRight x
            End If
            
            setColor parentOf(x), BLACK
            setColor parentOf(parentOf(x)), Red
            rotateLeft parentOf(parentOf(x))
        End If
    End If
Loop
mRoot.Color = BLACK

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getEnumerator() As Enumerator
Dim lEnum As Enumerator
Const ProcName As String = "getEnumerator"

On Error GoTo Err

Set lEnum = New Enumerator
lEnum.Initialise Me
Set getEnumerator = lEnum

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function leftOf( _
                ByVal p As DictionaryEntry) As DictionaryEntry
Const ProcName As String = "leftOf"

On Error GoTo Err

If Not p Is Nothing Then Set leftOf = p.Left

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function parentOf( _
                ByVal p As DictionaryEntry) As DictionaryEntry
Const ProcName As String = "parentOf"

On Error GoTo Err

If Not p Is Nothing Then Set parentOf = p.Parent

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function predecessor( _
                ByVal t As DictionaryEntry) As DictionaryEntry
Const ProcName As String = "predecessor"
On Error GoTo Err

Dim p As DictionaryEntry

If t Is Nothing Then

ElseIf Not t.Left Is Nothing Then
    Set p = t.Left
    Do While Not p.Right Is Nothing
        Set p = p.Right
    Loop
    Set predecessor = p
Else
    Set p = t.Parent
    
    Dim ch As DictionaryEntry
    Set ch = t
    
    Do While Not p Is Nothing
        If Not ch Is p.Left Then Exit Do
        Set ch = p
        Set p = p.Parent
    Loop
    Set predecessor = p
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function rightOf( _
                ByVal p As DictionaryEntry) As DictionaryEntry
Const ProcName As String = "rightOf"

On Error GoTo Err

If Not p Is Nothing Then Set rightOf = p.Right

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub rotateLeft(ByVal x As DictionaryEntry)
Dim y As DictionaryEntry

Const ProcName As String = "rotateLeft"

On Error GoTo Err

Set y = x.Right

x.Right = y.Left
If Not y.Left Is Nothing Then y.Left.Parent = x

If Not y Is Nothing Then y.Parent = x.Parent
If Not x.Parent Is Nothing Then
    If x Is x.Parent.Left Then
        x.Parent.Left = y
    Else
        x.Parent.Right = y
    End If
Else
    Set mRoot = y
End If

y.Left = x
If Not x Is Nothing Then x.Parent = y

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub rotateRight(ByVal x As DictionaryEntry)
Dim y As DictionaryEntry

Const ProcName As String = "rotateRight"

On Error GoTo Err

Set y = x.Left

x.Left = y.Right
If Not y.Right Is Nothing Then y.Right.Parent = x

If Not y Is Nothing Then y.Parent = x.Parent
If Not x.Parent Is Nothing Then
    If x Is x.Parent.Right Then
        x.Parent.Right = y
    Else
        x.Parent.Left = y
    End If
Else
    Set mRoot = y
End If

y.Right = x
If Not x Is Nothing Then x.Parent = y

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setColor( _
                ByVal p As DictionaryEntry, _
                ByVal Color As NodeColors)
Const ProcName As String = "setColor"

On Error GoTo Err

If Not p Is Nothing Then p.Color = Color

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function successor( _
                ByVal t As DictionaryEntry) As DictionaryEntry
Const ProcName As String = "successor"
On Error GoTo Err

Dim p As DictionaryEntry

If t Is Nothing Then

ElseIf Not t.Right Is Nothing Then
    Set p = t.Right
    Do While Not p.Left Is Nothing
        Set p = p.Left
    Loop
    Set successor = p
Else
    Set p = t.Parent
    
    Dim ch As DictionaryEntry
    Set ch = t
    
    Do While Not p Is Nothing
        If Not ch Is p.Right Then Exit Do
        Set ch = p
        Set p = p.Parent
    Loop
    Set successor = p
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ToString() As String
Dim jb As JSONBuilder
Dim en As Enumerator
Dim entry As DictionaryEntry

Const ProcName As String = "ToString"


On Error GoTo Err

Set jb = New JSONBuilder
Set en = EntryEnumerator

jb.BeginObject

Do While en.MoveNext
    Set entry = en.Current
    jb.EncodeName entry.Key
    jb.EncodeValue entry.Data
Loop

jb.EndObject

ToString = jb.JSONString
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



