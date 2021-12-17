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
' Represents a collection of key/value pairs, sorted on the key.
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

Private Sub Class_Initialize()
mKeyType = KeyTypeString
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Const ProcName As String = "IEnumerable_Enumerator"
On Error GoTo Err

Set IEnumerable_Enumerator = getEnumerator(Empty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IEnumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_NextItem"
On Error GoTo Err

Dim lCookie As EnumerationCookie
lCookie = pCookie
If lCookie.Current Is Nothing And Not lCookie.Next Is Nothing Then
    FindNext lCookie
ElseIf lCookie.Current Is Nothing Then
    FindFirst lCookie
Else
    FindNext lCookie
End If

If lCookie.Current Is Nothing Then
    Item = Empty
Else
    Dim lEntry As DictionaryEntry
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
Const ProcName As String = "IEnumerable_Remove"
On Error GoTo Err

Dim lCookie As EnumerationCookie
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
' the entries of the collection, ignoring changes to the <code>UpdateSequence</code>
' property, and deleting the entries as they are processed.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Friend Property Get DisconnectedEntryEnumerator() As Enumerator
Const ProcName As String = "DisconnectedEntryEnumerator"
On Error GoTo Err

Dim ee As New DictionaryEntriesEnum
ee.Initialise Me, Empty, mKeyType, pDeleteAsYouGo:=True

Dim en As IEnumerable
Set en = ee
Set DisconnectedEntryEnumerator = en.Enumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the entries of the collection.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get EntryEnumerator() As Enumerator
Const ProcName As String = "EntryEnumerator"
On Error GoTo Err

Dim ee As New DictionaryEntriesEnum
ee.Initialise Me, Empty, mKeyType
Set EntryEnumerator = ee.Enumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the entries of the collection, starting from a specified key.
'
' @param key
'   An expression that specifies the lowest key to be included in the iteration.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get EntryEnumeratorFromKey(ByVal Key As Variant) As Enumerator
Const ProcName As String = "EntryEnumeratorFromKey"
On Error GoTo Err

Dim ee As New DictionaryEntriesEnum
ee.Initialise Me, Key, mKeyType
Set EntryEnumeratorFromKey = ee.Enumerator

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

Set Enumerator = getEnumerator(Empty)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the data members of the collection, starting from a specified key.
'
' @param key
'   An expression that specifies the lowest key to be included in the iteration.
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get EnumeratorFromKey(ByVal Key As Variant) As Enumerator
Const ProcName As String = "EnumeratorFromKey"
On Error GoTo Err

Set EnumeratorFromKey = getEnumerator(Key)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get FirstItem() As Variant
Const ProcName As String = "FirstItem"
On Error GoTo Err

gAssert mCount > 0, "No items"

gSetVariant FirstItem, gFirstEntry(mRoot).Data

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns a specific item in the collection, identified by key.
'
' @remarks
'   If the key provided doesn't match any existing member of the
'   collection, an error occurs.
'
'   If the collection contains more than one item with the specified key, the
'   first such item is returned.
' @return
'   The required item.
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
' Updates a specific item in the collection, identified by key.
'
' @remarks
'   If the key provided doesn't match any existing member of the
'   collection, an error occurs.
'
'   If the collection contains more than one item with the specified key, the
'   first such item is updated.
' @param key
'   An expression that specifies the key of a member of the collection.
' @param Value
'   The new value to be assigned to the specified item.
'@/
Public Property Let Item( _
                ByVal Key As Variant, _
                ByVal Value As Variant)
Const ProcName As String = "Item"
On Error GoTo Err

Dim entry As DictionaryEntry
Dim Result As Long
Result = gFindEntry(Key, mRoot, mKeyType, entry)
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
Const ProcName As String = "KeyEnumerator"
On Error GoTo Err

Dim ke As New DictionaryKeysEnum
ke.Initialise Me

Dim en As IEnumerable
Set en = ke
Set KeyEnumerator = en.Enumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Root() As DictionaryEntry
Set Root = mRoot
End Property

Friend Property Get UpdateSequence() As Long
UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Adds an item to the collection.
'
' @remarks
'   An error occurs if a specified key duplicates the key for an existing member
'   of the collection, and the collection was not specified to allow duplicates
'   when it was created.
'
' @param item
'   The item to be added to the collection.
' @param key
'   A key that determines the position of the item in the collection.
'@/
Public Sub Add( _
                ByVal Item As Variant, _
                ByVal Key As Variant)
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

Dim cmp As Long
If Not mFirst Is Nothing Then
    cmp = gCompare(Key, mFirst.Key, mKeyType)
    If cmp < 0 Then Set mFirst = Nothing
End If

' find where node belongs
Dim Current As DictionaryEntry
Set Current = mRoot

Dim Parent As DictionaryEntry
Do While Not Current Is Nothing
    Set Parent = Current
    cmp = gCompare(Key, Current.Key, mKeyType)
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
' Removes all items from the collection asynchronously, and releases the memory
' they used.
'
' @remarks
'   Return from this method is immediate. New entries may be added to this
'   <code>SortedDictionary</code> object immediately.
'
'   The <code>Result</code> field of the <code>TaskCompletionEventData</code>
'   for this task is a <code>Properties</code> object containing details of the
'   number of dictionary entries released and other information.
'
'   Note that cancelling this operation using the returned
'   <code>TaskController</code> does not prevent release of the memory used by the
'   <code>SortedDictionary</code> object prior to the call to <code>ClearAsync</>.
'   The main thread will be busy while the remaining memory is released, so the
'   program may become unresponsive if there is a large amount remaining to be
'   released.
'
' @return
'   A <code>TaskController</code> object.
'@/
Public Function ClearAsync() As TaskController
Const ProcName As String = "Clear"
On Error GoTo Err

changed

Const HeuristicNumber As Long = 50
Const HeuristicTime As Long = 5

Dim t As New SortedDictClearanceTask
t.Initialise DisconnectedEntryEnumerator, HeuristicNumber, HeuristicTime
Set ClearAsync = gTaskManager.StartTask(t, PriorityLow)

mCount = 0
Set mRoot = Nothing
Set mFirst = Nothing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Indicates whether an item with the specified key already exists in the collection.
'
' @return
'   <code>True</code> if an item with the specified key already exists, otherwise
'   <code>False</code>.
' @param key
'   The relevant key.
'@/
Public Function Contains(ByVal Key As Variant) As Boolean
Const ProcName As String = "Contains"
On Error GoTo Err

If mCount = 0 Then Exit Function

Dim entry As DictionaryEntry
Contains = (gFindEntry(Key, mRoot, mKeyType, entry) = 0)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub FindFirst( _
                ByRef pCookie As EnumerationCookie)
Const ProcName As String = "FindFirst"
On Error GoTo Err

If mFirst Is Nothing Then
    Set mFirst = gFirstEntry(mRoot)
End If
Set pCookie.Current = mFirst
Set pCookie.Next = gSuccessor(mFirst, pDeleteAsYouGo:=False)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub FindNext( _
                ByRef pCookie As EnumerationCookie)
Const ProcName As String = "FindNext"
On Error GoTo Err

Set pCookie.Current = pCookie.Next
Set pCookie.Next = gSuccessor(pCookie.Current, pDeleteAsYouGo:=False)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal keyType As DictionaryKeyTypes, _
                ByVal allowDuplicates As Boolean)
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

Set NewEnum = getEnumerator(Empty)

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

If gFindEntry(Key, mRoot, mKeyType, entry) <= 0 Then
    ' either we found an entry with the supplied key, or
    ' the entry found is lower than the supplied key, so it must be the
    ' successor that we want
    Do
        Set entry = gSuccessor(entry, pDeleteAsYouGo:=False)
        If entry Is Nothing Then Exit Do
    Loop Until gCompare(Key, entry.Key, mKeyType) < 0
    
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
Result = gFindEntry(Key, mRoot, mKeyType, entry)

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
' Removes an item from the collection.
'
' @remarks
'   If the key provided doesn't match an existing member of the
'   collection, an error occurs.
'
'   If the collection contains more than one item with the specified key, the
'   first such item to be added is removed.
' @return
'   Returns the item being removed.
' @param key
'   Identifies the item to be removed from the collection.
'@/
Public Function Remove(ByVal Key As Variant) As Variant
Const ProcName As String = "Remove"
On Error GoTo Err

Dim entry As DictionaryEntry
If gFindEntry(Key, mRoot, mKeyType, entry) <> 0 Then
    ' raise the same error as the VB6 Collection object
    Err.Raise VBErrorCodes.VbErrInvalidProcedureCall, _
                , _
                "Invalid procedure call or argument"
End If

changed

If entry Is mFirst Then Set mFirst = Nothing
mCount = mCount - 1

gSetVariant Remove, entry.Data
deleteEntry entry, mRoot

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub RemoveCurrent( _
                ByRef pCookie As EnumerationCookie)
Const ProcName As String = "RemoveCurrent"
On Error GoTo Err

If IsEmpty(pCookie.Current) Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Current Item already deleted"

Dim lCurrentEntry As DictionaryEntry
Set lCurrentEntry = pCookie.Current

'If (Not lCurrentEntry.Left Is Nothing And _
'    Not lCurrentEntry.Right Is Nothing) _
'Then Set pCookie.Next = pCookie.Current

If lCurrentEntry Is mFirst Then Set mFirst = Nothing
mCount = mCount - 1

Set pCookie.Current = Nothing

deleteEntry lCurrentEntry, mRoot

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Gets a specific item in the collection, identified by key.
'
' @remarks
'   If the key provided doesn't match any existing member of the
'   collection, <code>False</code> is returned.
'
'   If the collection contains more than one item with the specified key, the
'   first such item is returned.
' @return
'   <code>True</code> if the required item is found. Otherwise <code>False</code>.
' @param Key
'   An expression that specifies the key of a member of the collection.
' @param item
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
Result = gFindEntry(Key, mRoot, mKeyType, entry)

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
    colorOf = Black
Else
    colorOf = p.Color
End If

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

Private Sub deleteEntry(ByVal p As DictionaryEntry, ByRef pRoot As DictionaryEntry)
Const ProcName As String = "deleteEntry"
On Error GoTo Err

' If strictly internal, copy successor's element to p and then make p
' point to successor.
If (Not p.Left Is Nothing And Not p.Right Is Nothing) Then
    Dim s As DictionaryEntry: Set s = gSuccessor(p, pDeleteAsYouGo:=False)
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
    ' Link replacement to parent
    replacement.Parent = p.Parent
    If p.Parent Is Nothing Then
        Set pRoot = replacement
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
    If (p.Color = Black) Then fixAfterDeletion replacement, pRoot
    
ElseIf (p.Parent Is Nothing) Then   ' return if we are the only node.
    Set pRoot = Nothing
Else  '  No children. Use self as phantom replacement and unlink.
    If (p.Color = Black) Then fixAfterDeletion p, pRoot

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

Private Sub fixAfterDeletion( _
                ByVal x As DictionaryEntry, _
                ByVal pRoot As DictionaryEntry)
Const ProcName As String = "fixAfterDeletion"
On Error GoTo Err

Do While (Not x Is pRoot) And colorOf(x) = Black
    Dim w As DictionaryEntry
    If x Is leftOf(ParentOf(x)) Then
        Set w = rightOf(ParentOf(x))
        If colorOf(w) = Red Then
            setColor w, Black
            setColor ParentOf(x), Red
            rotateLeft ParentOf(x)
            Set w = rightOf(ParentOf(x))
        End If
        
        If colorOf(leftOf(w)) = Black _
            And colorOf(rightOf(w)) = Black _
        Then
            setColor w, Red
            Set x = ParentOf(x)
        Else
            If colorOf(rightOf(w)) = Black Then
                setColor leftOf(w), Black
                setColor w, Red
                rotateRight w
                Set w = rightOf(ParentOf(x))
            End If
            setColor w, colorOf(ParentOf(x))
            setColor ParentOf(x), Black
            setColor rightOf(w), Black
            rotateLeft ParentOf(x)
            Set x = pRoot
        End If
    Else
        Set w = leftOf(ParentOf(x))
        If colorOf(w) = Red Then
            setColor w, Black
            setColor ParentOf(x), Red
            rotateRight ParentOf(x)
            Set w = leftOf(ParentOf(x))
        End If
        
        If colorOf(rightOf(w)) = Black _
            And colorOf(leftOf(w)) = Black _
        Then
            setColor w, Red
            Set x = ParentOf(x)
        Else
            If colorOf(leftOf(w)) = Black Then
                setColor leftOf(w), Black
                setColor w, Red
                rotateLeft w
                Set w = leftOf(ParentOf(x))
            End If
            setColor w, colorOf(ParentOf(x))
            setColor ParentOf(x), Black
            setColor leftOf(w), Black
            rotateRight ParentOf(x)
            Set x = pRoot
        End If
    End If
Loop
setColor x, Black

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fixAfterInsertion( _
                ByVal x As DictionaryEntry)
Const ProcName As String = "fixAfterInsertion"
On Error GoTo Err

x.Color = Red

Dim y As DictionaryEntry
Do While Not x Is Nothing
    If x Is mRoot Then Exit Do
    If x.Parent.Color = Black Then Exit Do
    
    If ParentOf(x) Is leftOf(ParentOf(ParentOf(x))) Then
        Set y = rightOf(ParentOf(ParentOf(x)))
        If colorOf(y) = Red Then

            ' uncle is Red
            setColor ParentOf(x), Black
            setColor y, Black
            setColor ParentOf(ParentOf(x)), Red
            Set x = ParentOf(ParentOf(x))
        Else

            ' uncle is Black
            If x Is rightOf(ParentOf(x)) Then
                Set x = ParentOf(x)
                rotateLeft x
            End If
            
            setColor ParentOf(x), Black
            setColor ParentOf(ParentOf(x)), Red
            rotateRight ParentOf(ParentOf(x))
        End If
    Else

        ' mirror image of above code
        Set y = leftOf(ParentOf(ParentOf(x)))
        If colorOf(y) = Red Then

            ' uncle is Red
            setColor ParentOf(x), Black
            setColor y, Black
            setColor ParentOf(ParentOf(x)), Red
            Set x = ParentOf(ParentOf(x))
        Else

            ' uncle is Black
            If x Is leftOf(ParentOf(x)) Then
                Set x = ParentOf(x)
                rotateRight x
            End If
            
            setColor ParentOf(x), Black
            setColor ParentOf(ParentOf(x)), Red
            rotateLeft ParentOf(ParentOf(x))
        End If
    End If
Loop
mRoot.Color = Black

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getEnumerator(ByVal Key As Variant) As Enumerator
Const ProcName As String = "getEnumerator"
On Error GoTo Err

Dim lInitialCookie As EnumerationCookie
Set lInitialCookie.Next = Nothing
Set lInitialCookie.Current = Nothing
If Not IsEmpty(Key) Then
    Dim lInitialEntry As DictionaryEntry
    If gFindEntry(Key, mRoot, mKeyType, lInitialEntry) < 0 Then Set lInitialEntry = gSuccessor(lInitialEntry, False)
    Set lInitialCookie.Next = lInitialEntry
End If

Dim lEnum As Enumerator
Set lEnum = New Enumerator
lEnum.Initialise Me, lInitialCookie
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

Private Function ParentOf( _
                ByVal p As DictionaryEntry) As DictionaryEntry
Const ProcName As String = "parentOf"
On Error GoTo Err

If Not p Is Nothing Then Set ParentOf = p.Parent

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

Private Sub rotateLeft(ByVal p As DictionaryEntry)
Const ProcName As String = "rotateLeft"
On Error GoTo Err

If p Is Nothing Then Exit Sub

Dim r As DictionaryEntry: Set r = p.Right

p.Right = r.Left
If Not r.Left Is Nothing Then r.Left.Parent = p

r.Parent = p.Parent
If p.Parent Is Nothing Then
    Set mRoot = r
ElseIf p Is p.Parent.Left Then
    p.Parent.Left = r
Else
    p.Parent.Right = r
End If

r.Left = p
p.Parent = r

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub rotateRight(ByVal p As DictionaryEntry)
Const ProcName As String = "rotateRight"
On Error GoTo Err

If p Is Nothing Then Exit Sub

Dim l As DictionaryEntry: Set l = p.Left

p.Left = l.Right
If Not l.Right Is Nothing Then l.Right.Parent = p

l.Parent = p.Parent
If p.Parent Is Nothing Then
    Set mRoot = l
ElseIf p Is p.Parent.Right Then
    p.Parent.Right = l
Else
    p.Parent.Left = l
End If

l.Right = p
p.Parent = l

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

Private Function ToString() As String
Const ProcName As String = "ToString"
On Error GoTo Err

Dim jb As JSONBuilder
Set jb = New JSONBuilder

Dim en As Enumerator
Set en = EntryEnumerator

jb.BeginObject

Do While en.MoveNext
    Dim entry As DictionaryEntry
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



