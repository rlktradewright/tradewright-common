VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumerableCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class is a more fully featured version of Visual Basic's built-in Collection
' class.
'
' As well as all the standard methods and properties, it also has a <code>Clear</code> method
' that removes all the items from the collection, a <code>Contains</code> method that indicates whether
' the collection contains an item with a specified key, and facilities for obtaining
' notifications when the collection is changed.
'
' Moreover the <code>Enumerator</code> property returns an <code>Enumerator</code> object
' that can be used to iterate through the collection without using the <code>For Each</code>
' construct.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IEnumerable

'@================================================================================
' Events
'@================================================================================

''
' Raised when an Item is added to or removed from this
' <code>EnumerableCollection</code> object, and when the collection is cleared.
'
' @remarks
'   The following members of the <code>CollectionChangeTypes</code> enum may
'   occur in the <code>ChangeType</code> field of the
'   <code>CollectionChangedEventData</code> parameter:
' <pre>
'    <tt>CollItemAdded</tt><br />
'    <tt>CollItemRemoved</tt><br />
'    <tt>CollCollectionCleared</tt>
' </pre>
' @param ev
'   contains details of the change.
'
'@/
Event CollectionChanged( _
                ev As CollectionChangeEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "EnumerableCollection"

'@================================================================================
' Member variables
'@================================================================================

Private mColl                               As Collection

Private mChangeListeners                    As Listeners

Private mUpdateSequence                     As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mColl = New Collection
Set mChangeListeners = New Listeners
End Sub

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

Private Function IEnumerable_NextItem( _
                Item As Variant, _
                pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_NextItem"

On Error GoTo Err

Dim lCookie As EnumerationCookie
Dim en As IEnumVARIANT
Dim numFetched As Long

If IsEmpty(pCookie) Then
    Set en = mColl.[_NewEnum]
    Set lCookie.Next = en
    lCookie.Current = 0
Else
    lCookie = pCookie
    Set en = lCookie.Next
End If

lCookie.Current = Abs(lCookie.Current) + 1

en.Next 1, Item, VarPtr(numFetched)

IEnumerable_NextItem = (numFetched <> 0)

pCookie = lCookie

Exit Function


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
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Current item already removed"
End If

Remove CLng(lCookie.Current)
lCookie.Current = -(lCookie.Current - 1) ' negative indicates item has been deleted

pCookie = lCookie

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IEnumerable_Skip( _
                ByVal numToSkip As Long, _
                pCookie As Variant) As Boolean

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

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the members of the collection.
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

'@================================================================================
' Methods
'@================================================================================

''
' Adds an Item to the collection.
'
' @remarks
'   Whether the before or after argument is a string expression or numeric expression,
'   it must refer to an existing member of the collection, or an error occurs.
'
'   An Error also occurs if a specified key duplicates the key for an existing member
'   of the collection.
'
' @param Item
'   An expression of any type that specifies the member to Add to the collection.
' @param key
'   A unique string expression that specifies a key string that can be used, instead
'   of a positional index, to access a member of the collection. If <code>Item</code> is
'   an object and <code>Key</code> is not supplied, the object may subsequently be removed
'   from the collection by passing it in the <code>pIndexOrItem</code> parameter to the
'   <code>Remove</code> method.
' @param before
'   An expression that specifies a relative position in the collection. The member
'   to be added is placed in the collection before the member identified by the before
'   argument. If a numeric expression, before must be a number from 1 to the Value of
'   the collection's Count property. If a string expression, before must correspond
'   to the key specified when the member being referred to was added to the collection.
'   You can specify a before position or an after position, but not both.
' @param after
'   An expression that specifies a relative position in the collection. The member
'   to be added is placed in the collection after the member identified by the after argument.
'   If numeric, after must be a number from 1 to the Value of the collection's Count
'   roperty. If a string, after must correspond to the key specified when the member
'   referred to was added to the collection. You can specify a before position or an
'   after position, but not both.
'@/
Public Sub Add(ByRef Item As Variant, _
                Optional ByRef Key As Variant, _
                Optional ByRef Before As Variant, _
                Optional ByRef After As Variant)
Const ProcName As String = "Add"
On Error GoTo Err

Change

If IsMissing(Key) And IsObject(Item) Then
    mColl.Add Item, gGetObjectKey(Item), Before, After
Else
    mColl.Add Item, Key, Before, After
End If
fireChange CollItemAdded, Item

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Registers an object to act as a listener for changes to the collection.
'
' @remarks
'   The listener object must implement the <code>CollectionChangeListener</code>
'   interface.
' @param listener
'   The object which is to act as a listener.
'@/
Public Sub AddCollectionChangeListener( _
                ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mChangeListeners.Add pListener

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

Set mColl = New Collection
fireChange CollCollectionCleared, Empty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Returns the number of items in the collection.
'
' @return
'   The number of items in the collection.
'@/
Public Function Count() As Long
Const ProcName As String = "Count"

On Error GoTo Err

Count = mColl.Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Indicates whether the collection contains an item that was added with the specified key, or the
' specified object.
'
' @return
'   <code>True</code> if either an item with the specified key, or the specified object, exists in the collection.
' @param Key
'   The key to be tested, or an object whose presence is to be tested.
'@/
Public Function Contains( _
                ByVal pKeyOrItem As Variant) As Boolean
Const ProcName As String = "Contains"
On Error GoTo Err

Dim lKey As String
If IsObject(pKeyOrItem) Then
    lKey = gGetObjectKey(pKeyOrItem)
ElseIf VarType(pKeyOrItem) = VbVarType.vbString Then
    lKey = pKeyOrItem
Else
    gAssert False, "Type mismatch", VBErrorCodes.VbErrTypeMismatch
End If

mColl.Item lKey
Contains = True

Exit Function

Err:
If Err.number = VBErrorCodes.VbErrInvalidProcedureCall Then Exit Function
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns a specific Item in the collection either by position or by key.
'
' @remarks
'   If the Value provided as index doesn’t match any existing member of the
'   collection, an error occurs.
'
'   The Item method is the default method for a collection. Therefore, the
'   following lines of code are equivalent:
'
' <pre>
'   Print MyCollection(1)
'   Print MyCollection.Item(1)
' </pre>
' @return
'   The required Item.
' @param pIndex
'   An expression that specifies the position of a member of the collection.
'   If a numeric expression, index must be a number from 1 to the Value of the
'   collection's Count property. If a string expression, index must
'   correspond to the key argument specified when the member referred to was
'   added to the collection.
'@/
Public Function Item( _
                ByRef pIndex As Variant) As Variant
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"

On Error GoTo Err

gSetVariant Item, mColl.Item(pIndex)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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
' Notifies listeners that an item in the collection has changed.
'
' @remarks
'   If the value provided as pIndex doesn’t match any existing member of the
'   collection, an error occurs.
'
'   If the member specified by pIndex is not an object, an error occurs.
'
' @param pIndex
'   An expression that specifies the position of a member of the collection.
'   If a numeric expression, index must be a number from 1 to the Value of the
'   collection's Count property. If a string expression, index must
'   correspond to the key argument specified when the member referred to was
'   added to the collection.
'@/
Public Function NotifyChange( _
                ByVal pIndex As Variant)
Const ProcName As String = "NotifyChange"
On Error GoTo Err

Dim lItem As Variant
gSetVariant lItem, mColl.Item(pIndex)

fireChange CollItemChanged, lItem

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Removes an lItem from the collection.
'
' @remarks
'   If the Value provided as pIndex doesn’t match an existing member of the
'   collection, an error occurs.
'
' @param pIndexOrItem
'   An expression that specifies the position of a member of the collection.
'   If a numeric expression, pIndex must be a number from 1 to the Value of
'   the collection's Count property. If a string expression, pIndex must
'   correspond to the key argument specified when the member referred to
'   was added to the collection. If an object, this object must have been
'   added to the collection without using a key.
'@/
Public Sub Remove( _
                ByRef pIndexOrItem As Variant)
Const ProcName As String = "Remove"
On Error GoTo Err

Dim lItem As Variant

Change

If IsObject(pIndexOrItem) Then
    Dim lKey As String
    lKey = gGetObjectKey(pIndexOrItem)
    gSetVariant lItem, mColl.Item(lKey)
    gAssert lItem Is pIndexOrItem, , VBErrorCodes.VbErrInvalidProcedureCall
    mColl.Remove gGetObjectKey(pIndexOrItem)
Else
    gSetVariant lItem, mColl.Item(pIndexOrItem)
    mColl.Remove pIndexOrItem
End If
fireChange CollItemRemoved, lItem

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Unregisters a listener.
'
' @param listener
'   The object which is to be unregistered.
'@/
Public Sub RemoveCollectionChangeListener( _
                ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Change()
mUpdateSequence = mUpdateSequence + 1
End Sub

Private Sub fireChange( _
                ByVal ChangeType As CollectionChangeTypes, _
                ByVal AffectedItem As Variant)
Const ProcName As String = "fireChange"
On Error GoTo Err

Dim ev As CollectionChangeEventData
Set ev.Source = Me
ev.ChangeType = ChangeType
gSetVariant ev.AffectedItem, AffectedItem

mChangeListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mChangeListeners.Count
    Dim lListener As ICollectionChangeListener
    Set lListener = mChangeListeners.GetListener(i)
    lListener.Change ev
Next
RaiseEvent CollectionChanged(ev)

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


