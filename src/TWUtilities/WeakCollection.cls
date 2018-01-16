VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeakCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


''
' Objects of this class hold sets of weak references to objects. A weak reference
' is a reference that is not taken into account by Visual Basic when deciding
' whether to terminate an object.
'
' Weak references are used to avoid circular references that can prevent unused
' objects being terminateed, leading to memory leaks. For example in a Parent-child
' relationship, if each object has a normal reference to the other, then neither
' will ever be terminateed unless at least one of them takes specific Action to set
' the reference to <code>Nothing</code>, which requires additional programming.
' But if one object maintains a weak reference to the other, the situation does
' not arise.
'
' The <code>WeakCollection</code> class may be useful where an object wishes to
' maintain a set of references to other objects such that circular references might
' arise if normal references were used.
'
' Note that if an object referenced in a weak collection is terminated, then
' any attempt to use the reference in the weak collection will cause a program
' crash. Therefore use of weak collections requires care to ensure that
' references to objects included in them are removed before the objects themselves
' are terminated.
'
' <code>WeakCollection</code> objects may be used with the Visual Basic <code>For Each</code>
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

'@================================================================================
' Constants
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mCollection As Collection

Private mUpdateSequence As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mCollection = New Collection
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Set IEnumerable_Enumerator = getEnumerator
End Property

Private Function IEnumerable_NextItem( _
                Item As Variant, _
                pCookie As Variant) As Boolean
Dim ref As WeakReference

If IsEmpty(pCookie) Then
    pCookie = 1&
Else
    pCookie = pCookie + 1
End If

If pCookie > mCollection.Count Then
    IEnumerable_NextItem = False
    Exit Function
End If

ref = mCollection.Item(pCookie)
Set Item = ref.Target

IEnumerable_NextItem = True
End Function

Private Sub IEnumerable_Remove(pCookie As Variant)
mCollection.Remove pCookie
End Sub

Private Function IEnumerable_Skip( _
                ByVal numToSkip As Long, _
                pCookie As Variant) As Boolean
If IsEmpty(pCookie) Then pCookie = 1&

pCookie = pCookie + numToSkip

If pCookie > mCollection.Count Then Exit Function

IEnumerable_Skip = True

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
' Returns an<code>Enumerator</code> object that can be used to iterate through
' the members of the collection.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get Enumerator() As Enumerator
Set Enumerator = getEnumerator
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Adds a weak reference to an object to the collection.
'
' @param Item
'   The object to be added.
' @param key
'   A value uniquely identifying the object in the collection.
' @param before
'   Specifies an existing Item in the collection before which
'   this object should be added.
' @param after
'   Specifies an existing Item in the collection after which
'   this object should be added.
'@/
Public Sub Add( _
                ByVal Item As Object, _
                Optional ByVal Key As Variant, _
                Optional ByVal Before As Variant, _
                Optional ByVal After As Variant)
Dim weakRef As WeakReference

Set weakRef = New WeakReference
weakRef.Initialise Item

mCollection.Add weakRef, Key, Before, After
changed
End Sub

''
' Returns the number of items in the collection.
'
' @return
'   The number of items in the collection.
'@/
Public Function Count() As Long
Count = mCollection.Count
End Function

''
' Returns a normal reference to the specified object.
'
' @return
'   A normal reference to the specified object.
' @param pIndex
'   Either the ordinal number of the Item in the collection,
'   or the key that was specified when it was added.
'@/
Public Function Item( _
                ByVal pIndex As Variant) As Object
Attribute Item.VB_UserMemId = 0
Dim weakRef As WeakReference

Set weakRef = mCollection.Item(pIndex)
Set Item = weakRef.Target
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = getEnumerator
End Function

''
' Removes the specified object from the collection.
'
' @param pIndex
'   Either the ordinal number of the Item in the collection,
'   or the key that was specified when it was added.
'@/
Public Sub Remove( _
                ByVal pIndex As Variant)
mCollection.Remove pIndex
changed
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub changed()
mUpdateSequence = mUpdateSequence + 1
End Sub

Private Function getEnumerator() As Enumerator
Dim lEnum As Enumerator
Set lEnum = New Enumerator
lEnum.Initialise Me
Set getEnumerator = lEnum
End Function


