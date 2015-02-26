VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLConfigItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class is an implementation of the <code>ConfigItems</code> interface that
' stores the configuration Data in an XML file.
'
' Objects of this class cannot be created using the <code>New</code> operator.
'
' Objects of this class must be used via their <code>ConfigItems</code> interface,
' since the default interface exposes no public properties or methods. For example:
'
' <pre>
' <tt>
' Dim cf As ConfigStoreProvider<br />
' Dim root As ConfigItem<br />
' Dim outputFilename As String<br />
' <br />
' Set cf = LoadConfigProviderFromXMLFile("C:\MyConfig.xml")<br />
' Set root = cf.rootItem    ' this is actually an XMLConfigItem object<br />
' outputFilename = root.childItems.Item("outputFilename").value<br />
' </tt>
' </pre>
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements ConfigItems
Implements Enumerable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type ConfigItemTableEntry
    ci              As ConfigItem
    Next            As Long
    Prev            As Long
End Type

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "XMLConfigItems"

Private Const growthFactor                  As Single = 2
Private Const HeadEntryIndex                As Long = 0
Private Const initialCapacity               As Long = 10
Private Const TailEntryIndex                As Long = 1

'@================================================================================
' Member variables
'@================================================================================


Private mConfigItems()                      As ConfigItemTableEntry
Private mConfigItemsIndex                   As Long
Private mConfigItemsIndexTree               As SortedDictionary

Private mConfigStoreProvider                As xmlConfigProvider

Private mElement                            As IXMLDOMElement

Private mIndent                             As Long

' always the last element, to ensure correct indentation
Private mEndTagSpacer                             As IXMLDOMNode

Private mUpdateSequence                     As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
initTable
End Sub

'@================================================================================
' ConfigItems Interface Members
'@================================================================================

Private Function ConfigItems_AddClone( _
                ByVal pItem As ConfigItem, _
                ByVal pDeep As Boolean) As ConfigItem
Const ProcName As String = "ConfigItems_AddClone"

On Error GoTo Err

Dim itemXML As XMLConfigItem
Set itemXML = pItem

Dim lClonedElement As IXMLDOMElement
Set lClonedElement = itemXML.DOMElement.cloneNode(pDeep)

If mEndTagSpacer Is Nothing Then
    Set mEndTagSpacer = mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent - 1, vbTab))
    mElement.appendChild mEndTagSpacer
End If

mElement.insertBefore mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent, vbTab)), mEndTagSpacer
mElement.insertBefore lClonedElement, mEndTagSpacer

Set ConfigItems_AddClone = AddItemFromNode(lClonedElement, Nothing)
If pItem.Comment <> "" Then ConfigItems_AddClone.Comment = pItem.Comment

changed

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ConfigItems_addItem( _
                ByVal Name As String, _
                Optional ByVal Value As String, _
                Optional ByVal Comment As String) As ConfigItem
Const ProcName As String = "ConfigItems_addItem"

On Error GoTo Err

Set ConfigItems_addItem = AddItem(Name, Value, Comment)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ConfigItems_Clear()
Const ProcName As String = "ConfigItems_Clear"

On Error GoTo Err

Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ConfigItems_Count() As Long
Const ProcName As String = "ConfigItems_Count"

On Error GoTo Err

ConfigItems_Count = Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get ConfigItems_Enumerator() As Enumerator
Const ProcName As String = "ConfigItems_Enumerator"

On Error GoTo Err

Set ConfigItems_Enumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function ConfigItems_insertItemBefore( _
                ByVal Name As String, _
                Optional ByVal Value As String, _
                Optional ByVal Comment As String, _
                Optional ByVal targetItem As ConfigItem) As ConfigItem
Const ProcName As String = "ConfigItems_insertItemBefore"

On Error GoTo Err

Set ConfigItems_insertItemBefore = InsertItemBefore(Name, _
                                                Value, _
                                                Comment, _
                                                targetItem)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function ConfigItems_Item( _
                ByVal Name As String) As ConfigItem
Const ProcName As String = "ConfigItems_Item"

On Error GoTo Err

Set ConfigItems_Item = Item(Name)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ConfigItems_moveItemBefore( _
                ByVal itemSpecifier As Variant, _
                Optional ByVal targetItem As ConfigItem)
Const ProcName As String = "ConfigItems_moveItemBefore"

On Error GoTo Err

MoveItemBefore itemSpecifier, targetItem

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function ConfigItems_NewEnum() As IUnknown
Const ProcName As String = "ConfigItems_NewEnum"

On Error GoTo Err

Set ConfigItems_NewEnum = getEnumerator

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub ConfigItems_Remove(ByVal itemSpecifier As Variant)
Const ProcName As String = "ConfigItems_Remove"

On Error GoTo Err

If IsObject(itemSpecifier) Then
    Dim obj As Object
    Set obj = itemSpecifier
    If Not TypeOf obj Is ConfigItem Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid object type"
    removeItemByIdentity obj
Else
    removeItemByName CStr(itemSpecifier)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As Enumerator
Const ProcName As String = "Enumerable_Enumerator"

On Error GoTo Err

Set Enumerable_Enumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function Enumerable_NextItem( _
                ByRef Item As Variant, _
                ByRef pCookie As Variant) As Boolean
Dim i As Long

Const ProcName As String = "Enumerable_NextItem"

On Error GoTo Err

If IsEmpty(pCookie) Then
    i = mConfigItems(HeadEntryIndex).Next
ElseIf CLng(pCookie) = &H7FFFFFFF Then
    Exit Function
Else
    i = mConfigItems(CLng(pCookie)).Next
End If

Do While i <> TailEntryIndex
    If Not mConfigItems(i).ci Is Nothing Then
        Set Item = mConfigItems(i).ci
        pCookie = i
        Enumerable_NextItem = True
        Exit Function
    End If
    i = mConfigItems(i).Next
Loop

pCookie = &H7FFFFFFF
Enumerable_NextItem = False

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub Enumerable_Remove(pCookie As Variant)
Const ProcName As String = "Enumerable_Remove"

On Error GoTo Err

removeItemByIdentity mConfigItems(CLng(pCookie)).ci

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function Enumerable_Skip( _
                ByVal numToSkip As Long, _
                ByRef pCookie As Variant) As Boolean
Dim i As Long

Const ProcName As String = "Enumerable_Skip"

On Error GoTo Err

If IsEmpty(pCookie) Then pCookie = mConfigItems(HeadEntryIndex).Next

i = CLng(pCookie)
Do While i <> TailEntryIndex
    If Not mConfigItems(i).ci Is Nothing Then
        pCookie = mConfigItems(i).Next
        numToSkip = numToSkip - 1
        If numToSkip = 0 Then
            Enumerable_Skip = True
            Exit Function
        End If
    End If
    i = mConfigItems(i).Next
Loop

pCookie = &H7FFFFFFF
Enumerable_Skip = False

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Property Get Enumerable_UpdateSequence() As Long
Const ProcName As String = "Enumerable_UpdateSequence"

On Error GoTo Err

Enumerable_UpdateSequence = mUpdateSequence

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

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
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

Friend Function AddItem( _
                ByVal Name As String, _
                Optional ByVal Value As String, _
                Optional ByVal Comment As String) As XMLConfigItem
Const ProcName As String = "addItem"

On Error GoTo Err

Set AddItem = addItemFromElement(mConfigStoreProvider.XMLDoc.createElement(Name), Value, Comment)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function AddItemFromNode( _
                ByVal element As IXMLDOMElement, _
                ByVal Comment As IXMLDOMComment) As XMLConfigItem

Const ProcName As String = "AddItemFromNode"

On Error GoTo Err

Set AddItemFromNode = New XMLConfigItem

mConfigItemsIndexTree.Add addEntry(AddItemFromNode, TailEntryIndex), element.nodeName

AddItemFromNode.InitialiseFromNode mConfigStoreProvider, element, Comment, mIndent + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Clear()
Dim i As Long
Dim ci As XMLConfigItem

Const ProcName As String = "Clear"

On Error GoTo Err

i = HeadEntryIndex
Do While i <> TailEntryIndex
    If Not mConfigItems(i).ci Is Nothing Then
        Set ci = mConfigItems(i).ci
        ci.Terminate
    End If
    i = mConfigItems(i).Next
Loop

initTable
changed

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Count() As Long
Const ProcName As String = "Count"

On Error GoTo Err

Count = mConfigItemsIndexTree.Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal confFile As xmlConfigProvider, _
                ByVal element As IXMLDOMElement, _
                ByVal indent As Long)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mConfigStoreProvider = confFile
Set mElement = element
mIndent = indent

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function InsertItemBefore( _
                ByVal Name As String, _
                Optional ByVal Value As String, _
                Optional ByVal Comment As String, _
                Optional ByVal targetItem As XMLConfigItem) As XMLConfigItem
Dim newElement As IXMLDOMElement
Dim targetIndex As Long
Dim targetPreSpacer As IXMLDOMText

Const ProcName As String = "insertItemBefore"

On Error GoTo Err

If mEndTagSpacer Is Nothing Then
    Set mEndTagSpacer = mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent - 1, vbTab))
    mElement.appendChild mEndTagSpacer
End If

If targetItem Is Nothing Then
    targetIndex = TailEntryIndex
Else
    targetIndex = findItemIndex(targetItem)
    If targetIndex = -1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Item to Insert before does not exist"
    Set targetPreSpacer = getPreSpacer(targetItem.DOMElement)
End If

Set newElement = mConfigStoreProvider.XMLDoc.createElement(Name)
If targetItem Is Nothing Then
    mElement.insertBefore mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent, vbTab)), mEndTagSpacer
    mElement.insertBefore newElement, mEndTagSpacer
ElseIf Not targetPreSpacer Is Nothing Then
    mElement.insertBefore mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent, vbTab)), targetPreSpacer
    mElement.insertBefore newElement, targetPreSpacer
Else
    mElement.insertBefore mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent, vbTab)), targetItem.DOMElement
    mElement.insertBefore newElement, targetItem.DOMElement
End If

Set InsertItemBefore = New XMLConfigItem
InsertItemBefore.Initialise mConfigStoreProvider, newElement, Value, Comment, mIndent + 1
    
mConfigItemsIndexTree.Add addEntry(InsertItemBefore, targetIndex), Name
changed

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function Item( _
                ByVal Name As String) As XMLConfigItem
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mConfigItems(mConfigItemsIndexTree.Item(Name)).ci

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub MoveItemBefore( _
                ByVal itemSpecifier As Variant, _
                Optional ByVal targetItem As XMLConfigItem)
Dim ci As XMLConfigItem
Dim sourceIndex As Long
Dim targetIndex As Long
Dim targetPreSpacer As IXMLDOMText

Const ProcName As String = "moveItemBefore"

On Error GoTo Err

If IsObject(itemSpecifier) Then
    Dim obj As Object
    Set obj = itemSpecifier
    If Not TypeOf obj Is XMLConfigItem Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid object type"
    Set ci = obj
Else
    Set ci = Item(CStr(itemSpecifier))
End If

sourceIndex = findItemIndex(ci)
If sourceIndex = -1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Item to move does not exist"

If targetItem Is Nothing Then
    targetIndex = TailEntryIndex
Else
    targetIndex = findItemIndex(targetItem)
    If targetIndex = -1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Item to move before does not exist"
    Set targetPreSpacer = getPreSpacer(targetItem.DOMElement)
End If

removePreSpacer ci.DOMElement
mElement.removeChild ci.DOMElement
If targetItem Is Nothing Then
    mElement.insertBefore mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent, vbTab)), mEndTagSpacer
    mElement.insertBefore ci.DOMElement, mEndTagSpacer
ElseIf Not targetPreSpacer Is Nothing Then
    mElement.insertBefore mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent, vbTab)), targetPreSpacer
    mElement.insertBefore ci.DOMElement, targetPreSpacer
Else
    mElement.insertBefore mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent, vbTab)), targetItem.DOMElement
    mElement.insertBefore ci.DOMElement, targetItem.DOMElement
End If

ci.RelocateComment

removeEntryFromList sourceIndex
addEntryToList sourceIndex, targetIndex

changed

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

Friend Sub removeItemByName( _
                ByVal Name As String)
Const ProcName As String = "removeItemByName"

On Error GoTo Err

removeItem mConfigItemsIndexTree.Remove(Name)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub removeItemByIdentity( _
                ByVal obj As ConfigItem)
Dim en As Enumerator
Const ProcName As String = "removeItemByIdentity"

On Error GoTo Err

Set en = mConfigItemsIndexTree.Enumerator

Do While en.MoveNext
    If mConfigItems(CLng(en.Current)).ci Is obj Then
        removeItem CLng(en.Current)
        Exit Do
    End If
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addEntry( _
                ByVal Item As XMLConfigItem, _
                ByVal followingItemIndex) As Long
Const ProcName As String = "addEntry"

On Error GoTo Err

If mConfigItemsIndex = UBound(mConfigItems) Then
    ReDim Preserve mConfigItems(growthFactor * (UBound(mConfigItems) + 1) - 1) As ConfigItemTableEntry
End If

mConfigItemsIndex = mConfigItemsIndex + 1
Set mConfigItems(mConfigItemsIndex).ci = Item

addEntryToList mConfigItemsIndex, followingItemIndex

addEntry = mConfigItemsIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function addItemFromElement( _
                ByVal pElement As IXMLDOMElement, _
                ByVal Value As String, _
                ByVal Comment As String) As XMLConfigItem
Const ProcName As String = "addItemFromElement"

On Error GoTo Err

If mEndTagSpacer Is Nothing Then
    Set mEndTagSpacer = mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent - 1, vbTab))
    mElement.appendChild mEndTagSpacer
End If

mElement.insertBefore mConfigStoreProvider.XMLDoc.createTextNode(vbNewLine & String(mIndent, vbTab)), mEndTagSpacer
mElement.insertBefore pElement, mEndTagSpacer

Dim lItem As New XMLConfigItem
lItem.Initialise mConfigStoreProvider, pElement, Value, Comment, mIndent + 1
    
mConfigItemsIndexTree.Add addEntry(lItem, TailEntryIndex), pElement.nodeName
changed

Set addItemFromElement = lItem

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub addEntryToList( _
                ByVal newIndex As Long, _
                ByVal followingIndex As Long)
Const ProcName As String = "addEntryToList"

On Error GoTo Err

mConfigItems(newIndex).Next = followingIndex
mConfigItems(newIndex).Prev = mConfigItems(followingIndex).Prev
mConfigItems(mConfigItems(followingIndex).Prev).Next = newIndex
mConfigItems(followingIndex).Prev = newIndex

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub changed()
Const ProcName As String = "changed"

On Error GoTo Err

mConfigStoreProvider.Dirty
mUpdateSequence = mUpdateSequence + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function findItemIndex( _
                ByVal Item As ConfigItem) As Long
Dim lIndex As Variant
Dim en As Enumerator

Const ProcName As String = "findItemIndex"

On Error GoTo Err

Set en = mConfigItemsIndexTree.Enumerator

Do While en.MoveNext
    lIndex = CLng(en.Current)
    If mConfigItems(CLng(lIndex)).ci Is Item Then
        findItemIndex = CLng(lIndex)
        Exit Function
    End If
Loop
findItemIndex = -1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

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

Private Function getPreSpacer(ByVal element As IXMLDOMElement) As IXMLDOMText
Const ProcName As String = "getPreSpacer"

On Error GoTo Err

If element.previousSibling Is Nothing Then
ElseIf TypeOf element.previousSibling Is IXMLDOMText Then
    Set getPreSpacer = element.previousSibling
    If Not gIswhitespace(getPreSpacer.nodeValue) Then
        Set getPreSpacer = Nothing
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub initTable()
Const ProcName As String = "initTable"

On Error GoTo Err

ReDim mConfigItems(initialCapacity + 1) As ConfigItemTableEntry ' NB: first two entries are dummies
mConfigItems(HeadEntryIndex).Next = TailEntryIndex
mConfigItems(TailEntryIndex).Prev = HeadEntryIndex
mConfigItemsIndex = TailEntryIndex + 1
Set mConfigItemsIndexTree = New SortedDictionary
mConfigItemsIndexTree.Initialise KeyTypeString, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeEntryFromList( _
                ByVal pIndex As Long)
Const ProcName As String = "removeEntryFromList"

On Error GoTo Err

mConfigItems(mConfigItems(pIndex).Prev).Next = mConfigItems(pIndex).Next
mConfigItems(mConfigItems(pIndex).Next).Prev = mConfigItems(pIndex).Prev

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeItem( _
                ByVal pIndex As Long)
Dim ci As XMLConfigItem
Const ProcName As String = "removeItem"

On Error GoTo Err

If pIndex <> -1 Then
    Set ci = mConfigItems(pIndex).ci
    removePreSpacer ci.DOMElement
    mElement.removeChild ci.DOMElement
    ci.Terminate
    Set mConfigItems(pIndex).ci = Nothing
    changed
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removePreSpacer(ByVal element As IXMLDOMElement)
Dim textEl As IXMLDOMText
Const ProcName As String = "removePreSpacer"

On Error GoTo Err

Set textEl = getPreSpacer(element)
If Not textEl Is Nothing Then
    textEl.parentNode.removeChild textEl
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


