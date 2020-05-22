VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DictionaryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This type is used internally by the SortedDictionary class to store the key and
' data of the entries in the collection.
'
' A <code>SortedDictionary</code>'s entries may be obtained by using the enumerator
' returned by the <code>SortedDictionary</code>'s <code>EntryEnumerator</code>
' property.
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


Private Const ModuleName                    As String = "DictionaryEntry"

'@================================================================================
' Member variables
'@================================================================================

Private mLeft                               As DictionaryEntry
Private mRight                              As DictionaryEntry
Private mParent                             As WeakReference
Private mColor                              As NodeColors
Private mKey                                As Variant
Private mData                               As Variant

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
mColor = Black
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Let Left(ByVal Value As DictionaryEntry)
Set mLeft = Value
End Property

Friend Property Get Left() As DictionaryEntry
Set Left = mLeft
End Property

Friend Property Let Right(ByVal Value As DictionaryEntry)
Set mRight = Value
End Property

Friend Property Get Right() As DictionaryEntry
Set Right = mRight
End Property

Friend Property Let Parent(ByVal Value As DictionaryEntry)
Const ProcName As String = "Parent"

On Error GoTo Err

Set mParent = gCreateWeakReference(Value)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Parent() As DictionaryEntry
Const ProcName As String = "Parent"

On Error GoTo Err

Set Parent = mParent.Target

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Color(ByVal Value As NodeColors)
mColor = Value
End Property

Friend Property Get Color() As NodeColors
Color = mColor
End Property

Friend Property Let Key(ByVal Value As Variant)
gSetVariant mKey, Value
End Property

''
' Returns the key of this entry.
'
' @return
'   The entry's key.
'@/
Public Property Get Key() As Variant
gSetVariant Key, mKey
End Property

''
' Sets the data of this entry.
'
' @param Value
'   The entry's data.
'@/
Public Property Let Data(ByVal Value As Variant)
gSetVariant mData, Value
End Property

''
' Returns the data of this entry.
'
' @return
'   The entry's data.
'@/
Public Property Get Data() As Variant
gSetVariant Data, mData
End Property

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


