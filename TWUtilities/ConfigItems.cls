VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This interface is implemented by classes that represent a list of configuration
' items.
'
' Note that a <code>ConfigItems</code> object can hold several configuration
' items with the same Name.
'
' Individual <code>ConfigItem</code> objects can be retrieved from a
' <code>ConfigItems</code> iteratively, in the order that they were added, using
' either <code>For Each</code> or the <code>Enumerator</code> object
' returned by the <code>Enumerator</code> property.
'
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


Private Const ModuleName                    As String = "ConfigItems"

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

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the members of the collection.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get Enumerator() As Enumerator
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Adds a new configuration Item to the end of the list.
'
' @return
'   A <code>ConfigItem</code> object representing the added configuration Item.
' @param Name
'   The new configuration Item's Name.
' @param Value
'   The new configuration Item's Value.
' @param Comment
'   The new configuration Item's Comment.
'
'@/
Public Function AddItem( _
                ByVal Name As String, _
                Optional ByVal value As String, _
                Optional ByVal Comment As String) As ConfigItem
End Function

''
' Adds a cloned configuration Item to the end of the list.
'
' @return
'   A <code>ConfigItem</code> object representing the clone that has been added.
' @param pItem
'   The configuration item to be cloned and added to the list.
' @param pDeep
'   If <code>True</code>, the clone that is added includes a recursive clone of all children
'   of the original configuration item.
'
'@/
Public Function AddClone( _
                ByVal pItem As ConfigItem, _
                ByVal pDeep As Boolean) As ConfigItem
End Function

''
' Removes all configuration items from the list.
'@/
Public Sub Clear()
End Sub

''
' Returns the number of configuration items in the list.
'
' @return
'   The number of configuration items in the list
'
'@/
Public Function Count() As Long
End Function

''
' Inserts a new configuration Item immediately before another specified Item in the list.
'
' @return
'   A <code>ConfigItem</code> object representing the added configuration Item.
' @param Name
'   The new configuration Item's Name.
' @param Value
'   The new configuration Item's Value.
' @param Comment
'   The new configuration Item's Comment.
' @param targetItem
'   The Item before which the new configuration Item is to be inserted. If this argument
'   is not supplied, or is <code>Nothing</code>, then the new Item is inserted at the end
'   of the list.
'
'
'@/
Public Function InsertItemBefore( _
                ByVal Name As String, _
                Optional ByVal value As String, _
                Optional ByVal Comment As String, _
                Optional ByVal targetItem As ConfigItem) As ConfigItem
End Function

''
' Returns the first configuration Item in the list with the specified Name.
'
' @remarks
'   Where there is more than one configuration Item in the list with the same
'   Name, only the first can be retrieved using this method.  To retrieve the
'   others, it is necessary to iterate through the list using either For Each
'   or an enumerator returned by the <code>Enumerator</code> property.
' @return
'   A <code>ConfigItem</code> object representing the first configuration Item
'   in the list that has the specified Name.
' @param Name
'   The Name of the required configuration Item.
'
'
'@/
Public Function Item( _
                ByVal Name As String) As ConfigItem
Attribute Item.VB_UserMemId = 0
End Function

''
' Moves an existing configuration item immediately before another specified Item in the list.
'
' @param itemSpecifier
'   Specifies which configuration item to be moved.
'
'   If the supplied Value is a <code>String</code>, then the first configuration
'   item in the list with that name is moved.
'
'   If the supplied Value is a configuration item, then that item is moved.
' @param targetItem
'   The item before which the configuration item is to be moved. If this argument
'   is not supplied, or is <code>Nothing</code>, then the new item is moved to the end
'   of the list.
'@/
Public Sub MoveItemBefore( _
                ByVal itemSpecifier As Variant, _
                Optional ByVal targetItem As ConfigItem)
End Sub

''
' Removes the specified configuration Item from the list.
'
' @param itemSpecifier
'   Specifies which configuration Item to be removed.
'
'   If the supplied Value is a <code>String</code>, then the first configuration
'   Item in the list with that Name is removed.
'
'   If the supplied Value is a configuration Item, then that Item is removed.
'
'@/
Public Sub Remove( _
                ByVal itemSpecifier As Variant)
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
End Function

'@================================================================================
' Helper Functions
'@================================================================================
