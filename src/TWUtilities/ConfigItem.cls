VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This interface is implemented by classes that represent configuration items.
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


Private Const ModuleName As String = "ConfigItem"

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
' Returns a <code>ConfigItems</code> object representing the configuration items
' below this one.
'
' @return
'   A <code>ConfigItems</code> object representing the configuration items
'   below this one in the hierarchy.
'
'@/
Public Property Get ChildItems() As ConfigItems
End Property

''
' Sets this <code>ConfigItem</code> object's Comment.
'
' @param Value
'   The Comment to be set for this configuration Item.
'
'@/
Public Property Let Comment(ByVal Value As String)
End Property

''
' Returns this <code>ConfigItem</code> object's Comment.
'
' @return
'   This configuration Item's Comment.
'
'@/
Public Property Get Comment() As String
End Property

''
' Returns this <code>ConfigItem</code> object's Name.
'
' @return
'   This configuration Item's Name.
'
'@/
Public Property Get Name() As String
End Property

''
' Sets this <code>ConfigItem</code> object's Value.
'
' @param Value
'   The Value to be set for this configuration Item.
'
'@/
Public Property Let Value(ByVal Value As String)
End Property

''
' Returns this <code>ConfigItem</code> object's Value.
'
' @return
'   This configuration Item's Value.
'
'@/
Public Property Get Value() As String
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Returns the specified attribute's Value.
'
' @return
'   The Value of the specified attribute.
' @param Name
'   The Name of the attribute whose Value is to be returned.
'
'@/
Public Function GetAttribute( _
                ByVal Name As String) As String
End Function

''
' Returns the specified attribute's Value, or the specified default Value if
' the attribute is not present.
'
' @return
'   The Value of the specified attribute, unless it is not present, in which case
'   the specified default is returned.
' @param Name
'   The Name of the attribute whose Value is to be returned.
' @param default
'   The Value to be returned if the specified attribute is not present.
'
'@/
Public Function GetDefaultableAttribute( _
                ByVal Name As String, _
                ByVal default As String) As String
End Function

''
' Removes the specified attribute.
'
' @param Name
'   The Name of the attribute to be removed.
'
'@/
Public Sub RemoveAttribute( _
                ByVal Name As String)
End Sub

''
' Sets the specified attribute's Value.
'
' @param Name
'   The Name of the attribute whose Value is to be set.
' @param Value
'   The new Value for the specified attribute.
'@/
Public Function SetAttribute( _
                ByVal Name As String, _
                ByVal Value As String) As String
End Function

'@================================================================================
' Helper Functions
'@================================================================================


