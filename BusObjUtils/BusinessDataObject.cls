VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusinessDataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' A <code>BusinessDataObject</code> is an abstraction of a data entity
' that encapsulates knowledge about some specific domain, and is held in permanent
' storage (for example in a database or in text files).
'
' <p>This interface specifies common mechanisms for modifying and deleting such objects.</p>
'
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

Private Const ModuleName As String = "BusinessDataObject"

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
' Returns an <code>ErrorList</code> object that can be used to receive
' notifications about validation errors in this <code>BusinessDataObject</code>.
'
' @return a <code>BusinessObjEventSource</code> object.
'
'@/
Public Property Get ErrorList() As ErrorList
End Property

''
' Returns a <code>BusinessObjEventSource</code> object that can be used to receive
' events from this <code>BusinessDataObject</code>.
'
' @return a <code>BusinessObjEventSource</code> object.
'
'@/
Public Property Get EventSource() As BusinessObjEventSource
End Property

''
' Returns the unique identifier for this <code>BusinessDataObject</code> object.
'
' The unique Id is permanently associated with the <code>BusinessDataObject</code> object
' and cannot be changed.
'
' @return the unique identifier for this <code>BusinessDataObject</code> object.
'
'@/
Public Property Get Id() As Long
End Property

''
' Indicates whether this <code>BusinessDataObject</code> object is to be, or has been,
' deleted from permanent storage (ie the <code>Delete</code> method has been called).
'
' @return returns <code>True</code> if the <code>Delete</code> method has been called,
'           and <code>False</code> otherwise.
'
'@/
Public Property Get IsDeleted() As Boolean
End Property

''
' Indicates whether this <code>BusinessDataObject</code> object has uncommitted modifications
' (ie changes have been made but the the <code>ApplyEdit</code> method has not yet been called).
'
' @return returns <code>True</code> if the <code>BusinessDataObject</code> object has
'                    uncommitted modifications.
'
'@/
Public Property Get IsDirty() As Boolean
End Property

''
' Indicates whether this <code>BusinessDataObject</code> object has not yet been
' written to permanent storage.
'
' @return returns <code>True</code> if the <code>BusinessDataObject</code> object has
'           not yet been written to permanent storage.'
'@/
Public Property Get IsNew() As Boolean
End Property

''
' Indicates whether this <code>BusinessDataObject</code> object complies with all
' relevant business rules.
'
' @return returns <code>True</code> if the <code>BusinessDataObject</code> object
'           complies with all relevant business rules.
'@/
Public Property Get IsValid() As Boolean
End Property

''
' Returns the unique Name for this <code>BusinessDataObject</code> object.
'
' The unique Name can be changed.
'
' @return the unique Name for this <code>BusinessDataObject</code> object.
'
'@/
Public Property Get Name() As String
End Property

''
' Sets the unique Name for this <code>BusinessDataObject</code> object.
'
' @param value the new unique Name for this <code>BusinessDataObject</code> object.
'
'@/
Public Property Let Name(ByVal value As String)
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Commits any changes to this <code>BusinessDataObject</code> object to permanent
' storage.
'
'
'@/
Public Sub ApplyEdit()
End Sub

''
' Discards any changes to this <code>BusinessDataObject</code> object that
' have not yet been committed to permanent storage.
'
'
'@/
Public Sub CancelEdit()
End Sub

''
' Deletes this <code>BusinessDataObject</code> object from permanent storage.
'
'
'@/
Public Sub Delete()
End Sub

''
' Ensures that this data object has the latest information from permanent storage.
'
' @remarks
'   If the data object has not yet been committed to permanent storage,
'   or if the data object has been deleted from permanent storage, an
'   <code>ErrorCodes.ErrIllegalStateException</code> error is raised.
'@/
Public Sub Refresh()

End Sub

'@================================================================================
' Helper Functions
'@================================================================================




