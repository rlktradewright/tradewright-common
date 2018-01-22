VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IEnumerable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Collection classes that use <code>Enumerator</code> objects to allow their contents
' to be iterated using <code>For Each</code> must implement this interface.
'
' The <code>For Each</code> syntax provides a convenient means of iterating over
' a set of items held within a collection object. In order to be usable with
' <code>For Each</code>, a collection class must provide a public method called
' NewEnum with the following signature:
'
' <pre>
' <code>Public Function NewEnum() As IUnknown</code>
' </pre>
'
' This method must have a procedure id of -4, and must return an object that
' implements the COM interface <code>IEnumVariant</code>. This interface provides
' methods that allow the Visual Basic runtime to retrieve the items in the collection.
'
' However, the specification of IEnumVariant does not lend itself to direct implementation
' in Visual Basic code. For collection classes that are based on a Visual Basic Collection
' object, this is not a problem since the NewEnum method can simply return the
' Collection object's hidden _NewEnum property, which yields an object with the
' required interface.
'
' A Collection class that cannot be based on a Collection object (for example
' because it can contain multiple items with the same key) can easily meet the
' requirements of NewEnum by creating an <code>Enumerator</code> object and returning this
' from the NewEnum method. The <code>Enumerator</code> object implements the
' <code>IEnumVariant</code> interface. In addition, the collection class must implement the
' <code>Enumerable</code> interface to enable the <code>Enumerator</code> object
' to retrieve the required items. The <code>Enumerator</code> object does this by
' calling the methods in the <code>Enumerable</code> interface.
'
' <h3>Cookies</h3>
' While a <code>For Each</code> loop is in Progress, the class that implements
' <code>Enumerable</code> needs to keep track of how far through the list it has
' enumerated. Information about this is stored in a variant called a <em>cookie</em>,
' which is passed back to the caller by the implementing object from the methods
' in this interface, and is presented unchanged to the implementing object on the
' next such call. Note that a pCookie must not contain an object.
'
' <h3>Acknowledgement</h3>
' I am indebted to Bryan Stafford for his article on implementing
' collection classes at <a href="http://www.mvps.org/vbvision/Super_Collections.htm">
' http://www.mvps.org/vbvision/Super_Collections.htm</a>. This article provided insights
' that facilitated implementation of the <code>Enumerator</code> class, though I took
' rather a different, more reusable, approach in the detail of the implementation.
' @remarks
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


Private Const ModuleName                    As String = "IEnumerable"

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
' An <code>Enumerator</code> object that can be used to manually iterate through the
' items in the collection object that implements <code>Enumerable</code>.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get Enumerator() As Enumerator

End Property

''
' Returns the current value of a variable that must be incremented each time
' the object that implements <code>Enumerable</code> is modified. This is
' used by the <code>Enumerator</code> object to invalidate itself if the
' object that implements <code>Enumerable</code> is modified after the
' creation of the <code>Enumerator</code> object.
'
' @return
'   A value that is changed for each update to the object that implements <code>Enumerable</code>.
'@/
Public Property Get UpdateSequence() As Long

End Property

'@================================================================================
' Methods
'@================================================================================

''
' Supplies the next item in the collection to the caller.
'
' @return
'   Returns <code>True</code> if the next item has been successfully supplied, and
'   <code>False</code> otherwise (ie there are no more items).
' @param item
'   A reference to a variant into which the next item is placed.
' @param pCookie
'   The current pCookie value. Before exiting from the call, the implementing object
'   should update this value to enable it to position correctly on the next call.
'
'   If the supplied value is <code>Empty</code>, the first item should be located.
'@/
Public Function NextItem( _
                ByRef Item As Variant, _
                ByRef pCookie As Variant) As Boolean
                
End Function

''
' Removes the current item (ie the item returned in the previous call to
' <code>nextItem</code>) from the collection.
'
' @param pCookie
'   The current pCookie value. Before exiting from the call, the implementing object
'   may need to update this value to enable it to position correctly on the next call.
'@/
Public Sub Remove( _
                ByRef pCookie As Variant)

End Sub

''
' Requests that the number of items specified be skipped over.
'
' @return
'   Returns <code>True</code> if the specified number of items have been skipped,
'   and <code>False</code> otherwise (ie the end of the list has been reached before
'   that number of items has been skipped).
' @param numToSkip
'   The number of items that are to be skipped.
' @param pCookie
'   The current pCookie value. Before exiting from the call, the implementing object
'   should update this value to enable it to position correctly on the next call.
'
'   If the supplied value is <code>Empty</code>, skipping must start from the
'   beginning of the list.
'@/
Public Function Skip( _
                ByVal numToSkip As Long, _
                ByRef pCookie As Variant) As Boolean
                
End Function

'@================================================================================
' Helper Functions
'@================================================================================


