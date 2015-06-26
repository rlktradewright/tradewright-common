VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Comparable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This interface defines the order for objects of classes that implement it.
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


Private Const ModuleName                    As String = "Comparable"

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

'@================================================================================
' Methods
'@================================================================================

''
' Compares this object with another object and returns a Value indicating whether
' this object is less than, equal to, or greater than the specified object.
'
' @remarks
'   The implementation of <code>compareTo</code> must ensure the following:
'<pre>
'   <p><code>sgn(obj1.compareTo(obj2)) = -sgn(obj2.compareTo(obj1))</code>;</p>
'
'   <p>if <code>obj1.compareTo(obj2) < 0</code> and <code>obj2.compareTo(obj3) < 0</code>,
'   then <code>obj1.compareTo(obj3) < 0</code>;</p>
'
'   <p>if <code>obj1.compareTo(obj2) = 0</code>, then <code>sgn(obj1.compareTo(obj3))
'   = sgn(obj2.compareTo(obj3))</code> for all obj3.</p>
'</pre>
'   If the specified object is <code>Nothing</code>, then an
'   <code>RrrorCodes.ErrNullPointerException</code> Error is raised.
' @return
'   A <code>Long</code> Value that is negative if this object is less than the
'   specified object, 0 if it is equal to it, and positive otherwise.
' @param obj
'   The object to which this object is to be compared.
'@/
Public Function CompareTo( _
                ByVal obj As Comparable) As Long
                
End Function

'@================================================================================
' Helper Functions
'@================================================================================


