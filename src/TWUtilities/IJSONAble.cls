VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IJSONable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Classes implement this interface as a standard way of providing a JOSN string
' representation of their instances.
'
' JSON is JavaScript Object Notation - see http://www.json.org and
' http://www.ietf.org/rfc/rfc4627.txt?number=4627. JSON is a lightweight, text-based,
' language-independent data interchange format.
'
' @remarks
'   Classes that implement <code>JSONAble</code> should normally also implement
'   <code>Stringable</code>. However there is no requirement that <code>JSONable.ToJSON</code>
'   and <code>Stringable.ToString</code> should return the same value.
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

Private Const ModuleName                            As String = "IJSONable"

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
' Generates a JSON string representation of the object, suitable for logging,
' diagnostic, or data exchange purposes.
'
' @remarks
'   Note that the <code>ParseJSON</code> global method returns an instance of the
'   class that implements <code>JSONable</code> if the following conditions are
'   true:
'   <ul>
'       <li>The first name/value pair in the JSON string has name '$ProgId'</li>
'       <li>The class has its <code>Instancing</code> property set to
'       <code>Public</code>;</li>
'       <li>For each name/value pair after the first in the JSON encoding, the class
'       has a writable property of the same name that accepts the encoded value.</li>
'   </ul>
'
' @return
'   A JSON string representation of the object.
'@/
Public Function ToJSON() As String

End Function

'@================================================================================
' Helper Functions
'@================================================================================


