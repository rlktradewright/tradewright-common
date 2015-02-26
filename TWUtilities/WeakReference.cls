VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' An object of this class holds a reference to another object that does not
' Result in that object's reference Count being increased. This means that the
' reference is not taken into account by Visual Basic when deciding whether to
' terminate the object.
'
' Weak references are used to avoid circular references that can prevent unused
' objects being terminateed, leading to memory leaks. For example in a Parent-child
' relationship, if each object has a normal reference to the other, then neither
' will ever be terminateed unless at least one of them takes specific Action to set
' the reference to <code>Nothing</code>, which requires additional programming.
' But if one object maintains a weak reference to the other, the situation does
' not arise.
'
' Note that if an object is terminated, then any attempt to use a weak reference
' to it will cause a program crash.
'
' Objects of this class cannot be created using the <code>New</code> operator. Use
' the <code>CreateWeakReference</code> global method.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

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

Private mObjectAddress As Long

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

Friend Sub Initialise(ByVal targetObject As Object)

If targetObject Is Nothing Then
    mObjectAddress = 0
Else
    mObjectAddress = ObjPtr(targetObject)
End If

End Sub

''
' Returns a full reference to the object to which this weak reference relates.
'
' @remarks
' @return
'   A full reference to the object to which this weak reference relates.
'@/
Public Property Get Target() As Object

Dim lObject As Object

If mObjectAddress = 0 Then
    Set Target = Nothing
Else
    CopyMemory VarPtr(lObject), VarPtr(mObjectAddress), 4
    Set Target = lObject
    ZeroMemory VarPtr(lObject), 4
End If

End Property

'@================================================================================
' Helper Functions
'@================================================================================

