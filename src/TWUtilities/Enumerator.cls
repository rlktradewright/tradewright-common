VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Enumerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class provides the mechanism that enables collection classes that are not
' based on the Visual Basic <code>Collection</code> object to be used with the
' <code>For Each</code> construct.
'
' Such a class should have a NewEnum method that returns an instance of this class, and
' must also implement the <code>Enumerable</code> interface.
'
' This class is most frequently called directly from the Visual Basic runtime during a
' <code>For Each</code> loop. However, it also provides facilities to enable manual iteration
' through the items of the collection object. This is useful where the iteration cannot be
' fitted into a <code>For Each</code> construct. You can obtain the <code>Enumerator</code>
' object from the <code>Enumerable</code> interface's <code>Enumerator</code> property.
'
' Instances of this class must be created using the global <code>CreateEnumerator</code>
' method.
'
' <h3>Example</h3>
' As an example, such a collection class should include code like the following:
' <pre>
' <code>
'    Implements IEnumerable
'
'    Private mUpdateSeq As Long
'
'    Public Property Get IEnumerable_Enumerator() As Enumerator
'        Set Enumerator = CreateEnumerator(Me)
'    End Property
'
'    Public Property Get IEnumerable_UpdateSequence() As Long
'        IEnumerable_UpdateSequence = mUpdateSeq
'    End Property
'
'    Private Function IEnumerable_NextItem(pItem As Variant, pCookie As Variant) As Boolean
'        .
'        . <em>Code to implement this method</em>
'        .
'    End Function
'
'    Private Sub IEnumerable_Remove(pCookie As Variant)
'        .
'        . <em>Code to implement this method</em>
'        .
'    End Sub
'
'    Private Function IEnumerable_Skip(ByVal pNumToSkip As Long, pCookie As Variant) As Boolean
'        .
'        . <em>Code to implement this method</em>
'        .
'    End Function
'
'    ' This function must have its Procedure ID set to -4 using
'    ' Tools &gt; Procedure Attributes.
'    Public Function NewEnum() As IUnknown
'        Set NewEnum = CreateEnumerator(Me)
'    End Function
' </code>
' </pre>
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IEnumVARIANT

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


Private Const ModuleName                    As String = "Enumerator"

'@================================================================================
' Member variables
'@================================================================================

Private mTarget                             As IEnumerable

Private mInitialCookie                      As Variant
Private mCookie                             As Variant

Private mCurrent                            As Variant

Private mUpdateSeq                          As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
If Not gRedirected Then
    Dim obj As IEnumVARIANT
    Set obj = Me    ' use the right interface
    gRedirectMethod ObjPtr(obj), 3, AddressOf GEnumerator.GetNext
    gRedirectMethod ObjPtr(obj), 4, AddressOf GEnumerator.Skip
    gRedirected = True
End If
mCookie = Empty
mUpdateSeq = MinLongValue
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

Private Sub IEnumVARIANT_Clone( _
                ByRef clonedEnum As TWWin32API.IEnumVARIANT)
Dim lEnumerator As New Enumerator
Const ProcName As String = "IEnumVARIANT_Clone"

On Error GoTo Err

lEnumerator.Initialise mTarget, mCookie
Set clonedEnum = lEnumerator

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IEnumVARIANT_Next( _
                ByVal numElementsRequested As Long, _
                ByRef items As Variant, _
                ByVal lpNumElementsFetched As Long)
' This method is redirected at initialisation
End Sub

Private Sub IEnumVARIANT_Reset()
Const ProcName As String = "IEnumVARIANT_Reset"

On Error GoTo Err

Reset

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IEnumVARIANT_Skip( _
                ByVal numElementsToSkip As Long)
' This method is redirected at initialisation
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

''
' Returns the current item in the collection object.
'
' @remarks
'   This property is only valid after a call to the <code>MoveNext</code> method
'   that returned <code>True</code>. Calling it at any other time causes
'   error <code>ErrorCodes.ErrIllegalStateException</code> to be raised.
' @return
'   A <code>Variant</code> containing the current item in the collection.
'@/
Public Property Get Current() As Variant
Const ProcName As String = "Current"

On Error GoTo Err

If IsEmpty(mCurrent) Then Err.Raise ErrorCodes.ErrIllegalStateException, , "No current Item"

gSetVariant Current, mCurrent

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function GetNext( _
                ByRef Item As Variant) As Boolean
Const ProcName As String = "GetNext"
On Error GoTo Err

If mTarget.UpdateSequence <> mUpdateSeq Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Collection object has been modified during enumeration of " & TypeName(mTarget)

If Not mTarget.NextItem(Item, mCookie) Then
    mCurrent = Empty
    GetNext = False
Else
    gSetVariant mCurrent, Item
    GetNext = True
End If

Exit Function

Err:
' Note that we call gNotifyUnhandledError rather than gHandleUnexpectedError to
' ensure that the error can't be swallowed by the application, and also because
' if it occurs during a For Each..Next loop, there is a very untidy failure that is
' difficult to diagnose
gNotifyUnhandledError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal Target As IEnumerable, _
                Optional ByVal pCookie As Variant)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mTarget = Target
If IsMissing(pCookie) Then
    mInitialCookie = Empty
    mCookie = Empty
Else
    gSetVariant mInitialCookie, pCookie
    gSetVariant mCookie, pCookie
End If
mUpdateSeq = Target.UpdateSequence

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Positions the iteration on the next item in the collection object. If there are no
' further objects in the collection, the returned value is <code>False</code>.
'
' @return
'   <code>True</code> if the next item exists, <code>False</code> otherwise.
'@/
Public Function MoveNext() As Boolean
Const ProcName As String = "MoveNext"
On Error GoTo Err

Dim Var As Variant
MoveNext = GetNext(Var)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
On Error GoTo Err

Set NewEnum = Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Removes the current item (ie the item returned in the previous call to
' <code>nextItem</code>) from the collection.
'
' @param pCookie
'   The current pCookie value. Before exiting from the call, the implementing object
'   should update this value to enable it to position correctly on the next call.
'@/
Public Sub Remove()
Const ProcName As String = "Remove"
On Error GoTo Err

If IsEmpty(mCurrent) Then Err.Raise ErrorCodes.ErrIllegalStateException, , "No current Item"
If mTarget.UpdateSequence <> mUpdateSeq Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Collection object has been modified during enumeration of " & TypeName(mTarget)

mTarget.Remove mCookie
mUpdateSeq = mTarget.UpdateSequence

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Resets the position in the iteration to just before the first item.
'
' @remarks
'   Note that there is no current item after calling this method. You must call
'   the <code>moveNext</code> method to make the first item in the collection current.
'@/
Public Sub Reset()
Const ProcName As String = "Reset"
On Error GoTo Err

If mTarget.UpdateSequence <> mUpdateSeq Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Collection object has been modified during enumeration of " & TypeName(mTarget)

If IsEmpty(mInitialCookie) Then
    mCookie = Empty
Else
    gSetVariant mCookie, mInitialCookie
End If

mCurrent = Empty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Skip( _
                ByVal numElementsToSkip As Long) As Boolean
Const ProcName As String = "Skip"
On Error GoTo Err

If mTarget.UpdateSequence <> mUpdateSeq Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Collection object has been modified during enumeration of " & TypeName(mTarget)

Skip = mTarget.Skip(numElementsToSkip, mCookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




