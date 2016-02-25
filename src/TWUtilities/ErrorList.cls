VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class provides a mechanism for maintaining a list of <code>ErrorItem</code>s,
' together with provision for being notified via events when an <code>ErrorItem</code>
' is added to or removed from the list.
' <p>It is useful in situations where validation of an object&rsquo;s State may Result in
' several simultaneous invalid conditions.</p>
'
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

''
' Raised when an <code>ErrorItem</code> is added to or removed from this
' <code>ErrorList</code> object, and when the list is cleared.
'
' @remarks
'   The following members of the <code>CollectionChangeTypes</code> enum may
'   occur in the <code>ChangeType</code> field of the
'   <code>CollectionChangedEventData</code> parameter:
' <pre>
'    <tt>CollItemAdded</tt><br />
'    <tt>CollItemRemoved</tt><br />
'    <tt>CollCollectionCleared</tt>
' </pre>
' @param ev
'   contains details of the change.
'
'@/
Event CollectionChanged( _
                ev As CollectionChangeEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName            As String = "ErrorList"

'@================================================================================
' Member variables
'@================================================================================

Private mErrorList                  As Collection

Private mChangeListeners            As Listeners

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mErrorList = New Collection
Set mChangeListeners = New Listeners
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

''
' Returns the number of <code>Erroritem</code>s currently in the list.
'
' @return The number of <code>Erroritem</code>s currently in the list
'
'@/
Public Property Get Count() As Long
Const ProcName As String = "Count"

On Error GoTo Err

Count = mErrorList.Count

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Registers an object to act as a listener for changes to the list.
'
' @remarks
'   The listener object must implement the <code>CollectionChangeListener</code>
'   interface.
' @param listener
'   The object which is to act as a listener.
'@/
Public Sub AddCollectionChangeListener( _
                ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "AddCollectionChangeListener"
On Error GoTo Err

mChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Declares whether a validation rule is violated or satisfied.
'
' <p>The validation rule is identified by a unique <code>String</code> Value.
' The calling application is responsible for assigning identifier values to rules.</p>
'
' <p>If the rule is violated, an <code>ErrorItem</code> is added to the list and the
' <code>ItemAdded</code> event is fired.</p>
' <p>If the rule is satisfied, any existing <code>ErrorItem</code> for that rule is
' removed from the list and the <code>ItemRemoved</code> event is fired. If there are
' now no items remaining in the list, the <code>NoItems</code> event is also fired.</p>
'
' @return
'   Returns <code>True</code> if the validation rule has been specified and <code>False</code>
'   if it has been violated.
' @param isValid
'   If <code>True</code>, indicates that the validation rule is satisfied.
'   If <code>False</code>, indicates that the validation rule is violated.
' @param ruleId
'   A <code>String</code> Value that identifies the relevant validation rule.
' @param params
'   An array of <code>Variant</code>s that may be useful to consumers of the
'   <code>ErrorItem</code>s. For example they could be values for substitution
'   in an Error message.
'@/
Public Function AssertRule( _
                ByVal isValid As Boolean, _
                ByVal RuleId As String, _
                ParamArray params()) As Boolean
Dim Item As ErrorItem
Dim lParams() As Variant
Dim i As Long

Const ProcName As String = "AssertRule"

On Error GoTo Err

If isValid Then
    AssertRule = True
    
    On Error Resume Next
    Set Item = mErrorList.Item(RuleId)
    On Error GoTo Err
    
    If Not Item Is Nothing Then
        mErrorList.Remove RuleId
        fireChange CollItemRemoved, Item
    End If
Else
    AssertRule = False
    
    On Error Resume Next
    Set Item = mErrorList.Item(RuleId)
    On Error GoTo Err
    
    If Item Is Nothing Then
        Set Item = New ErrorItem
        If UBound(params) >= 0 Then
            ReDim lParams(UBound(params)) As Variant
            For i = 0 To UBound(params)
                lParams(i) = params(i)
            Next
        End If
        Item.Initialise RuleId, lParams
        mErrorList.Add Item, RuleId
        fireChange CollItemAdded, Item
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

''
' Clears all <code>ErrorItem</code>s from the list
'@/
Public Sub Clear()
Const ProcName As String = "Clear"

On Error GoTo Err

Set mErrorList = New Collection
fireChange CollCollectionCleared, Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Returns a particular <code>ErrorItem</code> from the list.
'
' @return
'   The <code>ErrorItem</code> identified by the <code>index</code> argument.
' @param pIndex
'   Specifies which <code>ErrorItem</code> is to be returned.
'
'   This can be either the validation rule identifier, or a <code>Long</code> Value
'   specifying the ordinal number, of the required <code>ErrorItem</code>.
'@/
Public Function Item( _
                ByVal pIndex As Variant) As ErrorItem
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mErrorList.Item(pIndex)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = mErrorList.[_NewEnum]
End Function

''
' Unregisters a listener.
'
' @param listener
'   The object which is to be unregistered.
'@/
Public Sub RemoveCollectionChangeListener(ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub fireChange( _
                ByVal ChangeType As CollectionChangeTypes, _
                ByVal affectedObject As ErrorItem)
Const ProcName As String = "fireChange"
On Error GoTo Err

Dim ev As CollectionChangeEventData
Set ev.Source = Me
ev.ChangeType = ChangeType
If Not affectedObject Is Nothing Then Set ev.AffectedItem = affectedObject

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not mChangeListeners.Valid Then
    sInit = True
    sSomeListeners = mChangeListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As ICollectionChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If
RaiseEvent CollectionChanged(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub






