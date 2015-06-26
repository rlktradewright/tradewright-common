VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLConfigItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class is an implementation of the <code>ConfigItem</code> interface that
' stores the configuration Data in an XML file.
'
' Objects of this class cannot be created using the <code>New</code> operator.
'
' Objects of this class must be used via their <code>ConfigItem</code> interface,
' since the default interface exposes no public properties or methods. For example:
'
' <pre>
' <tt>
' Dim cf As ConfigStoreProvider<br />
' Dim root As IConfigItem<br />
' Dim outputFilename As String<br />
' <br />
' Set cf = LoadConfigProviderFromXMLFile("C:\MyConfig.xml")<br />
' Set root = cf.rootItem    ' this is actually an XMLConfigItem object<br />
' outputFilename = root.childItems.Item("outputFilename").value<br />
' </tt>
' </pre>
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IConfigItem

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


Private Const ModuleName As String = "XMLConfigItem"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigFile                         As xmlConfigProvider

Private mElement                            As IXMLDOMElement
Private mValueElement                       As IXMLDOMText
Private mCommentElement                     As IXMLDOMComment

Private mChildItems                         As XMLConfigItems

Private mIndent                             As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ConfigItem Interface Members
'@================================================================================

Private Property Get IConfigItem_childItems() As IConfigItems
Set IConfigItem_childItems = ChildItems
End Property

Private Property Let IConfigItem_Comment(ByVal RHS As String)
Comment = RHS
End Property

Private Property Get IConfigItem_Comment() As String
IConfigItem_Comment = Comment
End Property

Private Function IConfigItem_getAttribute(ByVal Name As String) As String
Const ProcName As String = "IConfigItem_getAttribute"

On Error GoTo Err

IConfigItem_getAttribute = GetAttribute(Name)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function IConfigItem_getDefaultableAttribute(ByVal Name As String, ByVal default As String) As String
On Error Resume Next
IConfigItem_getDefaultableAttribute = default
IConfigItem_getDefaultableAttribute = GetAttribute(Name)
If IConfigItem_getDefaultableAttribute = "" Then IConfigItem_getDefaultableAttribute = default
End Function

Private Property Get IConfigItem_Name() As String
IConfigItem_Name = Name
End Property

Private Sub IConfigItem_removeAttribute(ByVal Name As String)
Const ProcName As String = "IConfigItem_removeAttribute"

On Error GoTo Err

RemoveAttribute Name

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IConfigItem_setAttribute(ByVal Name As String, ByVal Value As String) As String
Const ProcName As String = "IConfigItem_setAttribute"

On Error GoTo Err

SetAttribute Name, Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Let IConfigItem_Value(ByVal RHS As String)
Value = RHS
End Property

Private Property Get IConfigItem_Value() As String
IConfigItem_Value = Value
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get ChildItems() As XMLConfigItems
Set ChildItems = mChildItems
End Property

Friend Property Let Comment( _
                ByVal Value As String)
Const ProcName As String = "Comment"

On Error GoTo Err

If mCommentElement Is Nothing Then
    Set mCommentElement = mConfigFile.XMLDoc.createComment(Value)
    mElement.parentNode.insertBefore mConfigFile.XMLDoc.createTextNode(vbNewLine & String(mIndent - 1, vbTab)), mElement
    mElement.parentNode.insertBefore mCommentElement, mElement
Else
    mCommentElement.nodeValue = Value
End If
mConfigFile.Dirty

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Comment() As String
Const ProcName As String = "Comment"

On Error GoTo Err

If Not mCommentElement Is Nothing Then
    Comment = mCommentElement.nodeValue
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get DOMElement() As IXMLDOMElement
Set DOMElement = mElement
End Property

Friend Property Get Name() As String
Const ProcName As String = "Name"

On Error GoTo Err

Name = mElement.nodeName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let Value(ByVal Value As String)
Const ProcName As String = "Value"

On Error GoTo Err

If mValueElement Is Nothing Then
    If Not gIswhitespace(Value) Then
        Set mValueElement = mConfigFile.XMLDoc.createTextNode(gXMLEncode(Value))
        mElement.insertBefore mValueElement, mElement.firstChild
        mConfigFile.Dirty
    End If
Else
    If gIswhitespace(Value) Then
        mElement.removeChild mValueElement
        Set mValueElement = Nothing
        mConfigFile.Dirty
    Else
        Value = gXMLEncode(Value)
        If mValueElement.nodeValue <> Value Then
            mValueElement.nodeValue = Value
            mConfigFile.Dirty
        End If
    End If
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Value() As String
Const ProcName As String = "Value"

On Error GoTo Err

If Not mValueElement Is Nothing Then
    Value = gXMLDecode(mValueElement.nodeValue)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function GetAttribute( _
                ByVal Name As String) As String
Const ProcName As String = "getAttribute"

On Error GoTo Err

GetAttribute = mElement.GetAttribute(Name)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal confFile As xmlConfigProvider, _
                ByVal element As IXMLDOMElement, _
                ByVal Value As String, _
                ByVal Comment As String, _
                ByVal indent As Long)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mConfigFile = confFile
Set mElement = element
mIndent = indent
If Value <> "" Then
    Set mValueElement = mConfigFile.XMLDoc.createTextNode(gXMLEncode(Value))
    mElement.appendChild mValueElement
End If

If Comment <> "" Then
    Set mCommentElement = mConfigFile.XMLDoc.createComment(Value)
    mElement.parentNode.insertBefore mConfigFile.XMLDoc.createTextNode(vbNewLine & String(mIndent - 1, vbTab)), mElement
    mElement.parentNode.insertBefore mCommentElement, mElement
End If

Set mChildItems = New XMLConfigItems
mChildItems.Initialise mConfigFile, mElement, mIndent

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub InitialiseFromNode( _
                ByVal confFile As xmlConfigProvider, _
                ByVal element As IXMLDOMElement, _
                ByVal Comment As IXMLDOMComment, _
                ByVal indent As Long)
Const ProcName As String = "InitialiseFromNode"

On Error GoTo Err

Dim node As IXMLDOMNode
Dim childComment As IXMLDOMComment
Dim Item As XMLConfigItem

Set mConfigFile = confFile
Set mElement = element
mIndent = indent
Set mCommentElement = Comment
    

Set mChildItems = New XMLConfigItems
mChildItems.Initialise mConfigFile, mElement, mIndent

For Each node In mElement.childNodes
    If TypeOf node Is IXMLDOMElement Then
        Set Item = mChildItems.AddItemFromNode(node, childComment)
        Set childComment = Nothing
    ElseIf TypeOf node Is IXMLDOMText Then
        ' the first non-white-space text node is the Value
        If mValueElement Is Nothing Then
            Set mValueElement = node
            If gIswhitespace(mValueElement.nodeValue) Then
                Set mValueElement = Nothing
            Else
            End If
        End If
    ElseIf TypeOf node Is IXMLDOMComment Then
        ' remember this for the Next Element
        Set childComment = node
    Else
        Debug.Assert False
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub RelocateComment()
Const ProcName As String = "RelocateComment"

On Error GoTo Err

If Not mCommentElement Is Nothing Then
    removePreSpacer mCommentElement
    mElement.parentNode.removeChild mCommentElement
    mElement.parentNode.insertBefore mConfigFile.XMLDoc.createTextNode(vbNewLine & String(mIndent - 1, vbTab)), mElement
    mElement.parentNode.insertBefore mCommentElement, mElement
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveAttribute( _
                ByVal Name As String)
Const ProcName As String = "removeAttribute"

On Error GoTo Err

mElement.RemoveAttribute Name
mConfigFile.Dirty

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function SetAttribute( _
                ByVal Name As String, _
                ByVal Value As String) As String
Dim currValue As String
Const ProcName As String = "setAttribute"

On Error GoTo Err

On Error Resume Next
currValue = mElement.GetAttribute(Name)
If Value <> currValue Then
    mElement.SetAttribute Name, Value
    mConfigFile.Dirty
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Terminate()
Dim ci As XMLConfigItem

Const ProcName As String = "Terminate"

On Error GoTo Err

On Error Resume Next

If Not mValueElement Is Nothing Then
    mElement.removeChild mValueElement
    Set mValueElement = Nothing
End If
If Not mCommentElement Is Nothing Then
    mElement.parentNode.removeChild mCommentElement
    Set mCommentElement = Nothing
End If

For Each ci In mChildItems
    ci.Terminate
Next

mElement.parentNode.removeChild mElement

Set mChildItems = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub removePreSpacer(ByVal element As IXMLDOMElement)
Dim textEl As IXMLDOMText
Const ProcName As String = "removePreSpacer"

On Error GoTo Err

If element.previousSibling Is Nothing Then
ElseIf TypeOf element.previousSibling Is IXMLDOMText Then
    Set textEl = element.previousSibling
    If gIswhitespace(textEl.nodeValue) Then
        textEl.parentNode.removeChild textEl
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
    
