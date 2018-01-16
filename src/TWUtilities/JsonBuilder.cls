VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "JSONBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Objects of this class are used to build JSON strings.
'
' JSON is JavaScript Object Notation - see http://www.json.org and
' http://www.ietf.org/rfc/rfc4627.txt?number=4627. JSON is a lightweight, text-based,
' language-independent data interchange format.
'
' A class that implements the <code>JSONable</code> interface can use a
' <code>JSONBuilder</code> object in the implementation of the <code>ToJSON</code>
' method to ensure that well-formed JSON is generated.
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

Private Const ModuleName                            As String = "JSONBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mSB                                         As StringBuilder

Private mNextJB                                     As JSONBuilder
Private mParent                                     As JSONBuilder

Private mStateEngine                                As StateEngine

Private mName                                       As String
Private mValue                                      As Variant

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStateEngine = New StateEngine
mStateEngine.Initialise GJSON.TableBuilder, _
                        JSONBuilderStates.Created
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
' Returns the JSON string resulting from a valid and complete sequence of calls to the
' various methods of this class.
'
' @remarks
'   Invoking this property before a valid JSON string has been built causes an
'   <code>ErrorCodes.ErrIllegalStateException</code> error to be raised.
' @return
'   The JSON string.
'@/
Public Property Get JSONString() As String
Const ProcName As String = "JSONString"

On Error GoTo Err

If mStateEngine.State <> JSONBuilderStates.Finished Then
    Err.Raise ErrorCodes.ErrIllegalStateException, , "JSON building is not yet complete"
End If

JSONString = mSB.ToString

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Encodes the JSON <code>begin-array</code> structural character.
'
'@/
Public Sub BeginArray()
Const ProcName As String = "BeginArray"

On Error GoTo Err

If mNextJB Is Nothing Then
    handleStimulus StimBeginArray
Else
    mNextJB.BeginArray
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Encodes the JSON <code>begin-object</code> structural character.
'
'@/
Public Sub BeginObject()
Const ProcName As String = "BeginObject"

On Error GoTo Err

If mNextJB Is Nothing Then
    handleStimulus StimBeginObject
Else
    mNextJB.BeginObject
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Encodes the name part of a name/value pair.
'
' @param pName
'   The name to be encoded.
'@/
Public Sub EncodeName( _
                ByVal pName As String)
Const ProcName As String = "EncodeName"
On Error GoTo Err

If mNextJB Is Nothing Then
    mName = pName
    handleStimulus StimName
Else
    mNextJB.EncodeName pName
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Encodes a name/value pair.
'
' @param pName
'   The name to be encoded.
' @param pValue
'   The value to be encoded.
'@/
Public Sub EncodeNameAndValue( _
                ByVal pName As String, _
                ByRef pValue As Variant, _
                Optional ByRef pDefaultValue As Variant)
Const ProcName As String = "EncodeName"
On Error GoTo Err

If gVariantEquals(pValue, pDefaultValue) Then Exit Sub

If mNextJB Is Nothing Then
    mName = pName
    handleStimulus StimName
Else
    mNextJB.EncodeName pName
End If

EncodeValue pValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Encodes the value part of a name/value pair, or an element of an array.
'
' @param pValue
'   The value to be encoded.
'@/
Public Sub EncodeValue( _
                ByRef pValue As Variant)
Const ProcName As String = "EncodeValue"
On Error GoTo Err

If mNextJB Is Nothing Then
    gSetVariant mValue, pValue
    handleStimulus StimValue
Else
    mNextJB.EncodeValue pValue
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Encodes the JSON <code>end-array</code> structural character.
'
'@/
Public Sub EndArray()
Const ProcName As String = "EndArray"

On Error GoTo Err

If mNextJB Is Nothing Then
    handleStimulus StimEndArray
Else
    mNextJB.EndArray
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Encodes the JSON <code>end-object</code> structural character.
'
'@/
Public Sub EndObject()
Const ProcName As String = "EndObject"

On Error GoTo Err

If mNextJB Is Nothing Then
    handleStimulus StimEndObject
Else
    mNextJB.EndObject
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal Parent As JSONBuilder, _
                ByVal sb As StringBuilder)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mParent = Parent
Set mSB = sb

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub UnChain()
Const ProcName As String = "UnChain"

On Error GoTo Err

Set mNextJB = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Append( _
                ByVal s As String)
Const ProcName As String = "Append"

On Error GoTo Err

If mSB Is Nothing Then
    Set mSB = New StringBuilder
    mSB.Initialise
End If
mSB.Append s

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub chainNewJB()
Const ProcName As String = "chainNewJB"

On Error GoTo Err

Set mNextJB = New JSONBuilder
mNextJB.Initialise Me, mSB

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub doActions(ByRef pActions() As Long)
Dim Var As Variant
Dim lAction As JSONBuilderActions

Const ProcName As String = "doActions"

On Error GoTo Err

For Each Var In pActions
    lAction = Var
    Select Case lAction
    Case SpecialActions.NoAction

    Case ActionEncodeBeginArray
        DoEncodeBeginArray
    Case ActionEncodeEndArray
        DoEncodeEndArray
    Case ActionEncodeBeginObject
        DoEncodeBeginObject
    Case ActionEncodeEndObject
        DoEncodeEndObject
    Case ActionEncodeNameSeparator
        Append GJSON.NameSeparator
    Case ActionEncodeValueSeparator
        Append GJSON.ValueSeparator
    Case ActionEncodeName
        gEncodeVariant mName, mSB
    Case ActionEncodeValue
        gEncodeVariant mValue, mSB
    Case ActionIncreaseLevel
        chainNewJB
    Case ActionDecreaseLevel
        If Not mParent Is Nothing Then
            mParent.UnChain
            Set mParent = Nothing
        End If
    Case Else
    Err.Raise ErrorCodes.ErrIllegalStateException, , "Invalid Action " & lAction
    End Select
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub DoEncodeBeginArray()
Const ProcName As String = "DoEncodeBeginArray"

On Error GoTo Err

    If mNextJB Is Nothing Then
        Append GJSON.BeginArray
    Else
        mNextJB.BeginArray
    End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub DoEncodeBeginObject()
Const ProcName As String = "DoEncodeBeginObject"

On Error GoTo Err

    If mNextJB Is Nothing Then
        Append GJSON.BeginObject
    Else
        mNextJB.BeginObject
    End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub DoEncodeEndArray()
Const ProcName As String = "DoEncodeEndArray"

On Error GoTo Err

    If mNextJB Is Nothing Then
        Append GJSON.EndArray
    Else
        mNextJB.EndArray
    End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub DoEncodeEndObject()
Const ProcName As String = "DoEncodeEndObject"

On Error GoTo Err

    If mNextJB Is Nothing Then
        Append GJSON.EndObject
    Else
        mNextJB.EndObject
    End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub handleStimulus(ByVal Stimulus As JSONBuilderStimuli)
Const ProcName As String = "handleStimulus"

On Error GoTo Err

doActions mStateEngine.NotifyStimulus(Stimulus)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


