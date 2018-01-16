VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class represents a set of <code>Parameter</code> objects.
'
' Duplicate parameter names are not permitted within the set.
'
' Objects of this class can be used with VB's <code>For Each</code> construct to
' enumerate through the items in the set. Applications should not generally depend
' on the items being returned in any particular order. However, after a call
' to the <code>SortByName</code> method, and before any subsequent additions to or
' removals from the list, items will be returned in ascending Name order.
'
' The <code>Parameters</code> class supports data binding. For example if a
' <code>Parameters</code> object is bound to a DataGrid control, any changes made
' to the items in the grid are automatically reflected in the object.
'
' Changes to the list are notified to registered <code>CollectionChangeListener</code>s,
' and via the <code>CollectionChanged</code> event. The following <code>CollectionChangeTypes</code>
' may be notified:
' <pre>
'    <tt>CollItemAdded</tt><br />
'    <tt>CollItemRemoved</tt><br />
'    <tt>CollOrderChanged</tt>
' </pre>
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IConfigurable
Implements IEnumerable
Implements IJSONable

'@================================================================================
' Events
'@================================================================================

''
' Raised when a <code>Parameter</code> is added to or removed from the list in this
' <code>Parameters</code> object, and when the value of a <code>Parameter</code>
' object in the list changes, and when the list is sorted.
'
' @remarks
'   The following members of the <code>CollectionChangeTypes</code> enum may
'   occur in the <code>ChangeType</code> field of the
'   <code>ev</code> parameter:
' <pre>
'    <tt>CollItemAdded</tt><br />
'    <tt>CollItemRemoved</tt><br />
'    <tt>CollOrderChanged</tt>
' </pre>
' @param ev
'   contains details of the change.
'
'@/
Event CollectionChanged( _
                ev As CollectionChangeEventData)
                
'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                                As String = "Parameters"

Private Const ConfigSectionParameter                    As String = "Parameter"

Private Const InitialSize                               As Long = 10

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mRS                                  As Recordset
Attribute mRS.VB_VarHelpID = -1

Private mParamsIndexCol                                 As Collection
Private mParams()                                       As Parameter
Private mParamsIndex                                    As Long

Private mUpdateSequence                                 As Long

Private mChangeListeners                                As Listeners

Private mConfig                                         As ConfigurationSection

Private mDeleting                                       As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_GetDataMember(DataMember As String, Data As Object)
Const ProcName As String = "Class_GetDataMember"
On Error GoTo Err

If mRS Is Nothing Then Set mRS = setupRecordset
Set Data = mRS

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Class_Initialize()
Initialise
Set mChangeListeners = New Listeners
End Sub

'@================================================================================
' Configurable Interface Members
'@================================================================================

Private Property Let IConfigurable_ConfigurationSection(ByVal RHS As ConfigurationSection)
Const ProcName As String = "IConfigurable_ConfigurationSection"
On Error GoTo Err

ConfigurationSection = RHS

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IConfigurable_LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "IConfigurable_LoadFromConfig"
On Error GoTo Err

LoadFromConfig pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub IConfigurable_RemoveFromConfig()
Const ProcName As String = "IConfigurable_RemoveFromConfig"
On Error GoTo Err

RemoveFromConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Const ProcName As String = "IEnumerable_Enumerator"
On Error GoTo Err

Set IEnumerable_Enumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function IEnumerable_NextItem( _
                ByRef pItem As Variant, _
                ByRef pCookie As Variant) As Boolean
Const ProcName As String = "IEnumerable_NextItem"
On Error GoTo Err

If IsEmpty(pCookie) Then
    pCookie = 1
Else
    pCookie = pCookie + 1
End If

If pCookie > mParamsIndex Then
    pCookie = &H7FFFFFFF
    Exit Function
End If

Set pItem = mParams(CLng(pCookie))

IEnumerable_NextItem = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub IEnumerable_Remove(pCookie As Variant)
Const ProcName As String = "IEnumerable_Remove"
On Error GoTo Err

removeParam CLng(pCookie), True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IEnumerable_Skip( _
                ByVal numToSkip As Long, _
                ByRef pCookie As Variant) As Boolean

Const ProcName As String = "IEnumerable_Skip"
On Error GoTo Err

If IsEmpty(pCookie) Then
    pCookie = 1
End If

If pCookie = &H7FFFFFFF Then Exit Function

pCookie = pCookie + numToSkip

If pCookie > mParamsIndex Then
    pCookie = &H7FFFFFFF
    Exit Function
End If

IEnumerable_Skip = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get IEnumerable_UpdateSequence() As Long
IEnumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' JSONable Interface Members
'@================================================================================

Private Function IJSONable_ToJSON() As String
Const ProcName As String = "IJSONable_ToJSON"
On Error GoTo Err

IJSONable_ToJSON = ToJSON

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' mRS Event Handlers
'@================================================================================

Private Sub mRS_FieldChangeComplete( _
                ByVal cFields As Long, _
                ByVal Fields As Variant, _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)
Const ProcName As String = "mRS_FieldChangeComplete"
On Error GoTo Err

If mRS.EditMode <> adEditInProgress Then Exit Sub

Dim lFields() As Object
lFields = Fields

Dim lVar As Variant
For Each lVar In lFields
    Dim lField As ADODB.Field
    Set lField = lVar
    If lField.Name = "Name" Then
        removeParam lField.OriginalValue, False
        SetParamValue lField.Value, mRS!Value, False
    End If
Next

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mRS_RecordChangeComplete( _
                ByVal adReason As ADODB.EventReasonEnum, _
                ByVal cRecords As Long, _
                ByVal pError As ADODB.Error, _
                adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)
Const ProcName As String = "mRS_RecordChangeComplete"
On Error GoTo Err

Select Case adReason
Case adRsnAddNew

Case adRsnDelete

Case adRsnUpdate
    If mDeleting Then
        mDeleting = False
    Else
        SetParamValue pRecordset!Name, pRecordset!Value, False
    End If
Case adRsnUndoUpdate

Case adRsnUndoAddNew

Case adRsnUndoDelete

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

Private Sub mRS_WillChangeRecord( _
                ByVal adReason As ADODB.EventReasonEnum, _
                ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, _
                ByVal pRecordset As ADODB.Recordset)
Const ProcName As String = "mRS_WillChangeRecord"
On Error GoTo Err

Select Case adReason
Case adRsnAddNew

Case adRsnDelete
    mDeleting = True
    removeParam pRecordset!Name, False
Case adRsnUpdate

Case adRsnUndoUpdate

Case adRsnUndoAddNew

Case adRsnUndoDelete

End Select

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Get arrayAddress() As Long
Const ProcName As String = "arrayAddress"
On Error GoTo Err

arrayAddress = VarPtr(mParams(1))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Supplies a <code>ConfigurationSection</code> object in which the list members are
' stored so that an identical <code>Parameters</code> object can be created later
' via a call to the global <code>LoadParametersFromConfig</code> method.
'
' @param Value
'   The <code>ConfigurationSection</code> object in which the list members are to be stored.
'@/
Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"
On Error GoTo Err

If Value Is mConfig Then Exit Property
Set mConfig = Value

storeSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the number of <code>Parameter</code> objects in the list.
'
' @return
'   The the number of <code>Parameter</code> objects in the list.
'@/
Public Property Get Count() As Long
Const ProcName As String = "Count"
On Error GoTo Err

Count = mParamsIndex

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an <code>Enumerator</code> object that can be used to iterate through
' the members of the collection.
'
' @return
'   An <code>Enumerator</code> object.
'@/
Public Property Get Enumerator() As Enumerator
Const ProcName As String = "Enumerator"
On Error GoTo Err

Set Enumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Clears the current parameters and then sets new parameters
' whose names and values are the keys and items
' of the supplied <code>Dictionary</code> object.
'
' @param Value
'   A <code>Dictionary</code> object.
'@/
Public Property Let ParamsDictionary(ByVal Value As Dictionary)
Const ProcName As String = "ParamsDictionary"
On Error GoTo Err

Clear

If Value Is Nothing Then Exit Property

Dim Keys() As Variant
Keys = Value.Keys

Dim Var As Variant
For Each Var In Keys
    SetParameterValue Var, Value.Item(Var)
Next

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Registers an object to act as a listener for changes to the list.
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

Friend Sub AddParameter( _
                ByVal param As Parameter, _
                ByVal addToConfig As Boolean, _
                ByVal pUpdateRS As Boolean)
Const ProcName As String = "AddParameter"
On Error GoTo Err

If mParamsIndex = UBound(mParams) Then
    ReDim Preserve mParams(1 To UBound(mParams) * 2) As Parameter
End If
mParamsIndex = mParamsIndex + 1
Set mParams(mParamsIndex) = param
mParamsIndexCol.Add CStr(mParamsIndex), param.Name
If pUpdateRS And Not mRS Is Nothing Then addParamToRecordset param, mRS

If addToConfig Then storeParamSettings param

changed

fireChange CollItemAdded, param

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

clearRecordset
Initialise

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'
' Creates a new <code>Parameters</code> object containing a set of <code>Parameter</code>
' objects with the same names and values as those in this object, and in the same order.
' @return
'   A Clone of this object.
'@/
Public Function Clone() As Parameters
Const ProcName As String = "Clone"
On Error GoTo Err

Set Clone = New Parameters
Dim i As Long
For i = 1 To mParamsIndex
    Dim param As Parameter
    Set param = mParams(i)
    Clone.SetParameterValue param.Name, param.Value
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Compares this object for equality with another <code>Parameters</code>
' object.
' @remarks
'   To be considered equal, both objects must have the same set of <code>Parameter</code>
'   objects (ie the same names and values), though not necessarily in the same order.
' @return
'   <code>True</code> if the objects are equal, and <code>False</code> if not.
' @param params
'   The object to which this object is to be compared.
'@/
Public Function Equals(params As Parameters) As Boolean
Const ProcName As String = "Equals"
On Error GoTo Err

If params Is Nothing Then Exit Function

If params Is Me Then
    Equals = True
    Exit Function
End If

If Me.Count <> params.Count Then Equals = False: Exit Function
If Me.Count = 0 Then Equals = True: Exit Function

If Me.Count = 1 Then
    If Me.Item(1).Equals(params.Item(1)) Then Equals = True: Exit Function
End If

' we don't want to do a permanent sort, so we'll copy the lists into temporary arrays
' and sort them
ReDim myTemp(1 To mParamsIndex) As Parameter
ReDim itsTemp(1 To mParamsIndex) As Parameter

CopyMemory VarPtr(myTemp(1)), VarPtr(mParams(1)), 4 * mParamsIndex
CopyMemory VarPtr(itsTemp(1)), params.arrayAddress, 4 * mParamsIndex

gSortTypedObjects VarPtr(myTemp(1)), mParamsIndex, False
gSortTypedObjects VarPtr(itsTemp(1)), mParamsIndex, False

Equals = True

Dim i As Long
For i = 1 To mParamsIndex
    If Not myTemp(i).Equals(itsTemp(i)) Then
        Equals = False
        Exit For
    End If
Next

ZeroMemory VarPtr(myTemp(1)), 4 * mParamsIndex
ZeroMemory VarPtr(itsTemp(1)), 4 * mParamsIndex

Erase myTemp
Erase itsTemp

changed

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the value of the specified parameter.
' @remarks
'   If this object does not contain a <code>Parameter</code> object with
'   the specified Name, the specified default value is returned.
' @return
'   The value of the specified parameter.
' @param Name
'   The Name of the parameter whose value is to be returned.
' @param defaultValue
'   This value is returned by the method if the specified parameter
'   is not included in the set. If this value is not specified, it defaults to
'   an empty string.
'@/
Public Function GetParameterValue( _
                ByVal Name As String, _
                Optional ByVal defaultValue As String) As String
Const ProcName As String = "GetParameterValue"
On Error GoTo Err

Dim lIndex As Long
lIndex = findParameterIndex(Name)

If lIndex = 0 Then
    GetParameterValue = defaultValue
Else
    GetParameterValue = mParams(lIndex).Value
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Returns the specified <code>Parameter</code> object.
'
' @return
'   The <code>Parameter</code> object specified by the <code>index</code> argument.
' @param pIndex
'   Either the parameter's Name or its ordinal position in the list.
' @see
'
'@/
Public Function Item( _
                ByVal pIndex As Variant) As Parameter
Const ProcName As String = "Item"
On Error GoTo Err

If gIsInteger(pIndex, 1, mParamsIndex) Then
    Set Item = mParams(pIndex)
ElseIf IsNumeric(pIndex) Then
    Err.Raise VBErrorCodes.VbErrInvalidProcedureCall, , "Invalid procedure call or argument"
Else
    Set Item = mParams(findParameterIndex(pIndex))
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub LoadFromConfig( _
                ByVal Config As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
On Error GoTo Err

gAssertArgument Not Config Is Nothing, "A ConfigurationSection object must be supplied"

Set mConfig = Config

Dim paramSect As ConfigurationSection
For Each paramSect In mConfig
    Dim param As Parameter
    Set param = New Parameter
    param.LoadFromConfig paramSect
    AddParameter param, False, True
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = getEnumerator
End Function

''
' Unregisters a collection change listener.
' @param listener
'   The listener object which is to be unregistered.
'@/
Public Sub RemoveCollectionChangeListener(ByVal pListener As ICollectionChangeListener)
Const ProcName As String = "RemoveCollectionChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Removes the <code>ConfigurationSection</code> object in which the list members have
' been stored from the configuration dataset.
'
' @remarks
'   No Error occurs if this method is called when no <code>ConfigurationSection</code>
'   object has been assigned to this <code>Parameters</code> object.
'@/
Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Removes the specified parameter.
'
' @param pIndex
'   Either the parameter's Name or its ordinal position in the list.
' @see
'
'@/
Public Sub RemoveParameter( _
                ByVal pIndex As Variant)
Const ProcName As String = "RemoveParameter"
On Error GoTo Err

removeParam pIndex, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Sets the value of the specified parameter.
' @remarks
'   If a parameter with this Name is already included, its value is
'   updated; otherwise a new parameter is added to the list.
' @return
'   The <code>Parameter</code> object whose value has been set.
' @param Name
'   The Name of the parameter whose value is to be set.
' @param Value
'   The value to be given to the specified parameter.
'@/
Public Function SetParameterValue( _
                ByVal Name As String, _
                ByVal Value As String) As Parameter
Const ProcName As String = "SetParameterValue"
On Error GoTo Err

Set SetParameterValue = SetParamValue(Name, Value, True)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function SetParamValue( _
                ByVal Name As String, _
                ByVal Value As String, _
                ByVal pUpdateRS As Boolean) As Parameter
Const ProcName As String = "SetParamValue"
On Error GoTo Err

If Name = "" Then Exit Function

changed

gAssertArgument Not IsNumeric(Name), "Name cannot be a number"

Dim lIndex As Long
lIndex = findParameterIndex(Name)

Dim param As Parameter
If lIndex = 0 Then
    Set param = New Parameter
    param.Name = Name
    param.Value = Value
    AddParameter param, True, pUpdateRS
Else
    Set param = mParams(lIndex)
    If param.Value <> Value Then
        updateParameterByIndex lIndex, Value, pUpdateRS
    End If
End If
Set SetParamValue = param

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Sets the value(s) of one or more specified parameters.
' @remarks
'   The parameter names and values are supplied in a <code>String</code>
'   with the following syntax:
'   <pre>
'   <tt>
'   <p>paramString: param [sep param]...</p>
'
'   <p>param: Name delim value</p>
'
'   <p>name: string</p>
'
'   <p>value: string</p>
'   </tt>
'   </pre>
'   where
'   <pre>
'   <p><tt>sep</tt> is a single Character used as a separator between parameters.
'   The default separator is a semi-colon, but an alternative can be specified
'   in the <code>parameterSeparator</code> argument.</p>
'
'   <p><tt>delim</tt> is a single Character used to delimit the Name
'   The default delimiter is =, but an alternative can be specified
'   in the <code>nameDelimiter</code> argument.</p>
'   </pre>
'
'   <p>If the parameter Name or value needs to contain a separator or delimiter
'   Character, this can be achieved by preceding it with an escape Character.
'   The default escape Character is \, but an alternative can be specified
'   in the <code>escapeCharacter</code> argument. Note that to include the
'   escape Character itself in the value or Name, it must be repeated. The
'   effect of the escape Character is to include the following Character in
'   the parameter Name or value verbatim.</p>
'
'   <p>For each specified parameter, if a parameter with this Name is
'   already included, its value is updated; otherwise a new parameter
'   is added to the list.</p>
' <h3>Examples</h3>
'   An example parameter string, using the default separator and delimiter
'   characters:
'   <pre>
'   <tt>Server=Delphi;Port=7496;ClientId=432</tt>
'   </pre>
'   An example incuding the escape Character in the value:
'   <pre>
'   <tt>SettingsFile=C:\\MyProg\\myprog.ini</tt>
'   </pre>
'   An example using the escape Character in the value in an unusual (and pointless!) way:
'   <pre>
'   <tt>My parameter=\a\b\c\d\e\f</tt>
'   </pre>
'   Note that in this case, the value would be interpreted as abcdef.
' @return
'   The <code>Parameter</code> objects whose values have been set.
' @param paramString
'   The parameters to be set, using the syntex described under Remarks.
' @param nameDelimiter
'   A single Character used to delimit the parameter Name.
' @param parameterSeparator
'   A single Character used to separate parameters.
' @param escapeCharacter
'   A single Character that causes the following Character to be included
'   verbatim in the parameter value in spite of any special meaning it may normally have.
'@/
Public Function SetParameterValues( _
                ByVal paramString As String, _
                Optional ByVal nameDelimiter As String = "=", _
                Optional ByVal parameterSeparator As String = ";", _
                Optional ByVal escapeCharacter As String = "\") As Parameter()
Const ProcName As String = "SetParameterValues"
On Error GoTo Err

ReDim params(3) As Parameter

Dim paramsIndex As Long
paramsIndex = -1

changed

Dim ps As ParameterStringParser
Set ps = gCreateParameterStringParser(paramString, _
                            nameDelimiter, _
                            parameterSeparator, _
                            escapeCharacter)
                            
Dim Name As String
Dim Value As String
Do While ps.GetParameter(Name, Value)
    If paramsIndex = UBound(params) Then
        ReDim Preserve params(2 * UBound(params) - 1) As Parameter
    End If
    paramsIndex = paramsIndex + 1
    Set params(paramsIndex) = SetParameterValue(Name, Value)
Loop

If paramsIndex >= 0 Then
    ReDim Preserve params(paramsIndex) As Parameter
Else
    Erase params
End If

SetParameterValues = params

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Sorts the parameters in the list in order of Name.
' @remarks
'   After calling this function, iterating through the list using the
'   For Each construct will return the parameters in ascending Name order.
'   Once another parameter has been added or one has been removed, this
'   may no longer be the case. The application must not make any assumptions
'   about the order after adding or removing a parameter.
' @param descending
'   If <code>True</code>, the parameters are sorted in descending Name order.
'@/
Public Sub SortByName( _
                Optional ByVal descending As Boolean)
Const ProcName As String = "SortByName"
On Error GoTo Err

changed

If mParamsIndex = 0 Then Exit Sub

gSortTypedObjects VarPtr(mParams(1)), mParamsIndex, descending

Set mParamsIndexCol = New Collection
Dim i As Long
For i = 1 To mParamsIndex
    mParamsIndexCol.Add CStr(i), mParams(i).Name
Next

If Not mRS Is Nothing Then
    clearRecordset
    populateRecordset mRS
End If

If Not mConfig Is Nothing Then mConfig.RemoveAllChildren
storeSettings

fireChange CollOrderChanged

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Returns a <code>Dictionary</code> object whose keys and items
' are the names and values of the members of the collection.
'
' @return
'   A <code>Dictionary</code> object.
'@/
Public Function ToDictionary() As Dictionary
Const ProcName As String = "ToDictionary"
On Error GoTo Err

Dim dict As New Dictionary
Dim i As Long
For i = 1 To mParamsIndex
    Dim param As Parameter
    Set param = mParams(i)
    dict.Add param.Name, param.Value
Next

Set ToDictionary = dict

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function ToJSON() As String
Const ProcName As String = "ToJSON"
On Error GoTo Err

Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "ParamsDictionary"
jb.EncodeValue ToDictionary

jb.EndObject

ToJSON = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addParamToRecordset( _
                ByVal pParam As Parameter, _
                ByVal pRS As Recordset)
Const ProcName As String = "addParamToRecordset"
On Error GoTo Err

pRS.AddNew
pRS!Name = pParam.Name
pRS!Value = pParam.Value
pRS.Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub changed()
mUpdateSequence = mUpdateSequence + 1
End Sub

Private Sub clearRecordset()
Const ProcName As String = "clearRecordset"
On Error GoTo Err

If mRS Is Nothing Then Exit Sub

Dim lRS As Recordset
Set lRS = mRS

' disable event handling
Set mRS = Nothing

If Not (lRS.BOF And lRS.EOF) Then
    lRS.MoveFirst
    Do While Not lRS.EOF
        lRS.Delete adAffectCurrent
        lRS.Update
        lRS.MoveNext
    Loop
End If

' re-enable event handling
Set mRS = lRS

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function createRecordset() As Recordset
Const ProcName As String = "createRecordset"
On Error GoTo Err

Dim lRS As New Recordset
lRS.CursorType = adOpenStatic
lRS.LockType = adLockOptimistic

lRS.Fields.Append "Name", adVarChar, 255
lRS.Fields.Append "Value", adVarChar, 255
lRS.CursorType = adOpenStatic
lRS.LockType = adLockOptimistic
lRS.Open

Set createRecordset = lRS

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function findParameterIndex( _
                ByVal Name As String) As Long

Const ProcName As String = "findParameterIndex"
On Error GoTo Err

On Error Resume Next
findParameterIndex = mParamsIndexCol(Name)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fireChange( _
                ByVal ChangeType As CollectionChangeTypes, _
                Optional ByVal affectedObject As Parameter)
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

Private Function getEnumerator() As Enumerator
Const ProcName As String = "getEnumerator"
On Error GoTo Err

Dim lEnum As Enumerator
Set lEnum = New Enumerator
lEnum.Initialise Me
Set getEnumerator = lEnum

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub Initialise()
Set mParamsIndexCol = New Collection
ReDim mParams(1 To InitialSize) As Parameter
mParamsIndex = 0
End Sub

Private Sub populateRecordset(ByVal pRS As Recordset)
Const ProcName As String = "populateRecordset"
On Error GoTo Err

Dim i As Long
For i = 1 To mParamsIndex
    Dim lParam As Parameter
    Set lParam = mParams(i)
    addParamToRecordset lParam, pRS
Next

If Not (pRS.BOF And pRS.EOF) Then pRS.MoveFirst

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeParameterByIndex( _
                ByVal pIndex As Long, _
                ByVal pUpdateRS As Boolean)
Const ProcName As String = "removeParameterByIndex"
On Error GoTo Err

Dim lRemovedParam As Parameter
Set lRemovedParam = mParams(pIndex)

mParamsIndexCol.Remove lRemovedParam.Name

Dim lVar As Variant
For Each lVar In mParamsIndexCol
    Dim i As Long
    i = lVar
    If i > pIndex Then
        mParamsIndexCol.Remove mParams(i).Name
        mParamsIndexCol.Add i - 1, mParams(i).Name
    End If
Next

For i = pIndex + 1 To mParamsIndex
    Set mParams(i - 1) = mParams(i)
    Set mParams(i) = Nothing
Next

mParamsIndex = mParamsIndex - 1

If pUpdateRS And Not mRS Is Nothing Then removeParamFromRecordset lRemovedParam, mRS
lRemovedParam.RemoveFromConfig
fireChange CollItemRemoved, lRemovedParam

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeParamFromRecordset( _
                ByVal pParam As Parameter, _
                ByVal pRS As Recordset)
Const ProcName As String = "removeParamFromRecordset"
On Error GoTo Err

pRS.MoveFirst
pRS.Find "Name='" & pParam.Name & "'"
gAssert Not mRS.EOF, "Record not found"
pRS.Delete adAffectCurrent
pRS.Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeParam( _
                ByVal pIndex As Variant, _
                ByVal pUpdateRS As Boolean)
Const ProcName As String = "removeParam"
On Error GoTo Err

changed

If gIsInteger(pIndex, 1, mParamsIndex) Then
    removeParameterByIndex pIndex, pUpdateRS
ElseIf IsNumeric(pIndex) Then
    Err.Raise VBErrorCodes.VbErrInvalidProcedureCall, , "Invalid procedure call or argument"
Else
    removeParameterByIndex findParameterIndex(pIndex), pUpdateRS
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function setupRecordset() As Recordset
Const ProcName As String = "setupRecordset"
On Error GoTo Err

Dim lRS As Recordset
Set lRS = createRecordset

populateRecordset lRS

Set setupRecordset = lRS

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub storeParamSettings( _
                ByVal param As Parameter)
Const ProcName As String = "storeParamSettings"
On Error GoTo Err

If Not mConfig Is Nothing Then param.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionParameter & "(" & gGenerateGUIDString & ")")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub storeSettings()
Const ProcName As String = "storeSettings"
On Error GoTo Err

Dim en As Enumerator
Set en = getEnumerator
Do While en.MoveNext
    Dim param As Parameter
    Set param = en.Current
    storeParamSettings param
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateParameterByIndex( _
                ByVal pIndex As Long, _
                ByVal Value As String, _
                ByVal pUpdateRS As Boolean)
Const ProcName As String = "updateParameterByIndex"
On Error GoTo Err

mParams(pIndex).Value = Value
If pUpdateRS And Not mRS Is Nothing Then updateParamInRecordset mParams(pIndex), mRS
fireChange CollItemChanged, mParams(pIndex)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub updateParamInRecordset( _
                ByVal pParam As Parameter, _
                ByVal pRS As Recordset)
Const ProcName As String = "updateParamInRecordset"
On Error GoTo Err

pRS.MoveFirst
pRS.Find "Name='" & pParam.Name & "'"
gAssert Not mRS.EOF, "Record not found"
pRS!Value = pParam.Value
pRS.Update

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




