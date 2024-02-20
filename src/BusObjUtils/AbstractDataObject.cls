VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AbstractDataObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
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

Public Event Dirty()
Public Event Clean()
Public Event Deleted()
Public Event Invalid()
Public Event Valid()

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                        As String = "AbstractDataObject"

'@================================================================================
' Member variables
'@================================================================================

Private mDataObjRef                             As WeakReference

Private mFactory                                As DataObjectFactory
Private mConnectionString                       As String
Private mConnection                                   As ADODB.Connection
Private mRs                                     As ADODB.Recordset

Private mNew                                    As Boolean
Private mDeleted                                As Boolean
Private mDirty                                  As Boolean

Private WithEvents mErrorList                   As ErrorList
Attribute mErrorList.VB_VarHelpID = -1

Private mBusRuleNameValidId                     As Long

Private mEventRaiser                            As BusinessObjEventRaiser

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mErrorList = New ErrorList
Set mEventRaiser = New BusinessObjEventRaiser
End Sub

'@===============================================================================
' mErrorList Event Handlers
'@===============================================================================

Private Sub mErrorList_ItemAdded( _
                ByVal Item As ErrorItem)
Const ProcName As String = "mErrorList_ItemAdded"
Dim failpoint As String
On Error GoTo Err

RaiseEvent Invalid
mEventRaiser.Invalid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, pFailpoint:=failpoint
End Sub

Private Sub mErrorList_NoItems()
Const ProcName As String = "mErrorList_NoItems"
Dim failpoint As String
On Error GoTo Err

RaiseEvent Valid
mEventRaiser.Valid

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName, pFailpoint:=failpoint
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ColumnMaxLength( _
                ByVal columnName As String) As Long
Const ProcName As String = "ColumnMaxLength"
Dim failpoint As String
On Error GoTo Err

Assert Not mRs Is Nothing, "Object has been deleted"

ColumnMaxLength = mRs.Fields(columnName).DefinedSize

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let ColumnValue( _
                ByVal columnName As String, _
                ByVal Value As Variant)
Dim fld As ADODB.Field

Const ProcName As String = "ColumnValue"
Dim failpoint As String
On Error GoTo Err

Assert Not IsDeleted, "Object has been deleted"

Set fld = mRs.Fields(columnName)
Select Case fld.Type
Case adBSTR, _
        adChar, _
        adVarChar, _
        adLongVarChar, _
        adWChar, _
        adVarWChar, _
        adLongVarWChar, _
        adVarBinary, _
        adLongVarBinary
    If Len(Value) > fld.DefinedSize Then AssertArgument False, "Value too long: Column=" & fld.Name & _
                "; max length=" & fld.DefinedSize & _
                "; value=" & IIf(IsNull(Value), "NULL", VariantToString(Value))
End Select

If IsNull(Value) Or IsNull(fld.Value) Or Value <> fld.Value Then
    fld.Value = Value
    Dirty
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ErrorList() As ErrorList
Const ProcName As String = "ErrorList"
Dim failpoint As String
On Error GoTo Err

Assert Not mRs Is Nothing, "Object has been deleted"

Set ErrorList = mErrorList

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get EventSource() As BusinessObjEventSource
Const ProcName As String = "EventSource"
Dim failpoint As String
On Error GoTo Err

Assert Not mRs Is Nothing, "Object has been deleted"

Set EventSource = mEventRaiser.EventSource

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Id() As Long
Const ProcName As String = "Id"
Dim failpoint As String
On Error GoTo Err

Id = GetColumnValue(GenericColumnId)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsValid() As Boolean
Const ProcName As String = "IsValid"
Dim failpoint As String
On Error GoTo Err

Assert Not mRs Is Nothing, "Object has been deleted"

IsValid = (mErrorList.Count = 0)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsDeleted() As Boolean
IsDeleted = (mDeleted Or mRs Is Nothing)
End Property

Public Property Get IsDirty() As Boolean
Const ProcName As String = "IsDirty"
Dim failpoint As String
On Error GoTo Err

Assert Not mRs Is Nothing, "Object has been deleted"

IsDirty = mDirty

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get IsNew() As Boolean
Const ProcName As String = "IsNew"
Dim failpoint As String
On Error GoTo Err

Assert Not mRs Is Nothing, "Object has been deleted"

IsNew = mNew

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Const ProcName As String = "Name"
Dim failpoint As String
On Error GoTo Err

Name = GetColumnValue(GenericColumnName)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Name(ByVal Value As String)
Const ProcName As String = "Name"
Dim failpoint As String
On Error GoTo Err

Value = Trim$(Value)
If Value = Name Then Exit Property
RuleBroken mBusRuleNameValidId, _
            Len(Value) = 0 Or _
            Len(Value) > ColumnMaxLength(GenericColumnName) Or _
            Not mFactory.LoadByName(Value) Is Nothing
ColumnValue(GenericColumnName) = Left$(Value, ColumnMaxLength(GenericColumnName))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ApplyEdit()
Const ProcName As String = "ApplyEdit"
Dim failpoint As String
On Error GoTo Err

Assert Not mRs Is Nothing, "Object has been deleted"

If mDeleted Then
    mFactory.Delete mRs("Id")
    Set mRs = Nothing
    RaiseEvent Deleted
    mEventRaiser.Deleted
    mDeleted = False
ElseIf mDirty Or mNew Then
    Assert IsValid, "Object is not valid for saving"

    save
    mNew = False

    Refresh
End If
Clean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub CancelEdit()
Const ProcName As String = "CancelEdit"
Dim failpoint As String
On Error GoTo Err

Assert Not mRs Is Nothing, "Object has been deleted"

mRs.CancelUpdate
mDeleted = False
If mRs.EOF Then
    mRs.AddNew
    mNew = True
    RuleBroken mBusRuleNameValidId, True
End If
Clean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Delete()
Const ProcName As String = "Delete"
Dim failpoint As String
On Error GoTo Err

Assert Not mRs Is Nothing, "Object has been deleted"

mDeleted = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function GetColumnValue( _
                ByVal columnName As String, _
                Optional ByVal defaultValue As Variant) As Variant
Const ProcName As String = "GetColumnValue"
Dim failpoint As String
On Error GoTo Err

Assert Not IsDeleted, "Object has been deleted"


GetColumnValue = gGetColumnValue(mRs, columnName, defaultValue)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal dataObj As BusinessDataObject, _
                ByVal busRuleNameValidId As Long, _
                ByVal factory As DataObjectFactory, _
                ByVal connectionString As String)
Set mDataObjRef = CreateWeakReference(dataObj)
mBusRuleNameValidId = busRuleNameValidId
Set mFactory = factory
mConnectionString = connectionString
End Sub

Public Sub InitialiseRs( _
                ByVal rs As ADODB.Recordset)
Const ProcName As String = "InitialiseRs"
Dim failpoint As String
On Error GoTo Err

Set mRs = rs

If mRs.EOF Then
    mRs.AddNew
    mNew = True
    RuleBroken mBusRuleNameValidId, True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Refresh()
Const ProcName As String = "Refresh"
Dim failpoint As String
On Error GoTo Err

Assert Not mNew, "Data object has not yet been committed to permanent storage"
Assert Not IsDeleted, "Object has been deleted"

mFactory.Refresh getDataObj

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function RuleBroken( _
                ByVal ruleId As Long, _
                ByVal isBroken As Boolean) As Boolean
Const ProcName As String = "RuleBroken"
Dim failpoint As String
On Error GoTo Err

Assert Not IsDeleted, "Object has been deleted"


mErrorList.AssertRule (Not isBroken), ruleId
RuleBroken = isBroken

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub Clean()
Const ProcName As String = "Clean"
Dim failpoint As String
On Error GoTo Err

mDirty = False
RaiseEvent Clean
mEventRaiser.Clean
mErrorList.clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Dirty()
Const ProcName As String = "Dirty"
Dim failpoint As String
On Error GoTo Err

If Not mDirty Then
    mDirty = True
    RaiseEvent Dirty
    mEventRaiser.Dirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getDataObj() As BusinessDataObject
Const ProcName As String = "getDataObj"
Dim failpoint As String
On Error GoTo Err

Set getDataObj = mDataObjRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub save()
Dim failpoint As String
Const ProcName As String = "save"

On Error GoTo Err

failpoint = 100

Debug.Print "AbstractDataObject: Connect to DB"
Set mConnection = New ADODB.Connection
mConnection.Open mConnectionString

failpoint = 200

Set mRs.ActiveConnection = mConnection
mRs.UpdateBatch
Set mRs.ActiveConnection = Nothing
Debug.Print "AbstractDataObject: close connection"
mConnection.Close

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub








