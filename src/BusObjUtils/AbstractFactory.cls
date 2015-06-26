VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AbstractDOFactory"
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

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                        As String = "AbstractDOFactory"

'@================================================================================
' Member variables
'@================================================================================

Private mFactoryRef                             As WeakReference

Private mConnectionString                       As String

Private mBaseTableName                          As String
Private mDetailsViewName                        As String

Private mConnection                             As ADODB.Connection
Private mRs                                     As ADODB.Recordset

Private mSpecifiers                             As FieldSpecifiers

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mSpecifiers = New FieldSpecifiers
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

Public Property Get FieldSpecifiers() As FieldSpecifiers
Set FieldSpecifiers = mSpecifiers
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddField( _
                ByVal Name As String, _
                ByVal dbColumnName As String, _
                ByVal width As Long, _
                ByVal align As FieldAlignments, _
                ByVal visible As Boolean)
Const ProcName As String = "AddField"
Dim failpoint As String
On Error GoTo Err

mSpecifiers.AddField Name, dbColumnName, width, align, visible

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Delete( _
                ByVal Id As Long)
Dim sql As String

Const ProcName As String = "Delete"
Dim failpoint As String
On Error GoTo Err

Debug.Print "AbstractDOFactory: Connect to DB"
Set mConnection = New ADODB.Connection
mConnection.Open mConnectionString

sql = "DELETE FROM " & mBaseTableName & _
        " WHERE ID=" & Id

mConnection.Execute sql, , adCmdText
Debug.Print "AbstractDOFactory: close connection"
mConnection.Close

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                

Friend Sub Initialise( _
                ByVal factory As DataObjectFactory, _
                ByVal connectionString As String, _
                ByVal baseTableName As String, _
                ByVal detailsViewName As String)
Const ProcName As String = "Initialise"
Dim failpoint As String
On Error GoTo Err

Set mFactoryRef = CreateWeakReference(factory)
mConnectionString = connectionString
mBaseTableName = baseTableName
mDetailsViewName = detailsViewName

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function LoadByName( _
                ByVal pName As String) As ADODB.Recordset
Dim where As String
Const ProcName As String = "LoadByName"
Dim failpoint As String
On Error GoTo Err

where = "Name='" & gCleanQueryArg(pName) & "'"
Set LoadByName = LoadDetails(where)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByNameAsync( _
                ByVal pName As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim where As String
Const ProcName As String = "LoadByNameAsync"
Dim failpoint As String
On Error GoTo Err

where = "Name='" & gCleanQueryArg(pName) & "'"
Set LoadByNameAsync = LoadDetailsAsync(where, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByID(ByVal Id As Long) As ADODB.Recordset
Dim where As String
Const ProcName As String = "LoadByID"
Dim failpoint As String
On Error GoTo Err

where = "ID=" & Id
Set LoadByID = LoadDetails(where)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByIDAsync( _
                ByVal Id As Long, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim where As String
Const ProcName As String = "LoadByIDAsync"
Dim failpoint As String
On Error GoTo Err

where = "ID=" & Id
Set LoadByIDAsync = LoadDetailsAsync(where, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByQuery(ByVal whereCondition As String) As ADODB.Recordset
Const ProcName As String = "LoadByQuery"
Dim failpoint As String
On Error GoTo Err

Set LoadByQuery = LoadDetails(whereCondition)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function LoadByQueryAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "LoadByQueryAsync"
Dim failpoint As String
On Error GoTo Err

Set LoadByQueryAsync = LoadDetailsAsync(whereCondition, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MakeNew() As ADODB.Recordset
Dim sql As String

Const ProcName As String = "MakeNew"
Dim failpoint As String
On Error GoTo Err

failpoint = 100

Debug.Print "AbstractDOFactory: Connect to DB"
Set mConnection = New ADODB.Connection
mConnection.Open mConnectionString

failpoint = 200

Set mRs = New Recordset
mRs.cursorlocation = adUseClient
sql = "SELECT * " & _
        "FROM " & mBaseTableName & _
        " WHERE 1=0"


mRs.Open sql, mConnection, adOpenStatic, adLockBatchOptimistic

Set MakeNew = mRs

Set mRs.ActiveConnection = Nothing
Debug.Print "AbstractDOFactory: close connection"
mConnection.Close

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MakeNewAsync( _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim t As New LoadTask
Const ProcName As String = "MakeNewAsync"
Dim failpoint As String
On Error GoTo Err

t.Initialise getFactory, _
            mConnectionString, _
            "SELECT * " & _
            "FROM " & mBaseTableName & _
            " WHERE 1=0"
Set MakeNewAsync = StartTask(t, PriorityNormal, , cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Query( _
                ByVal whereCondition As String, _
                ByRef FieldNames() As String) As DataObjectSummaries
Const ProcName As String = "Query"
Dim failpoint As String
On Error GoTo Err

Set Query = SearchWhere(whereCondition, FieldNames)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function QueryAsync( _
                ByVal whereCondition As String, _
                ByRef FieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "QueryAsync"
Dim failpoint As String
On Error GoTo Err

Set QueryAsync = SearchWhereAsync(whereCondition, FieldNames, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function QueryObjects( _
                ByVal whereCondition As String) As DataObjects
Const ProcName As String = "QueryObjects"
Dim failpoint As String
On Error GoTo Err

Set QueryObjects = SearchObjectsWhere(whereCondition)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function QueryObjectsAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "QueryObjectsAsync"
Dim failpoint As String
On Error GoTo Err

Set QueryObjectsAsync = SearchObjectsWhereAsync(whereCondition, cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Search( _
                ByVal criterion As String, _
                ByRef FieldNames() As String) As DataObjectSummaries
Const ProcName As String = "Search"
Dim failpoint As String
On Error GoTo Err

If criterion = "" Then
    Set Search = SearchWhere("", FieldNames)
Else
    Set Search = SearchWhere("Name LIKE '" & gCleanQueryArg(criterion) & "' ", _
                            FieldNames)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SearchAsync( _
                ByVal criterion As String, _
                ByRef FieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "SearchAsync"
Dim failpoint As String
On Error GoTo Err

If criterion = "" Then
    Set SearchAsync = SearchWhereAsync("", FieldNames, cookie, taskName)
Else
    Set SearchAsync = SearchWhereAsync("Name LIKE '" & gCleanQueryArg(criterion) & "' ", _
                            FieldNames, _
                            cookie, _
                            taskName)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SearchObjects( _
                ByVal criterion As String) As DataObjects
Const ProcName As String = "SearchObjects"
Dim failpoint As String
On Error GoTo Err

If criterion = "" Then
    Set SearchObjects = SearchObjectsWhere("")
Else
    Set SearchObjects = SearchObjectsWhere("Name LIKE '" & gCleanQueryArg(criterion) & "' ")
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SearchObjectsAsync( _
                ByVal criterion As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Const ProcName As String = "SearchObjectsAsync"
Dim failpoint As String
On Error GoTo Err

If criterion = "" Then
    Set SearchObjectsAsync = SearchObjectsWhereAsync("", cookie, taskName)
Else
    Set SearchObjectsAsync = SearchObjectsWhereAsync("Name LIKE '" & gCleanQueryArg(criterion) & "' ", cookie, taskName)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function generateColumnList( _
                ByRef FieldNames() As String) As String
Dim i As Long
Dim spec As FieldSpecifier

Const ProcName As String = "generateColumnList"
Dim failpoint As String
On Error GoTo Err

failpoint = 100

ReDim ar(0) As String
ar(0) = "Id"

ReDim Preserve ar(UBound(FieldNames) + 1) As String

'On Error Resume Next
For i = 0 To UBound(FieldNames)
    spec = mSpecifiers(FieldNames(i))
    ar(i + 1) = spec.dbColumnName
    AssertArgument ar(i + 1) <> "", "Invalid field Name: " & FieldNames(i)
Next
generateColumnList = Join(ar, ",")

Exit Function

Err:

If Err.Number <> ErrorCodes.ErrIllegalArgumentException Then
    generateColumnList = "*"
    Exit Function
End If
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFactory() As DataObjectFactory
Const ProcName As String = "getFactory"
Dim failpoint As String
On Error GoTo Err

Set getFactory = mFactoryRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function LoadDetails(ByVal where As String) As Recordset
Dim sql As String

Const ProcName As String = "LoadDetails"
Dim failpoint As String
On Error GoTo Err

failpoint = 100

Debug.Print "AbstractDOFactory: Connect to DB"
Set mConnection = New ADODB.Connection
mConnection.Open mConnectionString

failpoint = 200

Set mRs = New Recordset
mRs.cursorlocation = adUseClient
sql = "SELECT * " & _
        "FROM " & mDetailsViewName & _
        " WHERE " & where

mRs.Open sql, mConnection, adOpenStatic, adLockBatchOptimistic

failpoint = 300

If Not mRs.EOF Then
    Set LoadDetails = mRs
Else
    Set LoadDetails = Nothing
End If

Set mRs.ActiveConnection = Nothing
Debug.Print "AbstractDOFactory: close connection"
mConnection.Close
Set mRs = Nothing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function LoadDetailsAsync( _
                ByVal where As String, _
                ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
Dim t As New LoadTask
Const ProcName As String = "LoadDetailsAsync"
Dim failpoint As String
On Error GoTo Err

t.Initialise getFactory, _
            mConnectionString, _
            "SELECT * " & _
            "FROM " & mDetailsViewName & _
            " WHERE " & where
Set LoadDetailsAsync = StartTask(t, PriorityNormal, , cookie, taskName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function normaliseFieldnames( _
                ByRef FieldNames() As String) As String()
Dim max As Long: max = -1

Const ProcName As String = "normaliseFieldnames"
Dim failpoint As String
On Error GoTo Err

On Error Resume Next
max = UBound(FieldNames)
On Error GoTo Err

If max = -1 Then
    normaliseFieldnames = mSpecifiers.FieldNames
Else
    normaliseFieldnames = FieldNames
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function SearchWhere( _
                ByVal criterion As String, _
                ByRef FieldNames() As String) As DataObjectSummaries
Dim sql As String
Dim where As String
Dim fnames() As String

Const ProcName As String = "SearchWhere"
Dim failpoint As String
On Error GoTo Err

fnames = normaliseFieldnames(FieldNames)

where = IIf(criterion = "", "", "WHERE " & criterion)

failpoint = 100

Debug.Print "AbstractDOFactory: Connect to DB"
Set mConnection = New ADODB.Connection
mConnection.Open mConnectionString

failpoint = 200

Set mRs = New Recordset
mRs.cursorlocation = adUseClient
sql = "SELECT " & generateColumnList(fnames) & _
        " FROM " & mDetailsViewName & " " & _
        where & _
        " ORDER BY Name"
        
If gLogger.IsLoggable(LogLevelMediumDetail) Then gLogger.Log "Search " & "Query string=" & sql, ProcName, ModuleName, LogLevelDetail

failpoint = 300

If gLogger.IsLoggable(LogLevelMediumDetail) Then gLogger.Log "Search " & "open record set", ProcName, ModuleName, LogLevelDetail

mRs.Open sql, mConnection, adOpenStatic, adLockBatchOptimistic

Set mRs.ActiveConnection = Nothing
Debug.Print "AbstractDOFactory: close connection"
mConnection.Close

Set SearchWhere = gBuildSummariesFromRecordset(mRs, fnames, mSpecifiers)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function SearchWhereAsync( _
                ByVal criterion As String, _
                ByRef FieldNames() As String, _
                ByVal cookie As Variant, _
                ByVal taskName As String) As TaskController
Dim t As New QueryTask
Dim where As String
Dim fnames() As String

Const ProcName As String = "SearchWhereAsync"
Dim failpoint As String
On Error GoTo Err

fnames = normaliseFieldnames(FieldNames)

where = IIf(criterion = "", "", "WHERE " & criterion)

t.Initialise mConnectionString, _
            "SELECT " & generateColumnList(fnames) & _
            " FROM " & mDetailsViewName & " " & _
            where & _
            " ORDER BY Name", _
            fnames, _
            mSpecifiers
Set SearchWhereAsync = StartTask(t, PriorityNormal, taskName, cookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function SearchObjectsWhere( _
                ByVal criterion As String) As DataObjects
Dim sql As String
Dim where As String

Const ProcName As String = "SearchObjectsWhere"
Dim failpoint As String
On Error GoTo Err

where = IIf(criterion = "", "", "WHERE " & criterion)

failpoint = 100

Debug.Print "AbstractDOFactory: Connect to DB"
Set mConnection = New ADODB.Connection
mConnection.Open mConnectionString

failpoint = 200

Set mRs = New Recordset
mRs.cursorlocation = adUseClient
sql = "SELECT * FROM " & mDetailsViewName & " " & _
        where & _
        " ORDER BY Name"
        
If gLogger.IsLoggable(LogLevelMediumDetail) Then gLogger.Log "SearchObjectsWhere " & "Query string=" & sql, ProcName, ModuleName, LogLevelDetail

failpoint = 300

If gLogger.IsLoggable(LogLevelMediumDetail) Then gLogger.Log "SearchObjectsWhere " & "open record set", ProcName, ModuleName, LogLevelDetail

mRs.Open sql, mConnection, adOpenStatic, adLockBatchOptimistic

Set mRs.ActiveConnection = Nothing
Debug.Print "AbstractDOFactory: close connection"
mConnection.Close

Set SearchObjectsWhere = gBuildDataObjectsFromRecordset(getFactory, mRs)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function SearchObjectsWhereAsync( _
                ByVal criterion As String, _
                ByVal cookie As Variant, _
                ByVal taskName As String) As TaskController
Dim t As New QueryObjectsTask
Dim where As String

Const ProcName As String = "SearchObjectsWhereAsync"
Dim failpoint As String
On Error GoTo Err

where = IIf(criterion = "", "", "WHERE " & criterion)

t.Initialise getFactory, _
            mConnectionString, _
            "SELECT * FROM " & mDetailsViewName & " " & _
            where & _
            " ORDER BY Name"
Set SearchObjectsWhereAsync = StartTask(t, PriorityNormal, taskName, cookie)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

