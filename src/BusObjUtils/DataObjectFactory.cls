VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DataObjectFactory"
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

Private Const ModuleName                        As String = "DataObjectFactory"

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

Public Property Get FieldSpecifiers() As FieldSpecifiers
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Delete(ByVal Id As Long)
End Sub

Public Function LoadByName(ByVal pName As String) As BusinessDataObject
End Function

Public Function LoadByNameAsync(ByVal pName As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
End Function

Public Function LoadByID(ByVal Id As Long) As BusinessDataObject
End Function

Public Function LoadByIDAsync(ByVal Id As Long, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
End Function

Public Function LoadByQuery(ByVal whereCondition As String) As BusinessDataObject
End Function

Public Function LoadByQueryAsync(ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
End Function

''
' Creates a new <code>BusinessDataObject</code> from the contents of a recordset.
'
' @remarks
'   This method is not intended for use by applications.
' @param rs
'   A recordset containing the appropriate data to Initialise the new
'   <code>BusinessDataObject</code> .
'@/
Public Function MakeNewFromRecordset( _
                ByVal rs As ADODB.Recordset) As BusinessDataObject
End Function

Public Function MakeNew() As BusinessDataObject
End Function

Public Function MakeNewAsync( _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
End Function

Public Function Query( _
                ByVal whereCondition As String, _
                ByRef FieldNames() As String) As DataObjectSummaries
End Function

Public Function QueryAsync( _
                ByVal whereCondition As String, _
                ByRef FieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
End Function

Public Function QueryObjects( _
                ByVal whereCondition As String) As DataObjects
End Function

Public Function QueryObjectsAsync( _
                ByVal whereCondition As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
End Function

''
' Ensures that the specified data object has the latest information from permanent storage.
'
' @remarks
'   If the data object has not yet been committed to permanent storage,
'   or if the data object has been deleted from permanent storage, an
'   <code>ErrorCodes.ErrIllegalStateException</code> error is raised.
' @param dataObj
'   The <code>BusinessDataObject</code> object to be refreshed.
'@/
Public Sub Refresh( _
                ByVal dataObj As BusinessDataObject)
End Sub

Public Function Search( _
                ByVal criterion As String, _
                ByRef FieldNames() As String) As DataObjectSummaries
End Function

Public Function SearchObjects( _
                ByVal criterion As String) As DataObjects
End Function

Public Function SearchAsync( _
                ByVal criterion As String, _
                ByRef FieldNames() As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
End Function

Public Function SearchObjectsAsync( _
                ByVal criterion As String, _
                Optional ByVal cookie As Variant, _
                Optional ByVal taskName As String) As TaskController
End Function

'@================================================================================
' Helper Functions
'@================================================================================

