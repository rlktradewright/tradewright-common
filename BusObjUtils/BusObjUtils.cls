VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BusObjUtils"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
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

Public Enum ConditionalOperators
    CondOpEqual
    CondOpGreaterThan
    CondOpGreaterThanOrEqual
    CondOpLessThan
    CondOpLessThanOrEqual
    CondOpLike
End Enum

Public Enum FieldAlignments
    FieldAlignNone
    FieldAlignLeft
    FieldALignCentre
    FieldAlignRight
End Enum

Public Enum LogicalOperators
    LogicalOpNone
    LogicalOpAND
    LogicalOpOR
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type FieldSpecifier
    Name                As String
    align               As FieldAlignments
    width               As Long
    visible             As Boolean
    dbColumnName        As String
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "BusObjUtils"

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

Public Function CreateAbstractDataObject( _
                ByVal dataObj As BusinessDataObject, _
                ByVal busRuleNameValidId As Long, _
                ByVal factory As DataObjectFactory, _
                ByVal connectionString As String) As AbstractDataObject
Const ProcName As String = "CreateAbstractDataObject"
Dim failpoint As String
On Error GoTo Err

Set CreateAbstractDataObject = New AbstractDataObject
CreateAbstractDataObject.Initialise dataObj, busRuleNameValidId, factory, connectionString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function CreateAbstractDOFactory( _
                ByVal factory As DataObjectFactory, _
                ByVal connectionString As String, _
                ByVal baseTableName As String, _
                ByVal detailsViewName As String) As AbstractDOFactory
Const ProcName As String = "CreateAbstractDOFactory"
Dim failpoint As String
On Error GoTo Err

Set CreateAbstractDOFactory = New AbstractDOFactory
CreateAbstractDOFactory.Initialise factory, connectionString, baseTableName, detailsViewName

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


