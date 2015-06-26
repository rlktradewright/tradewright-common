VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FieldSpecifiers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                    As String = "FieldSpecifiers"

'@================================================================================
' Member variables
'@================================================================================

Private mCol                                As Collection

Private mFieldNames() As String
Private mFieldNamesIndex As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mCol = New Collection
ReDim mFieldNames(3) As String
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

Public Property Get FieldNames() As String()
Const ProcName As String = "FieldNames"
Dim failpoint As String
On Error GoTo Err

If mFieldNamesIndex = 0 Then Exit Property

If UBound(mFieldNames) <> mFieldNamesIndex - 1 Then
    ReDim Preserve mFieldNames(mFieldNamesIndex - 1) As String
End If
FieldNames = mFieldNames

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function AddField( _
                ByVal Name As String, _
                ByVal dbColumnName As String, _
                ByVal width As Long, _
                ByVal align As FieldAlignments, _
                ByVal visible As Boolean) As FieldSpecifier
Dim spec As FieldSpecifier

Const ProcName As String = "AddField"
Dim failpoint As String
On Error GoTo Err

spec.Name = Name
spec.align = align
spec.width = width
spec.visible = visible
spec.dbColumnName = dbColumnName

AddField = AddFieldSpecifier(spec)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function AddFieldSpecifier( _
                ByRef specifier As FieldSpecifier) As FieldSpecifier
Const ProcName As String = "AddFieldSpecifier"
Dim failpoint As String
On Error GoTo Err

If mFieldNamesIndex > UBound(mFieldNames) Then
    ReDim Preserve mFieldNames(2 * (UBound(mFieldNames) + 1) - 1) As String
End If

mFieldNames(mFieldNamesIndex) = specifier.Name
mFieldNamesIndex = mFieldNamesIndex + 1

mCol.Add specifier, specifier.Name
AddFieldSpecifier = specifier

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Count() As Long
Const ProcName As String = "Count"
Dim failpoint As String
On Error GoTo Err

Count = mCol.Count

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function Item( _
                ByVal index As Variant) As FieldSpecifier
Attribute Item.VB_UserMemId = 0
Const ProcName As String = "Item"
Dim failpoint As String
On Error GoTo Err

Item = mCol.Item(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Const ProcName As String = "NewEnum"
Dim failpoint As String
On Error GoTo Err

Set NewEnum = mCol.[_NewEnum]

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




'@================================================================================
' Helper Functions
'@================================================================================


