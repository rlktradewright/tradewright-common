VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtProps"
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

Public Enum ExtendedPropertyHostChangeTypes
    StylesChanged = 1
    ParentChanged
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type ExtendedPropertyChangedEventData
    ExtendedProperty    As ExtendedProperty
    Source              As Object
    OldValue            As Variant
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ExtProps"

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

Public Function RegisterExtendedProperty( _
                ByVal pName As String, _
                ByVal pType As VbVarType, _
                ByVal pTypename As String, _
                Optional ByVal pDefaultValue As Variant, _
                Optional ByVal pMetadata As Variant, _
                Optional ByVal pValidatorPointer As Long) As ExtendedProperty
Const ProcName As String = "RegisterExtendedProperty"
On Error GoTo Err

If IsMissing(pDefaultValue) Then pDefaultValue = Empty
Set RegisterExtendedProperty = New ExtendedProperty
RegisterExtendedProperty.Initialise pName, pType, pTypename, pDefaultValue, pMetadata, pValidatorPointer

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




