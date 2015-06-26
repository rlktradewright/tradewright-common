VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtendedProperty"
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

Private Const ModuleName                            As String = "ExtendedProperty"

'@================================================================================
' Member variables
'@================================================================================

Private mPropertyName                               As String
Private mPropertyType                               As VbVarType
Private mTypename                                   As String
Private mDefaultValue                               As Variant
Private mMetadata                                   As Variant
Private mValidatorPointer                           As Long

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

Public Property Get DefaultValue() As Variant
Const ProcName As String = "DefaultValue"
On Error GoTo Err

If ValueIsObject Then
    If Not IsEmpty(mDefaultValue) Then
        Set DefaultValue = mDefaultValue
    Else
        Set DefaultValue = Nothing
    End If
Else
    If Not IsEmpty(mDefaultValue) Then DefaultValue = mDefaultValue
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Metadata() As Variant
Const ProcName As String = "Metadata"
On Error GoTo Err

If Not IsEmpty(mMetadata) Then
    gSetVariant Metadata, mMetadata
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get PropertyName() As String
PropertyName = mPropertyName
End Property

Public Property Get ValueIsObject() As Boolean
ValueIsObject = (((mPropertyType And vbArray) = 0) And (mPropertyType = vbObject))
End Property

Public Property Get PropertyType() As VbVarType
PropertyType = mPropertyType
End Property

Public Property Get PropertyTypeName() As String
PropertyTypeName = mTypename
End Property

Public Property Get Validator() As PropertyValidator
Const ProcName As String = "Validator"
On Error GoTo Err

Set Validator = gGetValidator(mValidatorPointer)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal pName As String, _
                ByVal pType As VbVarType, _
                ByVal pTypename As String, _
                ByVal pDefaultValue As Variant, _
                ByVal pMetadata As Variant, _
                ByVal pValidatorPointer As Long)
Const ProcName As String = "Initialise"
On Error GoTo Err

mPropertyName = pName
mPropertyType = pType
mTypename = pTypename

If Not IsEmpty(pDefaultValue) Then
    If ValueIsObject Then
        Set mDefaultValue = pDefaultValue
    Else
        mDefaultValue = pDefaultValue
    End If
End If

If Not IsEmpty(pMetadata) Then
    gSetVariant mMetadata, pMetadata
End If

mValidatorPointer = pValidatorPointer

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




