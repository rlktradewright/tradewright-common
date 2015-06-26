VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ExtendedPropertyHost"
Attribute VB_GlobalNameSpace = False
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

Implements IChangeListener
Implements IExtPropertyChangedListener

'@================================================================================
' Events
'@================================================================================

Event Change(ByRef pEv As ChangeEventData)
Event ExtendedPropertyChanged(ByRef pEv As ExtendedPropertyChangedEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ExtendedPropertyHost"

'@================================================================================
' Member variables
'@================================================================================

Private mPropertyValues                             As Collection
Private mPropertyOverrideValues                     As Collection
Private mStyle                                      As ExtendedPropertyHost

Private mParent                                     As ExtendedPropertyHost
Attribute mParent.VB_VarHelpID = -1

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mPropertyValues = New Collection
Set mPropertyOverrideValues = New Collection
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
fireChange ev.ChangeType
End Sub

'@================================================================================
' IExtPropertyChangedListener Interface Members
'@================================================================================

Private Sub IExtPropertyChangedListener_ExtendedPropertyChanged(ByRef pEv As ExtendedPropertyChangedEventData)
fireExtendedPropertyChanged pEv.ExtendedProperty, pEv.OldValue
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Parent() As ExtendedPropertyHost
Set Parent = mParent
End Property

Public Property Let Style(ByVal Value As ExtendedPropertyHost)
Const ProcName As String = "Style"
On Error GoTo Err

If Not mStyle Is Nothing Then mStyle.RemoveExtendedPropertyChangedListener Me

Set mStyle = Value
If Not mStyle Is Nothing Then
    mStyle.AddExtendedPropertyChangedListener Me
    mStyle.AddChangeListener Me
End If
fireChange StylesChanged

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Style() As ExtendedPropertyHost
Set Style = mStyle
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub AddChangeListener( _
            ByVal pListener As IChangeListener)
gAddChangeListener Me, pListener
End Sub

Public Sub AddExtendedPropertyChangedListener( _
            ByVal pListener As IExtPropertyChangedListener)
gAddExtendedPropertyChangedListener Me, pListener
End Sub

Public Sub ClearOverrideValue( _
                ByVal pExtProp As ExtendedProperty)
Const ProcName As String = "ClearOverrideValue"
On Error GoTo Err

Dim lOldValue As Variant
gSetVariant lOldValue, GetValue(pExtProp)
If clearTheOverrideValue(pExtProp) Then fireExtendedPropertyChanged pExtProp, lOldValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub ClearValue( _
                ByVal pExtProp As ExtendedProperty)
Const ProcName As String = "ClearValue"
On Error GoTo Err

Dim lOldValue As Variant
gSetVariant lOldValue, GetValue(pExtProp)
If clearTheValue(pExtProp) Then fireExtendedPropertyChanged pExtProp, lOldValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetDefaultValue( _
                ByVal pExtProp As ExtendedProperty) As Variant
Const ProcName As String = "GetDefaultValue"
On Error GoTo Err

If Not IsEmpty(pExtProp.DefaultValue) Then
    If pExtProp.ValueIsObject Then
        Set GetDefaultValue = pExtProp.DefaultValue
    Else
        GetDefaultValue = pExtProp.DefaultValue
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetLocalValue( _
                ByVal pExtProp As ExtendedProperty) As Variant
Const ProcName As String = "GetLocalValue"
On Error GoTo Err

If pExtProp.ValueIsObject Then
    Set GetLocalValue = Nothing
    Set GetLocalValue = mPropertyValues(pExtProp.PropertyName)
Else
    GetLocalValue = mPropertyValues(pExtProp.PropertyName)
End If

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetNonDefaultValue( _
                ByVal pExtProp As ExtendedProperty) As Variant
Const ProcName As String = "GetNonDefaultValue"
On Error GoTo Err

If pExtProp.ValueIsObject Then
    Set GetNonDefaultValue = Nothing
    
    Set GetNonDefaultValue = mPropertyOverrideValues(pExtProp.PropertyName)
    
    If GetNonDefaultValue Is Nothing Then Set GetNonDefaultValue = mPropertyValues(pExtProp.PropertyName)
    
    If GetNonDefaultValue Is Nothing Then If Not mStyle Is Nothing Then Set GetNonDefaultValue = mStyle.GetNonDefaultValue(pExtProp)
    If GetNonDefaultValue Is Nothing Then If Not mParent Is Nothing Then Set GetNonDefaultValue = mParent.GetNonDefaultValue(pExtProp)
Else
    GetNonDefaultValue = mPropertyOverrideValues(pExtProp.PropertyName)
    
    If IsEmpty(GetNonDefaultValue) Then GetNonDefaultValue = mPropertyValues(pExtProp.PropertyName)
    
    If IsEmpty(GetNonDefaultValue) Then If Not mStyle Is Nothing Then GetNonDefaultValue = mStyle.GetNonDefaultValue(pExtProp)
    If IsEmpty(GetNonDefaultValue) Then If Not mParent Is Nothing Then GetNonDefaultValue = mParent.GetNonDefaultValue(pExtProp)
End If

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetOverrideValue( _
                ByVal pExtProp As ExtendedProperty) As Variant
Const ProcName As String = "GetOverrideValue"
On Error GoTo Err

If pExtProp.ValueIsObject Then
    Set GetOverrideValue = Nothing
    Set GetOverrideValue = mPropertyOverrideValues(pExtProp.PropertyName)
Else
    GetOverrideValue = mPropertyOverrideValues(pExtProp.PropertyName)
End If

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetValue( _
                ByVal pExtProp As ExtendedProperty) As Variant
Const ProcName As String = "GetValue"
On Error GoTo Err

If pExtProp.ValueIsObject Then
    Set GetValue = Me.GetNonDefaultValue(pExtProp)
    If GetValue Is Nothing Then Set GetValue = GetDefaultValue(pExtProp)
    If GetValue Is Nothing Then If Not mStyle Is Nothing Then Set GetValue = mStyle.GetDefaultValue(pExtProp)
    If GetValue Is Nothing Then If Not mParent Is Nothing Then Set GetValue = mParent.GetDefaultValue(pExtProp)
Else
    GetValue = Me.GetNonDefaultValue(pExtProp)
    If IsEmpty(GetValue) Then GetValue = GetDefaultValue(pExtProp)
    If IsEmpty(GetValue) Then If Not mStyle Is Nothing Then GetValue = mStyle.GetDefaultValue(pExtProp)
    If IsEmpty(GetValue) Then If Not mParent Is Nothing Then GetValue = mParent.GetDefaultValue(pExtProp)
End If

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsPropertySet( _
                ByVal pExtProp As ExtendedProperty) As Boolean
Const ProcName As String = "IsPropertySet"
On Error GoTo Err

Dim val As Variant

If pExtProp.ValueIsObject Then
    Set val = GetLocalValue(pExtProp)
    If Not val Is Nothing Then IsPropertySet = True
Else
    val = GetLocalValue(pExtProp)
    If Not IsEmpty(val) Then IsPropertySet = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValueValid( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewValue As Variant) As Boolean
Const ProcName As String = "IsValueValid"
On Error GoTo Err

On Error Resume Next
checkValueIsValid pExtProp, pNewValue
If Err.Number = 0 Then IsValueValid = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub RemoveChangeListener( _
            ByVal pListener As IChangeListener)
gRemoveChangeListener Me, pListener
End Sub

Public Sub RemoveExtendedPropertyChangedListener( _
            ByVal pListener As IExtPropertyChangedListener)
gRemoveExtendedPropertyChangedListener Me, pListener
End Sub

Public Sub SetParent( _
                ByVal pExtProp As ExtendedPropertyHost, _
                ByVal pNotifyChanges As Boolean)
Const ProcName As String = "SetParent"
On Error GoTo Err

If Not mParent Is Nothing Then mParent.RemoveExtendedPropertyChangedListener Me
Set mParent = pExtProp
If pNotifyChanges Then
    If Not mParent Is Nothing Then
        mParent.AddExtendedPropertyChangedListener Me
        mParent.AddChangeListener Me
    End If
    fireChange ParentChanged
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function SetOverrideValue( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewValue As Variant) As Boolean
Const ProcName As String = "SetOverrideValue"
On Error GoTo Err

Dim lOldValue As Variant

checkValueIsValid pExtProp, pNewValue

If newValueIsDifferent(pExtProp, GetOverrideValue(pExtProp), pNewValue) Then
    gSetVariant lOldValue, GetValue(pExtProp)
    clearTheOverrideValue pExtProp
    mPropertyOverrideValues.Add pNewValue, pExtProp.PropertyName
    fireExtendedPropertyChanged pExtProp, lOldValue
    SetOverrideValue = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function SetValue( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewValue As Variant) As Boolean
Const ProcName As String = "SetValue"
On Error GoTo Err

Dim lOldValue As Variant

checkValueIsValid pExtProp, pNewValue

If newValueIsDifferent(pExtProp, GetLocalValue(pExtProp), pNewValue) Then
    gSetVariant lOldValue, GetValue(pExtProp)
    clearTheValue pExtProp
    mPropertyValues.Add pNewValue, pExtProp.PropertyName
    fireExtendedPropertyChanged pExtProp, lOldValue
    SetValue = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function canCoerceTypes(ByVal pToType As VbVarType, _
                ByVal pFromType As VbVarType) As Boolean
Select Case pToType
Case vbInteger, vbString
    Select Case pFromType
    Case vbInteger
        canCoerceTypes = True
    End Select
Case vbLong
    Select Case pFromType
    Case vbInteger, vbLong, vbString
        canCoerceTypes = True
    End Select
Case vbSingle
    Select Case pFromType
    Case vbInteger, vbLong, vbSingle, vbCurrency, vbString
        canCoerceTypes = True
    End Select
Case vbDouble
    Select Case pFromType
    Case vbInteger, vbLong, vbSingle, vbCurrency, vbDouble, vbDate, vbString
        canCoerceTypes = True
    End Select
Case vbCurrency
    Select Case pFromType
    Case vbInteger, vbLong, vbCurrency, vbString
        canCoerceTypes = True
    End Select
Case vbDate
    Select Case pFromType
    Case vbInteger, vbLong, vbSingle, vbDouble, vbDate, vbString
        canCoerceTypes = True
    End Select
Case vbString
    Select Case pFromType
    Case vbString, vbBoolean, vbByte, vbCurrency, vbDate, vbDecimal, vbDouble, vbInteger, vbLong, vbSingle
        canCoerceTypes = True
    End Select
Case vbObject
    Select Case pFromType
    Case vbObject
        canCoerceTypes = True
    End Select
Case vbBoolean
    Select Case pFromType
    Case vbBoolean, vbString
        canCoerceTypes = True
    End Select
Case vbByte
    Select Case vbByte
    Case vbObject, vbString
        canCoerceTypes = True
    End Select
Case vbUserDefinedType
    Select Case vbUserDefinedType
    Case vbObject
        canCoerceTypes = True
    End Select
Case Else

End Select
End Function

Private Sub checkValueIsValid( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewValue As Variant)
Const ProcName As String = "checkValueIsValid"
On Error GoTo Err

If IsEmpty(pNewValue) Or IsNull(pNewValue) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "A non-null value must be supplied"

If pExtProp.ValueIsObject Then
    If pNewValue Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "A non-null value must be supplied"
End If

If VarType(pNewValue) = pExtProp.PropertyType Then
ElseIf Not canCoerceTypes(pExtProp.PropertyType, VarType(pNewValue)) Then
    Err.Raise ErrorCodes.ErrClassCastException, , "Invalid type"
ElseIf pExtProp.PropertyType <> vbUserDefinedType Then
ElseIf TypeName(pNewValue) <> pExtProp.PropertyTypeName Then
    Err.Raise ErrorCodes.ErrClassCastException, , "Invalid type"
End If

Dim lValidator As PropertyValidator
Set lValidator = pExtProp.Validator
If Not lValidator Is Nothing Then lValidator.ValidateProperty pNewValue

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function clearTheOverrideValue( _
                ByVal pExtProp As ExtendedProperty) As Boolean
Const ProcName As String = "clearTheOverrideValue"
On Error GoTo Err

On Error Resume Next
mPropertyOverrideValues.Remove pExtProp.PropertyName
If Err.Number = 0 Then clearTheOverrideValue = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function clearTheValue( _
                ByVal pExtProp As ExtendedProperty) As Boolean
Const ProcName As String = "clearTheValue"
On Error GoTo Err

On Error Resume Next
mPropertyValues.Remove pExtProp.PropertyName
If Err.Number = 0 Then clearTheValue = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub fireChange( _
                ByVal pValue As ExtendedPropertyHostChangeTypes)
RaiseEvent Change(gFireChange(Me, pValue))
End Sub

Private Sub fireExtendedPropertyChanged( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pOldValue As Variant)
RaiseEvent ExtendedPropertyChanged(gFireExtendedPropertyChanged(Me, pExtProp, pOldValue))
End Sub

Private Function newValueIsDifferent( _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pCurrvalue As Variant, _
                ByVal pValue As Variant) As Boolean
Const ProcName As String = "newValueIsDifferent"
On Error GoTo Err

If pExtProp.ValueIsObject Then
    If pCurrvalue Is Nothing Then
        If IsEmpty(pValue) Or pValue Is Nothing Then
            newValueIsDifferent = False
        Else
            newValueIsDifferent = True
        End If
    ElseIf IsEmpty(pValue) Then
        newValueIsDifferent = True
    ElseIf pCurrvalue Is pValue Then
        newValueIsDifferent = False
    Else
        newValueIsDifferent = True
    End If
Else
    If IsEmpty(pCurrvalue) Then
        If IsEmpty(pValue) Then
            newValueIsDifferent = False
        Else
            newValueIsDifferent = True
        End If
    ElseIf IsEmpty(pValue) Then
        newValueIsDifferent = True
    ElseIf pCurrvalue = pValue Then
        newValueIsDifferent = False
    Else
        newValueIsDifferent = True
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
