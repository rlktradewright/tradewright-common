Attribute VB_Name = "GListeners"
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

Private Const ModuleName                            As String = "GListeners"

'@================================================================================
' Member variables
'@================================================================================

Private mChangeListeners                            As New Collection
Private mExtendedPropertyChangedListeners           As New Collection


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

Public Sub gAddChangeListener( _
            ByVal pSource As Object, _
            ByVal pListener As IChangeListener)
Const ProcName As String = "gAddChangeListener"
On Error GoTo Err

Dim lKey As String
lKey = GetObjectKey(pSource)

Dim lListeners As Listeners
Set lListeners = mChangeListeners(lKey)

If lListeners Is Nothing Then
    Set lListeners = New Listeners
    mChangeListeners.Add lListeners, lKey
End If

lListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gAddExtendedPropertyChangedListener( _
            ByVal pSource As Object, _
            ByVal pListener As IExtPropertyChangedListener)
Const ProcName As String = "gAddExtendedPropertyChangedListener"
On Error GoTo Err

Dim lKey As String
lKey = GetObjectKey(pSource)

Dim lListeners As Listeners
Set lListeners = mExtendedPropertyChangedListeners(lKey)
If lListeners Is Nothing Then
    Set lListeners = New Listeners
    mExtendedPropertyChangedListeners.Add lListeners, lKey
End If

lListeners.Add pListener

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gFireChange( _
            ByVal pSource As Object, _
            ByVal pValue As Long) As ChangeEventData
Const ProcName As String = "gFireChange"
On Error GoTo Err

Dim ev As ChangeEventData
Set ev.Source = pSource
ev.ChangeType = pValue
gFireChange = ev

Dim lListeners As Listeners
Set lListeners = mChangeListeners(GetObjectKey(pSource))
If lListeners Is Nothing Then Exit Function

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not lListeners.Valid Then
    sInit = True
    sSomeListeners = lListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IChangeListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.Change ev
    Next
End If

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gFireExtendedPropertyChanged( _
                ByVal pSource As Object, _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pOldValue As Variant) As ExtendedPropertyChangedEventData
Const ProcName As String = "gFireExtendedPropertyChanged"
On Error GoTo Err

Dim ev As ExtendedPropertyChangedEventData
Set ev.Source = pSource
Set ev.ExtendedProperty = pExtProp
gSetVariant ev.OldValue, pOldValue
gFireExtendedPropertyChanged = ev

Dim lListeners As Listeners
Set lListeners = mExtendedPropertyChangedListeners(GetObjectKey(pSource))
If lListeners Is Nothing Then Exit Function

Static sInit As Boolean
Static sCurrentListeners() As Object
Static sSomeListeners As Boolean

If Not sInit Or Not lListeners.Valid Then
    sInit = True
    sSomeListeners = lListeners.GetCurrentListeners(sCurrentListeners)
End If
If sSomeListeners Then
    Dim lListener As IExtPropertyChangedListener
    Dim i As Long
    For i = 0 To UBound(sCurrentListeners)
        Set lListener = sCurrentListeners(i)
        lListener.ExtendedPropertyChanged ev
    Next
End If

Exit Function

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gRemoveChangeListener( _
            ByVal pSource As Object, _
            ByVal pListener As IChangeListener)
Const ProcName As String = "gRemoveChangeListener"
On Error GoTo Err

Dim lKey As String
lKey = GetObjectKey(pSource)

Dim lListeners As Listeners
Set lListeners = mChangeListeners(lKey)
If Not lListeners Is Nothing Then
    lListeners.Remove pListener
    If lListeners.Count = 0 Then mChangeListeners.Remove lKey
End If

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gRemoveExtendedPropertyChangedListener( _
            ByVal pSource As Object, _
            ByVal pListener As IExtPropertyChangedListener)
Const ProcName As String = "gRemoveExtendedPropertyChangedListener"
On Error GoTo Err

Dim lKey As String
lKey = GetObjectKey(pSource)

Dim lListeners As Listeners
Set lListeners = mExtendedPropertyChangedListeners(lKey)
If Not lListeners Is Nothing Then
    lListeners.Remove pListener
    If lListeners.Count = 0 Then mExtendedPropertyChangedListeners.Remove lKey
End If

Exit Sub

Err:
If Err.Number = VBErrorCodes.VbErrInvalidProcedureCall Then Resume Next
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




