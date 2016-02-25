Attribute VB_Name = "Globals"
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

Public Const ProjectName                            As String = "GraphObj40"
Private Const ModuleName                            As String = "Globals"

Public Const ConfigSettingBasedOn                   As String = "&BasedOn"
Public Const ConfigSettingName                      As String = "&Name"
Public Const ConfigSettingStyleType                 As String = "&StyleType"

Public Const Pi                                     As Double = 3.14159265358979
Public Const InverseSqrtOf2                         As Double = 0.707106781186547

'@================================================================================
' Member variables
'@================================================================================

Private mChangeListeners                            As New Collection

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

Public Function gGetConfigName(ByVal pExtProp As ExtendedProperty) As String
Dim lMetadata As GraphicExtPropMetadata
Set lMetadata = pExtProp.Metadata
gGetConfigName = lMetadata.ConfigName
End Function

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

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

Public Function gSetProperty( _
                ByVal pExtHost As ExtendedPropertyHost, _
                ByVal pExtProp As ExtendedProperty, _
                ByVal pNewValue As Variant, _
                Optional ByRef pPrevValue As Variant) As Boolean
Const ProcName As String = "setProperty"
On Error GoTo Err

If Not IsMissing(pPrevValue) Then gSetVariant pPrevValue, pExtHost.GetLocalValue(pExtProp)

gSetProperty = pExtHost.SetValue(pExtProp, pNewValue)

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Sub gValidateBrush(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidateBrush"
If Not TypeOf pValue Is IBrush Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must be of type IBrush"
End If
End Sub

Public Sub gValidateLayer(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidateLayer"
Dim lValue As Long
lValue = CLng(pValue)
If lValue < LayerNumbers.LayerMin Or lValue > LayerNumbers.LayerMax Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must be a member of the LayerNumbers enum"
End If
End Sub

Public Sub gValidatePen(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidatePen"
If Not TypeOf pValue Is Pen Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must be of type Pen"
End If
End Sub

Public Sub gValidatePolygonNumberOfSides(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidatePolygonNumberOfSides"
On Error GoTo Err

If CLng(pValue) <= 2 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid number of sides"


Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub gValidatePosition(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidatePosition"
If Not TypeOf pValue Is Point Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must be of type Point"
End If
End Sub

Public Sub gValidateSize(ByVal pThis As Object, ByVal pValue As Variant)
Const ProcName As String = "gValidateSize"
If Not TypeOf pValue Is Size Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, _
            ProjectName & "." & ModuleName & ":" & ProcName, _
            "Value must be of type Size"
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




