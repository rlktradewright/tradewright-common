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

Public Const ProjectName                    As String = "BusObjUtils40"
Private Const ModuleName                    As String = "Globals"

Public Const ConnectCompletionTimeoutMillisecs  As Long = 3000

Public Const InfoTypeBusObjUtils            As String = "tradewright.busobjutils"

Public Const GenericColumnId                As String = "ID"
Public Const GenericColumnName              As String = "NAME"

'@================================================================================
' Member variables
'@================================================================================

Private mSqlBadWords()                      As Variant

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

Public Property Get gLogger() As FormattingLogger
Static sLogger As FormattingLogger
If sLogger Is Nothing Then
    Set sLogger = CreateFormattingLogger(InfoTypeBusObjUtils, ProjectName)
End If

Set gLogger = sLogger
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gBuildDataObjectFromRecordset( _
                ByVal factory As DataObjectFactory, _
                ByVal rs As ADODB.Recordset) As BusinessDataObject
Const ProcName As String = "gBuildDataObjectFromRecordset"
Dim failpoint As String
On Error GoTo Err

Dim bookmark As Variant
bookmark = rs.bookmark

Dim cloneRS As ADODB.Recordset
Set cloneRS = rs.Clone
cloneRS.bookmark = bookmark

Set gBuildDataObjectFromRecordset = factory.MakeNewFromRecordset(cloneRS)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBuildDataObjectsFromRecordset( _
                ByVal factory As DataObjectFactory, _
                ByVal rs As ADODB.Recordset) As DataObjects
Const ProcName As String = "gBuildDataObjectsFromRecordset"
Dim failpoint As String
On Error GoTo Err

Dim dataObjs As New DataObjects

Do While Not rs.EOF
    dataObjs.Add gBuildDataObjectFromRecordset(factory, rs)
    rs.MoveNext
Loop

Set gBuildDataObjectsFromRecordset = dataObjs

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBuildSummaryFromRecordset( _
                ByVal rs As ADODB.Recordset, _
                ByRef FieldNames() As String, _
                ByVal specifiers As FieldSpecifiers) As DataObjectSummary
Const ProcName As String = "gBuildSummaryFromRecordset"
Dim failpoint As String
On Error GoTo Err

Dim lDOSummaryBuilder As New DOSummaryBuilder
lDOSummaryBuilder.Id = rs("ID")

Dim i As Long
For i = 0 To UBound(FieldNames)
    Dim colValue As String
    colValue = getFormattedColumnValue(rs, FieldNames(i), specifiers)
    lDOSummaryBuilder.FieldValue(FieldNames(i)) = colValue
Next
    
Set gBuildSummaryFromRecordset = lDOSummaryBuilder.DataObjectSummary

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gBuildSummariesFromRecordset( _
                ByVal rs As ADODB.Recordset, _
                ByRef FieldNames() As String, _
                ByVal specifiers As FieldSpecifiers) As DataObjectSummaries
Const ProcName As String = "gBuildSummariesFromRecordset"
Dim failpoint As String
On Error GoTo Err

Dim summsBuilder As DOSummariesBuilder
Set summsBuilder = New DOSummariesBuilder

Dim summs As DataObjectSummaries
Set summs = summsBuilder.DataObjectSummaries

Dim i As Long
For i = 0 To UBound(FieldNames)
    summsBuilder.AddFieldDetails specifiers(FieldNames(i))
Next

Do While Not rs.EOF
    summsBuilder.Add gBuildSummaryFromRecordset(rs, FieldNames, specifiers)
    
    rs.MoveNext
Loop

Set gBuildSummariesFromRecordset = summs

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCleanQueryArg( _
                ByRef inString) As String
Const ProcName As String = "gCleanQueryArg"
Dim failpoint As String
On Error GoTo Err

Static initialised As Boolean
If Not initialised Then
    mSqlBadWords = Array("'", "select", "drop", ";", "--", "insert", "Delete", "xp_")
    initialised = True
End If

gCleanQueryArg = inString

Dim i As Long
For i = 0 To UBound(mSqlBadWords)
    gCleanQueryArg = Replace(gCleanQueryArg, mSqlBadWords(i), "")
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGenerateConnectionErrorMessages( _
                ByVal pConnection As ADODB.Connection) As String
Const ProcName As String = "gGenerateConnectionErrorMessages"
Dim failpoint As String
On Error GoTo Err

Dim errMsg As String

Dim lError As ADODB.Error
For Each lError In pConnection.Errors
    errMsg = "--------------------" & vbCrLf & _
            gGenerateErrorMessage(lError)
Next
pConnection.Errors.clear
gGenerateConnectionErrorMessages = errMsg

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGenerateErrorMessage( _
                ByVal pError As ADODB.Error)
Const ProcName As String = "gGenerateErrorMessage"
Dim failpoint As String
On Error GoTo Err

gGenerateErrorMessage = _
        "Error " & pError.Number & ": " & pError.Description & vbCrLf & _
        "    Source: " & pError.Source & vbCrLf & _
        "    SQL state: " & pError.SQLState & vbCrLf & _
        "    Native error: " & pError.NativeError & vbCrLf

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetColumnValue( _
                ByVal rs As Recordset, _
                ByVal columnName As String, _
                ByVal defaultValue As Variant) As Variant
Const ProcName As String = "gGetColumnValue"
On Error GoTo Err

Dim fld As ADODB.Field
Set fld = rs.Fields(columnName)

Dim value As Variant
value = Nz(fld.value, defaultValue)
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
    gGetColumnValue = Trim(value)
Case adDBTime
    ' ensure any date part is removed
    If IsDate(value) Then
        gGetColumnValue = CDate(value - Int(value))
    End If
Case Else
    gGetColumnValue = value
End Select

Exit Function

Err:
If Err.Number = 3265 Then Err.Source = Err.Source & " (" & columnName & ")"
gHandleUnexpectedError ProcName, ModuleName
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

Public Function gIsStateSet( _
                ByVal value As Long, _
                ByVal stateToTest As ADODB.ObjectStateEnum) As Boolean
Const ProcName As String = "gIsStateSet"
Dim failpoint As String
On Error GoTo Err

gIsStateSet = ((value And stateToTest) = stateToTest)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getFormattedColumnValue( _
                ByVal rs As Recordset, _
                ByVal fieldName As String, _
                ByVal specifiers As FieldSpecifiers) As String
Const ProcName As String = "getFormattedColumnValue"
Dim failpoint As String
On Error GoTo Err

Dim spec As FieldSpecifier
spec = specifiers(fieldName)

Dim valtype As Long
valtype = rs(spec.dbColumnName).Type

Dim value As Variant
value = gGetColumnValue(rs, spec.dbColumnName, Empty)
If valtype = adDBTime Or _
    valtype = adDate Or _
    valtype = adDBDate Or _
    valtype = adDBTimeStamp Then
    If IsDate(value) Then
        If Int(value) = value Then
            getFormattedColumnValue = FormatTimestamp(CDate(value), TimestampDateOnlyISO8601)
        ElseIf Int(value) = 0 Then
            getFormattedColumnValue = FormatTimestamp(CDate(value), TimestampTimeOnlyISO8601)
        Else
            getFormattedColumnValue = FormatTimestamp(CDate(value), TimestampDateAndTimeISO8601)
        End If
    Else
        getFormattedColumnValue = Trim$(CStr(value))
    End If
Else
    getFormattedColumnValue = Trim$(CStr(value))
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


