VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ParameterStringParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class parses a parameter string into its Name/Value pairs.
'
' Note the <code>New</code> operator cannot be used to create an instance of this
' class. Instead, use the <code>CreateParameterStringParser</code> global method.
'
' <h3>Parameter String Format</h3>
'   The parameter names and values are supplied in a <code>String</code>
'   with the following syntax:
'   <pre>
'   <tt>
'   <p>paramString: param [sep param]...</p>
'
'   <p>param: Name delim value</p>
'
'   <p>name: string</p>
'
'   <p>value: string</p>
'   </tt>
'   </pre>
'   where
'   <pre>
'   <p><tt>sep</tt> is a single Character used as a separator between parameters.
'   The default separator is a semi-colon, but an alternative can be specified
'   in the <code>parameterSeparator</code> argument to <code>CreateParameterStringParser</code>.</p>
'
'   <p><tt>delim</tt> is a single Character used to delimit the Name
'   The default delimiter is =, but an alternative can be specified
'   in the <code>nameDelimiter</code> argument to <code>CreateParameterStringParser</code>.</p>
'   </pre>
'
'   <p>If the parameter Name or Value needs to contain a separator or delimiter
'   Character, this can be achieved by preceding it with an escape Character.
'   The default escape Character is \, but an alternative can be specified
'   in the <code>escapeCharacter</code> argument to <code>CreateParameterStringParser</code>.
'   Note that to include the escape Character itself in the Value or Name,
'   it must be repeated. The effect of the escape Character is to include the
'   following Character in the parameter Name or Value verbatim.</p>
'
'   <p>For each specified parameter, if a parameter with this Name is
'   already included, its Value is updated; otherwise a new parameter
'   is added to the list.</p>
' <h3>Examples</h3>
'   An example parameter string, using the default separator and delimiter
'   characters:
'   <pre>
'   <tt>Server=Delphi;Port=7496;ClientId=432</tt>
'   </pre>
'   An example incuding the escape Character in the Value:
'   <pre>
'   <tt>SettingsFile=C:\\MyProg\\myprog.ini</tt>
'   </pre>
'   An example using the escape Character in the Value in a unusual (and pointless!) way:
'   <pre>
'   <tt>My parameter=\a\b\c\d\e\f</tt>
'   </pre>
'   Note that in this case, the Value would be interpreted as abcdef.
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


Private Const ModuleName As String = "ParameterStringParser"

'@================================================================================
' Member variables
'@================================================================================

Private mArIn() As Byte

Private mNameDelimiter As Byte
Private mparameterSeparator As Byte
Private mEscapeCharacter As Byte


Private mCurrPosn As Long
Private mArOut() As Byte
Private mArIndex As Long

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

Friend Sub Initialise( _
                ByVal Value As String, _
                ByVal nameDelimiter As String, _
                ByVal parameterSeparator As String, _
                ByVal escapeCharacter As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

If Len(nameDelimiter) > 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "nameDelimiter must be a single Character"
If Len(parameterSeparator) > 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "parameterSeparator must be a single Character"
If Len(escapeCharacter) > 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "escapeCharacter must be a single Character"

mArIn = Trim$(Value)
mNameDelimiter = Asc(nameDelimiter)
mparameterSeparator = Asc(parameterSeparator)
mEscapeCharacter = Asc(escapeCharacter)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Gets the Name and Value for the Next parameter in the parameter string.
'
' @remarks
' @return
'   <code>True</code> if the parameter Name and Value were available. <code>False</code>
'   if there are no more parameters in the parameter string.
' @param Name
'   On return to the caller, this argument contains the parameter Name. The contents
'   prior to the call are overwritten. If the method returns <code>False</code>, the
'   contents on return to the caller are undefined.
' @param Value
'   On return to the caller, this argument contains the parameter Value. The contents
'   prior to the call are overwritten. If the method returns <code>False</code>, the
'   contents on return to the caller are undefined.
' @see
'
'@/
Public Function GetParameter( _
                ByRef Name As String, _
                ByRef Value As String) As Boolean

Const ProcName As String = "GetParameter"

On Error GoTo Err

If mCurrPosn >= UBound(mArIn) Then Exit Function

ReDim mArOut((UBound(mArIn) + 1 - mCurrPosn) / 2) As Byte
mArIndex = 0

Name = getToken(mNameDelimiter, True)

ReDim mArOut((UBound(mArIn) + 1 - mCurrPosn) / 2) As Byte
mArIndex = 0

Value = getToken(mparameterSeparator, False)

GetParameter = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getToken( _
                ByVal delimiter As Byte, _
                ByVal delimiterIsMandatory As Boolean) As String
Const ProcName As String = "getToken"

On Error GoTo Err

Do While mCurrPosn <= UBound(mArIn)
    If mArIn(mCurrPosn) = mEscapeCharacter Then
        If mCurrPosn < UBound(mArIn) - 1 Then
            mArOut(mArIndex) = mArIn(mCurrPosn + 2)
            mArIndex = mArIndex + 1
            mCurrPosn = mCurrPosn + 4
        Else
            mCurrPosn = mCurrPosn + 2
        End If
    ElseIf mArIn(mCurrPosn) = delimiter Then
        mCurrPosn = mCurrPosn + 2
        getToken = convertArrayToString
        Exit Do
    Else
        mArOut(mArIndex) = mArIn(mCurrPosn)
        mArIndex = mArIndex + 1
        mCurrPosn = mCurrPosn + 2
    End If
Loop

If getToken = "" Then
    If delimiterIsMandatory Then
        Err.Raise ErrorCodes.ErrIllegalStateException, , "Syntax Error: no Name delimiter"
    Else
        getToken = convertArrayToString
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function convertArrayToString() As String
Const ProcName As String = "convertArrayToString"

On Error GoTo Err

If mArIndex > 0 Then
    ReDim Preserve mArOut(mArIndex - 1) As Byte
    convertArrayToString = Trim$(StrConv(mArOut, vbUnicode))
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


