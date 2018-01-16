VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigurationPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

Private Const ModuleName                            As String = "SectionPathParser"

'@================================================================================
' Member variables
'@================================================================================

Private mPath                                       As String

Private mParsed                                     As Boolean

Private mRelative                                   As Boolean

Private mParentPath                                 As String
Private mTailPath                                   As String
Private mTailName                                   As String
Private mInstanceQualifier                          As String

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

Friend Property Get InstanceQualifier() As String
If Not mParsed Then parse
InstanceQualifier = mInstanceQualifier
End Property

Friend Property Get Path() As String
Path = mPath
End Property

Friend Property Get ParentPath() As String
Const ProcName As String = "ParentPath"

On Error GoTo Err

If Not mParsed Then parse
ParentPath = mParentPath

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get TailName() As String
Const ProcName As String = "TailName"

On Error GoTo Err

If Not mParsed Then parse
TailName = mTailName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the final element of the configuration path.
'
' @remarks
'   For an attribute- or value path, the prefixing "&" or "." character is included.
' @return
'   The final element of the configuration path.
'@/
Friend Property Get TailPath() As String
Const ProcName As String = "TailPath"

On Error GoTo Err

If Not mParsed Then parse
TailPath = mTailPath

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal Path As String)
Const ProcName As String = "Initialise"

On Error GoTo Err

mPath = Path
If IsValidValuePath Then
ElseIf IsValidAttributePath Then
ElseIf IsValidSectionPath Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Not a valid configuration path"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function IsRelative() As Boolean
Const ProcName As String = "IsRelative"

On Error GoTo Err

If Not mParsed Then parse
IsRelative = mRelative

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function IsValidAttributePath() As Boolean
Static isValid As Variant
Const ProcName As String = "IsValidAttributePath"

On Error GoTo Err

If IsEmpty(isValid) Then
    gRegExp.Global = True
    gRegExp.Pattern = "^/?([a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s|-)*\))?(/[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s|-)*\))?)*)?\&[a-zA-Z0-9]\w*$"
    isValid = gRegExp.Test(mPath)
End If
IsValidAttributePath = isValid

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function IsValidValuePath() As Boolean
Static isValid As Variant
Const ProcName As String = "IsValidValuePath"

On Error GoTo Err

If IsEmpty(isValid) Then
    gRegExp.Global = True
    gRegExp.Pattern = "^/?([a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s|-)*\))?(/[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s|-)*\))?)*)?\.[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s|-)*\))?$"
    isValid = gRegExp.Test(mPath)
End If
IsValidValuePath = isValid

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function IsValidSectionPath() As Boolean
Static isValid As Variant
Const ProcName As String = "IsValidSectionPath"

On Error GoTo Err

If IsEmpty(isValid) Then
    gRegExp.Global = True
    gRegExp.Pattern = "^/?[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s|-)*\))?(/[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s|-)*\))?)*$"
    isValid = gRegExp.Test(mPath)
End If
IsValidSectionPath = isValid

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub parse()
Dim matches As MatchCollection
Dim lMatch As Match
Dim i As Long

Const ProcName As String = "parse"

On Error GoTo Err

If Not mParsed Then
    mRelative = (Left$(mPath, 1) <> ConfigSectionPathSeparator)
    gRegExp.Pattern = "[a-zA-Z0-9]\w*(\([a-zA-Z0-9](\w|\s|-)*\))?"
    gRegExp.Global = True
    Set matches = gRegExp.Execute(mPath)
    
    If Not mRelative And matches.Count > 1 Then mParentPath = ConfigSectionPathSeparator
    For i = 0 To matches.Count - 1
        Set lMatch = matches(i)
        If i = matches.Count - 1 Then
            mTailPath = lMatch.Value
        Else
            mParentPath = mParentPath & IIf(i <> 0, ConfigSectionPathSeparator, "") & lMatch.Value
        End If
    Next
    mParsed = True

    If IsValidSectionPath Then
        gRegExp.Pattern = "([^\(]+)?"
        Set matches = gRegExp.Execute(mTailPath)
        Set lMatch = matches(0)
        mTailName = lMatch.Value
        
        gRegExp.Pattern = "(\([a-zA-Z0-9](\w|\s|-)*\))"
        Set matches = gRegExp.Execute(mTailPath)
        If matches.Count = 1 Then
            Set lMatch = matches(0)
            mInstanceQualifier = Mid$(lMatch.Value, 2, Len(lMatch.Value) - 2)
        End If
    ElseIf IsValidAttributePath Then
        mTailName = mTailPath
        mTailPath = AttributePathNameSeparator & mTailPath
    Else
        mTailName = mTailPath
        mTailPath = ValuePathNameSeparator & mTailPath
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub


