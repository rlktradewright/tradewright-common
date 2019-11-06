Attribute VB_Name = "GJSON"
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

Public Enum JSONBuilderStates
    Created = 1
    BuildingObject
    BuildingFirstNameValuePair
    BuildingNameValuePair
    ExpectingName
    BuildingArray
    ExpectingValue
    Finished
End Enum

Public Enum JSONBuilderActions
    ActionEncodeBeginArray = 1
    ActionEncodeEndArray = 2
    ActionEncodeBeginObject = 4
    ActionEncodeEndObject = 8
    ActionEncodeNameSeparator = &H10&
    ActionEncodeValueSeparator = &H20&
    ActionEncodeName = &H40&
    ActionEncodeValue = &H80&
    ActionIncreaseLevel = &H100&
    ActionDecreaseLevel = &H200&
End Enum

Public Enum JSONBuilderStimuli
    StimBeginArray = 1
    StimEndArray
    StimBeginObject
    StimEndObject
    StimName
    StimValue
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GJSON"

Public Const BeginArray                            As String = "["
Public Const BeginObject                           As String = "{"
Public Const EndArray                              As String = "]"
Public Const EndObject                             As String = "}"
Public Const NameSeparator                         As String = ":"
Public Const ValueSeparator                        As String = ","

Private Const ShortCommentStart                     As String = "//"
Private Const LongCommentStart                      As String = "/*"
Private Const LongCommentEnd                        As String = "*/"

Private Const EscSingleQuote                        As String = "\'"
Private Const EscQuotation                          As String = "\"""
Private Const EscRevSolidus                         As String = "\\"
Private Const EscSolidus                            As String = "\/"
Private Const EscBackspace                          As String = "\b"
Private Const EscFormFeed                           As String = "\f"
Private Const EscLineFeed                           As String = "\n"
Private Const EscCarrReturn                         As String = "\r"
Private Const EscTab                                As String = "\t"
Private Const EscHexDigits                          As String = "\u"

Private Const QuotationMark                         As String = """"
Private Const SingleQuoteMark                       As String = "'"

Private Const ValueTrue                             As String = "true"
Private Const ValueFalse                            As String = "false"
Private Const ValueNull                             As String = "null"

Private Const ProgIdName                            As String = "$ProgId"

'@================================================================================
' Member variables
'@================================================================================

Private mCurrPosn                                   As Long

Private mTableBuilder                               As StateTableBuilder

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

Public Property Get TableBuilder() As StateTableBuilder
If mTableBuilder Is Nothing Then
    Set mTableBuilder = New StateTableBuilder
    buildStateTable
End If
Set TableBuilder = mTableBuilder
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gEncode( _
                ByVal Value As Variant, _
                ByVal sb As StringBuilder)
Const ProcName As String = "gEncodeVariant"
On Error GoTo Err

If IsArray(Value) Then
    encodeArray Value, sb
ElseIf IsObject(Value) Then
    encodeObjectVariant Value, sb
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value to be encoded must be an array or object variant"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gEncodeVariant( _
                ByVal Value As Variant, _
                ByVal sb As StringBuilder)
Const ProcName As String = "gEncodeVariant"
On Error GoTo Err

If IsArray(Value) Then
    encodeArray Value, sb
ElseIf IsObject(Value) Then
    encodeObjectVariant Value, sb
Else
    encodeNonObjectVariant Value, sb
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Public Sub gParse( _
                ByVal pInputString As String, _
                ByRef pResult As Variant)
Const ProcName As String = "gParse"
On Error GoTo Err

mCurrPosn = 1

skipWhitespace pInputString

If tryChars(pInputString, BeginObject) Then
    parseObject pInputString, pResult
ElseIf tryChars(pInputString, BeginArray) Then
    pResult = parseArray(pInputString)
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Expected ""{"" or ""["" at position " & mCurrPosn
End If

Dim nextChar As String
If getNextChar(pInputString, nextChar) Then
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Unexpected non-white-space character(s) at position " & mCurrPosn
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addItemToArray( _
                ByVal Value As Variant, _
                ByRef ar() As Variant, _
                ByRef pIndex As Long)
Const ProcName As String = "addItemToArray"
On Error GoTo Err

If pIndex > UBound(ar) Then ReDim Preserve ar(2 * (UBound(ar) + 1) - 1) As Variant
gSetVariant ar(pIndex), Value
pIndex = pIndex + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub buildStateTable()

'=======================================================================
'                       State:      Created
'=======================================================================

mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.Created, _
            JSONBuilderStimuli.StimBeginArray, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.BuildingArray, _
            JSONBuilderActions.ActionEncodeBeginArray
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.Created, _
            JSONBuilderStimuli.StimBeginObject, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.BuildingObject, _
            JSONBuilderActions.ActionEncodeBeginObject

'=======================================================================
'                       State:      BuildingArray
'=======================================================================

mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingArray, _
            JSONBuilderStimuli.StimEndArray, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.Finished, _
            JSONBuilderActions.ActionEncodeEndArray _
            + JSONBuilderActions.ActionDecreaseLevel
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingArray, _
            JSONBuilderStimuli.StimValue, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingValue, _
            JSONBuilderActions.ActionEncodeValue
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingArray, _
            JSONBuilderStimuli.StimBeginArray, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingValue, _
            JSONBuilderActions.ActionIncreaseLevel, _
            JSONBuilderActions.ActionEncodeBeginArray
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingArray, _
            JSONBuilderStimuli.StimBeginObject, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingValue, _
            JSONBuilderActions.ActionIncreaseLevel, _
            JSONBuilderActions.ActionEncodeBeginObject
            
'=======================================================================
'                       State:      ExpectingValue
'=======================================================================

mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.ExpectingValue, _
            JSONBuilderStimuli.StimEndArray, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.Finished, _
            JSONBuilderActions.ActionEncodeEndArray, _
            JSONBuilderActions.ActionDecreaseLevel
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.ExpectingValue, _
            JSONBuilderStimuli.StimValue, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingValue, _
            JSONBuilderActions.ActionEncodeValueSeparator, _
            JSONBuilderActions.ActionEncodeValue
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.ExpectingValue, _
            JSONBuilderStimuli.StimBeginArray, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingValue, _
            JSONBuilderActions.ActionEncodeValueSeparator, _
            JSONBuilderActions.ActionIncreaseLevel, _
            JSONBuilderActions.ActionEncodeBeginArray
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.ExpectingValue, _
            JSONBuilderStimuli.StimBeginObject, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingValue, _
            JSONBuilderActions.ActionEncodeValueSeparator, _
            JSONBuilderActions.ActionIncreaseLevel, _
            JSONBuilderActions.ActionEncodeBeginObject
            
'=======================================================================
'                       State:      BuildingObject
'=======================================================================

mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingObject, _
            JSONBuilderStimuli.StimEndObject, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.Finished, _
            JSONBuilderActions.ActionEncodeEndObject, _
            JSONBuilderActions.ActionDecreaseLevel
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingObject, _
            JSONBuilderStimuli.StimName, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.BuildingFirstNameValuePair, _
            JSONBuilderActions.ActionEncodeName
            
'=======================================================================
'                       State:      BuildingFirstNameValuePair
'=======================================================================

mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingFirstNameValuePair, _
            JSONBuilderStimuli.StimValue, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingName, _
            JSONBuilderActions.ActionEncodeNameSeparator, _
            JSONBuilderActions.ActionEncodeValue
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingFirstNameValuePair, _
            JSONBuilderStimuli.StimBeginArray, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingName, _
            JSONBuilderActions.ActionEncodeNameSeparator, _
            JSONBuilderActions.ActionIncreaseLevel, _
            JSONBuilderActions.ActionEncodeBeginArray
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingFirstNameValuePair, _
            JSONBuilderStimuli.StimBeginObject, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingName, _
            JSONBuilderActions.ActionEncodeNameSeparator, _
            JSONBuilderActions.ActionIncreaseLevel, _
            JSONBuilderActions.ActionEncodeBeginObject
            
'=======================================================================
'                       State:      ExpectingName
'=======================================================================

mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.ExpectingName, _
            JSONBuilderStimuli.StimEndObject, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.Finished, _
            JSONBuilderActions.ActionEncodeEndObject, _
            JSONBuilderActions.ActionDecreaseLevel
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.ExpectingName, _
            JSONBuilderStimuli.StimName, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.BuildingNameValuePair, _
            JSONBuilderActions.ActionEncodeValueSeparator, _
            JSONBuilderActions.ActionEncodeName
            
'=======================================================================
'                       State:      BuildingNameValuePair
'=======================================================================

mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingNameValuePair, _
            JSONBuilderStimuli.StimValue, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingName, _
            JSONBuilderActions.ActionEncodeNameSeparator, _
            JSONBuilderActions.ActionEncodeValue
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingNameValuePair, _
            JSONBuilderStimuli.StimBeginArray, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingName, _
            JSONBuilderActions.ActionEncodeNameSeparator, _
            JSONBuilderActions.ActionIncreaseLevel, _
            JSONBuilderActions.ActionEncodeBeginArray
            
mTableBuilder.AddStateTableEntry _
            JSONBuilderStates.BuildingNameValuePair, _
            JSONBuilderStimuli.StimBeginObject, _
            SpecialConditions.NoConditions, _
            SpecialConditions.NoConditions, _
            JSONBuilderStates.ExpectingName, _
            JSONBuilderActions.ActionEncodeNameSeparator, _
            JSONBuilderActions.ActionIncreaseLevel, _
            JSONBuilderActions.ActionEncodeBeginObject
            
mTableBuilder.StateTableComplete
End Sub

Private Sub encodeArray( _
                ByVal Value As Variant, _
                ByVal sb As StringBuilder)
Const ProcName As String = "encodeArray"
On Error GoTo Err

Dim baseType As VbVarType
baseType = VarType(Value) And (Not VbVarType.vbArray)

sb.Append BeginArray

Dim doneFirst As Boolean
Dim Var As Variant
For Each Var In Value
    If Not doneFirst Then
        doneFirst = True
    Else
        sb.Append ValueSeparator
    End If
    If baseType = VbVarType.vbString Then
        encodeString CStr(Var), sb
    Else
        gEncodeVariant Var, sb
    End If
Next
sb.Append EndArray

Exit Sub

Err:
If Err.number = 92 Then
    sb.Append EndArray
    Exit Sub
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub encodeEnumerableObject( _
                ByVal Value As Variant, _
                ByVal sb As StringBuilder)
Const ProcName As String = "encodeEnumerableObject"
On Error GoTo Err

Dim doneFirst As Boolean
doneFirst = False
sb.Append BeginArray

Dim Var As Variant
For Each Var In Value
    If Not doneFirst Then
        doneFirst = True
    Else
        sb.Append ValueSeparator
    End If
    sb.Append NameSeparator
    gEncodeVariant Var, sb
Next
sb.Append EndArray

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub encodeDictionary( _
                ByVal Value As Variant, _
                ByVal sb As StringBuilder)
Const ProcName As String = "encodeDictionary"
On Error GoTo Err

Dim dict As Dictionary: Set dict = Value

Dim Keys() As Variant
Keys = dict.Keys

sb.Append BeginObject

Dim doneFirst As Boolean
Dim Var As Variant
For Each Var In Keys
    If Not doneFirst Then
        doneFirst = True
    Else
        sb.Append ValueSeparator
    End If
    gEncodeVariant Var, sb
    sb.Append NameSeparator
    gEncodeVariant dict.Item(Var), sb
Next
sb.Append EndObject

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub encodeJSONableObject( _
                ByVal Value As Variant, _
                ByVal sb As StringBuilder)
Const ProcName As String = "encodeJSONableObject"
On Error GoTo Err

Dim obj As IJSONable: Set obj = Value
sb.Append obj.ToJSON

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub encodeNonObjectVariant( _
                ByVal Value As Variant, _
                ByVal sb As StringBuilder)
Const ProcName As String = "encodeNonObjectVariant"
On Error GoTo Err

If IsArray(Value) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Argument must not be an array variant"

Select Case VarType(Value)
Case VbVarType.vbBoolean
    sb.Append IIf(Value, "true", "false")
Case VbVarType.vbString
    encodeString CStr(Value), sb
Case VbVarType.vbCurrency, _
        VbVarType.vbDecimal, _
        VbVarType.vbError, _
        VbVarType.vbInteger, _
        VbVarType.vbLong, _
        VbVarType.vbSingle
    sb.Append CStr(Value)
Case VbVarType.vbDouble
    ' we use gDoubleToString() rather than CStr because the latter does
    ' not properly handle MaxDouble - it gives a string that cannot be
    ' round-tripped
    sb.Append gDoubleToString(Value)
Case VbVarType.vbByte
    sb.Append CStr(Value)
Case VbVarType.vbDataObject
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Variants of type vbDataObject are not supported"
Case VbVarType.vbDate
    If Int(Value) = Value Then
        encodeString gFormatTimestamp(Value, TimestampDateOnlyISO8601), sb
    ElseIf Int(Value) = 0 Then
        encodeString gFormatTimestamp(Value, TimestampTimeOnlyISO8601 + TimestampNoMillisecs), sb
    Else
        encodeString gFormatTimestamp(Value, TimestampDateAndTimeISO8601 + TimestampNoMillisecs), sb
    End If
Case VbVarType.vbEmpty
    sb.Append "null"
Case VbVarType.vbNull
    sb.Append "null"
Case VbVarType.vbObject
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Argument must not be an object with no default property"
Case VbVarType.vbUserDefinedType
    Err.Raise ErrorCodes.ErrUnsupportedOperationException, , "Variants of type vbUserDefinedType are not supported"
End Select

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
                
End Sub

Private Sub encodeNonJSONableObject( _
                ByVal Value As Variant, _
                ByVal sb As StringBuilder)
Const ProcName As String = "encodeNonJSONableObject"
On Error GoTo Err

Dim baseType As VbVarType
baseType = VarType(Value) And (Not VbVarType.vbArray)

' object has no natural JSON string representation
If baseType <> VbVarType.vbObject Then
    ' this means the value has a default property - we'll use
    ' that value instead
    sb.Append BeginObject
    sb.Append "DefaultProp"
    sb.Append NameSeparator
    gEncodeVariant Value, sb
    sb.Append EndObject
Else
    ' nothing we can find out about this object
    sb.Append BeginObject
    sb.Append "***No info available***"
    sb.Append NameSeparator
    encodeString "", sb
    sb.Append EndObject
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub encodeNothing( _
                ByVal sb As StringBuilder)
sb.Append BeginObject
sb.Append EndObject
End Sub

Private Sub encodeObjectVariant( _
                ByVal Value As Variant, _
                ByVal sb As StringBuilder)
Const ProcName As String = "encodeObjectVariant"
On Error GoTo Err

Dim baseType As VbVarType
baseType = VarType(Value) And (Not VbVarType.vbArray)

If Value Is Nothing Then
    encodeNothing sb
ElseIf TypeOf Value Is IJSONable Then
    encodeJSONableObject Value, sb
ElseIf TypeOf Value Is Dictionary Then
    encodeDictionary Value, sb
ElseIf TypeOf Value Is Collection Or _
    TypeOf Value Is IEnumerable _
Then
    encodeEnumerableObject Value, sb
Else
    encodeNonJSONableObject Value, sb
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub encodeString( _
                ByVal str As String, _
                ByVal sb As StringBuilder)
Const ProcName As String = "encodeString"
On Error GoTo Err

sb.Append QuotationMark

Dim i As Long
For i = 1 To Len(str)
    Dim ch As String
    ch = Mid$(str, i, 1)
    
    Select Case ch
    Case QuotationMark
        sb.Append "\"""
    Case "\"
        sb.Append "\\"
    Case "/"            ' we escape "/" in case it is followed by "*" or "/" which
                        ' could be misinterpreted as comment markers on re-parsing
        sb.Append "\/"
    Case vbBack
        sb.Append "\\"
    Case vbFormFeed
        sb.Append "\f"
    Case vbLf
        sb.Append "\n"
    Case vbCr
        sb.Append "\r"
    Case vbTab
        sb.Append "\t"
    Case Else
        Dim a As Integer
        a = AscW(ch)
        If a > 31 And a < 127 Then
            sb.Append ch
        ElseIf a >= 0 Or a < 65535 Then
            Dim ah As String
            ah = Hex(a)
            sb.Append "\u"
            sb.Append String(4 - Len(ah), "0")
            sb.Append ah
        End If
    End Select
Next

sb.Append QuotationMark

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getNextChar( _
                ByVal inputString As String, _
                ByRef nextChar As String) As Boolean
Const ProcName As String = "getNextChar"
On Error GoTo Err

If mCurrPosn > Len(inputString) Then Exit Function

If tryChars(inputString, ShortCommentStart) Then
    skipShortComment inputString
ElseIf tryChars(inputString, LongCommentStart) Then
    skipLongComment inputString
End If

nextChar = Mid$(inputString, mCurrPosn, 1)
mCurrPosn = mCurrPosn + 1
getNextChar = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseArray( _
                ByVal inputString As String) As Variant()
Const ProcName As String = "parseArray"
On Error GoTo Err

ReDim ar(3) As Variant

skipWhitespace inputString

Dim lIndex As Long

If tryChars(inputString, EndArray) Then
Else
    Dim Var As Variant
    parseValue inputString, Var
    addItemToArray Var, ar, lIndex
    
    Do While Not tryChars(inputString, EndArray)

        If Not tryChars(inputString, ValueSeparator) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Expected "","" at position " & mCurrPosn
        
        skipWhitespace inputString
        
        parseValue inputString, Var
        addItemToArray Var, ar, lIndex
    Loop
End If

skipWhitespace inputString

If lIndex > 0 Then
    ReDim Preserve ar(lIndex - 1) As Variant
    parseArray = ar
End If
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseDigit( _
                ByVal inputString As String) As String
Const ProcName As String = "parseDigit"
On Error GoTo Err

Dim nextChar As String
If Not getNextChar(inputString, nextChar) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Expected digit at position " & mCurrPosn
If nextChar < "0" Or nextChar > "9" Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Expected digit at position " & mCurrPosn

parseDigit = nextChar

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseDigits( _
                ByVal inputString As String) As String
Const ProcName As String = "parseDigits"
On Error GoTo Err

Do While peekDigit(inputString) <> ""
    parseDigits = parseDigits & parseDigit(inputString)
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseInteger( _
                ByVal inputString As String) As String
Const ProcName As String = "parseinteger"
On Error GoTo Err

parseInteger = parseDigit(inputString)

If parseInteger = "0" Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Expected non-zero digit at position " & mCurrPosn

parseInteger = parseInteger & parseDigits(inputString)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseName(ByVal inputString As String) As String
Const ProcName As String = "parseName"
On Error GoTo Err

skipWhitespace inputString

If tryChars(inputString, QuotationMark) Then
    parseName = parseString(inputString, QuotationMark)
    Exit Function
ElseIf tryChars(inputString, SingleQuoteMark) Then
    parseName = parseString(inputString, SingleQuoteMark)
    Exit Function
End If

Do
    If peekChars(inputString, NameSeparator) Then Exit Do
    
    Dim nextChar As String
    If Not getNextChar(inputString, nextChar) Then Exit Do
    
    parseName = parseName & nextChar
Loop

parseName = Trim$(parseName)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Sub parseNameValuePair( _
                ByVal inputString As String, _
                ByRef Name As String, _
                ByRef Value As Variant)
Const ProcName As String = "parseNameValuePair"
On Error GoTo Err

Name = parseName(inputString)

If Not tryChars(inputString, NameSeparator) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Expected "":"" at position " & mCurrPosn
   
parseValue inputString, Value
skipWhitespace inputString

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function parseNumber( _
                ByVal inputString As String)
Const ProcName As String = "parseNumber"
On Error GoTo Err

skipWhitespace inputString

Dim Value As String
If tryChars(inputString, "-") Then Value = "-"

If tryChars(inputString, "0") Then
    Value = Value & "0"
Else
    Value = Value & parseInteger(inputString)
End If

Dim isDouble As Boolean
If tryChars(inputString, ".") Then
    isDouble = True
    Value = Value & "."
    
    Value = Value & parseDigit(inputString)
    
    Value = Value & parseDigits(inputString)
End If

If tryChars(inputString, "E") Or _
    tryChars(inputString, "e") _
Then
    isDouble = True
    Value = Value & "e"
    
    If tryChars(inputString, "-") Then
        Value = Value & "-"
    ElseIf tryChars(inputString, "+") Then
        Value = Value & "+"
    End If
    
    Value = Value & parseDigit(inputString)
    
    Value = Value & parseDigits(inputString)
End If

If isDouble Then
    parseNumber = CDbl(Value)
Else
    parseNumber = CLng(Value)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub parseObject( _
                ByVal pInputString As String, _
                ByRef pResult As Variant)
Const ProcName As String = "parseObject"
On Error GoTo Err

skipWhitespace pInputString

If tryChars(pInputString, EndObject) Then
    Set pResult = Nothing
    Exit Sub
End If

Dim Name As String
Dim Value As Variant
parseNameValuePair pInputString, Name, Value

Dim usingDict As Boolean
Dim dict As Dictionary
Dim obj As Object
If UCase$(Name) = UCase$(ProgIdName) Then
    Set obj = CreateObject(Value)
    Set pResult = obj
Else
    Set dict = New Dictionary
    Set pResult = dict
    usingDict = True
    dict.Add Name, Value
End If

Do
    If tryChars(pInputString, EndObject) Then Exit Do
    
    If Not tryChars(pInputString, ValueSeparator) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Expected "","" at position " & mCurrPosn
    
    skipWhitespace pInputString
    parseNameValuePair pInputString, Name, Value
    If usingDict Then
        dict.Add Name, Value
    Else
        CallByName obj, Name, VbLet, Value
    End If
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function parseString( _
                ByVal inputString As String, _
                ByVal delimiter As String) As String
Const ProcName As String = "parseString"
On Error GoTo Err

Dim sb As New StringBuilder: sb.Initialise
skipWhitespace inputString

Do
    Dim nextChar As String
    If tryChars(inputString, delimiter) Then
        parseString = sb.ToString
        skipWhitespace inputString
        Exit Function
    ElseIf tryChars(inputString, EscSingleQuote) Then
        sb.Append SingleQuoteMark
    ElseIf tryChars(inputString, EscQuotation) Then
        sb.Append QuotationMark
    ElseIf tryChars(inputString, EscRevSolidus) Then
        sb.Append "\"
    ElseIf tryChars(inputString, EscSolidus) Then
        sb.Append "/"
    ElseIf tryChars(inputString, EscQuotation) Then
        sb.Append QuotationMark
    ElseIf tryChars(inputString, EscSingleQuote) Then
        sb.Append "'"
    ElseIf tryChars(inputString, EscBackspace) Then
        sb.Append vbBack
    ElseIf tryChars(inputString, EscFormFeed) Then
        sb.Append vbFormFeed
    ElseIf tryChars(inputString, EscLineFeed) Then
        sb.Append vbLf
    ElseIf tryChars(inputString, EscCarrReturn) Then
        sb.Append vbCr
    ElseIf tryChars(inputString, EscTab) Then
        sb.Append vbTab
    ElseIf tryChars(inputString, EscHexDigits) Then
        sb.Append ChrW("&h" & parseUnicode(inputString))
    ElseIf getNextChar(inputString, nextChar) Then
        sb.Append nextChar
    Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Expected " & delimiter & " at position " & mCurrPosn
    End If
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function parseUnicode(ByVal inputString As String) As String
Const ProcName As String = "parseUnicode"
On Error GoTo Err

Dim i As Long
For i = 1 To 4
    Dim nextChar As String
    If Not getNextChar(inputString, nextChar) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Expected hex character at position " & mCurrPosn
    
    If (nextChar >= "0" And nextChar <= "9") Or _
        (nextChar >= "a" And nextChar <= "f") Or _
        (nextChar >= "A" And nextChar <= "F") _
    Then
    Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid hex char at position " & mCurrPosn
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub parseValue( _
                ByVal pInputString As String, _
                ByRef pResult As Variant)
Const ProcName As String = "parseValue"
On Error GoTo Err

skipWhitespace pInputString

If tryChars(pInputString, BeginObject) Then
    parseObject pInputString, pResult
ElseIf tryChars(pInputString, BeginArray) Then
    pResult = parseArray(pInputString)
ElseIf tryChars(pInputString, QuotationMark) Then
    pResult = parseString(pInputString, QuotationMark)
ElseIf tryChars(pInputString, SingleQuoteMark) Then
    pResult = parseString(pInputString, SingleQuoteMark)
ElseIf tryChars(pInputString, ValueTrue) Then
    pResult = True
ElseIf tryChars(pInputString, ValueFalse) Then
    pResult = False
ElseIf tryChars(pInputString, ValueNull) Then
    pResult = Null
Else
    pResult = parseNumber(pInputString)
End If

skipWhitespace pInputString

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function peekChars( _
                ByVal inputString As String, _
                ByVal charsToTry As String) As Boolean
Const ProcName As String = "peekChars"
On Error GoTo Err

If Len(inputString) + 1 - mCurrPosn < Len(charsToTry) Then Exit Function

If Mid$(inputString, mCurrPosn, Len(charsToTry)) = charsToTry Then peekChars = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function peekDigit( _
                ByVal inputString As String) As String
Const ProcName As String = "peekDigit"
On Error GoTo Err

Dim nextChar As String
nextChar = peekNextChar(inputString)
If nextChar >= "0" And nextChar <= "9" Then peekDigit = nextChar

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function peekNextChar( _
                ByVal inputString As String) As String
Const ProcName As String = "peekNextChar"
On Error GoTo Err

If mCurrPosn > Len(inputString) Then Exit Function

peekNextChar = Mid$(inputString, mCurrPosn, 1)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub skipLongComment( _
                ByVal inputString As String)
Const ProcName As String = "skipLongComment"
On Error GoTo Err

Do While mCurrPosn <= Len(inputString) - 1
    If Mid$(inputString, mCurrPosn, 2) = LongCommentEnd Then Exit Do
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub skipShortComment( _
                ByVal inputString As String)
Const ProcName As String = "skipShortComment"
On Error GoTo Err

Do While mCurrPosn <= Len(inputString)
    If tryChars(inputString, vbCr) Then
        tryChars inputString, vbLf
        Exit Do
    ElseIf tryChars(inputString, vbLf) Then
        Exit Do
    End If
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub skipWhitespace( _
                ByVal inputString As String)
Const ProcName As String = "skipWhitespace"
On Error GoTo Err

Do While tryChars(inputString, " ") Or _
        tryChars(inputString, vbTab) Or _
        tryChars(inputString, vbCr) Or _
        tryChars(inputString, vbLf)
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function tryChars( _
                ByVal inputString As String, _
                ByVal charsToTry As String) As Boolean
Const ProcName As String = "tryChars"
On Error GoTo Err

If peekChars(inputString, charsToTry) Then
    mCurrPosn = mCurrPosn + Len(charsToTry)
    tryChars = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

