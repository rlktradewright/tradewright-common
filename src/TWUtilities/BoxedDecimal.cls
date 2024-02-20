VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BoxedDecimal"
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

Implements IComparable
Implements IStringable
Implements IJSONable

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

Private Const ModuleName                            As String = "BoxedDecimal"

Private Const FlagIsInteger                         As Long = &H1
Private Const FlagIsPositiveInfinity                As Long = &H2
Private Const FlagIsNegativeInfinity                As Long = &H4

'@================================================================================
' Member variables
'@================================================================================

Private mValue                                      As Variant
Private mFlags                                      As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' IComparable Interface Members
'@================================================================================

Private Function IComparable_compareTo(ByVal obj As IComparable) As Long
Const ProcName As String = "IComparable_CompareTo"
On Error GoTo Err

gAssertArgument obj Is Not Null, "obj argument is Nothing"

Dim d As BoxedDecimal: Set d = obj

If Me.DecimalValue < d.DecimalValue Then
    IComparable_compareTo = -1
ElseIf Me.DecimalValue > d.DecimalValue Then
    IComparable_compareTo = 1
Else
    IComparable_compareTo = 0
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' IJSONable Interface Members
'@================================================================================

Private Function IJSONable_ToJSON() As String
IJSONable_ToJSON = ToJSON
End Function

'@================================================================================
' IStringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
IStringable_ToString = gDecimalToString(Me)
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let DecimalValue(ByVal Value As Variant)
gAssert IsEmpty(mValue), "Value already set"
Initialise Value
End Property

Public Property Get DecimalValue() As Variant
Attribute DecimalValue.VB_UserMemId = 0
DecimalValue = mValue
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function Add(d As BoxedDecimal) As BoxedDecimal
gAssertArgument Not d Is Nothing, "A value must be supplied"
If Me Is gDecimalZero Then
    Set Add = d
ElseIf d Is gDecimalZero Then
    Set Add = Me
Else
    Set Add = New BoxedDecimal
    Add.Initialise Me.DecimalValue + d.DecimalValue
End If
End Function

Public Function Div(d As BoxedDecimal) As BoxedDecimal
gAssertArgument Not d Is Nothing, "A value must be supplied"
If Me Is gDecimalZero Then
    Set Div = gDecimalZero
Else
    Set Div = New BoxedDecimal
    Div.Initialise Me.DecimalValue / d.DecimalValue
End If
End Function

Public Function EQ(d As BoxedDecimal) As Boolean
If d Is Nothing Then
    EQ = False
ElseIf d Is Me Then
    EQ = True
ElseIf Me.DecimalValue = d.DecimalValue Then
    EQ = True
Else
    EQ = False
End If
End Function

Public Function GE(d As BoxedDecimal) As Boolean
gAssertArgument Not d Is Nothing, "A value must be supplied"
GE = Me.DecimalValue >= d.DecimalValue
End Function

Public Function GT(d As BoxedDecimal) As Boolean
gAssertArgument Not d Is Nothing, "A value must be supplied"
GT = Me.DecimalValue > d.DecimalValue
End Function

Friend Sub Initialise(pValue As Variant)
    mValue = CDec(pValue)
    If (mValue - Fix(mValue)) = 0 Then mFlags = FlagIsInteger
End Sub

Public Function IsInteger() As Boolean
IsInteger = (mFlags And FlagIsInteger) = FlagIsInteger
End Function

Public Function LE(d As BoxedDecimal) As Boolean
gAssertArgument Not d Is Nothing, "A value must be supplied"
LE = Me.DecimalValue <= d.DecimalValue
End Function

Public Function LT(d As BoxedDecimal) As Boolean
gAssertArgument Not d Is Nothing, "A value must be supplied"
LT = Me.DecimalValue < d.DecimalValue
End Function

Public Function Mult(d As BoxedDecimal) As BoxedDecimal
Set Mult = New BoxedDecimal
Mult.Initialise Me.DecimalValue * d.DecimalValue
End Function

Public Function NE(d As BoxedDecimal) As Boolean
If d Is Nothing Then
    NE = True
ElseIf d Is Me Then
    NE = False
ElseIf Me.DecimalValue = d.DecimalValue Then
    NE = False
Else
    NE = True
End If
End Function

Public Function Subtr(d As BoxedDecimal) As BoxedDecimal
Set Subtr = New BoxedDecimal
Subtr.Initialise Me.DecimalValue - d.DecimalValue
End Function

Public Function ToJSON() As String
Dim jb As JSONBuilder
Set jb = New JSONBuilder

jb.BeginObject

jb.EncodeName "$ProgID"
jb.EncodeValue ProjectName & "." & ModuleName

jb.EncodeName "DecimalValue"
jb.EncodeValue mValue

jb.EndObject

ToJSON = jb.JSONString
End Function

Public Function ToString() As String
ToString = gDecimalToString(Me)
End Function

'@================================================================================
' Helper Functions
'@================================================================================





