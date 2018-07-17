VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SimpleConditionBuilder"
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

Private Const ModuleName                    As String = "SimpleConditionBuilder"

'@================================================================================
' Member variables
'@================================================================================

Private mConditionString As String

Private mBracketCount As Long
Private mExpectingLogicalOperator As Boolean

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

Public Property Get conditionString() As String
conditionString = mConditionString
End Property

Public Property Get expectingLogicalOperator() As Boolean
expectingLogicalOperator = mExpectingLogicalOperator
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub addTerm( _
                ByVal columnName As String, _
                ByVal operator As ConditionalOperators, _
                ByVal Value As String, _
                Optional ByVal logicalOperator As LogicalOperators = LogicalOpNone, _
                Optional ByVal negate As Boolean)
Const ProcName As String = "addTerm"
Dim failpoint As String
On Error GoTo Err

Assert mExpectingLogicalOperator Or logicalOperator = LogicalOpNone, "No logical operator expected"

Assert Not (mExpectingLogicalOperator And logicalOperator = LogicalOpNone), "Logical operator expected"


mExpectingLogicalOperator = True

Dim arg As String
arg = gCleanQueryArg(Value)

If mConditionString <> "" Then
    mConditionString = mConditionString & logicalOperatorToString(logicalOperator)
End If

If negate Then mConditionString = mConditionString & " NOT"

mConditionString = mConditionString & " " & columnName
mConditionString = mConditionString & conditionalOperatorToString(operator)
mConditionString = mConditionString & " '" & arg & "'"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
                
End Sub

Public Sub clear()
Const ProcName As String = "clear"
Dim failpoint As String
On Error GoTo Err

mConditionString = ""
mBracketCount = 0
mExpectingLogicalOperator = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub endPredicate()
Const ProcName As String = "endPredicate"
Dim failpoint As String
On Error GoTo Err

mBracketCount = mBracketCount + 1
Assert mBracketCount >= 0, "No predicate started"
Assert mExpectingLogicalOperator, "Empty predicate"

mConditionString = mConditionString & ")"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub startPredicate( _
                Optional ByVal negate As Boolean)
Const ProcName As String = "startPredicate"
Dim failpoint As String
On Error GoTo Err

mBracketCount = mBracketCount + 1
mExpectingLogicalOperator = False
If negate Then mConditionString = mConditionString & " NOT"
mConditionString = mConditionString & " ("

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function conditionalOperatorToString( _
                ByVal operator As ConditionalOperators) As String
Const ProcName As String = "conditionalOperatorToString"
Dim failpoint As String
On Error GoTo Err

Select Case operator
Case CondOpEqual
    conditionalOperatorToString = " ="
Case CondOpGreaterThan
    conditionalOperatorToString = " >"
Case CondOpGreaterThanOrEqual
    conditionalOperatorToString = " >="
Case CondOpLessThan
    conditionalOperatorToString = " <"
Case CondOpLessThanOrEqual
    conditionalOperatorToString = " <="
Case CondOpLike
    conditionalOperatorToString = " LIKE"
Case Else
    AssertArgument False, "Invalid conditional operator"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function logicalOperatorToString( _
                ByVal logicalOperator As LogicalOperators) As String
Const ProcName As String = "logicalOperatorToString"
Dim failpoint As String
On Error GoTo Err

Select Case logicalOperator
Case LogicalOpNone
    logicalOperatorToString = ""
Case LogicalOpAND
    logicalOperatorToString = " AND"
Case LogicalOpOR
    logicalOperatorToString = " OR"
Case Else
    AssertArgument False, "Invalid logical operator"
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

