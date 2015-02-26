VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ErrorItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Encapsulates information about a violated validation rule.
'
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


Private Const ModuleName As String = "ErrorItem"

'@================================================================================
' Member variables
'@================================================================================

Private mRuleId As String
Private mParameters() As Variant

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

''
' Returns an array of <code>Variant</code>s associated with the validation rule
' violation.
'
' <p>These are actually the values specified in the <code>parameters</code> argument
' when the AssertRule method was called for the rule corresponding to this
' <code>ErrorItem</code>.</p>
'
' @return An array of <code>Variant</code>s.
' @param Name description
' @see
'
'@/
Public Property Get Parameters() As Variant()
Parameters = mParameters
End Property

''
' Returns the identifier for the validation rule whose violation gave rise to this
' <code>ErrorItem</code>.
'
' @return The identifier for the validation rule
'
'@/
Public Property Get RuleId() As String
RuleId = mRuleId
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal RuleId As String, _
                ByRef Parameters() As Variant)
mRuleId = RuleId
mParameters = Parameters
End Sub
                

'@================================================================================
' Helper Functions
'@================================================================================


