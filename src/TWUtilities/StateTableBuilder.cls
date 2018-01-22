VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StateTableBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class is used to create state tables for use with the <code>StateEngine</code>
' class.
'
' A state table consists of a set of <em>state table entries</em>. Each state table
' entry specifies an initial State, a stimulus that may occur in that State,
' the conditions under which it may occur, the new State that is entered when
' that stimulus occurs under those particular conditions, and a list of actions that must
' be performed when this happens.
'
' There may be several state table entries that have the same initial State and stimulus,
' but that have different conditions.
'
' States are specified by non-negative <code>Long</code> values assigned by the application. There
' is a generic special State called the <em>error state</em> which is entered when
' the State engine encounters an unexpected stimulus.
'
' Stimuli are specified by non-negative <code>Long</code> values assigned by the application.
'
' Conditions are bit values assigned by the application (ie each value must be a
' power of two so that it can be represented by a single bit). There is a limit of 32
' conditions because they are passed using 32-bit values.
'
' Note that there are two predefined condition masks:
' <ul>
' <li><code>SpecialConditions.NoConditions</code>: this is useful when specifying state
' table entries for which no condition need be true or no conditions need be false. It
' is more documentary than using the literal value 0</li>
' <li><code>SpecialConditions.AllConditions</code>: there are few uses for this value but
' it is defined for completeness.</li>
' </ul>
'
' Actions are specified by non-negative <code>Long</code> values assigned by the application.
'
' A state table, once created, is read-only. Therefore a single <code>StateTableBuilder</code>
' object can be used for all <code>StateEngine</code> objects that have the same state table
' definition.
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                As String = "StateTableBuilder"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mStateTable() As StateTableEntry
Private mNumStateTableEntries As Long

Private mInitialisationComplete As Boolean

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

Friend Property Get StateTable() As StateTableEntry()
gAssert mInitialisationComplete, "state table not yet fully initialised"

StateTable = mStateTable
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Adds a state table entry to the state table under construction by this
' <code>StateTableBuilder</code> object.
'
' @remarks
'   No checks are made to detect entries that have identical values for
'   <code>state</code>. <code>stimulus</code>, <code>conditionsTrue</code>
'   and <code>conditionsFalse</code>.
'
'   Once the <code>stateTableComplete</code> method has been called, attempting
'  to add further state table entries causes error
'   <code>ErrorCodes.ErrIllegalStateException</code> to be raised.
' @return
'   An array of Action identifiers that are to be performed when this
'   entry is triggered.
' @param pState
'   The entry State.
' @param pstimulus
'   The stimulus.
' @param pconditionsTrue
'   A bit-mask of conditions that must be true for this entry to be relevant.
' @param pconditionsFalse
'   A bit-mask of conditions that must be false for this entry to be relevant.
' @param pnewState
'   The State to be entered when this entry is triggered.
' @param Actions
'   An array of actions to be carried out when this entry is triggered.
'@/
Public Function AddStateTableEntry( _
    ByVal pState As Long, _
    ByVal pStimulus As Long, _
    ByVal pConditionsTrue As Long, _
    ByVal pConditionsFalse As Long, _
    ByVal pNewState As Long, _
    ParamArray Actions() As Variant)

Dim theActions() As Long
Dim StateTableEntry As StateTableEntry
Dim Action As Variant
Dim i As Long

Const ProcName As String = "AddStateTableEntry"

On Error GoTo Err

gAssertArgument pState >= 0, "State may not be negative"
gAssertArgument pNewState >= 0 Or pNewState = SpecialStates.StateError, "pNewState must be non-negative or SpecialStates.StateError"
gAssert Not mInitialisationComplete, "Can't Add more state table entries"

StateTableEntry.State = pState
StateTableEntry.Stimulus = pStimulus
StateTableEntry.NewState = pNewState
StateTableEntry.ConditionsFalse = pConditionsFalse
StateTableEntry.ConditionsTrue = pConditionsTrue

If UBound(Actions) >= 0 Then
    ReDim theActions(UBound(Actions)) As Long

    For Each Action In Actions
        theActions(i) = CLng(Action)
        i = i + 1
    Next
Else
    ReDim theActions(0) As Long
    theActions(0) = SpecialActions.NoAction
End If

StateTableEntry.Actions = theActions

If mNumStateTableEntries = 0 Then
    ReDim Preserve mStateTable(5) As StateTableEntry
ElseIf mNumStateTableEntries > UBound(mStateTable) Then
    ReDim Preserve mStateTable(mNumStateTableEntries + 5) As StateTableEntry
End If

mStateTable(mNumStateTableEntries) = StateTableEntry

mNumStateTableEntries = mNumStateTableEntries + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

''
' Indicates that all required state table entries have been added.
'
' @remarks
'   Once this method has been called, adding further state table entries
'   causes error <code>ErrorCodes.ErrIllegalStateException</code> to be raised.
' @return
'   description
' @param Name
'   description
' @see
'
'@/
Public Sub StateTableComplete()
Const ProcName As String = "StateTableComplete"

On Error GoTo Err

ReDim Preserve mStateTable(mNumStateTableEntries - 1) As StateTableEntry
mInitialisationComplete = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




