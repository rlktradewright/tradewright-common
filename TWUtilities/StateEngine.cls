VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StateEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' A simple implememtation of a finite State engine.
'
' Instances of this class cannot be created using the <code>New</code> operator.
' Instead, use the <code>CreateStateEngine</code> global method.
'
' See the documentation for the <code>StateTableBuilder</code> class for information
' about states, stimuli, conditions and actions.
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

Private Const ModuleName                As String = "StateEngine"

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
Private mState As Long
Private mConditions As Long

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
' Gets the current Conditions.
'
' @return
'   The current Conditions.
'@/
Public Property Get Conditions() As Long
Conditions = mConditions
End Property

''
' Gets the current State.
'
' @return
'   The current State.
'@/
Public Property Get State() As Long
State = mState
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Clears one or more conditions.
'
' @param Value
'   A bit-signficant Value specifying the conditions to be cleared.
'@/
Public Sub ClearConditions( _
                ByVal Value As Long)
mConditions = mConditions And (Not Value)
End Sub

''
' Sets one or more conditions.
'
' @param Value
'   A bit-signficant Value specifying the conditions to be set.
'@/
Public Sub SetConditions( _
                ByVal Value As Long)
mConditions = mConditions Or Value
End Sub

Friend Sub Initialise( _
                ByVal TableBuilder As StateTableBuilder, _
                ByVal initialState As Long)
Const ProcName As String = "Initialise"

On Error GoTo Err

mStateTable = TableBuilder.StateTable
mState = initialState

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Notifies the occurrence of a stimulus, potentially resulting in a change of State
' and returning a list of actions to be taken.
'
' @remarks
'   Attempting to notify a stimulus that is not expected (as specified in the State
'   table) causes Error <code>ErrorCodes.ErrIllegalArgumentException</code> to be raised, and
'   the Error State is entered.
'
'   Attempting to notify any stimulus when the State engine is in the Error State
'   causes Error <code>ErrorCodes.ErrIllegalStateException</code> to be raised, and the
'   State engine remains in the Error State.
' @return
'   An array of Action identifiers.
' @param Value
'   The stimulus to be notified.
'@/
Public Function NotifyStimulus( _
                ByVal Value As Long) As Long()
Const ProcName As String = "NotifyStimulus"
On Error GoTo Err



Dim i As Long
Dim validStimulus As Boolean
Dim entryState As Long

If mState = StateError Then Err.Raise ErrorCodes.ErrIllegalStateException, , "State engine is in Error State"

If Value < 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value must be non-negative"

entryState = mState

For i = 0 To UBound(mStateTable)
    If mStateTable(i).State = mState And _
        (mStateTable(i).Stimulus = Value Or mStateTable(i).Stimulus = SpecialStimuli.StimulusAll) And _
        (((mStateTable(i).ConditionsTrue And mConditions) = mStateTable(i).ConditionsTrue) Or _
            (mStateTable(i).ConditionsTrue = SpecialConditions.NoConditions)) And _
        (((mStateTable(i).ConditionsFalse And mConditions) = 0) Or _
            (mStateTable(i).ConditionsFalse = SpecialConditions.NoConditions)) _
    Then
        mState = mStateTable(i).NewState
        If mState = StateError Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Stimulus " & Value & " not permitted in State " & entryState & " in these conditions"
        NotifyStimulus = mStateTable(i).Actions
        validStimulus = True
        Exit For
    End If
Next

If Not validStimulus Then
    mState = StateError
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Stimulus " & Value & " not permitted in State " & entryState
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


