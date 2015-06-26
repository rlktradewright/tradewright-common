Attribute VB_Name = "GExtendedEventManager"
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

Private Const ModuleName                            As String = "GExtendedEventManager"

'@================================================================================
' Member variables
'@================================================================================

Private mClassEventCollections                      As New Collection

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

Public Function gRegisterExtendedEvent( _
                ByVal pName As String, _
                ByVal pMode As ExtendedEventModes, _
                ByVal pOwnerClassName As String) As ExtendedEvent
Const ProcName As String = "gRegisterExtendedEvent"
On Error GoTo Err

Dim lEventsForClass As Collection

Set lEventsForClass = addEventsForClass(pOwnerClassName)

On Error Resume Next
Set gRegisterExtendedEvent = lEventsForClass(pName)
On Error GoTo Err

If gRegisterExtendedEvent Is Nothing Then
    Set gRegisterExtendedEvent = New ExtendedEvent
    gRegisterExtendedEvent.Initialise pName, pMode
    lEventsForClass.Add gRegisterExtendedEvent, pName
Else
    If gRegisterExtendedEvent.Mode <> pMode Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                ProjectName & "." & ModuleName & ":" & ProcName, _
                "Already registered for this class with different properties"
    
    End If
End If

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName, pProjectName:=ProjectName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addEventsForClass(ByVal pOwnerClassName As String) As Collection
On Error Resume Next
Set addEventsForClass = mClassEventCollections(pOwnerClassName)
If addEventsForClass Is Nothing Then
    Set addEventsForClass = New Collection
    mClassEventCollections.Add addEventsForClass, pOwnerClassName
End If
End Function


