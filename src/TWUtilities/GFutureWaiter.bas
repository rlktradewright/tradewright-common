Attribute VB_Name = "GFutureWaiter"
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

Private Const ModuleName                            As String = "GFutureWaiter"

'@================================================================================
' Member variables
'@================================================================================

Private mFutures                                    As SortedDictionary
Private mDatas                                      As SortedDictionary

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

Public Property Get gFutures() As SortedDictionary
If mFutures Is Nothing Then
    Set mFutures = New SortedDictionary
    mFutures.Initialise KeyTypeString, True
End If
Set gFutures = mFutures
End Property

Public Property Get gDatas() As SortedDictionary
If mDatas Is Nothing Then
    Set mDatas = New SortedDictionary
    mDatas.Initialise KeyTypeString, True
End If
Set gDatas = mDatas
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gDiagnosticLog( _
                ByVal pDiagnosticID As String, _
                ByVal pMessage As String, _
                ByVal pProcName As String, _
                ByVal pModuleName As String)

Dim s As String: s = "ID=" & pDiagnosticID & ": " & pMessage
'''Debug.Print pModuleName & "::" & pProcName & ": " & s
gLogger.Log s, pProcName, pModuleName, LogLevelDetail
End Sub

Public Function gGetFutureStateAsString(ByVal pFuture As IFuture) As String
If pFuture.IsAvailable Then
    gGetFutureStateAsString = "Available"
ElseIf pFuture.IsCancelled Then
    gGetFutureStateAsString = "Cancelled"
ElseIf pFuture.IsFaulted Then
    gGetFutureStateAsString = "Errored"
ElseIf pFuture.IsPending Then
    gGetFutureStateAsString = "Pending"
End If
End Function

'@================================================================================
' Helper Functions
'@================================================================================




