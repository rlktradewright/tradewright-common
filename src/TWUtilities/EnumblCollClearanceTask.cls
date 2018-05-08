VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "EnumblCollClearanceTask"
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

Implements ITask

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

Private Const ModuleName                            As String = "EnumblCollClearanceTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mCollEn                                     As Enumerator

Private mHeuristicNumber                            As Long
Private mHeuristicTime                              As Long

Private mStartTime                                  As Date

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' ITask Interface Members
'@================================================================================

Private Sub ITask_Cancel()

End Sub

Private Sub ITask_Run()
Const ProcName As String = "ITask_Run"
On Error GoTo Err

Static sTotalNumberReleased As Long
Static sYieldCount As Long

Dim results As Parameters

If mStartTime = 0 Then mStartTime = gGetTimestamp

If mTaskContext.CancelPending Then
    gLogger.Log "Released (cancelled) " & sTotalNumberReleased & " entries in " & sYieldCount & " yields", ProcName, ModuleName, LogLevelHighDetail
    Set results = generateResults(False, sTotalNumberReleased, sYieldCount)
    mTaskContext.Finish results, True
    Exit Sub
End If

Static sNumberToRelease As Long

Static sLoopCount As Long

If sLoopCount Mod 50 = 0 Then
    sNumberToRelease = calibrate(mCollEn, mHeuristicNumber, mHeuristicTime, sTotalNumberReleased, sYieldCount)
ElseIf Not releaseEntries(mCollEn, sNumberToRelease, sTotalNumberReleased, sYieldCount) Then
    sNumberToRelease = 0
End If

If sNumberToRelease = 0 Then
    gLogger.Log "Released " & sTotalNumberReleased & " entries in " & sYieldCount & " yields", ProcName, ModuleName, LogLevelHighDetail
    Set results = generateResults(False, sTotalNumberReleased, sYieldCount)
    mTaskContext.Finish results, False
End If

sLoopCount = sLoopCount + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Let ITask_TaskContext(ByVal Value As TaskContext)
Set mTaskContext = Value
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

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
                ByVal pCollEn As Enumerator, _
                ByVal pHeuristicNumber As Long, _
                ByVal pHeuristicTime As Long)
Set mCollEn = pCollEn
mHeuristicNumber = pHeuristicNumber
mHeuristicTime = pHeuristicTime
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calibrate( _
                ByVal pCollEn As Enumerator, _
                ByVal pNumberToRelease As Long, _
                ByVal pHeuristicTime As Long, _
                ByRef pTotalNumberReleased As Long, _
                ByRef pYieldCount As Long) As Long
Const ProcName As String = "calibrate"
On Error GoTo Err

Dim lPrevCpuTime As Double: lPrevCpuTime = gGetCurrentProcessCpuTime
Dim lAligned As Boolean

Dim lNumberReleased As Long
Do While releaseEntries(mCollEn, pNumberToRelease, pTotalNumberReleased, pYieldCount)
    lNumberReleased = lNumberReleased + pNumberToRelease
    Dim lCurrCpuTime As Double: lCurrCpuTime = gGetCurrentProcessCpuTime
    Dim lCpuTimeMillisecs As Double: lCpuTimeMillisecs = 1000 * (lCurrCpuTime - lPrevCpuTime)
    If lCpuTimeMillisecs > 0 Then
        If Not lAligned Then
            lNumberReleased = 0
            lPrevCpuTime = lCurrCpuTime
            lAligned = True
        Else
            gLogger.Log "Time to release " & lNumberReleased & " entries: " & lCpuTimeMillisecs, ProcName, ModuleName, LogLevelHighDetail
            calibrate = CLng(pHeuristicTime / lCpuTimeMillisecs * lNumberReleased)
            gLogger.Log "Number of entries to release per yield: " & calibrate, ProcName, ModuleName, LogLevelHighDetail
            Exit Function
        End If
    End If
Loop

calibrate = 0

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function generateResults( _
                ByVal pCancelled As Boolean, _
                ByVal pNumberReleased As Long, _
                ByVal pNumberOfYields As Long) As Parameters
Const ProcName As String = "generateResults"
On Error GoTo Err

Dim results As New Parameters
results.SetParameterValue "Cancelled", pCancelled
results.SetParameterValue "Number of entries released", pNumberReleased
results.SetParameterValue "Number of timeslots", pNumberOfYields

Dim lDuration As Double: lDuration = 86400# * (gGetTimestamp - mStartTime)
If lDuration < 0.001 Then
    results.SetParameterValue "Duration (microsecs)", Format(lDuration * 1000000#, "0")
ElseIf lDuration < 1# Then
    results.SetParameterValue "Duration (millisecs)", Format(lDuration * 1000#, "0.000")
Else
    results.SetParameterValue "Duration (secs)", Format(lDuration, "0.000")
End If

Set generateResults = results

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function releaseEntries( _
                ByVal pCollEn As Enumerator, _
                ByVal pNumberToRelease As Long, _
                ByRef pTotalNumberReleased As Long, _
                ByRef pYieldCount As Long) As Boolean
Const ProcName As String = "releaseEntries"
On Error GoTo Err

Dim lNumberReleased As Long

Do While pCollEn.MoveNext
    pCollEn.Remove
    
    lNumberReleased = lNumberReleased + 1
    If lNumberReleased >= pNumberToRelease Then
        releaseEntries = True
        pTotalNumberReleased = pTotalNumberReleased + lNumberReleased
        pYieldCount = pYieldCount + 1
        Exit Function
    End If
Loop

pTotalNumberReleased = pTotalNumberReleased + lNumberReleased
pYieldCount = pYieldCount + 1

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




