VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SortedDictClearanceTask"
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

Private Const ModuleName                            As String = "SortedDictClearanceTask"

'@================================================================================
' Member variables
'@================================================================================

Private mTaskContext                                As TaskContext

Private mEntriesEn                                  As Enumerator

Private mHeuristicNumber                            As Long
Private mHeuristicTime                              As Long

Private mExpFactor                                  As Double

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

If mTaskContext.CancelPending Then
    mTaskContext.Finish Empty, True
    Exit Sub
End If

Static sTotalNumberReleased As Long
Static sNumberToRelease As Long
Static sYieldCount As Long

Static sLoopCount As Long

If sLoopCount Mod 50 = 0 Then
    sNumberToRelease = calibrate(mEntriesEn, mHeuristicNumber, mHeuristicTime, sTotalNumberReleased, sYieldCount)
ElseIf Not releaseEntries(mEntriesEn, sNumberToRelease, sTotalNumberReleased, sYieldCount) Then
    sNumberToRelease = 0
End If

If sNumberToRelease = 0 Then
    gLogger.Log "Released " & sTotalNumberReleased & " entries in " & sYieldCount & " yields", ProcName, ModuleName, LogLevelHighDetail
    mTaskContext.Finish Empty, False
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
                ByVal pEntriesEn As Enumerator, _
                ByVal pHeuristicNumber As Long, _
                ByVal pHeuristicTime As Long, _
                ByVal pSmoothingPeriods As Long)
Set mEntriesEn = pEntriesEn
mHeuristicNumber = pHeuristicNumber
mHeuristicTime = pHeuristicTime
mExpFactor = 2 / (pSmoothingPeriods + 1)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calibrate( _
                ByVal pEntriesEn As Enumerator, _
                ByVal pNumberToRelease As Long, _
                ByVal pHeuristicTime As Long, _
                ByRef pTotalNumberReleased As Long, _
                ByRef pYieldCount As Long) As Long
Const ProcName As String = "calibrate"
On Error GoTo Err

Dim lPrevCpuTime As Double: lPrevCpuTime = gGetCurrentProcessCpuTime
Dim lAligned As Boolean

Dim lNumberReleased As Long
Do While releaseEntries(mEntriesEn, pNumberToRelease, pTotalNumberReleased, pYieldCount)
    lNumberReleased = lNumberReleased + mHeuristicNumber
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

Private Function releaseEntries( _
                ByVal pEntriesEn As Enumerator, _
                ByVal pNumberToRelease As Long, _
                ByRef pTotalNumberReleased As Long, _
                ByRef pYieldCount As Long) As Boolean
Const ProcName As String = "releaseEntries"
On Error GoTo Err

Dim lNumberReleased As Long

Do While pEntriesEn.MoveNext
    Dim lEntry As DictionaryEntry
    Set lEntry = pEntriesEn.Current
    
    lEntry.Data = 0
    lEntry.Key = 0
    
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


