VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DeferredPainter"
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

Implements DeferredAction

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

Public Enum DeferredPaintActions
    DeferredPaintRepaintNone
    DeferredPaintRepaintChangedAreas
    DeferredPaintRepaintAll
    DeferredPaintChangeGaugeandRepaint
    DeferredPaintScaleAndRepaint
    DeferredPaintCalculateAndRepaint
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "DeferredRegionPainter"

'@================================================================================
' Member variables
'@================================================================================

Private mControllerRef As WeakReference

Private mDeferredPaintPending As Boolean

Private mDeferredPaintAction As DeferredPaintActions

Private mAreasToRedraw() As TRectangle
Private mAreasToRedrawIndex As Long

Private mDeferredPaintDelay As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mAreasToRedraw(7) As TRectangle
End Sub

'================================================================================
' DeferredAction Interface Members
'================================================================================

Private Sub DeferredAction_run(ByVal Data As Variant)
Static et As ElapsedTimer

Const ProcName As String = "DeferredAction_run"

On Error GoTo Err

Dim s As String
Dim numObjectsRepainted As Long

#If trace Then
    gTracer.EnterProcedure pInfo:="", pProcedureName:=ProcName, pModuleName:=ModuleName
#End If

If controllerObj.IsCleared Then Exit Sub

If gIsInDev Then
    If et Is Nothing Then Set et = New ElapsedTimer
    et.StartTiming

    s = "Deferred paint: "
    
    Select Case mDeferredPaintAction
    Case DeferredPaintRepaintNone
        s = s & "Repaint none: "
    Case DeferredPaintRepaintChangedAreas
        s = s & "Repaint " & CStr(mAreasToRedrawIndex + 1) & " changed areas; "
    Case DeferredPaintRepaintAll
        s = s & "Repaint all: "
    Case DeferredPaintChangeGaugeandRepaint
        s = s & "Change gauge and repaint: "
    Case DeferredPaintScaleAndRepaint
        s = s & "Scale and repaint: "
    Case DeferredPaintCalculateAndRepaint
        s = s & "Calculate and repaint: "
    End Select
End If

Select Case mDeferredPaintAction
Case DeferredPaintRepaintNone

Case DeferredPaintRepaintChangedAreas
    numObjectsRepainted = controllerObj.RepaintAreas(mAreasToRedraw, mAreasToRedrawIndex + 1)
Case DeferredPaintRepaintAll
    controllerObj.PaintAll
Case DeferredPaintChangeGaugeandRepaint
    controllerObj.ChangeGaugeAndRepaint
Case DeferredPaintScaleAndRepaint
    controllerObj.ScaleAndRepaint
Case DeferredPaintCalculateAndRepaint
    controllerObj.CalculateAndRepaint
End Select


If gIsInDev Then
    If mDeferredPaintAction = DeferredPaintRepaintChangedAreas Then s = s & CStr(numObjectsRepainted) & " objects; "
    s = s & Format(et.ElapsedTimeMicroseconds, "0 microsecs")
    Debug.Print s
End If

doDeferredPaint DeferredPaintRepaintNone

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddAreaToRedraw( _
                ByRef areaToRedraw As TRectangle)
Dim i As Long
Dim handled As Boolean

Const ProcName As String = "AddareaToRedraw"

On Error GoTo Err

If Not areaToRedraw.isValid Then Exit Sub

For i = 0 To mAreasToRedrawIndex
    If TRectangleContainsRect(mAreasToRedraw(i), areaToRedraw) Then
        handled = True
        Exit For
    ElseIf TRectangleContainsRect(areaToRedraw, mAreasToRedraw(i)) Then
        If Not handled Then
            mAreasToRedraw(i) = areaToRedraw
            handled = True
        Else
            mAreasToRedraw(i).isValid = False
        End If
    End If
Next

If Not handled Then
    Debug.Print "Add area to redraw: " & TRectangleToString(areaToRedraw)
    mAreasToRedrawIndex = mAreasToRedrawIndex + 1
    If mAreasToRedrawIndex > UBound(mAreasToRedraw) Then
        ReDim Preserve mAreasToRedraw(2 * (UBound(mAreasToRedraw) + 1) - 1) As TRectangle
        If gLogger.IsLoggable(LogLevelHighDetail) Then _
            gLogger.Log "Increased mAreasToRedraw size to", ProcName, ModuleName, LogLevelHighDetail, CStr(UBound(mAreasToRedraw) + 1)
    End If
    mAreasToRedraw(mAreasToRedrawIndex) = areaToRedraw
End If

doDeferredPaint DeferredPaintRepaintChangedAreas

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub CalculateAndRepaint()
Const ProcName As String = "CalculateAndRepaint"

On Error GoTo Err

doDeferredPaint DeferredPaintCalculateAndRepaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ChangeGaugeAndRepaint()
Const ProcName As String = "ChangeGaugeAndRepaint"

On Error GoTo Err

doDeferredPaint DeferredPaintChangeGaugeandRepaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pController As Controller, _
                ByVal pDeferredPaintDelay As Long)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mControllerRef = CreateWeakReference(pController)
mDeferredPaintDelay = pDeferredPaintDelay
doDeferredPaint DeferredPaintRepaintNone

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RepaintAll()
Const ProcName As String = "RepaintAll"

On Error GoTo Err

doDeferredPaint DeferredPaintRepaintAll

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub ScaleAndRepaint()
Const ProcName As String = "ScaleAndRepaint"

On Error GoTo Err

doDeferredPaint DeferredPaintScaleAndRepaint

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub doDeferredPaint(ByVal pAction As DeferredPaintActions)
Const ProcName As String = "DeferredPaint"

On Error GoTo Err

If pAction = DeferredPaintRepaintNone Then
    mDeferredPaintAction = DeferredPaintRepaintNone
    mDeferredPaintPending = False
Else
    Select Case mDeferredPaintAction
    Case DeferredPaintRepaintNone
        mDeferredPaintAction = pAction
    Case DeferredPaintRepaintChangedAreas
        If pAction = DeferredPaintRepaintAll Or _
            pAction = DeferredPaintChangeGaugeandRepaint Or _
            pAction = DeferredPaintScaleAndRepaint Or _
            pAction = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = pAction
    Case DeferredPaintRepaintAll
        If pAction = DeferredPaintChangeGaugeandRepaint Or _
            pAction = DeferredPaintScaleAndRepaint Or _
            pAction = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = pAction
    Case DeferredPaintChangeGaugeandRepaint
        If pAction = DeferredPaintScaleAndRepaint Or _
            pAction = DeferredPaintCalculateAndRepaint _
        Then mDeferredPaintAction = pAction
    Case DeferredPaintScaleAndRepaint
        If pAction = DeferredPaintCalculateAndRepaint Then mDeferredPaintAction = pAction
    Case DeferredPaintCalculateAndRepaint
        
    End Select
    If Not mDeferredPaintPending Then
        mDeferredPaintPending = True
        DeferAction Me, Empty, mDeferredPaintDelay, ExpiryTimeUnitMilliseconds
    End If
End If

If pAction <> DeferredPaintRepaintChangedAreas Then mAreasToRedrawIndex = -1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Function controllerObj() As Controller
Const ProcName As String = "controllerObj"

On Error GoTo Err

Set controllerObj = mControllerRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


