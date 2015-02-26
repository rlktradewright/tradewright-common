VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphObjSeriesCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Private Const ModuleName                            As String = "GraphObjSeriesCollection"

'@================================================================================
' Member variables
'@================================================================================

Private mGraphObjSeriesTable()                      As IGraphicObjectSeries
Private mGraphObjSeriesIndex                        As Long
Private mModelRef                                  As WeakReference

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mGraphObjSeriesTable(3) As IGraphicObjectSeries
End Sub

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

Friend Function Add( _
                ByVal pGraphObjSeries As IGraphicObjectSeries, _
                ByVal pLayer As Long, _
                ByVal pName As String, _
                ByVal pController As Controller) As IGraphicObjectSeries
Const ProcName As String = "Add"

On Error GoTo Err

Set Add = pGraphObjSeries

mGraphObjSeriesIndex = mGraphObjSeriesIndex + 1
If mGraphObjSeriesIndex > UBound(mGraphObjSeriesTable) Then
    ReDim Preserve mGraphObjSeriesTable(2 * (UBound(mGraphObjSeriesTable) + 1) - 1) As IGraphicObjectSeries
    If gLogger.IsLoggable(LogLevelHighDetail) Then _
        gLogger.Log "Increased mGraphObjSeriesTable size to", ProcName, ModuleName, LogLevelHighDetail, CStr(UBound(mGraphObjSeriesTable) + 1)
End If

Set mGraphObjSeriesTable(mGraphObjSeriesIndex) = pGraphObjSeries

If pName = "" Then pName = "Series" & mGraphObjSeriesIndex

pGraphObjSeries.Initialise pName, _
                        pLayer, _
                        mGraphObjSeriesIndex, _
                        pController

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function Clear( _
                ByVal pSeries As IGraphicObjectSeries) As TaskController
Const ProcName As String = "Remove"
On Error GoTo Err

Dim t As New GraphObjSeriesRemovalTask

t.Initialise pSeries, ModelObj, True
Set Clear = StartTask(t, PriorityLow)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Finish()
Dim i As Long
Dim t As GraphObjSeriesRemovalTask

' remove all graphic object references from the various collections
Const ProcName As String = "Finish"

On Error GoTo Err

For i = 0 To mGraphObjSeriesIndex
    If Not mGraphObjSeriesTable(i) Is Nothing Then
        mGraphObjSeriesTable(i).Finish
        
        Set t = New GraphObjSeriesRemovalTask
        t.Initialise mGraphObjSeriesTable(i), ModelObj, False
        StartTask t, PriorityLow

        Set mGraphObjSeriesTable(i) = Nothing
    End If
Next
Erase mGraphObjSeriesTable

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub Initialise(ByVal pModel As LayeredGraphicsModel)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mModelRef = CreateWeakReference(pModel)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function Item( _
                ByVal index As Long) As IGraphicObjectSeries
Const ProcName As String = "Item"

On Error GoTo Err

Set Item = mGraphObjSeriesTable(index)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Remove( _
                ByVal pSeries As IGraphicObjectSeries)
Const ProcName As String = "Remove"

On Error GoTo Err

Clear pSeries

Set mGraphObjSeriesTable(pSeries.SeriesID) = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function ModelObj() As LayeredGraphicsModel
Const ProcName As String = "ModelObj"

On Error GoTo Err

Set ModelObj = mModelRef.Target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


