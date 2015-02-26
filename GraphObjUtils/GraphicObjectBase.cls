VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicObjectBase"
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

Private Const ModuleName                            As String = "GraphicObjectBase"

'@================================================================================
' Member variables
'@================================================================================

Private mGraphObjRef                                As WeakReference

Private WithEvents mEPhost                          As ExtendedPropertyHost
Attribute mEPhost.VB_VarHelpID = -1

Private mEditController                             As EditController

Private mSeries                                     As IGraphicObjectSeries

Private mGraphicsUpdateNumber                       As Long
Private mSeriesUpdateNumber                         As Long

Private mBoundingRect                               As TRectangle

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mEditController = New EditController
TRectangleInitialise mBoundingRect

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mEPHost Event Handlers
'@================================================================================

Private Sub mEPhost_Change(pEv As ChangeEventData)
Const ProcName As String = "mEPhost_Change"
On Error GoTo Err

Dim lChangeType As ExtendedPropertyHostChangeTypes
lChangeType = pEv.ChangeType
If lChangeType = ParentChanged Or lChangeType = StylesChanged Then initiateRedraw True, True

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Sub mEPHost_ExtendedPropertyChanged(pEv As ExtendedPropertyChangedEventData)
Dim lMetadata As GraphicExtPropMetadata

Const ProcName As String = "mEPHost_ExtendedPropertyChanged"
On Error GoTo Err

Set lMetadata = pEv.ExtendedProperty.Metadata
If mEditController.Editing Then
    If lMetadata.AffectsPosition Or lMetadata.AffectsSize Then
        mEditController.SetImageToBeCleared
        mEditController.SetBoundaryToBeRecalculated
    ElseIf lMetadata.AffectsPaintingRegion Then
        mEditController.SetImageToBeCleared
    ElseIf lMetadata.AffectsRender Then
        mEditController.SetRenderRequired
    End If
Else
    If lMetadata.AffectsPosition Or lMetadata.AffectsSize Then
        initiateRedraw True, True
    ElseIf lMetadata.AffectsPaintingRegion Then
        initiateRedraw True, False
    ElseIf lMetadata.AffectsRender Then
        initiateRedraw False, False
    End If
End If
Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get BoundingRectangle() As TRectangle
Const ProcName As String = "BoundingRectangle"
On Error GoTo Err

If needToRecalculateBoundingRect Then mBoundingRect = mSeries.Renderer.GetBoundingRectangle(graphObj)
BoundingRectangle = mBoundingRect

Exit Property

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Property

Public Property Get Series() As IGraphicObjectSeries
Set Series = mSeries
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ApplyChanges()
Const ProcName As String = "ApplyChanges"
On Error GoTo Err

If mEditController.RenderRequired Then initiateRedraw mEditController.ImageToBeCleared, mEditController.BoundaryToBeRecalculated
mEditController.EndEdit

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Draw(pArea As TRectangle)
Const ProcName As String = "Draw"
On Error GoTo Err

mSeries.Renderer.Render graphObj

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Public Sub Edit()
mEditController.BeginEdit
End Sub

Public Sub Initialise( _
                ByVal pGraphObj As IGraphicObject, _
                ByVal pSeries As IGraphicObjectSeries, _
                ByVal pEPHost As ExtendedPropertyHost)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mGraphObjRef = CreateWeakReference(pGraphObj)
Set mSeries = pSeries
Set mEPhost = pEPHost

mEPhost.SetParent mSeries.ExtendedPropertyHost, False

mGraphicsUpdateNumber = mSeries.GraphicsUpdateNumber
mSeriesUpdateNumber = mSeries.SeriesUpdateNumber

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function graphObj() As IGraphicObject
Set graphObj = mGraphObjRef.Target
End Function

Private Sub initiateRedraw( _
                ByVal pUndrawCurrentImage As Boolean, _
                ByVal pRecalcBoundaries As Boolean)
Const ProcName As String = "initiateRedraw"
Dim failpoint As String
On Error GoTo Err

If graphObj.Handle = 0 Then
    ' haven't yet been added to model
    Exit Sub
End If
If pUndrawCurrentImage Then mSeries.Controller.UndrawObject graphObj

If graphObj.NoDraw Then Exit Sub

If pRecalcBoundaries Then
    mBoundingRect = mSeries.Renderer.GetBoundingRectangle(graphObj)
End If
mSeries.LayeredGraphicsModel.UpdateGraphicObject graphObj

Exit Sub

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

Private Function needToRecalculateBoundingRect() As Boolean
Dim lGraphicsUpdateNumber As Long
Dim lSeriesUpdateNumber As Long

Const ProcName As String = "needToRecalculateBoundingRect"
On Error GoTo Err

lSeriesUpdateNumber = mSeries.SeriesUpdateNumber
If mSeriesUpdateNumber <> lSeriesUpdateNumber Then
    ' the series (or its style) have experienced one or more
    ' changes. We don't know what these were, so we need to recalculate
    needToRecalculateBoundingRect = True
    mSeriesUpdateNumber = lSeriesUpdateNumber
    mGraphicsUpdateNumber = mSeries.GraphicsUpdateNumber
    Exit Function
End If

lGraphicsUpdateNumber = mSeries.GraphicsUpdateNumber
If mGraphicsUpdateNumber = lGraphicsUpdateNumber Then
    ' neither the series nor the graphics have had any changes
    needToRecalculateBoundingRect = False
    Exit Function
End If

If mGraphicsUpdateNumber = (lGraphicsUpdateNumber - 1) Then
    ' there has only been one change to the graphics since last time this function was called
    If gIsFlagSet(mSeries.LastGraphicsChange, GraphicsChangeTypes.GraphicsGaugeChangedX) Or _
        gIsFlagSet(mSeries.LastGraphicsChange, GraphicsChangeTypes.GraphicsGaugeChangedY) _
    Then
        ' the gauge has changed, so we need to recalculate
        needToRecalculateBoundingRect = True
    Else
        ' only the graphics boundary has changed, no need to recalculate
        needToRecalculateBoundingRect = False
    End If
    mGraphicsUpdateNumber = lGraphicsUpdateNumber
    Exit Function
End If

' there have been more than one changes to the graphics. We can't find out what they
' are so we need to recalculate
needToRecalculateBoundingRect = True
mGraphicsUpdateNumber = lGraphicsUpdateNumber

Exit Function

Err:
gHandleUnexpectedError pReRaise:=True, pLog:=False, pProcedureName:=ProcName, pModuleName:=ModuleName

End Function


