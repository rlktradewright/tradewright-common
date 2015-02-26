VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelectionManager"
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

Private Const ModuleName                            As String = "SelectionManager"

'@================================================================================
' Member variables
'@================================================================================

Private mSelectedObject                             As IGraphicObject
Private mModel                                      As LayeredGraphicsModel

Private mIsSelectionEnabled                         As Boolean

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

Friend Property Get SelectedObject() As IGraphicObject
Set SelectedObject = mSelectedObject
End Property

Friend Property Let IsSelectionEnabled(ByVal Value As Boolean)
mIsSelectionEnabled = Value
End Property

Friend Property Get IsSelectionEnabled() As Boolean
IsSelectionEnabled = mIsSelectionEnabled
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub ClearSelection()
Const ProcName As String = "ClearSelection"

On Error GoTo Err

deSelectObjectOrSeries mSelectedObject
Set mSelectedObject = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function ClickSelectedObjectOrSeries() As Boolean
Const ProcName As String = "ClickSelectedObjectOrSeries"

On Error GoTo Err

If mSelectedObject Is Nothing Then
    ClickSelectedObjectOrSeries = False
ElseIf mSelectedObject.IsSelectable Then
    mSelectedObject.Click
    ClickSelectedObjectOrSeries = True
ElseIf mModel.GetGraphicObjectSeries(mSelectedObject.SeriesID).IsSelectable Then
    mModel.GetGraphicObjectSeries(mSelectedObject.SeriesID).Click
    ClickSelectedObjectOrSeries = True
Else
    ClickSelectedObjectOrSeries = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function DblClickSelectedObjectOrSeries() As Boolean
Const ProcName As String = "DblClickSelectedObjectOrSeries"

On Error GoTo Err

If mSelectedObject Is Nothing Then
    DblClickSelectedObjectOrSeries = False
ElseIf mSelectedObject.IsSelectable Then
    mSelectedObject.DblCLick
    DblClickSelectedObjectOrSeries = True
ElseIf mModel.GetGraphicObjectSeries(mSelectedObject.SeriesID).IsSelectable Then
    mModel.GetGraphicObjectSeries(mSelectedObject.SeriesID).DblCLick
    DblClickSelectedObjectOrSeries = True
Else
    DblClickSelectedObjectOrSeries = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise(ByVal pModel As LayeredGraphicsModel)
Set mModel = pModel
End Sub

Friend Sub SetSelectedObjectOrSeries( _
                ByVal pCurrX As Single, _
                ByVal pCurrY As Single)
Dim newSelectedObject As IGraphicObject

Const ProcName As String = "setSelectedObjectOrSeries"

On Error GoTo Err

If Not mIsSelectionEnabled Then Exit Sub

Set newSelectedObject = mModel.GetNextSelectableObjectOrSeriesAtPoint(mSelectedObject, pCurrX, pCurrY)

If newSelectedObject Is Nothing Then
    deSelectObjectOrSeries mSelectedObject
    Set mSelectedObject = Nothing
ElseIf newSelectedObject Is mSelectedObject Then
    
Else
    deSelectObjectOrSeries mSelectedObject
    Set mSelectedObject = newSelectedObject
    selectObjectOrSeries newSelectedObject
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub deSelectObjectOrSeries(ByVal GraphObj As IGraphicObject)
Const ProcName As String = "deSelectObjectOrSeries"

On Error GoTo Err

If GraphObj Is Nothing Then
ElseIf GraphObj.IsSelectable Then
    GraphObj.IsSelected = False
ElseIf mModel.GetGraphicObjectSeries(GraphObj.SeriesID).IsSelectable Then
    mModel.GetGraphicObjectSeries(GraphObj.SeriesID).IsSelected = False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub selectObjectOrSeries(ByVal GraphObj As IGraphicObject)
Const ProcName As String = "selectObjectOrSeries"

On Error GoTo Err

If GraphObj.IsSelectable Then
    GraphObj.IsSelected = True
ElseIf mModel.GetGraphicObjectSeries(GraphObj.SeriesID).IsSelectable Then
    mModel.GetGraphicObjectSeries(GraphObj.SeriesID).IsSelected = True
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




