VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphicObjectSeriesBase"
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

Event RepaintRegion(ByVal pRecalcBoundaries As Boolean)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GraphicObjectSeriesBase"

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mEPhost                          As ExtendedPropertyHost
Attribute mEPhost.VB_VarHelpID = -1

Private mEditController                             As EditController

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

Set mEPhost = New ExtendedPropertyHost
Set mEditController = New EditController

Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' mEPHost Event Handlers
'@================================================================================

Private Sub mEPhost_Change(pEv As ChangeEventData)
Const ProcName As String = "mEPhost_Change"
On Error GoTo Err

Dim lChangeType As ExtendedPropertyHostChangeTypes
lChangeType = pEv.ChangeType
If lChangeType = ParentChanged Or lChangeType = StylesChanged Then RaiseEvent RepaintRegion(True)

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
        mEditController.SetBoundaryToBeRecalculated
    ElseIf lMetadata.AffectsPaintingRegion Then
        mEditController.SetRenderRequired
    ElseIf lMetadata.AffectsRender Then
        mEditController.SetRenderRequired
    End If
Else
    If lMetadata.AffectsPosition Or lMetadata.AffectsSize Then
        RaiseEvent RepaintRegion(True)
    ElseIf lMetadata.AffectsPaintingRegion Then
        RaiseEvent RepaintRegion(False)
    ElseIf lMetadata.AffectsRender Then
        RaiseEvent RepaintRegion(False)
    End If
End If
Exit Sub

Err:
gNotifyUnhandledError pProcedureName:=ProcName, pModuleName:=ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get ExtendedPropertyHost() As ExtendedPropertyHost
Set ExtendedPropertyHost = mEPhost
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub ApplyChanges()
If mEditController.RenderRequired Then RaiseEvent RepaintRegion(mEditController.BoundaryToBeRecalculated)
mEditController.EndEdit
End Sub

Public Sub Edit()
mEditController.BeginEdit
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




