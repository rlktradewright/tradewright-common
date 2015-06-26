VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ColumnResizer"
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

Private Const ModuleName                    As String = "ColumnResizer"

'@================================================================================
' Declares
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mCol                                As Long
Private mFirstX                             As Long
Private mLastX                              As Long
Private mMinX                               As Long
Private mColWidth                           As Long
Private mHeight                             As Long

Private mResizeLine                         As line

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

Public Property Get Col() As Long
Const ProcName As String = "col"
On Error GoTo Err

Col = mCol

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function endResize() As Long
Const ProcName As String = "endResize"
On Error GoTo Err

mResizeLine.Visible = False
Screen.MousePointer = MousePointerConstants.vbDefault
If mFirstX <> &H80000000 Then endResize = (mLastX - mFirstX)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Initialise( _
                ByVal pCol As Long, _
                ByVal pColLeft As Long, _
                ByVal pColWidth As Long, _
                ByVal pHeight As Long, _
                ByVal pResizeLine As line)
Const ProcName As String = "Initialise"
On Error GoTo Err

mCol = pCol
mMinX = pColLeft
mColWidth = pColWidth
mHeight = pHeight
Set mResizeLine = pResizeLine

Screen.MousePointer = MousePointerConstants.vbSizeWE
mFirstX = &H80000000

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub moveTo( _
                ByVal X As Long)
Const ProcName As String = "moveTo"
On Error GoTo Err

X = X
If X < mMinX Then X = mMinX
If X = mLastX Then Exit Sub

If mFirstX = &H80000000 Then mFirstX = X
mLastX = X
DrawLine mLastX - mFirstX + mMinX + mColWidth

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub
                

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub DrawLine( _
                ByVal X As Long)
Const ProcName As String = "DrawLine"
On Error GoTo Err

mResizeLine.Visible = False
mResizeLine.x1 = X
mResizeLine.y1 = 0
mResizeLine.x2 = X
mResizeLine.y2 = mHeight
mResizeLine.ZOrder 0
mResizeLine.Visible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


