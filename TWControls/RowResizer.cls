VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RowResizer"
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

Private Const ModuleName                    As String = "RowResizer"

'@================================================================================
' Declares
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mRow                                As Long
Private mFirstY                             As Long
Private mLastY                              As Long
Private mMinY                               As Long
Private mRowHeight                          As Long
Private mWidth                              As Long

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

Public Property Get Row() As Long
Row = mRow
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function endResize() As Long
Const ProcName As String = "endResize"
On Error GoTo Err

mResizeLine.Visible = False
Screen.MousePointer = MousePointerConstants.vbDefault
If mFirstY <> &H80000000 Then endResize = (mLastY - mFirstY)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Initialise( _
                ByVal pRow As Long, _
                ByVal pRowTop As Long, _
                ByVal pRowHeight As Long, _
                ByVal pWidth As Long, _
                ByVal pResizeLine As line)
Const ProcName As String = "Initialise"
On Error GoTo Err

mRow = pRow
mMinY = pRowTop
mRowHeight = pRowHeight
mWidth = pWidth
Set mResizeLine = pResizeLine

Screen.MousePointer = MousePointerConstants.vbSizeNS
mFirstY = &H80000000

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub moveTo( _
                ByVal Y As Long)
Const ProcName As String = "moveTo"
On Error GoTo Err

Debug.Print "Row resizer move to " & Y
Y = Y
If Y < mMinY Then Y = mMinY
If Y = mLastY Then Exit Sub

If mFirstY = &H80000000 Then mFirstY = Y
mLastY = Y
DrawLine mLastY - mFirstY + mMinY + mRowHeight

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub
                

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub DrawLine( _
                ByVal Y As Long)
Const ProcName As String = "DrawLine"
On Error GoTo Err

mResizeLine.Visible = False
mResizeLine.y1 = Y
mResizeLine.y2 = Y
mResizeLine.ZOrder 0
mResizeLine.Visible = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub




