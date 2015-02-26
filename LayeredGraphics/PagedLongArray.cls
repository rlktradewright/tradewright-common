VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PagedLongArray"
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
' Constants
'@================================================================================

Private Const ModuleName                            As String = "PagedLongArray"

Const MaxPageItemIndex                              As Long = &H3F&

Const MaxPageIndex                                  As Long = &H1FFF&
Const PageTableDivisor                              As Long = 1 + MaxPageItemIndex
Const PageTableMask                                 As Long = MaxPageIndex * PageTableDivisor

Const MaxSegmentIndex                               As Long = &H1FFF&
Const SegmentTableDivisor                           As Long = (1 + MaxPageItemIndex) * (1 + MaxPageIndex)
Const SegmentTableMask                              As Long = &HFFF80000

'@================================================================================
' Types
'@================================================================================

Private Type Page
    Data(MaxPageItemIndex) As Long
End Type

Private Type PageTable
    PageIndex(MaxPageIndex) As Long
End Type

'@================================================================================
' Member variables
'@================================================================================

Private mSegmentTable() As Long

Private mPageTables() As PageTable
Private mPageTablesIndex As Long

Private mPages() As Page
Private mPagesIndex As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mSegmentTable(MaxSegmentIndex) As Long

ReDim mPageTables(1 To 1) As PageTable
mPageTablesIndex = 0

ReDim mPages(1 To 64) As Page
mPagesIndex = 0

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

Public Sub SetEntry(ByVal pAddress As Long, ByVal Value As Long)
Dim lSegment As Long
Dim lPage As Long
Dim lItem As Long

lSegment = ((pAddress And SegmentTableMask) / SegmentTableDivisor) And MaxSegmentIndex
lPage = (pAddress And PageTableMask) / PageTableDivisor
lItem = pAddress And MaxPageItemIndex

If mSegmentTable(lSegment) = 0 Then
    mPageTablesIndex = mPageTablesIndex + 1
    If mPageTablesIndex > UBound(mPageTables) Then ReDim Preserve mPageTables(1 To 2 * UBound(mPageTables)) As PageTable
    mSegmentTable(lSegment) = mPageTablesIndex
End If

Dim lPageTableIndex As Long
lPageTableIndex = mSegmentTable(lSegment)
If mPageTables(lPageTableIndex).PageIndex(lPage) = 0 Then
    mPagesIndex = mPagesIndex + 1
    If mPagesIndex > UBound(mPages) Then ReDim Preserve mPages(1 To 2 * UBound(mPages)) As Page
    mPageTables(lPageTableIndex).PageIndex(lPage) = mPagesIndex
End If

Dim lPageIndex As Long
lPageIndex = mPageTables(lPageTableIndex).PageIndex(lPage)
mPages(lPageIndex).Data(lItem) = Value
End Sub

Public Function GetEntry(ByVal pAddress As Long) As Long
Dim lSegment As Long
Dim lPage As Long
Dim lItem As Long

lSegment = ((pAddress And SegmentTableMask) / SegmentTableDivisor) And MaxSegmentIndex
lPage = (pAddress And PageTableMask) / PageTableDivisor
lItem = pAddress And MaxPageItemIndex

If mSegmentTable(lSegment) = 0 Then Exit Function

Dim lPageTableIndex As Long
lPageTableIndex = mSegmentTable(lSegment)
If mPageTables(lPageTableIndex).PageIndex(lPage) = 0 Then Exit Function

Dim lPageIndex As Long
lPageIndex = mPageTables(lPageTableIndex).PageIndex(lPage)
GetEntry = mPages(lPageIndex).Data(lItem)
End Function

'@================================================================================
' Helper Functions
'@================================================================================





