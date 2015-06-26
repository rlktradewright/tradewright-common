VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GraphObjLists"
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

Implements Enumerable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type StripTableEntry
    GraphobjList                As LinkedList
    Cycle                       As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "GraphObjLists"

'@================================================================================
' Member variables
'@================================================================================

Private mStripTableIndex                            As PagedLongArray

Private mStripTable()                               As StripTableEntry
Private mStripTableMaxIndex                         As Long

Private mCycle                                      As Long

Private mMinUsedStripNumber                         As Long
Private mMaxUsedStripNumber                         As Long

Private mStripWidth                                 As Double
Private mStripOrigin                                As Double

Private mVisitOnce                                  As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mStripTableIndex = New PagedLongArray
ReDim mStripTable(1 To 256) As StripTableEntry
mMinUsedStripNumber = MaxLong
mMaxUsedStripNumber = MinLong
End Sub

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As Enumerator
Set Enumerable_Enumerator = Enumerator
End Property

Private Function Enumerable_NextItem(Item As Variant, cookie As Variant) As Boolean
Const ProcName As String = "Enumerable_NextItem"
On Error GoTo Err

Dim lCookie As GraphObjListCookie
Dim lIndex As Long

If IsEmpty(cookie) Then
    Enumerable_NextItem = False
Else
    lCookie = cookie
    If lCookie.Current < lCookie.From Then
        lCookie.Current = lCookie.From
    Else
        lCookie.Current = lCookie.Current + 1
    End If
        
    Item = Empty
    cookie = Empty
    Enumerable_NextItem = False
    
    Do While lCookie.Current <= lCookie.To
        lIndex = mStripTableIndex.GetEntry(lCookie.Current)
        If filterValidIndex(lIndex) Then
            Set Item = mStripTable(lIndex).GraphobjList
            cookie = lCookie
            Enumerable_NextItem = True
            Exit Do
        End If
        lCookie.Current = lCookie.Current + 1
    Loop
End If

Exit Function


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub Enumerable_Remove(cookie As Variant)
Const ProcName As String = "Enumerable_Remove"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, cookie As Variant) As Boolean
Const ProcName As String = "Enumerable_Skip"
On Error GoTo Err

Err.Raise ErrorCodes.ErrUnsupportedOperationException

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get Enumerable_UpdateSequence() As Long

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Enumerator( _
                Optional ByVal pFrom As Double = MinusInfinityDouble, _
                Optional ByVal pTo As Double = PlusInfinityDouble) As Enumerator
Dim lCookie As GraphObjListCookie

If pFrom = MinusInfinityDouble Then
    lCookie.From = mMinUsedStripNumber
Else
    lCookie.From = calcStripNumber(pFrom)
End If

If pTo = PlusInfinityDouble Then
    lCookie.To = mMaxUsedStripNumber
Else
    lCookie.To = calcStripNumber(pTo, True)
End If

lCookie.Current = MinLong

Set Enumerator = CreateEnumerator(Me, lCookie)
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function Add( _
                ByVal pIndex As Long, _
                ByVal pMinX As Double, _
                ByVal pMaxX As Double) As Variant
Const ProcName As String = "Add"
On Error GoTo Err

Dim i As Long
Dim lMembs As ListMemberships

calcStripNumbers lMembs, pMinX, pMaxX
Set lMembs.StripListMemberships = New LinkedList

addToLists pIndex, lMembs

Add = lMembs
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function Adjust( _
                ByVal pIndex As Long, _
                ByVal pMinX As Double, _
                ByVal pMaxX As Double, _
                ByVal pGraphObjListsCookie As Variant) As Variant
Const ProcName As String = "Adjust"
On Error GoTo Err

Dim lNewMembs As ListMemberships
Dim lPrevMembs As ListMemberships

If IsEmpty(pGraphObjListsCookie) Then
    Adjust = Add(pIndex, pMinX, pMaxX)
    Exit Function
End If

lPrevMembs = pGraphObjListsCookie

calcStripNumbers lNewMembs, pMinX, pMaxX

If lNewMembs.MinStrip = lPrevMembs.MinStrip And _
    lNewMembs.MaxStrip = lPrevMembs.MaxStrip _
Then
    ' object has not moved so keep previous membership info
    Adjust = lPrevMembs
Else
    gClearListMembership lPrevMembs.StripListMemberships
    Set lNewMembs.StripListMemberships = lPrevMembs.StripListMemberships
    
    addToLists pIndex, lNewMembs
    Adjust = lNewMembs
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub EndNonDuplicationCycle()
Const ProcName As String = "EndNonDuplicationCycle"
On Error GoTo Err

Assert mVisitOnce, "Not currently in non-duplication cycle"

mVisitOnce = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pStripWidth As Double, _
                ByVal pStripOrigin As Double)
mStripWidth = pStripWidth
mStripOrigin = pStripOrigin
End Sub

Friend Sub Remove( _
                ByVal pGraphObjListsCookie As Variant)
Const ProcName As String = "Remove"
On Error GoTo Err

Dim lMembs As ListMemberships

lMembs = pGraphObjListsCookie
gClearListMembership lMembs.StripListMemberships

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub StartNonDuplicationCycle()
Const ProcName As String = "StartNonDuplicationCycle"
On Error GoTo Err

Assert Not mVisitOnce, "Already in non-duplication cycle"

mCycle = mCycle + 1
mVisitOnce = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub addToLists( _
            ByVal pIndex As Long, _
            ByRef pMembs As ListMemberships)
Const ProcName As String = "addToLists"
On Error GoTo Err

Dim lStripIndex As Long
Dim i As Long

For i = pMembs.MinStrip To pMembs.MaxStrip
    lStripIndex = getStripTableEntryIndex(i)
    If mStripTable(lStripIndex).GraphobjList Is Nothing Then Set mStripTable(lStripIndex).GraphobjList = New LinkedList
    gAddListMembership pMembs.StripListMemberships, pIndex, mStripTable(lStripIndex).GraphobjList
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function allocateStrip() As Long
Const ProcName As String = "allocateStrip"
On Error GoTo Err

mStripTableMaxIndex = mStripTableMaxIndex + 1
If mStripTableMaxIndex > UBound(mStripTable) Then ReDim Preserve mStripTable(1 To 2 * UBound(mStripTable)) As StripTableEntry
allocateStrip = mStripTableMaxIndex

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function calcStripNumber( _
                ByVal pX As Double, _
                Optional pDecrementIfInteger As Boolean = False) As Long
Dim lStrip As Double
Const ProcName As String = "calcStripNumber"
On Error GoTo Err

lStrip = (pX - mStripOrigin) / mStripWidth
If pDecrementIfInteger Then
    If IsInteger(lStrip) Then
        calcStripNumber = Int(lStrip) - 1
    Else
        calcStripNumber = Int(lStrip)
    End If
Else
    calcStripNumber = Int(lStrip)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub calcStripNumbers( _
                ByRef pMembs As ListMemberships, _
                ByVal pMinX As Double, _
                ByVal pMaxX As Double)
Const ProcName As String = "calcStripNumbers"
On Error GoTo Err

pMembs.MinStrip = calcStripNumber(pMinX)
If pMembs.MinStrip < mMinUsedStripNumber Then mMinUsedStripNumber = pMembs.MinStrip

pMembs.MaxStrip = calcStripNumber(pMaxX, True)
If pMembs.MaxStrip > mMaxUsedStripNumber Then mMaxUsedStripNumber = pMembs.MaxStrip

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getStripTableEntryIndex(ByVal pStripNumber As Long) As Long
Const ProcName As String = "getStripTableEntryIndex"
On Error GoTo Err

getStripTableEntryIndex = mStripTableIndex.GetEntry(pStripNumber)
If getStripTableEntryIndex = 0 Then
    getStripTableEntryIndex = allocateStrip
    mStripTableIndex.SetEntry pStripNumber, getStripTableEntryIndex
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function filterValidIndex(ByVal pIndex As Long) As Boolean
Const ProcName As String = "filterValidIndex"
On Error GoTo Err

If pIndex = 0 Then
    filterValidIndex = False
ElseIf mStripTable(pIndex).GraphobjList Is Nothing Then
    filterValidIndex = False
ElseIf mStripTable(pIndex).GraphobjList.Count = 0 Then
    filterValidIndex = False
ElseIf mVisitOnce Then
    If mStripTable(pIndex).Cycle = mCycle Then
        filterValidIndex = False
    Else
        mStripTable(pIndex).Cycle = mCycle
        filterValidIndex = True
    End If
Else
    filterValidIndex = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function
