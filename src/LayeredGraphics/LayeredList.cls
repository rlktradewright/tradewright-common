VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LayeredList"
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

Implements IEnumerable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type LayerTableEntry
    List        As LinkedList
    Cycle       As Long
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "LayeredList"

'@================================================================================
' Member variables
'@================================================================================

Private mLayerTable(LayerNumbers.LayerMin To LayerNumbers.LayerMax) As LayerTableEntry

Private mUpdateSequence         As Long

Private mCycle                  As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get IEnumerable_Enumerator() As Enumerator
Set IEnumerable_Enumerator = Enumerator
End Property

Private Function IEnumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean
Dim lCookie As LayeredListCookie

Const ProcName As String = "IEnumerable_NextItem"
On Error GoTo Err

If IsEmpty(pCookie) Then
    IEnumerable_NextItem = False
Else
    lCookie = pCookie
    getCookie lCookie
    If lCookie.En Is Nothing Then
        IEnumerable_NextItem = False
    ElseIf lCookie.En.MoveNext Then
        IEnumerable_NextItem = True
    ElseIf getNextCookie(lCookie) Then
        IEnumerable_NextItem = lCookie.En.MoveNext
    Else
        IEnumerable_NextItem = False
    End If
End If

If IEnumerable_NextItem Then
    Item = lCookie.En.Current
    pCookie = lCookie
Else
    Item = Empty
    pCookie = Empty
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub getCookie(ByRef pCookie As LayeredListCookie)
If pCookie.En Is Nothing Then getNextCookie pCookie
End Sub

Private Sub IEnumerable_Remove(cookie As Variant)
Dim lCookie As LayeredListCookie
Const ProcName As String = "IEnumerable_Remove"
On Error GoTo Err

lCookie = cookie
lCookie.En.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function IEnumerable_Skip(ByVal numToSkip As Long, cookie As Variant) As Boolean

End Function

Private Property Get IEnumerable_UpdateSequence() As Long
IEnumerable_UpdateSequence = mUpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get Enumerator( _
                Optional ByVal pStartLayer As LayerNumbers = LayerNumbers.LayerMin, _
                Optional ByVal pStopLayer As LayerNumbers = LayerNumbers.LayerMax, _
                Optional ByVal pUpLayers As Boolean = True) As Enumerator
Dim lCookie As LayeredListCookie

If pStartLayer < LayerNumbers.LayerMin Then pStartLayer = LayerNumbers.LayerMin
If pStartLayer > LayerNumbers.LayerMax Then pStartLayer = LayerNumbers.LayerMax
If pStopLayer < LayerNumbers.LayerMin Then pStopLayer = LayerNumbers.LayerMin
If pStopLayer > LayerNumbers.LayerMax Then pStopLayer = LayerNumbers.LayerMax

With lCookie
    .CurrLayer = incrementLayer(pStartLayer, IIf(pUpLayers, -1, 1))
    .StopLayer = pStopLayer
    .UpLayers = pUpLayers
End With

Set Enumerator = CreateEnumerator(Me, lCookie)
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Function AddItem( _
                ByVal pItem As Long, _
                ByVal pLayer As LayerNumbers) As Long
If mLayerTable(pLayer).Cycle <> mCycle Then
    mLayerTable(pLayer).Cycle = mCycle
    If Not mLayerTable(pLayer).List Is Nothing Then mLayerTable(pLayer).List.Clear
End If
If mLayerTable(pLayer).List Is Nothing Then Set mLayerTable(pLayer).List = New LinkedList
AddItem = mLayerTable(pLayer).List.AddItem(pItem)
End Function

Friend Function ReInitialise()
mCycle = mCycle + 1
End Function

Friend Sub RemoveItem( _
                ByVal pLayer As LayerNumbers, _
                ByVal pItemIndex As Long)
If pItemIndex = NullIndex Then Exit Sub
Assert mLayerTable(pLayer).Cycle = mCycle And Not mLayerTable(pLayer).List Is Nothing, "Layer is empty"

mLayerTable(pLayer).List.RemoveEntry pItemIndex
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getNextCookie( _
                ByRef pCookie As LayeredListCookie) As Boolean
Const ProcName As String = "getNextCookie"
On Error GoTo Err

Set pCookie.En = Nothing
Do While getNextLayer(pCookie)
    If Not mLayerTable(pCookie.CurrLayer).List Is Nothing Then
        If mLayerTable(pCookie.CurrLayer).Cycle <> mCycle Then
            mLayerTable(pCookie.CurrLayer).Cycle = mCycle
            mLayerTable(pCookie.CurrLayer).List.Clear
        ElseIf mLayerTable(pCookie.CurrLayer).List.Count <> 0 Then
            Set pCookie.En = mLayerTable(pCookie.CurrLayer).List.Enumerator
            Exit Do
        End If
    End If
Loop

getNextCookie = Not (pCookie.En Is Nothing)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

'Private Function getNextLayer(ByRef pCookie As LayeredListCookie) As Boolean
'If pCookie.CurrLayer = -1 Then
'    pCookie.CurrLayer = pCookie.StartLayer
'    getNextLayer = True
'ElseIf pCookie.UpLayers Then
'    pCookie.CurrLayer = pCookie.CurrLayer + 1
'    If pCookie.CurrLayer = pCookie.StopLayer Then
'        getNextLayer = False
'    Else
'        If pCookie.CurrLayer > LayerNumbers.LayerMax Then pCookie.CurrLayer = LayerNumbers.LayerMin
'        getNextLayer = True
'    End If
'Else
'    pCookie.CurrLayer = pCookie.CurrLayer - 1
'    If pCookie.CurrLayer = pCookie.StopLayer Then
'        getNextLayer = False
'    Else
'        If pCookie.CurrLayer < LayerNumbers.LayerMin Then pCookie.CurrLayer = LayerNumbers.LayerMax
'        getNextLayer = True
'    End If
'End If
'End Function

Private Function getNextLayer(ByRef pCookie As LayeredListCookie) As Boolean
If pCookie.UpLayers Then
    pCookie.CurrLayer = incrementLayer(pCookie.CurrLayer, 1)
    getNextLayer = (pCookie.CurrLayer <= pCookie.StopLayer)
    If pCookie.CurrLayer > LayerNumbers.LayerMax Then pCookie.CurrLayer = LayerNumbers.LayerMin
Else
    pCookie.CurrLayer = incrementLayer(pCookie.CurrLayer, -1)
    getNextLayer = (pCookie.CurrLayer >= pCookie.StopLayer)
    If pCookie.CurrLayer < LayerNumbers.LayerMin Then pCookie.CurrLayer = LayerNumbers.LayerMax
End If
End Function

Public Function incrementLayer(ByVal Value As LayerNumbers, ByVal pIncrement As Integer) As LayerNumbers
Dim newVal As Long
newVal = Value + pIncrement
If newVal = LayerNumbers.LayerInvisible Then newVal = newVal + Sgn(pIncrement)
incrementLayer = newVal
End Function


