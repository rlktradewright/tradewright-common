VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LayeredGraphicsModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'================================================================================
' Interfaces
'================================================================================

Implements ITask

'================================================================================
' Events
'================================================================================

'================================================================================
' Enums
'================================================================================

'================================================================================
' Constants
'================================================================================


Private Const ModuleName                As String = "Model"

'================================================================================
' Types
'================================================================================

Private Type GraphicObjectTableEntry
    GraphObj            As IGraphicObject
    GraphObjListsCookie As Variant
    InScopeListIndex    As Long
    VisibleListIndex    As Long
    Cycle               As Long
End Type

'================================================================================
' Member variables
'================================================================================

Private mControllerRef As WeakReference

Private mGraphicObjects()                   As GraphicObjectTableEntry
Private mGraphicObjectsIndex                As Long

Private mNumberOfObjects                    As Long

Private mGraphObjSeriesCollection           As GraphObjSeriesCollection

Private mInScopeList                        As LinkedList

Private mVisibleList                        As LayeredList

Private mTaskContext                        As TaskContext

Private mGraphObjCycle                      As Long

Private mGraphObjLists                      As GraphObjLists

'================================================================================
' Class Event Handlers
'================================================================================

Private Sub Class_Initialize()

ReDim mGraphicObjects(1023) As GraphicObjectTableEntry

Set mInScopeList = New LinkedList

'Set mExtendedList = New LinkedList

Set mVisibleList = New LayeredList

Set mGraphObjSeriesCollection = New GraphObjSeriesCollection
mGraphObjSeriesCollection.Initialise Me


End Sub

Private Sub Class_Terminate()
Debug.Print "LayeredGraphicsModel terminated"
End Sub

'================================================================================
' Task Interface Members
'================================================================================

Private Sub ITask_cancel()
mTaskContext.Finish Empty, True
End Sub

Private Sub ITask_run()
Const ProcName As String = "ITask_run"
On Error GoTo Err

Static i As Long
Dim j As Long

If i <= mGraphicObjectsIndex Then
    If Not mGraphicObjects(i).GraphObj Is Nothing Then
        mGraphicObjects(i).GraphObj.Finish
        Set mGraphicObjects(i).GraphObj = Nothing
        mNumberOfObjects = mNumberOfObjects - 1
    End If
    i = i + 1
Else
    Erase mGraphicObjects

    mTaskContext.Finish 0, False
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Private Property Let ITask_TaskContext(ByVal RHS As TaskContext)
Set mTaskContext = RHS
End Property

Private Property Get ITask_TaskName() As String
ITask_TaskName = mTaskContext.Name
End Property

'================================================================================
' Properties
'================================================================================

Friend Property Get NumberOfObjects() As Long
NumberOfObjects = mNumberOfObjects
End Property

'================================================================================
' Methods
'================================================================================

Public Function AddGraphicObjectSeries( _
                ByVal graphObjSeries As IGraphicObjectSeries, _
                Optional ByVal Layer As LayerNumbers = -1, _
                Optional ByVal Name As String) As IGraphicObjectSeries
Const ProcName As String = "AddGraphicObjectSeries"
On Error GoTo Err

Set AddGraphicObjectSeries = mGraphObjSeriesCollection.Add( _
                                            graphObjSeries, _
                                            GetNextUserLayer(Layer), _
                                            Name, _
                                            controllerObj)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function AddGraphicObject( _
                ByVal pGraphObj As IGraphicObject) As Long
Const ProcName As String = "AddGraphicObject"
On Error GoTo Err

Dim lHandle As Long

lHandle = allocateGraphicObjectTableEntry

pGraphObj.Handle = lHandle

Set mGraphicObjects(lHandle).GraphObj = pGraphObj

Dim lLayer As Long
lLayer = pGraphObj.Layer

Dim lBoundingRect As TRectangle
Dim lScopeRect As TRectangle
lScopeRect = controllerObj.GetScopeRect
lBoundingRect = pGraphObj.BoundingRectangleInArea(lScopeRect)
If lBoundingRect.isValid Then mGraphicObjects(lHandle).GraphObjListsCookie = mGraphObjLists.Add(lHandle, lBoundingRect.Left, lBoundingRect.Right)

If graphicObjectIsInScope(lHandle, lScopeRect) Then
    addEntryToInScopeList lHandle
    If graphicObjectIsVisible(lHandle, lScopeRect) Then
        addEntryToVisibleList lHandle, lLayer
    End If
End If
mNumberOfObjects = mNumberOfObjects + 1
AddGraphicObject = lHandle

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub CalculateInscopeGaugeDependentObjects( _
                ByRef pScopeRect As TRectangle)
Const ProcName As String = "CalculateInscopeGaugeDependentObjects"
On Error GoTo Err

Dim En As Enumerator
Set En = mInScopeList.Enumerator
Do While En.MoveNext
    Dim lIndex As Long
    lIndex = CLng(En.Current)
    With mGraphicObjects(lIndex)
        If .GraphObj.GaugeDependent Then
            If graphicObjectIsVisible(lIndex, pScopeRect) Then
                addEntryToVisibleList lIndex, .GraphObj.Layer
            End If
        End If
    End With
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub CalcVerticalBoundsForGaugeDependentObjects( _
                ByRef pScopeRect As TRectangle, _
                ByRef pScaleMin As Double, _
                ByRef pScaleMax As Double)
Const ProcName As String = "CalcVerticalBoundsForGaugeDependentObjects"
On Error GoTo Err

Dim En As Enumerator
Set En = mInScopeList.Enumerator

Do While En.MoveNext
    Dim lIndex As Long
    lIndex = CLng(En.Current)
    With mGraphicObjects(lIndex)
        If Not .GraphObj.NoDraw And _
            .GraphObj.IncludeInAutoscale And _
            .GraphObj.GaugeDependent _
        Then
            Dim objMax As Double
            Dim objMin As Double
            calculateObjectHighAndLowInScope lIndex, _
                                            pScopeRect, _
                                            objMax, _
                                            objMin
            If objMax > pScaleMax Then
                pScaleMax = objMax
            End If
            If objMin < pScaleMin Then
                pScaleMin = objMin
            End If
        End If
    End With
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function CalcVerticalBoundsForNonGaugeDependentObjects( _
                ByRef pScopeRect As TRectangle, _
                ByRef pScaleMin As Double, _
                ByRef pScaleMax As Double) As Boolean
Const ProcName As String = "CalcVerticalBoundsForNonGaugeDependentObjects"
On Error GoTo Err

Dim En As Enumerator
Set En = mInScopeList.Enumerator

Do While En.MoveNext
    Dim lIndex As Long
    lIndex = CLng(En.Current)
    With mGraphicObjects(lIndex)
        If Not .GraphObj.NoDraw And _
            .GraphObj.IncludeInAutoscale _
        Then
            If .GraphObj.GaugeDependent Then
                CalcVerticalBoundsForNonGaugeDependentObjects = True
            Else

                Dim objMax As Double
                Dim objMin As Double
                calculateObjectHighAndLowInScope lIndex, _
                                                        pScopeRect, _
                                                        objMax, _
                                                        objMin
                If objMax > pScaleMax Then
                    pScaleMax = objMax
                End If
                If objMin < pScaleMin Then
                    pScaleMin = objMin
                End If
            End If
        End If
    End With
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Clear()
Const ProcName As String = "Clear"
On Error GoTo Err

StartTask Me, PriorityLow

mGraphObjSeriesCollection.Finish

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function ClearGraphicObjectSeries( _
                ByVal pGraphObjSeries As IGraphicObjectSeries) As TaskController
Const ProcName As String = "ClearGraphicObjectSeries"
On Error GoTo Err

Set ClearGraphicObjectSeries = mGraphObjSeriesCollection.Clear(pGraphObjSeries)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub DetermineInScopeObjects( _
                ByRef pScopeRect As TRectangle)
Const ProcName As String = "DetermineInScopeObjects"
On Error GoTo Err

Dim numberExamined As Long
Dim numberInScope As Long

Dim et As New ElapsedTimer
et.StartTiming

clearInScopeObjectList

Dim lGraphObjListEn As Enumerator
Set lGraphObjListEn = mGraphObjLists.Enumerator(pScopeRect.Left, pScopeRect.Right)

mGraphObjCycle = mGraphObjCycle + 1

Do While lGraphObjListEn.MoveNext
    Dim lGraphObjList As LinkedList
    Set lGraphObjList = lGraphObjListEn.Current
    
    Dim lGraphObjEn As Enumerator
    Set lGraphObjEn = lGraphObjList.Enumerator
    Do While lGraphObjEn.MoveNext
        Dim lIndex As Long
        lIndex = CLng(lGraphObjEn.Current)
        With mGraphicObjects(lIndex)
            If .Cycle <> mGraphObjCycle Then
                .Cycle = mGraphObjCycle
                numberExamined = numberExamined + 1
                If graphicObjectIsInScope(lIndex, pScopeRect) Then
                    addEntryToInScopeList lIndex
                    numberInScope = numberInScope + 1
                End If
            End If
        End With
    Loop
Loop

Debug.Print "DetermineInScopeObjects: total=" & numberExamined & " time=" & et.ElapsedTimeMicroseconds / 1000 & "ms"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub DetermineVisibleObjects( _
                ByRef pScopeRect As TRectangle)
Const ProcName As String = "DetermineVisibleObjects"
On Error GoTo Err

Dim numberExamined As Long
Dim numberVisible As Long

Dim et As New ElapsedTimer
et.StartTiming

clearVisibleObjectList

Dim En As Enumerator
Set En = mInScopeList.Enumerator
Do While En.MoveNext
    Dim lIndex As Long
    lIndex = CLng(En.Current)
    numberExamined = numberExamined + 1
    If graphicObjectIsVisible(lIndex, pScopeRect) Then
        addEntryToVisibleList lIndex, mGraphicObjects(lIndex).GraphObj.Layer
        numberVisible = numberVisible + 1
    End If
Loop

Debug.Print "DetermineVisibleObjects: total=" & numberExamined & " visible=" & numberVisible & " time=" & et.ElapsedTimeMicroseconds / 1000 & "ms"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function GetFirstVisibleObjectAtPoint( _
                ByVal X As Single, _
                ByVal Y As Single) As IGraphicObject
Const ProcName As String = "GetFirstVisibleObjectAtPoint"
On Error GoTo Err

Dim En As Enumerator
Set En = mVisibleList.Enumerator(MaxLayer, , False)

Do While En.MoveNext
    Dim lIndex As Long
    lIndex = CLng(En.Current)
    With mGraphicObjects(lIndex)
        If .GraphObj.HitTest(X, Y) Then
            Set GetFirstVisibleObjectAtPoint = .GraphObj
            ' we've found an object/series to select
            
            ' NB we can't use Exit For, because we are
            ' within a With statement, and it causes the table lock
            ' not to be released - see http://support.microsoft.com/kb/187553
            GoTo ExitFor
        End If
    End With
Loop

ExitFor:

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetGraphicObject( _
                ByVal pHandle As Long) As IGraphicObject
Const ProcName As String = "GetGraphicObject"
On Error GoTo Err

Set GetGraphicObject = mGraphicObjects(pHandle).GraphObj

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetGraphicObjectSeries( _
                ByVal pSeriesID As Long) As IGraphicObjectSeries
Set GetGraphicObjectSeries = mGraphObjSeriesCollection.Item(pSeriesID)
End Function

Friend Function GetNextSelectableObjectOrSeriesAtPoint( _
                ByVal currObject As IGraphicObject, _
                ByVal X As Single, _
                ByVal Y As Single) As IGraphicObject
Const ProcName As String = "GetNextSelectableObjectOrSeriesAtPoint"
On Error GoTo Err

Dim currObjStillSelectable As Boolean

Dim En As Enumerator
Set En = mVisibleList.Enumerator(IIf(currObject Is Nothing, MaxLayer, currObject.Layer), , False)
Do While En.MoveNext
    Dim lIndex As Long
    lIndex = CLng(En.Current)
    With mGraphicObjects(lIndex)
        If isObjectOrSeriesSelectable(.GraphObj) Then
            If .GraphObj.HitTest(X, Y) Then
                If .GraphObj Is currObject Then
                    currObjStillSelectable = True
                Else
                    Set GetNextSelectableObjectOrSeriesAtPoint = .GraphObj
                    ' we've found an object/series to select
                    
                    ' NB we can't use Exit For, because we are
                    ' within a With statement, and it causes the table lock
                    ' not to be released - see http://support.microsoft.com/kb/187553
                    GoTo ExitFor
                    
                End If
            End If
        End If
    End With
Loop

ExitFor:

If GetNextSelectableObjectOrSeriesAtPoint Is Nothing And _
    currObjStillSelectable _
Then
    Set GetNextSelectableObjectOrSeriesAtPoint = currObject
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetNextUserLayer(Optional ByVal pLayer As Long = -1) As Long
Const ProcName As String = "GetNextUserLayer"
On Error GoTo Err

Static sNextLayer As Long

AssertArgument pLayer = -1 Or (pLayer >= LayerMin And pLayer <= LayerMax), "Value is not a valid layer number"

If sNextLayer = 0 Then sNextLayer = LayerNumbers.LayerLowestUser

If pLayer = -1 Then
    GetNextUserLayer = sNextLayer
    If sNextLayer <> LayerNumbers.LayerHighestUser Then
        sNextLayer = sNextLayer + 1
    End If
Else
    GetNextUserLayer = pLayer
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetAllVisibleObjects() As IEnumerable
Const ProcName As String = "GetAllVisibleObjects"
On Error GoTo Err

Set GetAllVisibleObjects = mVisibleList

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function GetObjectsInAreas( _
                ByRef areasToRedraw() As TRectangle, _
                ByVal pCount As Long, _
                ByVal pLowestLayer As Long) As IEnumerable
Const ProcName As String = "GetObjectsInAreas"
On Error GoTo Err

Static sResultList As LayeredList

If sResultList Is Nothing Then
    Set sResultList = New LayeredList
Else
    sResultList.ReInitialise
End If

If gLogger.IsLoggable(LogLevelHighDetail) Then
    Static et As ElapsedTimer
    Static sCount As Long
    Static sTime As Single
    If et Is Nothing Then Set et = New ElapsedTimer
    et.StartTiming
End If

mGraphObjCycle = mGraphObjCycle + 1
mGraphObjLists.StartNonDuplicationCycle

Dim i As Long
For i = 0 To pCount - 1
    If areasToRedraw(i).isValid Then
        Dim lGraphObjListEn As Enumerator
        Set lGraphObjListEn = mGraphObjLists.Enumerator(areasToRedraw(i).Left, areasToRedraw(i).Right)
        Do While lGraphObjListEn.MoveNext
            Dim lGraphObjList As LinkedList
            Set lGraphObjList = lGraphObjListEn.Current
            
            Dim lGraphObjEn As Enumerator
            Set lGraphObjEn = lGraphObjList.Enumerator
            Do While lGraphObjEn.MoveNext
                Dim lIndex As Long
                lIndex = CLng(lGraphObjEn.Current)
                With mGraphicObjects(lIndex)
                    If .Cycle <> mGraphObjCycle Then
                        .Cycle = mGraphObjCycle
                        If .GraphObj.Layer >= pLowestLayer Then
                            Dim j As Long
                            For j = 0 To pCount - 1
                                If .GraphObj.BoundingRectangleInArea(areasToRedraw(j)).isValid Then
                                    sResultList.AddItem lIndex, .GraphObj.Layer
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                End With
            Loop
        Loop
    End If
Next

mGraphObjLists.EndNonDuplicationCycle

Set GetObjectsInAreas = sResultList

If gLogger.IsLoggable(LogLevelHighDetail) Then
    sTime = sTime + et.ElapsedTimeMicroseconds
    sCount = sCount + 1
    If sCount = 10 Then
        gLogger.Log "Average time: ", ProcName, ModuleName, LogLevelHighDetail, sTime / sCount
        sCount = 0
        sTime = 0
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pController As Controller, _
                ByVal pStripWidth As Double, _
                ByVal pStripOrigin As Double)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mControllerRef = CreateWeakReference(pController)

Set mGraphObjLists = New GraphObjLists
mGraphObjLists.Initialise pStripWidth, pStripOrigin

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RecalcBoundaryForObjectsInSeries( _
                ByVal SeriesID As Long)
Const ProcName As String = "RecalcBoundaryForObjectsInSeries"
On Error GoTo Err

Dim En As Enumerator
Set En = mVisibleList.Enumerator
Do While En.MoveNext
    Dim lIndex As Long
    lIndex = CLng(En.Current)
    With mGraphicObjects(lIndex)
        If .GraphObj.SeriesID = SeriesID Then
            Dim boundingRect As TRectangle
            ' force recalculation of the bounding GDI_RECT
            boundingRect = .GraphObj.BoundingRectangle
        End If
    End With
Loop

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveGraphicObject( _
                ByVal pGraphObj As IGraphicObject)
Const ProcName As String = "RemoveGraphicObject"
On Error GoTo Err

RemoveGraphicObjectEx pGraphObj, True, True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveGraphicObjectEx( _
                ByVal pGraphObj As IGraphicObject, _
                ByVal undraw As Boolean, _
                ByVal removeFromSeries As Boolean)
Const ProcName As String = "RemoveGraphicObjectEx"
On Error GoTo Err

If pGraphObj.Handle < 0 Then Exit Sub    ' check if object has already been removed from model

If Not controllerObj.SelectedObject Is Nothing Then
    If pGraphObj Is controllerObj.SelectedObject Then controllerObj.ClearSelection
End If

If removeFromSeries And pGraphObj.SeriesID <> 0 Then mGraphObjSeriesCollection.Item(pGraphObj.SeriesID).RemoveObject pGraphObj

If undraw Then controllerObj.UndrawObject pGraphObj

removeGraphObj pGraphObj

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveGraphicObjectSeries( _
                ByVal pSeries As IGraphicObjectSeries)
Const ProcName As String = "RemoveGraphicObjectSeries"
On Error GoTo Err

AssertArgument Not pSeries Is Nothing, "The supplied argument is Nothing"

mGraphObjSeriesCollection.Remove pSeries

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub UpdateGraphicObject( _
                ByVal pGraphObj As IGraphicObject)
Const ProcName As String = "UpdateGraphicObject"
On Error GoTo Err

Dim lHandle As Long
lHandle = pGraphObj.Handle

' check if object has been removed from the model
If lHandle < 0 Then Exit Sub

Dim lLayer As Long
lLayer = pGraphObj.Layer

Dim lScopeRect As TRectangle
lScopeRect = controllerObj.GetScopeRect

Dim lBoundingRect As TRectangle
lBoundingRect = pGraphObj.BoundingRectangleInArea(lScopeRect)
If lBoundingRect.isValid Then mGraphicObjects(lHandle).GraphObjListsCookie = mGraphObjLists.Adjust(lHandle, lBoundingRect.Left, lBoundingRect.Right, mGraphicObjects(lHandle).GraphObjListsCookie)

If graphicObjectIsInScope(lHandle, lScopeRect) Then
    addEntryToInScopeList lHandle
    If graphicObjectIsVisible(lHandle, lScopeRect) Then
        addEntryToVisibleList lHandle, lLayer
    Else
        removeEntryFromVisibleList lHandle, lLayer
    End If
Else
    removeEntryFromInScopeList lHandle
    removeEntryFromVisibleList lHandle, lLayer
End If

controllerObj.OnObjectChanged pGraphObj
    
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'================================================================================
' Helper Functions
'================================================================================

Private Sub addEntryToInScopeList(ByVal pIndex As Long)
Const ProcName As String = "addEntryToInScopeList"
On Error GoTo Err

If mGraphicObjects(pIndex).InScopeListIndex <> NullIndex Then Exit Sub

mGraphicObjects(pIndex).InScopeListIndex = mInScopeList.AddItem(pIndex)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub addEntryToVisibleList(ByVal pIndex As Long, ByVal pLayer As Long)
Const ProcName As String = "addEntryToVisibleList"
On Error GoTo Err

If mGraphicObjects(pIndex).VisibleListIndex <> NullIndex Then Exit Sub

mGraphicObjects(pIndex).VisibleListIndex = mVisibleList.AddItem(pIndex, pLayer)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function allocateGraphicObjectTableEntry() As Long
Const ProcName As String = "allocateGraphicObjectTableEntry"
On Error GoTo Err

mGraphicObjectsIndex = mGraphicObjectsIndex + 1
allocateGraphicObjectTableEntry = mGraphicObjectsIndex

If mGraphicObjectsIndex > UBound(mGraphicObjects) Then
    ReDim Preserve mGraphicObjects(UBound(mGraphicObjects) + 500) As GraphicObjectTableEntry
    If gLogger.IsLoggable(LogLevelHighDetail) Then _
        gLogger.Log "Increased mGraphicObjects size to", ProcName, ModuleName, LogLevelHighDetail, CStr(UBound(mGraphicObjects) + 1)
End If

mGraphicObjects(allocateGraphicObjectTableEntry).InScopeListIndex = NullIndex
mGraphicObjects(allocateGraphicObjectTableEntry).VisibleListIndex = NullIndex
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub calculateObjectHighAndLowInScope( _
                ByVal pIndex As Long, _
                ByRef pScopeRect As TRectangle, _
                ByRef high As Double, _
                ByRef low As Double)
Const ProcName As String = "calculateObjectHighAndLowInScope"
On Error GoTo Err

high = MinusInfinityDouble
low = PlusInfinityDouble

With mGraphicObjects(pIndex)
    Dim verticalRange As TInterval
    verticalRange = .GraphObj.VerticalRangeInXinterval(TRectangleGetXInterval(pScopeRect))
    If verticalRange.isValid Then
        high = verticalRange.endValue
        low = verticalRange.startValue
    End If
End With

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearInScopeObjectList()
Const ProcName As String = "clearInScopeObjectList"
On Error GoTo Err

Dim En As Enumerator
Set En = mInScopeList.Enumerator
Do While En.MoveNext
    Dim lIndex As Long
    lIndex = CLng(En.Current)
    mGraphicObjects(lIndex).InScopeListIndex = NullIndex
    En.Remove
Loop

mInScopeList.Clear

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearVisibleObjectList()
Const ProcName As String = "clearVisibleObjectList"
On Error GoTo Err

Dim En As Enumerator
Set En = mVisibleList.Enumerator
Do While En.MoveNext
    Dim lIndex As Long
    lIndex = CLng(En.Current)
    mGraphicObjects(lIndex).GraphObj.Visible = False
    mGraphicObjects(lIndex).VisibleListIndex = NullIndex
    En.Remove
Loop

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

Private Function graphicObjectIsInScope( _
                ByVal pIndex As Long, _
                ByRef pScopeRect As TRectangle) As Boolean
Const ProcName As String = "graphicObjectIsInScope"
On Error GoTo Err

With mGraphicObjects(pIndex)
    If Not .GraphObj.BoundingRectangle.isValid Then Exit Function
    graphicObjectIsInScope = TIntervalOverlaps(TRectangleGetXInterval(pScopeRect), TRectangleGetXInterval(.GraphObj.BoundingRectangle))
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function graphicObjectIsVisible( _
                ByVal pIndex As Long, _
                ByRef pScopeRect As TRectangle) As Boolean
Const ProcName As String = "graphicObjectIsVisible"
On Error GoTo Err

With mGraphicObjects(pIndex).GraphObj
    If Not .BoundingRectangle.isValid Or .NoDraw Then Exit Function
    graphicObjectIsVisible = TRectangleOverlaps(.BoundingRectangleInArea(pScopeRect), pScopeRect)
End With

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isObjectOrSeriesSelectable(ByVal pGraphObj As IGraphicObject) As Boolean
Const ProcName As String = "isObjectOrSeriesSelectable"
On Error GoTo Err

If pGraphObj.IsSelectable Then
    isObjectOrSeriesSelectable = True
ElseIf mGraphObjSeriesCollection.Item(pGraphObj.SeriesID).IsSelectable Then
    isObjectOrSeriesSelectable = True
Else
    isObjectOrSeriesSelectable = False
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub removeEntryFromInScopeList( _
                ByVal pIndex As Long)
Const ProcName As String = "removeEntryFromInScopeList"
On Error GoTo Err

mInScopeList.RemoveEntry mGraphicObjects(pIndex).InScopeListIndex
mGraphicObjects(pIndex).InScopeListIndex = NullIndex
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeEntryFromVisibleList( _
                ByVal pIndex As Long, _
                ByVal pLayer As Long)
Const ProcName As String = "removeEntryFromVisibleList"
On Error GoTo Err


mVisibleList.RemoveItem pLayer, mGraphicObjects(pIndex).VisibleListIndex
mGraphicObjects(pIndex).VisibleListIndex = NullIndex
mGraphicObjects(pIndex).GraphObj.Visible = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub removeGraphObj( _
                ByVal pGraphObj As IGraphicObject)
Const ProcName As String = "removeGraphObj"
On Error GoTo Err

Dim lHandle As Long
lHandle = pGraphObj.Handle

mGraphObjLists.Remove mGraphicObjects(lHandle).GraphObjListsCookie
removeEntryFromInScopeList lHandle
removeEntryFromVisibleList lHandle, mGraphicObjects(lHandle).GraphObj.Layer

Set mGraphicObjects(lHandle).GraphObj = Nothing
pGraphObj.Handle = -1
pGraphObj.Finish

mNumberOfObjects = mNumberOfObjects - 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub






