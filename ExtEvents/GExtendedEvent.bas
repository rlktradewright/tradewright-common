Attribute VB_Name = "GExtendedEvent"
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

Private Const ModuleName                            As String = "GExtendedEvent"

'@================================================================================
' Member variables
'@================================================================================

Private mEventListenerCollections                   As New Collection

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

'@================================================================================
' Methods
'@================================================================================

Public Sub gAddListener( _
                ByVal pSource As IExtendedEventsSource, _
                ByVal pListener As IExtendedEventListener, _
                ByVal pExtendedEvent As ExtendedEvent, _
                ByVal pNotifyHandledEvents As Boolean)
Dim lEntry As ExtendedEventListenerEntry
Set lEntry.Listener = pListener
lEntry.NotifyHandledEvents = pNotifyHandledEvents
addListenerCollectionForSource(addListenerCollectionsForEvent(pExtendedEvent), pSource).Add lEntry, Hex$(ObjPtr(pListener))
End Sub

Public Sub gFire( _
                ByVal pSource As IExtendedEventsSource, _
                ByVal pValue As Variant, _
                ByVal pExtendedEvent As ExtendedEvent)
Dim ev As ExtendedEventData
Dim lListenerCollections As Collection
Dim lListeners As Collection
Dim var As Variant
Dim lEntry As ExtendedEventListenerEntry

On Error Resume Next

Set ev.ExtendedEvent = pExtendedEvent
Set ev.Source = pSource
Set ev.OriginalSource = pSource
If IsObject(pValue) Then
    Set ev.Data = pValue
Else
    ev.Data = pValue
End If

Set lListenerCollections = getListenerCollectionsForEvent(pExtendedEvent)
If lListenerCollections Is Nothing Then Exit Sub

Set lListeners = getListenerCollectionForSource(lListenerCollections, pSource)
If Not lListeners Is Nothing Then
    For Each var In lListeners
        lEntry = var
        lEntry.Listener.Notify ev
    Next
End If

If Not pSource.Parent Is Nothing Then fire ev, pSource.Parent
End Sub

Public Sub gRemoveListener( _
                ByVal pSource As IExtendedEventsSource, _
                ByVal pListener As IExtendedEventListener, _
                ByVal pExtendedEvent As ExtendedEvent)
Dim lListenerCollections As Collection
Dim lListeners As Collection

On Error Resume Next

Set lListenerCollections = getListenerCollectionsForEvent(pExtendedEvent)
If lListenerCollections Is Nothing Then Exit Sub

Set lListeners = getListenerCollectionForSource(lListenerCollections, pSource)
If Not lListeners Is Nothing Then
    lListeners.Remove Hex$(ObjPtr(pListener))
    If lListeners.Count = 0 Then mEventListenerCollections.Remove Hex$(ObjPtr(pSource))
End If

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function addListenerCollectionsForEvent( _
                ByVal pExtendedEvent As ExtendedEvent) As Collection
Set addListenerCollectionsForEvent = getListenerCollectionsForEvent(pExtendedEvent)
If addListenerCollectionsForEvent Is Nothing Then
    Set addListenerCollectionsForEvent = New Collection
    mEventListenerCollections.Add addListenerCollectionsForEvent, Hex$(ObjPtr(pExtendedEvent))
End If
End Function

Private Function addListenerCollectionForSource( _
                ByVal pListenerCollectionsForEvent As Collection, _
                ByVal pSource As IExtendedEventsSource) As Collection
Set addListenerCollectionForSource = getListenerCollectionForSource(pListenerCollectionsForEvent, pSource)
If addListenerCollectionForSource Is Nothing Then
    ' we need to include the source object as well as the new collection
    ' in the entry, as otherwise this collection could be used later for a different
    ' object if there is no call to RemoveExtendedEventListener when the original
    ' source object is disposed (this is because we use ObjPtr as the key, and there is
    ' nothing to prevent a new object being allocated at the same address as one that has
    ' been disposed). Keeping the source object in the entry means that the source object
    ' will not disappear unless RemoveExtendedEventListener is called.
    Dim lEntry As ExtendedEventListenerCollectionEntry
    Set lEntry.Collection = New Collection
    Set lEntry.Source = pSource
    Set addListenerCollectionForSource = lEntry.Collection
    pListenerCollectionsForEvent.Add lEntry, Hex$(ObjPtr(pSource))
End If
End Function

Private Sub fire( _
                ByRef ev As ExtendedEventData, _
                ByVal pSource As IExtendedEventsSource)
Dim lListenerCollections As Collection
Dim lListeners As Collection
Dim var As Variant
Dim lEntry As ExtendedEventListenerEntry

On Error Resume Next

Set ev.Source = pSource

Set lListenerCollections = getListenerCollectionsForEvent(ev.ExtendedEvent)
If lListenerCollections Is Nothing Then Exit Sub

Set lListeners = getListenerCollectionForSource(lListenerCollections, ev.Source)
If lListeners Is Nothing Then Exit Sub

For Each var In lListeners
    lEntry = var
    If Not ev.Handled Or lEntry.NotifyHandledEvents Then lEntry.Listener.Notify ev
Next

If Not pSource.Parent Is Nothing Then fire ev, pSource.Parent

End Sub

Private Function getListenerCollectionsForEvent( _
                ByVal pExtendedEvent As ExtendedEvent) As Collection
On Error Resume Next
Set getListenerCollectionsForEvent = mEventListenerCollections(Hex$(ObjPtr(pExtendedEvent)))
End Function

Private Function getListenerCollectionForSource( _
                ByVal pListenerCollectionsForEvent As Collection, _
                ByVal pSource As IExtendedEventsSource) As Collection
On Error Resume Next
Dim lEntry As ExtendedEventListenerCollectionEntry
lEntry = pListenerCollectionsForEvent(Hex$(ObjPtr(pSource)))
Set getListenerCollectionForSource = lEntry.Collection
End Function



