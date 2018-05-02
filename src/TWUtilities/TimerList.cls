VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimerList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Maintains a list of TimerListEntry objects and raises events when they expire."
Option Explicit

''
' This class maintains a list of <code>TimerListItem</code> objects and raises
' events when they expire.
'
' There are two types of <code>TimerListItem</code> object:
' <ul>
'   <li><em>Realtime timer lists</em>:</li> these are driven by the computer's
'   realtime clock;
'   <li><em>Simulated-time timer lists</em>:</li> these are driven by timestamps
'   supplied by the application, and enable historical time or speeded-up time
'   to be simulated.
' </ul>
'
' There is a single global realtime timer list object per process. It is created
' automatically and can be accessed using the <code>GetGlobalTimerList</code>
' global method.
'
' Any number of simulated-time timer list objects may be created using the
' <code>CreateSimulatedTimerList</code> global method.
'
' To determine when a <code>TimerListItem</code> has expired, either handle the
' <code>TimerListItem</code>'s <code>StateChange</code> event, or register a
' listener object using the <code>TimerListItem</code>'s
' <code>addStateChangeListener</code> method.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "TimerList"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private WithEvents mTimer As IntervalTimer
Attribute mTimer.VB_VarHelpID = -1
Private mList As SortedDictionary

Private mIsSimulated As Boolean

Private WithEvents mSimulatedClock As Clock
Attribute mSimulatedClock.VB_VarHelpID = -1
Private mClock As Clock
Attribute mClock.VB_VarHelpID = -1

Private mNextExpiry As Date

Private mItemsToNotify() As TimerListItem
Private mItemsToNotifyIndex As Long

'Private mIsGlobal As Boolean

Private mInhibitReschedule As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mList = New SortedDictionary
mList.Initialise KeyTypeDate, True
Set mTimer = New IntervalTimer
mNextExpiry = MaxDateValue
ReDim mItemsToNotify(15) As TimerListItem
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' mSimulatedClock Event Handlers
'@================================================================================

Private Sub mSimulatedClock_Tick()
Const ProcName As String = "mSimulatedClock_Tick"
On Error GoTo Err

SetTime mSimulatedClock.Timestamp

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' mTimer Event Handlers
'@================================================================================

Private Sub mTimer_TimerExpired(ev As TimerExpiredEventData)
Const ProcName As String = "mTimer_TimerExpired"
On Error GoTo Err

Dim Timestamp As Date

mNextExpiry = MaxDateValue
Timestamp = gGetTimestampUtc
CheckTimer Timestamp
reschedule

Exit Sub

Err:
gNotifyUnhandledError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Public Property Get Clock() As Clock
Set Clock = mClock
End Property

Public Property Get IsSimulated() As Boolean
IsSimulated = mIsSimulated
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Creates a new <code>TimerListItem</code> object.
'
' @remarks
' @return
'   The<code>TimerListItem</code> object that has been created.
' @param pData
'   This can be any pData that can be held in a Variant.
' @param pExpiryTime
'   The time at which the new<code>TimerListItem</code> is to expire. The values
'   allowed depend on the value of the<code>units</code> argument.
'
'   If <code>units</code> is <code>ExpiryTimeUnits.ExpiryTimeUnitDateTime</code>, then
'   the value can be any expression that yields a valid date. For a realtime timerlist,
'   this date must be in the future (interpreted as local time). For a simulated-time
'   timer list, this date must be later than the most recent time supplied via the
'   <code>setTime</code> method.
'
'   If <code>units</code> has any other value, this argument must be a positive numeric
'   value representing the amount of time (in the relevant units) to expiry.
' @param pUnits
'   Specifies the pUnits in which the<code>expiryTime</code> argument is expressed.
'@/
Public Function Add( _
                ByVal pData As Variant, _
                ByVal pExpiryTime As Variant, _
                Optional ByVal pUnits As ExpiryTimeUnits = ExpiryTimeUnits.ExpiryTimeUnitSeconds) As TimerListItem
Attribute Add.VB_Description = "Creates a new TimerListItem object."
Const ProcName As String = "Add"
On Error GoTo Err

Dim lExpiryDate As Date
Dim lEntry As New TimerListItem

If pUnits = ExpiryTimeUnits.ExpiryTimeUnitDateTime Then
    If Not IsDate(pExpiryTime) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "pExpiryTime is not a valid date"
    If CDate(pExpiryTime) <= IIf(mIsSimulated, mClock.Timestamp, gGetTimestamp) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "lExpiry time is before present"

    
    If mIsSimulated Then
        lExpiryDate = CDate(pExpiryTime)
    Else
        lExpiryDate = gLocalToUtc(CDate(pExpiryTime))
    End If
Else
    If Not IsNumeric(pExpiryTime) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "pExpiryTime is not a valid number"
    
    Dim lExpiry As Double
    lExpiry = CDbl(pExpiryTime)
    
    If lExpiry <= 0 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "pExpiryTime must not be negative or zero"

    Dim lExpiryIntervalMillisecs As Long
    If pUnits = ExpiryTimeUnits.ExpiryTimeUnitMilliseconds Then
        lExpiryIntervalMillisecs = lExpiry
    ElseIf pUnits = ExpiryTimeUnits.ExpiryTimeUnitSeconds Then
        lExpiryIntervalMillisecs = lExpiry * 1000
    ElseIf pUnits = ExpiryTimeUnits.ExpiryTimeUnitMinutes Then
        lExpiryIntervalMillisecs = lExpiry * 60 * 1000
    ElseIf pUnits = ExpiryTimeUnits.ExpiryTimeUnitHours Then
        lExpiryIntervalMillisecs = lExpiry * 60 * 60 * 1000
    ElseIf pUnits = ExpiryTimeUnits.ExpiryTimeUnitDays Then
        lExpiryIntervalMillisecs = lExpiry * 24 * 60 * 60 * 1000
    Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, , "pUnits argument invalid"
    End If
    
    If mIsSimulated Then
        lExpiryDate = mClock.Timestamp + (lExpiryIntervalMillisecs / 86400000)
    Else
        lExpiryDate = gGetTimestampUtc + (lExpiryIntervalMillisecs / 86400000)
    End If
End If

lEntry.Initialise lExpiryDate, pData, Me

'Debug.Print "add lEntry: " & gFormatTimestamp(lExpiryDate, TimestampTimeOnly)
mList.Add lEntry, lExpiryDate
reschedule

Set Add = lEntry

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

''
' Removes all items from this <code>TimerList</code>.
'
' @remarks
'   All items have their status set to
'   <code>TimerListItemStateCancelled</code>, and are prevented from expiring.
'@/
Public Sub Clear()
Attribute Clear.VB_Description = "Removes all TimerListItems from the TimerList."
Const ProcName As String = "Clear"
On Error GoTo Err

If mList.Count = 0 Then Exit Sub
    
Dim en As Enumerator
Set en = mList.Enumerator

Do While en.MoveNext
    Dim tli As TimerListItem
    Set tli = en.Current
    tli.State = TimerListItemStateCancelled
    en.Remove
Loop

reschedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pIsSimulated As Boolean, _
                ByVal pClock As Clock)
mIsSimulated = pIsSimulated
Set mClock = pClock
If mIsSimulated Then Set mSimulatedClock = mClock
End Sub

''
' Removes the specified <code>TimerListItem</code> object from the list.
'
' @remarks
'   Removing an item from the list causes its status to be set to
'   <code>TimerListItemStateCancelled</code>, and prevents it from expiring.
' @param tli
'   The <code>TimerListItem</code> object to be removed.
'@/
Public Sub Remove( _
                ByVal tli As TimerListItem)
Attribute Remove.VB_Description = "Removes a TimerListItem object from the list."
' Mark the specified TimerListItem object as disabled. We
' don't Remove it from the list here because this sub may
' be called while we are in checktimer, which would damage the
' tree structure
Const ProcName As String = "Remove"
On Error GoTo Err

tli.State = TimerListItemStateCancelled
reschedule

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub CheckTimer(ByVal currentTime As Date)
Const ProcName As String = "CheckTimer"
On Error GoTo Err

If mList.Count = 0 Then Exit Sub

mInhibitReschedule = True

mItemsToNotifyIndex = 0

Dim en As Enumerator
Set en = mList.Enumerator
Do While en.MoveNext
    Dim currItem As TimerListItem
    Set currItem = en.Current
    If currItem.ExpiryTime > currentTime Then Exit Do
    If currItem.State = TimerListItemStatePending Then
        ' we can't notify the item at this point because we are iterating the list, and
        ' the item's owner might add another item which would corrupt the list. So add it
        ' to a collection and notify it when we've Finished iterating the list
        If mItemsToNotifyIndex > UBound(mItemsToNotify) Then
            ReDim Preserve mItemsToNotify(2 * (UBound(mItemsToNotify) + 1) - 1) As TimerListItem
            If gLogger.IsLoggable(LogLevelHighDetail) Then _
                gLogger.Log "Increased mItemsToNotify size", ProcName, ModuleName, LogLevelHighDetail, CStr(UBound(mItemsToNotify) + 1)
        End If
        Set mItemsToNotify(mItemsToNotifyIndex) = currItem
        mItemsToNotifyIndex = mItemsToNotifyIndex + 1
    End If
    en.Remove
Loop

Set currItem = Nothing

' now notify expired items
If mItemsToNotifyIndex > 0 Then
    Dim i As Long
    For i = 0 To mItemsToNotifyIndex - 1
        mItemsToNotify(i).Notify
        Set mItemsToNotify(i) = Nothing
    Next
End If

mInhibitReschedule = False

Exit Sub

Err:
mInhibitReschedule = False
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub reschedule()
Const ProcName As String = "reschedule"
On Error GoTo Err

If mIsSimulated Then Exit Sub
If mInhibitReschedule Then Exit Sub

If mList.Count = 0 Then
    mNextExpiry = MaxDateValue
    mTimer.StopTimer
Else
    Dim en As Enumerator
    Set en = mList.Enumerator
    
    Dim firstExpiry As Date
    Do While en.MoveNext
        Dim timerItem As TimerListItem
        Set timerItem = en.Current
        If timerItem.State = TimerListItemStatePending Then
            firstExpiry = timerItem.ExpiryTime
            Exit Do
        End If
        en.Remove
    Loop
    
    If mList.Count = 0 Then
        mNextExpiry = MaxDateValue
        mTimer.StopTimer
        Exit Sub
    End If
    
    If firstExpiry = mNextExpiry And _
        mTimer.State = TimerRunning _
    Then
        Exit Sub
    End If
    
    mTimer.StopTimer
    
    mNextExpiry = firstExpiry
    
    Dim interval As Long
    interval = Int((mNextExpiry - gGetTimestampUtc) * 86400 * 1000 + 0.9999)
    If interval <= 0 Then interval = 1
        ' this could happen if we've just gone past the
        ' expiry time for the first item in the list
    
    mTimer.Initialise interval, 0, False, Empty
    mTimer.StartTimer
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Notifies the <code>TimerList</code> object of the current time.
'
' The <code>TimerList</code> object uses this to determine if any
' <code>TimerListItem</code> objects have expired.
'
' @remarks
'   Calling this method on a realtime <code>TimerList</code> object causes
'   error <code>ErrorCodes.ErrIllegalStateException</code> to be raised.
' @param theTime
'   The time to be set.
'@/
Private Sub SetTime(ByVal theTime As Date)
Const ProcName As String = "SetTime"
On Error GoTo Err

CheckTimer theTime

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub


