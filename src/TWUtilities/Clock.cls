VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'@================================================================================
' Description
'@================================================================================

''
' The Clock class represents a clock running in a specified time zone.
'
' A <code>Clock</code> object fires a Tick event each second, on the second. It has
' a property that enables very accurate timestamps in the relevant time zone to be
' obtained.
'
' There are two types of <code>Clock</code> object, realtime and simulated-time:
' <DIR>
' <LI>A realtime
' <code>Clock</code> is driven by the computer's internal clock, and its
' accuracy depends entirely on how well the computer's clock is synchronized with
' repect to actual time. There is a single global realtime <code>Clock</code> object
' for each timezone known to the system.</LI>
' <LI>A simulated-time <code>Clock</code> is driven by the application, ie the application
' supplies the current time. This is useful for applications which are purporting to
' run at some time in the past or future, and (optionally) where the passage of time
' is accelerated (ie speeded up or slowed down) by some factor. Timestamps supplied
' by a simulated-time <code>Clock</code> are calculated by multiplying the real
' elapsed time since the last-notified simulated-time by the acceleration factor, and
' adding that to the last-notified simulated-time. Note that the frequency of the
' <code>Tick</code> event for simulated-time <code>Clock</code>s is once every real
' second (ie the ticks are not speeded up in accordance with the clock rate).</LI>
' </DIR>
'
' To access a realtime <code>Clock</code> object, use the <code>GetClock</code> global
' method.
'
' To create a simulated-time <code>Clock</code> object, use the <code>CreateSimulatedClock</code>
' global method.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IJSONable
Implements IStringable

'@================================================================================
' Events
'@================================================================================

''
' Fired at the start of each (realtime) second.
'@/
Event Tick()

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "Clock"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mSimulateTime                       As Boolean
Private mTimePaused                         As Boolean
Private mTimestamp                          As Date
Private mReferenceTimeUTC                   As Date
Private mRate                               As Single

Private mTimeZone                           As TimeZone

Private mTimeHasBeenSet                     As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' JSONable Interface Members
'@================================================================================

Private Function IJSONable_ToJSON() As String
Const ProcName As String = "IJSONable_ToJSON"

On Error GoTo Err

IJSONable_ToJSON = ToString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Stringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
Const ProcName As String = "IStringable_ToString"

On Error GoTo Err

IStringable_ToString = ToString
IStringable_ToString = "Timezone=" & gVariantToString(mTimeZone) & _
                    "; Simulated=" & mSimulateTime & _
                    IIf(mSimulateTime, "; Paused=" & mTimePaused & "; Rate=" & mRate, "") & _
                    "; Curr time=" & gFormatTimestamp(Timestamp, TimestampDateAndTimeISO8601)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

''
' Indicates whether this is a simulated-time <code>Clock.</code>
'
' @return Returns <code>True</code> for a simulated-time <code>Clock</code>, and <code>False</code> otherwise.
'
'@/
Public Property Get IsSimulated() As Boolean
IsSimulated = mSimulateTime
End Property

''
' For a simulated-time <code>Clock</code>, specifies the acceleration factor that is applied
' to the passage of time.
'
' @remarks
'   This property is only valid for simulated-time <code>Clock</code> objects.
'   Attempting to set it for a realtime <code>Clock</code> results in error
'   <code>ErrUnsupportedOperationException</code>.
'
'   The value may be changed at any time.
'
' @param Value
'   Indicates the factor by which simulated-time is to be accelerated for
'   this <code>Clock</code> object. Values greater than 1 speed up time.
'   Values less than 1 slow down time. A zero value is used to indicate that the
'   passage of real time is not to be taken into account when calculating timestamps:
'   in this case, the <code>timestamp</code> property always returns the most recent
'   value passed in the <code>setTime</code> method.
'
'   Negative values are not permitted and cause an <code>ErrIllegalArgumentException</code>
'   error.
'
'@/
Public Property Let Rate( _
                ByVal Value As Single)
Const ProcName As String = "rate"
On Error GoTo Err

gAssert mSimulateTime, "Can't set clock rate for non-simulated clock"
gAssertArgument Value >= 0, "Value cannot be negative."

If mRate = 0 Then
    mRate = Value
    SetTime mTimestamp
Else
    mRate = Value
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' For a simulated <code>Clock</code>, returns the acceleration factor that is applied
' to the passage of time.
'
' @remarks
'   This property is only valid for simulated-time <code>Clock</code> objects.
'   For realtime <code>Clock</code> objects the value is undefined.
'
' @return
'   Returns the factor by which simulated-time is accelerated for
'   this <code>Clock</code> object.
'
'@/
Public Property Get Rate() As Single
Rate = mRate
End Property

''
' Returns an accurate Timestamp for this <code>Clock</code>, in the relevant time zone.
'
' @remarks
'   If this is a simulated-time <code>Clock</code>, the value is calculated by multiplying the
'   elapsed time since the most recent call to the <code>setTime</code> method by the acceleration
'   factor specified in the <code>rate</code> property and adding the result to the
'   time specified in the most recent call to <code>setTime</code>.
'
' @return
'   The current time for this <code>Clock</code>, in the relevant time zone.
' @param Name description
' @see
'
'@/
Public Property Get Timestamp() As Date
Const ProcName As String = "Timestamp"
On Error GoTo Err

If mSimulateTime Then
    gAssert mTimeHasBeenSet, "Time has not been set for this clock"
    If mRate = 0 Then
        Timestamp = mTimestamp
    ElseIf mTimePaused Then
        Timestamp = mTimeZone.ConvertDateUTCToTZ(mTimestamp)
    Else
        Timestamp = mTimeZone.ConvertDateUTCToTZ(mTimestamp + mRate * (gGetTimestampUtc - mReferenceTimeUTC))
    End If
Else
    Timestamp = mTimeZone.ConvertDateUTCToTZ(gGetTimestampUtc)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns an accurate UTC Timestamp for this <code>Clock</code>.
'
' @remarks
'   The returned value is the same as that obtained by calling <code>ConvertDateTzToUTC</code>
'   on the value returned by the <code>Timestamp</code> property, but this property is more efficient.
'
' @return
'   The current UTC time for this <code>Clock</code>.
' @param Name description
' @see
'
'@/
Public Property Get TimestampUTC() As Date
Const ProcName As String = "Timestamp"
On Error GoTo Err

If mSimulateTime Then
    gAssert mTimeHasBeenSet, "Time has not been set for this clock"
    If mRate = 0 Then
        TimestampUTC = mTimeZone.ConvertDateTzToUTC(mTimestamp)
    ElseIf mTimePaused Then
        TimestampUTC = mTimestamp
    Else
        TimestampUTC = mTimestamp + mRate * (gGetTimestampUtc - mReferenceTimeUTC)
    End If
Else
    TimestampUTC = gGetTimestampUtc
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the<code>TimeZone</code> object for this<code>Clock</code>.
'
' @remarks
' @return
'   The<code>TimeZone</code> object for the time zone in which this clock runs.
'
'@/
Public Property Get TimeZone() As TimeZone
Set TimeZone = mTimeZone
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                 ByVal simulateTime As Boolean, _
                 ByVal Rate As Single, _
                 ByVal pTimezonename As String)
Const ProcName As String = "Initialise"

On Error GoTo Err

mSimulateTime = simulateTime
mRate = Rate

Set mTimeZone = GTimeZone.gGetTimeZone(pTimezonename)

If mSimulateTime Then
    mTimestamp = MinDateValue
    mReferenceTimeUTC = gGetTimestampUtc
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

''
' Suspends the passage of time (simulated-time clocks only).
'
' @remarks
'   After calling <code>pauseTime</code>, the <code>timestamp</code> property
'   will return the same value until <code>resumeTime</code> is called.
'
'   A call to this method for a realtime <code>Clock</code> results in error
'   <code>ErrUnsupportedOperationException</code>.
'@/
Public Sub PauseTime()
Const ProcName As String = "PauseTime"
On Error GoTo Err

gAssert mSimulateTime, "Clock is not simulated"

If mRate = 0 Then Exit Sub

mTimestamp = TimestampUTC
mTimePaused = True

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Resumes the passage of time (simulated-time clocks only).
'
' @remarks
'   A call to this method for a realtime <code>Clock</code> results in error
'   <code>ErrUnsupportedOperationException</code>.
'@/
Public Sub ResumeTime()
Const ProcName As String = "ResumeTime"
On Error GoTo Err

gAssert mSimulateTime, "Clock is not simulated"

If mRate = 0 Then Exit Sub

mReferenceTimeUTC = gGetTimestampUtc
mTimePaused = False

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Sets the current time (simulated-time clocks only), and raises a <code>Tick</code> event.
'
' @remarks
'   A call to this method for a realtime <code>Clock</code> results in error
'   <code>ErrUnsupportedOperationException</code>.
' @param pTimestamp
'   The current time for this <code>Clock</code>. The time supplied must be in
'   the <code>Clock</code>'s time zone
' @see
'
'@/
Public Sub SetTime(ByVal pTimestamp As Date)
Const ProcName As String = "SetTime"
On Error GoTo Err

gAssert mSimulateTime, "Clock is not simulated"
gAssertArgument pTimestamp >= MinDateValue And pTimestamp <= MaxDateValue, "Invalid date"

If mRate > 0 Then
    mTimestamp = mTimeZone.ConvertDateTzToUTC(pTimestamp)
    mReferenceTimeUTC = gGetTimestampUtc
Else
    mTimestamp = pTimestamp ' don't convert to UTC for efficiency
End If

mTimeHasBeenSet = True

RaiseEvent Tick

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Sets the current UTC time (simulated-time clocks only), and raises a <code>Tick</code> event.
'
' @remarks
'   A call to this method for a realtime <code>Clock</code> results in error
'   <code>ErrUnsupportedOperationException</code>.
' @param pTimestampUTC
'   The current time for this <code>Clock</code>. The time supplied must be in UTC
' @see
'
'@/
Public Sub SetTimeUTC(ByVal pTimestampUTC As Date)
Const ProcName As String = "SetTimeUTC"
On Error GoTo Err

gAssert mSimulateTime, "Clock is not simulated"
gAssertArgument pTimestampUTC >= MinDateValue And pTimestampUTC <= MaxDateValue, "Invalid date"

If mRate > 0 Then
    mTimestamp = pTimestampUTC
    mReferenceTimeUTC = gGetTimestampUtc
Else
    mTimestamp = mTimeZone.ConvertDateUTCToTZ(pTimestampUTC)
End If

mTimeHasBeenSet = True

RaiseEvent Tick

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Tick()
RaiseEvent Tick
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function ToString() As String
Dim jb As JSONBuilder
Const ProcName As String = "ToString"

On Error GoTo Err

Set jb = New JSONBuilder

jb.BeginObject
jb.EncodeName "Timezone"
jb.EncodeValue mTimeZone

jb.EncodeName "Simulated"
jb.EncodeValue mSimulateTime
If mSimulateTime Then
    jb.EncodeName "Paused"
    jb.EncodeValue mTimePaused
    jb.EncodeName "Rate"
    jb.EncodeValue mRate
End If
jb.EncodeName "Curr time"
jb.EncodeValue gFormatTimestamp(Timestamp, TimestampDateAndTimeISO8601)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

