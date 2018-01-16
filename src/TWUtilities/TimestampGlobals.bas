Attribute VB_Name = "TimestampGlobals"
Option Explicit

''
' Description here
'
' @remarks
' @see
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


Private Const ModuleName                    As String = "TimestampGlobals"

Private Const OneHour                       As Double = 3600# * 1000000#

Private Const TicksPerDay                   As Currency = 86400000
Private Const TicksPerSec                   As Currency = 10000000

Public Const VbDateZero                     As Currency = 9435312000000#

'@================================================================================
' External Declarations
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mBaseTime As Double
Private mBaseTimeUTC As Double

Private mPerfFreq As Double
Private mStartPerfCounter As Currency

Private mLocalTimeZone As TimeZone

Private mDaysPerMonth(11) As Long
Private mDaysPerMonthLeap(11) As Long

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

Public Function gFileTimeToVbDate( _
                ByVal pFileTime As Currency) As Date
Dim days As Currency
Dim ticksSinceMidnight As Currency
Dim seconds As Double

Const ProcName As String = "gFileTimeToVbDate"

On Error GoTo Err

pFileTime = pFileTime - VbDateZero
days = Int(pFileTime / TicksPerDay)
ticksSinceMidnight = pFileTime - days * TicksPerDay
seconds = CDbl(ticksSinceMidnight * 10000) / TicksPerSec

gFileTimeToVbDate = CDbl(days) + seconds / 86400

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
                
End Function

Public Function gFormatTimestamp( _
                ByVal Timestamp As Date, _
                Optional ByVal formatOption As TimestampFormats = TimestampDateAndTime, _
                Optional ByVal formatString As String = "yyyymmddhhnnss", _
                Optional ByVal tz As TimeZone) As String
Dim timestampDays As Long
Dim timestampSecs As Double
Dim timestampAsDate As Date
Dim Milliseconds As Long
Dim includeTimezone As Boolean
Dim noMillisecs As Boolean

'includeTimezone = formatOption And TimestampFormats.TimestampIncludeTimezone

'formatOption = formatOption And (Not TimestampFormats.TimestampIncludeTimezone)

Const ProcName As String = "gFormatTimestamp"

On Error GoTo Err

noMillisecs = formatOption And TimestampFormats.TimestampNoMillisecs

formatOption = formatOption And (Not TimestampFormats.TimestampNoMillisecs)

timestampDays = Int(Timestamp)
timestampSecs = Int((Timestamp - Int(Timestamp)) * 86400) / 86400#
timestampAsDate = CDate(CDbl(timestampDays) + timestampSecs)
Milliseconds = CLng((Timestamp - timestampAsDate) * 86400# * 1000#)

If Milliseconds >= 1000& Then
    Milliseconds = Milliseconds - 1000&
    timestampSecs = timestampSecs + (1# / 86400#)
    timestampAsDate = CDate(CDbl(timestampDays) + timestampSecs)
End If

If noMillisecs Then
    ' round to the nearest second
    timestampSecs = timestampSecs + (Milliseconds / (86400# * 1000#))
    timestampAsDate = CDate(CDbl(timestampDays) + timestampSecs)
End If

Select Case formatOption
Case TimestampFormats.TimestampTimeOnly
    gFormatTimestamp = Format(timestampAsDate, "hhnnss") & _
                        IIf(noMillisecs, "", Format(Milliseconds, "\.000"))
Case TimestampFormats.TimestampDateOnly
    gFormatTimestamp = Format(timestampAsDate, "yyyymmdd")
Case TimestampFormats.TimestampDateAndTime
    gFormatTimestamp = Format(timestampAsDate, "yyyymmddhhnnss") & _
                        IIf(noMillisecs, "", Format(Milliseconds, "\.000"))
Case TimestampFormats.TimestampTimeOnlyISO8601
    gFormatTimestamp = Format(timestampAsDate, "hh:nn:ss") & _
                        IIf(noMillisecs, "", Format(Milliseconds, "\.000"))
Case TimestampFormats.TimestampDateOnlyISO8601
    gFormatTimestamp = Format(timestampAsDate, "yyyy-mm-dd")
Case TimestampFormats.TimestampDateAndTimeISO8601
    gFormatTimestamp = Format(timestampAsDate, "yyyy-mm-dd hh:nn:ss") & _
                        IIf(noMillisecs, "", Format(Milliseconds, "\.000"))
Case TimestampFormats.TimestampTimeOnlyLocal
    gFormatTimestamp = FormatDateTime(timestampAsDate, vbLongTime) & _
                        IIf(noMillisecs, "", Format(Milliseconds, "\.000"))
Case TimestampFormats.TimestampDateOnlyLocal
    gFormatTimestamp = FormatDateTime(timestampAsDate, vbShortDate)
Case TimestampFormats.TimestampDateAndTimeLocal
    gFormatTimestamp = FormatDateTime(timestampAsDate, vbShortDate) & " " & _
                        FormatDateTime(timestampAsDate, vbLongTime) & _
                        IIf(noMillisecs, "", Format(Milliseconds, "\.000"))
Case TimestampFormats.TimestampCustom
    gFormatTimestamp = Format(timestampAsDate, formatString) & _
                        IIf(noMillisecs, "", Format(Milliseconds, "\.000"))
End Select

If includeTimezone Then
    ' to be implemented !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function gGetTimestamp() As Date
Dim et As Double

Const ProcName As String = "gGetTimestamp"

On Error GoTo Err

et = ElapsedTimeMicroseconds
gGetTimestamp = mBaseTime + et / 86400000000#

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetTimestampUtc() As Date
Dim et As Double
Const ProcName As String = "gGetTimestampUtc"

On Error GoTo Err

et = ElapsedTimeMicroseconds
gGetTimestampUtc = mBaseTimeUTC + et / 86400000000#

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gInit()
Const ProcName As String = "gInit"
On Error GoTo Err

initDaysInMonth

getLocalTimeZone

gSetBaseTimes

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gLocalToUtc( _
                ByVal localDate As Date) As Date
Const ProcName As String = "gLocalToUtc"

On Error GoTo Err

gLocalToUtc = mLocalTimeZone.ConvertDateTzToUTC(localDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gNthWeekdayOfMonth( _
                ByRef monthFirstDay As SYSTEMTIME, _
                ByVal dayOfWeek As Long, _
                ByVal n As Long) As SYSTEMTIME

Dim st As SYSTEMTIME
Dim filetime As Currency

Const ProcName As String = "gNthWeekdayOfMonth"

On Error GoTo Err

st = monthFirstDay

SystemTimeToFileTime st, filetime
FileTimeToSystemTime filetime, st   ' now st.DayOfWeek is set, so we know what day of the
                                    ' week the first day of the month is
st.Day = (((7 + dayOfWeek - st.dayOfWeek) Mod 7) + 1) + 7 * (n - 1)
If isLeapYear(st.Year) Then
    If st.Day > mDaysPerMonthLeap(st.Month - 1) Then st.Day = st.Day - 7
Else
    If st.Day > mDaysPerMonth(st.Month - 1) Then st.Day = st.Day - 7
End If

gNthWeekdayOfMonth = st

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Sub gSetBaseTimes()
Dim firstSysTime As Currency
Dim newSysTime As Currency
Dim newSysTimeLocal As Currency
Dim lPerfFreq As Currency

Const ProcName As String = "gSetBaseTimes"

On Error GoTo Err

QueryPerformanceFrequency lPerfFreq
mPerfFreq = lPerfFreq

QueryPerformanceCounter mStartPerfCounter
GetSystemTimeAsFileTime firstSysTime
Do
    GetSystemTimeAsFileTime newSysTime
Loop Until newSysTime <> firstSysTime
'Debug.Print "Microseconds to get base time: " & format(ElapsedTimeMicroseconds, "0.000")

' start measuring elapsed time from here
QueryPerformanceCounter mStartPerfCounter

' newSysTime is now our base time in UTC
mBaseTimeUTC = gFileTimeToVbDate(newSysTime)

FileTimeToLocalFileTime newSysTime, newSysTimeLocal
mBaseTime = gFileTimeToVbDate(newSysTimeLocal)

'Debug.Print "Microseconds to convert base time to VB format: " & format(ElapsedTimeMicroseconds, "0.000")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gSystemTimeToVbDate( _
                ByRef sysTime As SYSTEMTIME) As Date
Dim filetime As Currency
Const ProcName As String = "gSystemTimeToVbDate"

On Error GoTo Err

SystemTimeToFileTime sysTime, filetime
gSystemTimeToVbDate = gFileTimeToVbDate(filetime)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gTerm()
Const ProcName As String = "gTerm"

On Error GoTo Err

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gUtcToLocal( _
                ByVal utcDate As Date) As Date
Const ProcName As String = "gUtcToLocal"

On Error GoTo Err

gUtcToLocal = mLocalTimeZone.ConvertDateUTCToTZ(utcDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gVbDateToFileTime( _
                pDate As Date) As Currency

' we round this to an integral value because the Windows APIs FileTimeToSystemTime
' and SystemTimeToFileTime work at the millisecond level. Not doing so causes some
' times to be incorrectly converted between timezones - eg 16:31 ends up as 16:30:59.999.
Const ProcName As String = "gVbDateToFileTime"

On Error GoTo Err

gVbDateToFileTime = Int(pDate * TicksPerDay + VbDateZero + 0.4999)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function calcElapsedTime() As Double
Dim perfCounter As Currency
Dim diff As Currency

QueryPerformanceCounter perfCounter
diff = perfCounter - mStartPerfCounter
calcElapsedTime = (1000000# * CDbl(diff)) / mPerfFreq
End Function

Private Function ElapsedTimeMicroseconds() As Double
Static lastEt As Double

Const ProcName As String = "ElapsedTimeMicroseconds"

On Error GoTo Err

If mBaseTime = 0# Or lastEt >= OneHour Then gSetBaseTimes

ElapsedTimeMicroseconds = calcElapsedTime

If ElapsedTimeMicroseconds < lastEt Then
    ' this happens on WinXP VMs under HyperV, where the counter is
    ' reset every 20 minutes
    gSetBaseTimes
    ElapsedTimeMicroseconds = calcElapsedTime
End If

lastEt = ElapsedTimeMicroseconds

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub getLocalTimeZone()
Const ProcName As String = "getLocalTimeZone"

On Error GoTo Err

Set mLocalTimeZone = gGetTimeZone("")

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initDaysInMonth()
mDaysPerMonth(0) = 31
mDaysPerMonth(1) = 28
mDaysPerMonth(2) = 31
mDaysPerMonth(3) = 30
mDaysPerMonth(4) = 31
mDaysPerMonth(5) = 30
mDaysPerMonth(6) = 31
mDaysPerMonth(7) = 31
mDaysPerMonth(8) = 30
mDaysPerMonth(9) = 31
mDaysPerMonth(10) = 30
mDaysPerMonth(11) = 31

mDaysPerMonthLeap(0) = 31
mDaysPerMonthLeap(1) = 29
mDaysPerMonthLeap(2) = 31
mDaysPerMonthLeap(3) = 30
mDaysPerMonthLeap(4) = 31
mDaysPerMonthLeap(5) = 30
mDaysPerMonthLeap(6) = 31
mDaysPerMonthLeap(7) = 31
mDaysPerMonthLeap(8) = 30
mDaysPerMonthLeap(9) = 31
mDaysPerMonthLeap(10) = 30
mDaysPerMonthLeap(11) = 31
End Sub

Private Function isLeapYear( _
                ByVal pYear As Long) As Boolean
Const ProcName As String = "isLeapYear"

On Error GoTo Err

If (pYear Mod 4) <> 0 Then
ElseIf (pYear Mod 400) = 0 Then
    isLeapYear = True
ElseIf (pYear Mod 100) = 0 Then
Else
    isLeapYear = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function




