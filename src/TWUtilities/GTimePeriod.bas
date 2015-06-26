Attribute VB_Name = "GTimePeriod"
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

Private Const ProjectName                   As String = "TimeframeUtils26"
Private Const ModuleName                    As String = "GTimePeriod"

Public Const TimePeriodNameSecond As String = "Second"
Public Const TimePeriodNameMinute As String = "Minute"
Public Const TimePeriodNameHour As String = "Hourly"
Public Const TimePeriodNameDay As String = "Daily"
Public Const TimePeriodNameWeek As String = "Weekly"
Public Const TimePeriodNameMonth As String = "Monthly"
Public Const TimePeriodNameYear As String = "Yearly"

Public Const TimePeriodNameSeconds As String = "Seconds"
Public Const TimePeriodNameMinutes As String = "Minutes"
Public Const TimePeriodNameHours As String = "Hours"
Public Const TimePeriodNameDays As String = "Days"
Public Const TimePeriodNameWeeks As String = "Weeks"
Public Const TimePeriodNameMonths As String = "Months"
Public Const TimePeriodNameYears As String = "Years"
Public Const TimePeriodNameVolumeIncrement As String = "Volume"
Public Const TimePeriodNameTickVolumeIncrement As String = "Tick Volume"
Public Const TimePeriodNameTickIncrement As String = "Ticks Movement"

Public Const TimePeriodShortNameSeconds As String = "s"
Public Const TimePeriodShortNameMinutes As String = "m"
Public Const TimePeriodShortNameHours As String = "h"
Public Const TimePeriodShortNameDays As String = "D"
Public Const TimePeriodShortNameWeeks As String = "W"
Public Const TimePeriodShortNameMonths As String = "M"
Public Const TimePeriodShortNameYears As String = "Y"
Public Const TimePeriodShortNameVolumeIncrement As String = "V"
Public Const TimePeriodShortNameTickVolumeIncrement As String = "TV"
Public Const TimePeriodShortNameTickIncrement As String = "T"

'@================================================================================
' Member variables
'@================================================================================

Private mTimePeriods                        As New Collection

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

Public Function gCalcMonthStartDate( _
                ByVal pDate As Date) As Date
Const ProcName As String = "gCalcMonthStartDate"

On Error GoTo Err

gCalcMonthStartDate = gCalcMonthStartDateFromMonthNumber(Month(pDate), pDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalcMonthStartDateFromMonthNumber( _
                ByVal monthNumber As Long, _
                ByVal baseDate As Date) As Date
Dim yearStart As Date

Const ProcName As String = "gCalcMonthStartDateFromMonthNumber"

On Error GoTo Err

yearStart = DateAdd("d", 1 - DatePart("y", baseDate), baseDate)
gCalcMonthStartDateFromMonthNumber = DateAdd("m", monthNumber - 1, yearStart)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalcWeekStartDate( _
                ByVal pDate As Date) As Date
Const ProcName As String = "gCalcWeekStartDate"

On Error GoTo Err

Dim theDate As Long
Dim weekNum As Long

theDate = Int(CDbl(pDate))
weekNum = DatePart("ww", theDate, vbMonday, vbFirstFullWeek)
If weekNum >= 52 And Month(theDate) = 1 Then
    ' this must be part of the final week of the previous year
    theDate = DateAdd("yyyy", -1, theDate)
End If
gCalcWeekStartDate = gCalcWeekStartDateFromWeekNumber(weekNum, theDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalcWeekStartDateFromWeekNumber( _
                ByVal weekNumber As Long, _
                ByVal baseDate As Date) As Date
Dim yearStart As Date
Dim week1Date As Date
Dim dow1 As Long    ' day of week of 1st jan of base year

Const ProcName As String = "gCalcWeekStartDateFromWeekNumber"

On Error GoTo Err

yearStart = DateAdd("d", 1 - DatePart("y", baseDate), baseDate)

dow1 = DatePart("w", yearStart, vbMonday)

If dow1 = 1 Then
    week1Date = yearStart
Else
    week1Date = DateAdd("d", 8 - dow1, yearStart)
End If

gCalcWeekStartDateFromWeekNumber = DateAdd("ww", weekNumber - 1, week1Date)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalcWorkingDayDate( _
                ByVal dayNumber As Long, _
                ByVal baseDate As Date) As Date
Dim yearStart As Date
Dim yearEnd As Date
Dim doy As Long

Dim wd1 As Long     ' weekdays in first week (excluding weekend)
Dim we1 As Long     ' weekend days at start of first week

Dim dow1 As Long    ' day of week of 1st jan of base year

' number of whole weeks after the first week
Dim numWholeWeeks As Long

Const ProcName As String = "gCalcWorkingDayDate"

On Error GoTo Err

yearStart = DateAdd("d", 1 - DatePart("y", baseDate), baseDate)

Do While dayNumber < 0
    yearEnd = yearStart - 1
    yearStart = DateAdd("yyyy", -1, yearStart)
    dayNumber = dayNumber + gCalcWorkingDayNumber(yearEnd) + 1
Loop

dow1 = DatePart("w", yearStart, vbMonday)

If dow1 = 7 Then
    ' Sunday
    wd1 = 0
    we1 = 1
ElseIf dow1 = 6 Then
    ' Saturday
    wd1 = 0
    we1 = 2
Else
    wd1 = 5 - dow1 + 1
    we1 = 2
End If

If dayNumber <= wd1 Then
    doy = dayNumber
ElseIf dayNumber - wd1 <= 5 Then
    doy = we1 + dayNumber
Else
    numWholeWeeks = Int((dayNumber - wd1) / 5) - 1
    doy = wd1 + we1 + IIf(numWholeWeeks > 0, 7 * numWholeWeeks + 5, 5) + IIf(((dayNumber - wd1) Mod 5) > 0, ((dayNumber - wd1) Mod 5) + 2, 0)
End If

gCalcWorkingDayDate = DateAdd("d", doy - 1, yearStart)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCalcWorkingDayNumber( _
                ByVal pDate As Date) As Long
Dim doy As Long     ' day of year
Dim woy As Long     ' week of year
Dim wd1 As Long     ' weekdays in first week (excluding weekend)
'Dim we1 As Long     ' weekend days at start of first week
Dim wdN As Long     ' weekdays in last week

Dim dow1 As Long    ' day of week of 1st jan
Dim dow As Long     ' day of week of sUpplied date

Const ProcName As String = "gCalcWorkingDayNumber"

On Error GoTo Err

doy = DatePart("y", pDate, vbMonday)
woy = DatePart("ww", pDate, vbMonday)
dow = DatePart("w", pDate, vbMonday)
dow1 = DatePart("w", pDate - doy + 1, vbMonday)

If dow1 = 7 Then
    ' Sunday
    wd1 = 0
'    we1 = 1
ElseIf dow1 = 6 Then
    ' Saturday
    wd1 = 0
'    we1 = 2
Else
    wd1 = 5 - dow1 + 1
'    we1 = 0
End If

If dow = 7 Or dow = 6 Then
    wdN = 5
Else
    wdN = dow
End If

gCalcWorkingDayNumber = wd1 + 5 * (woy - 2) + wdN

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function gGetTimePeriod( _
                ByVal Length As Long, _
                ByVal Units As TimePeriodUnits) As TimePeriod
Dim tp As TimePeriod

Const ProcName As String = "gGetTimePeriod"

On Error GoTo Err

If Length < 1 And Units <> TimePeriodNone Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Length cannot be < 1"
If Length <> 0 And Units = TimePeriodNone Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Length must be zero for a null timeperiod"

Select Case Units
    Case TimePeriodNone

    Case TimePeriodSecond

    Case TimePeriodMinute

    Case TimePeriodHour

    Case TimePeriodDay

    Case TimePeriodWeek

    Case TimePeriodMonth

    Case TimePeriodYear

    Case TimePeriodTickMovement

    Case TimePeriodTickVolume

    Case TimePeriodVolume
Case Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid Units argument"
End Select

Set tp = New TimePeriod
tp.Initialise Length, Units


' now ensure that only a single object for each timeperiod exists
On Error Resume Next
Set gGetTimePeriod = mTimePeriods(tp.ToString)
On Error GoTo Err

If gGetTimePeriod Is Nothing Then
    mTimePeriods.Add tp, tp.ToString
    Set gGetTimePeriod = tp
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function gTimePeriodUnitsFromString( _
                timeUnits As String) As TimePeriodUnits

Const ProcName As String = "gTimePeriodUnitsFromString"

On Error GoTo Err

Select Case UCase$(timeUnits)
Case UCase$(TimePeriodNameSecond), UCase$(TimePeriodNameSeconds), "SEC", "SECS", "S"
    gTimePeriodUnitsFromString = TimePeriodSecond
Case UCase$(TimePeriodNameMinute), UCase$(TimePeriodNameMinutes), "MIN", "MINS", "M"
    gTimePeriodUnitsFromString = TimePeriodMinute
Case UCase$(TimePeriodNameHour), UCase$(TimePeriodNameHours), "HR", "HRS", "H"
    gTimePeriodUnitsFromString = TimePeriodHour
Case UCase$(TimePeriodNameDay), UCase$(TimePeriodNameDays), "D", "DY", "DYS"
    gTimePeriodUnitsFromString = TimePeriodDay
Case UCase$(TimePeriodNameWeek), UCase$(TimePeriodNameWeeks), "W", "WK", "WKS"
    gTimePeriodUnitsFromString = TimePeriodWeek
Case UCase$(TimePeriodNameMonth), UCase$(TimePeriodNameMonths), "MTH", "MNTH", "MTHS", "MNTHS", "MM"
    gTimePeriodUnitsFromString = TimePeriodMonth
Case UCase$(TimePeriodNameYear), UCase$(TimePeriodNameYears), "YR", "YRS", "Y", "YY", "YS"
    gTimePeriodUnitsFromString = TimePeriodYear
Case UCase$(TimePeriodNameVolumeIncrement), "VOL", "V"
    gTimePeriodUnitsFromString = TimePeriodVolume
Case UCase$(TimePeriodNameTickVolumeIncrement), "TICKVOL", "TICK VOL", "TICKVOLUME", "TV"
    gTimePeriodUnitsFromString = TimePeriodTickVolume
Case UCase$(TimePeriodNameTickIncrement), "TICK", "TICKS", "TCK", "TCKS", "T", "TM", "TICKSMOVEMENT", "TICKMOVEMENT"
    gTimePeriodUnitsFromString = TimePeriodTickMovement
Case Else
    gTimePeriodUnitsFromString = TimePeriodNone
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gTimePeriodUnitsToString( _
                timeUnits As TimePeriodUnits) As String

Select Case timeUnits
Case TimePeriodSecond
    gTimePeriodUnitsToString = TimePeriodNameSeconds
Case TimePeriodMinute
    gTimePeriodUnitsToString = TimePeriodNameMinutes
Case TimePeriodHour
    gTimePeriodUnitsToString = TimePeriodNameHours
Case TimePeriodDay
    gTimePeriodUnitsToString = TimePeriodNameDays
Case TimePeriodWeek
    gTimePeriodUnitsToString = TimePeriodNameWeeks
Case TimePeriodMonth
    gTimePeriodUnitsToString = TimePeriodNameMonths
Case TimePeriodYear
    gTimePeriodUnitsToString = TimePeriodNameYears
Case TimePeriodVolume
    gTimePeriodUnitsToString = TimePeriodNameVolumeIncrement
Case TimePeriodTickVolume
    gTimePeriodUnitsToString = TimePeriodNameTickVolumeIncrement
Case TimePeriodTickMovement
    gTimePeriodUnitsToString = TimePeriodNameTickIncrement
End Select
End Function

Public Function gTimePeriodUnitsToShortString( _
                timeUnits As TimePeriodUnits) As String

Select Case timeUnits
Case TimePeriodSecond
    gTimePeriodUnitsToShortString = TimePeriodShortNameSeconds
Case TimePeriodMinute
    gTimePeriodUnitsToShortString = TimePeriodShortNameMinutes
Case TimePeriodHour
    gTimePeriodUnitsToShortString = TimePeriodShortNameHours
Case TimePeriodDay
    gTimePeriodUnitsToShortString = TimePeriodShortNameDays
Case TimePeriodWeek
    gTimePeriodUnitsToShortString = TimePeriodShortNameWeeks
Case TimePeriodMonth
    gTimePeriodUnitsToShortString = TimePeriodShortNameMonths
Case TimePeriodYear
    gTimePeriodUnitsToShortString = TimePeriodShortNameYears
Case TimePeriodVolume
    gTimePeriodUnitsToShortString = TimePeriodShortNameVolumeIncrement
Case TimePeriodTickVolume
    gTimePeriodUnitsToShortString = TimePeriodShortNameTickVolumeIncrement
Case TimePeriodTickMovement
    gTimePeriodUnitsToShortString = TimePeriodShortNameTickIncrement
End Select
End Function

'@================================================================================
' Helper Functions
'@================================================================================


