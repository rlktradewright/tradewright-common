VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeZone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class provides the means to convert times from one timezone to another.
'
' Instances of this class cannot be created using the <code>New</code> operator. Instead,
' use the <code>GetTimeZone</code> global method. There is a single global
' <code>GetTimeZone</code> object for each timezone known to the system.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IJSONable
Implements Stringable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

Private Type TzConversionInfo
    standardDate            As Date
    daylightDate            As Date
    standardBias            As Long
    daylightBias            As Long
End Type

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "TimeZone"

'@================================================================================
' Member variables
'@================================================================================

Private mBaseTimeZoneInfo                   As TIME_ZONE_INFORMATION

Private mDynamicTimeZoneInfo()              As TIME_ZONE_INFORMATION

Private mFirstYear                          As Long
Private mLastYear                           As Long

Private mDisplayName                        As String

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' JSONable Interface Members
'@================================================================================

Private Function IJSONable_ToJSON() As String
IJSONable_ToJSON = ToString
End Function

'@================================================================================
' Stringable Interface Members
'@================================================================================

Private Function Stringable_ToString() As String
Stringable_ToString = ToString
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Friend Property Get BaseTimeZoneInfo() As TIME_ZONE_INFORMATION
BaseTimeZoneInfo = mBaseTimeZoneInfo
End Property

''
' Returns the daylight saving time Name for this timezone.
'
' @return
'   The daylight saving time Name for this timezone.
'@/
Public Property Get DstName() As String
Const ProcName As String = "DstName"

On Error GoTo Err

DstName = gTrimNull(StrConv(mBaseTimeZoneInfo.DaylightName, vbUnicode))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the daylight saving time offset (in minutes) for this timezone.
'
' @return
'   The daylight saving time offset for this timezone.
'@/
Public Property Get DstOffset() As Long
DstOffset = -mBaseTimeZoneInfo.daylightBias
End Property

''
' Returns the display Name for this timezone.
'
' @return
'   The display Name for this timezone.
'@/
Public Property Get displayName() As String
displayName = mDisplayName
End Property

''
' Returns the offset from UTC (in minutes) for this timezone.
'
' @return
'   The offset from UTC for this timezone.
'@/
Public Property Get Offset() As Long
Offset = -mBaseTimeZoneInfo.Bias
End Property

''
' Returns the standard Name for this timezone.
'
' @return
'   The standard Name for this timezone.
'@/
Public Property Get StandardName() As String
Const ProcName As String = "StandardName"

On Error GoTo Err

StandardName = gTrimNull(StrConv(mBaseTimeZoneInfo.StandardName, vbUnicode))

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Converts a date in the computer's local timezone to this timezone.
'
' @return
'   The specified date converted to the timezone represented by this
'   <code>TimeZone</code> object.
' @param LocalDate
'   A date in the computer's local timezone.
'@/
Public Function ConvertDateLocalToTz( _
                ByVal localDate As Date) As Date
Const ProcName As String = "ConvertDateLocalToTz"

On Error GoTo Err

If Me Is gGetCurrentTimeZone Then
    ConvertDateLocalToTz = localDate
    Exit Function
End If
ConvertDateLocalToTz = ConvertDateUTCToTZ(gLocalToUtc(localDate))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a date in this timezone to the computer's local timezone.
'
' @return
'   The specified date converted to the local timezone.
' @param tzDate
'   A date in the timezone represented by this <code>TimeZone</code> object.
'@/
Public Function ConvertDateTzToLocal( _
                ByVal tzDate As Date) As Date
Const ProcName As String = "ConvertDateTzToLocal"

On Error GoTo Err

If Me Is gGetCurrentTimeZone Then
    ConvertDateTzToLocal = tzDate
    Exit Function
End If
ConvertDateTzToLocal = gUtcToLocal(ConvertDateTzToUTC(tzDate))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a date in this timezone to another timezone.
'
' @return
'   The specified date converted to the required timezone.
' @param pDate
'   A date in the this timezone.
' @param tz
'   The timezone to which the date is to be converted.
'@/
Public Function ConvertDateTzToTz( _
                ByVal pDate As Date, _
                ByVal tz As TimeZone) As Date
Const ProcName As String = "ConvertDateTzToTz"

On Error GoTo Err

If Me Is tz Then
    ConvertDateTzToTz = pDate
    Exit Function
End If
ConvertDateTzToTz = tz.ConvertDateUTCToTZ(localTimeToSystemTime(pDate))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a date in this timezone to UTC.
'
' @return
'   The specified date converted to UTC.
' @param LocalDate
'   A date in the timezone represented by this <code>TimeZone</code> object.
'@/
Public Function ConvertDateTzToUTC( _
                ByVal localDate As Date) As Date
Const ProcName As String = "ConvertDateTzToUTC"

On Error GoTo Err

ConvertDateTzToUTC = localTimeToSystemTime(localDate)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Converts a date from UTC to this timezone.
'
' @return
'   The specified date converted to the timezone represented by
'   this <code>TimeZone</code> object.
' @param utcDate
'   A UTC date.
'@/
Public Function ConvertDateUTCToTZ( _
                ByVal utcDate As Date) As Date
Dim utcFileTime As Currency
Dim localFileTime As Currency
Dim sysUTC As SYSTEMTIME
Dim sysLocal As SYSTEMTIME

Const ProcName As String = "ConvertDateUTCToTZ"

On Error GoTo Err

utcFileTime = gVbDateToFileTime(utcDate)
FileTimeToSystemTime utcFileTime, sysUTC
SystemTimeToTzSpecificLocalTime getTimeZoneInfoForYear(sysUTC.Year), sysUTC, sysLocal
SystemTimeToFileTime sysLocal, localFileTime
ConvertDateUTCToTZ = gFileTimeToVbDate(localFileTime)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
                
End Function

Friend Function getTimeZoneInfoForYear( _
                ByVal pYear As Long) As TIME_ZONE_INFORMATION

Const ProcName As String = "getTimeZoneInfoForYear"

On Error GoTo Err

If mFirstYear = 0 Then
    getTimeZoneInfoForYear = mBaseTimeZoneInfo
ElseIf pYear < mFirstYear Then
    getTimeZoneInfoForYear = mBaseTimeZoneInfo
ElseIf pYear > mLastYear Then
    getTimeZoneInfoForYear = mBaseTimeZoneInfo
Else
    getTimeZoneInfoForYear = mDynamicTimeZoneInfo(pYear - mFirstYear)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Friend Sub Initialise( _
                ByVal pTimezonename As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

If UCase$(pTimezonename) = "UTC" Then
    mBaseTimeZoneInfo.DaylightName(0) = Asc("U")
    mBaseTimeZoneInfo.DaylightName(1) = Asc("T")
    mBaseTimeZoneInfo.DaylightName(2) = Asc("C")
    mBaseTimeZoneInfo.StandardName(0) = Asc("U")
    mBaseTimeZoneInfo.StandardName(1) = Asc("T")
    mBaseTimeZoneInfo.StandardName(2) = Asc("C")
    mDisplayName = "UTC"
Else
    mBaseTimeZoneInfo = gGetTimeZoneInformation(pTimezonename, _
                                            mDisplayName, _
                                            mFirstYear, _
                                            mLastYear, _
                                            mDynamicTimeZoneInfo)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function localTimeToSystemTime( _
                ByVal localTime As Date) As Date

Static tzConvInfo As TzConversionInfo
Static sYear As Long
Dim tzi As TIME_ZONE_INFORMATION
Dim st As SYSTEMTIME
Dim lYear As Long
Dim utcTime As Date

Const ProcName As String = "localTimeToSystemTime"

On Error GoTo Err

lYear = Year(localTime)

If lYear <> sYear Then
    
    tzi = getTimeZoneInfoForYear(lYear)

    If tzi.standardDate.Month = 0 Then
        tzConvInfo.standardBias = tzi.Bias
    Else
        
        If tzi.daylightDate.Month = 0 Then Err.Raise ErrorCodes.ErrRuntimeException, , "Daylightdate month should be non-zero"
        
        ' calculate the daylight saving time start date
        st = tzi.daylightDate
        st.Year = lYear
        st.Day = 1
        
        tzConvInfo.daylightDate = gSystemTimeToVbDate( _
                                            gNthWeekdayOfMonth(st, _
                                                                tzi.daylightDate.dayOfWeek, _
                                                                tzi.daylightDate.Day))
        
        ' calculate the standard time start date
        st = tzi.standardDate
        st.Year = lYear
        st.Day = 1
        
        tzConvInfo.standardDate = gSystemTimeToVbDate( _
                                            gNthWeekdayOfMonth(st, _
                                                                tzi.standardDate.dayOfWeek, _
                                                                tzi.standardDate.Day))
                    
        tzConvInfo.daylightBias = tzi.daylightBias
        tzConvInfo.standardBias = tzi.Bias
        
    End If
    
    sYear = lYear
    
End If

utcTime = localTime + tzConvInfo.standardBias / 1440

If tzConvInfo.standardDate = 0 Then
    ' no daylight time = nothing to do

ElseIf tzConvInfo.daylightDate < tzConvInfo.standardDate Then
    ' northern hemisphere
    If CDbl(localTime + tzConvInfo.daylightBias / 1440) >= tzConvInfo.daylightDate And _
        localTime < tzConvInfo.standardDate _
    Then
        ' we are in DST
        utcTime = utcTime + tzConvInfo.daylightBias / 1440
    End If
Else
    ' northern hemisphere
    If localTime < tzConvInfo.standardDate Or _
        CDbl(localTime + tzConvInfo.daylightBias / 1440) >= tzConvInfo.daylightDate _
    Then
        ' we are in DST
        utcTime = utcTime + tzConvInfo.daylightBias / 1440
    End If
End If

localTimeToSystemTime = utcTime

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Function ToString() As String
Dim jb As JSONBuilder
Const ProcName As String = "ToString"

On Error GoTo Err

Set jb = New JSONBuilder
jb.BeginObject
jb.EncodeName "DisplayName"
jb.EncodeValue mDisplayName
jb.EndObject
ToString = jb.JSONString

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

