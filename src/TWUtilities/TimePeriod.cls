VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimePeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
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

Implements IStringable

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
Private Const ModuleName                    As String = "Class1"

'@================================================================================
' Member variables
'@================================================================================

Private mLength                             As Long
Private mUnits                              As TimePeriodUnits

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Stringable Interface Members
'@================================================================================

Private Function IStringable_ToString() As String
IStringable_ToString = ToString
End Function

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Length() As Long
Length = mLength
End Property

Public Property Get Name() As String
Name = ToString
End Property

Public Property Get ShortName() As String
ShortName = ToShortString
End Property

Public Property Get Units() As TimePeriodUnits
Units = mUnits
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal Length As Long, _
                ByVal Units As TimePeriodUnits)
mLength = Length
mUnits = Units
End Sub

Public Function ToString() As String
Static Value As String
If Value = "" Then
    Select Case mUnits
    Case TimePeriodNone
        Value = "0"
    Case TimePeriodSecond
        If mLength = 1 Then
            Value = mLength & " " & TimePeriodNameSecond
        Else
            Value = mLength & " " & TimePeriodNameSeconds
        End If
    Case TimePeriodMinute
        If mLength = 1 Then
            Value = mLength & " " & TimePeriodNameMinute
        Else
            Value = mLength & " " & TimePeriodNameMinutes
        End If
    Case TimePeriodHour
        If mLength = 1 Then
            Value = TimePeriodNameHour
        Else
            Value = mLength & " " & TimePeriodNameHours
        End If
    Case TimePeriodDay
        If mLength = 1 Then
            Value = TimePeriodNameDay
        Else
            Value = mLength & " " & TimePeriodNameDays
        End If
    Case TimePeriodWeek
        If mLength = 1 Then
            Value = TimePeriodNameWeek
        Else
            Value = mLength & " " & TimePeriodNameWeeks
        End If
    Case TimePeriodMonth
        If mLength = 1 Then
            Value = TimePeriodNameMonth
        Else
            Value = mLength & " " & TimePeriodNameMonths
        End If
    Case TimePeriodYear
        If mLength = 1 Then
            Value = TimePeriodNameYear
        Else
            Value = mLength & " " & TimePeriodNameYears
        End If
    Case TimePeriodVolume
        Value = mLength & " " & TimePeriodNameVolumeIncrement
    Case TimePeriodTickVolume
        Value = mLength & " " & TimePeriodNameTickVolumeIncrement
    Case TimePeriodTickMovement
        Value = mLength & " " & TimePeriodNameTickIncrement
    End Select
End If
ToString = Value
End Function

Public Function ToShortString() As String
Static Value As String
If Value = "" Then
    If mUnits = TimePeriodNone Then
        Value = "0"
    Else
        Value = mLength & gTimePeriodUnitsToShortString(mUnits)
    End If
End If
ToShortString = Value
End Function

'@================================================================================
' Helper Functions
'@================================================================================



