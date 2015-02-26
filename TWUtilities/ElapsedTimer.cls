VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ElapsedTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Objects of this class are used to measure elapsed times."
Option Explicit

''
' This class is used for measuring elapsed time to sub-microsecond accuracy.
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

'@================================================================================
' Member variables
'@================================================================================

Private mPerfFreq As Double
Private mStartPerfCounter As Currency

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Dim lPerfFreq As Currency
QueryPerformanceFrequency lPerfFreq
mPerfFreq = lPerfFreq
End Sub

'@================================================================================
' xxxx Interface Members
'@================================================================================

'@================================================================================
' xxxx Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

''
' Returns the elapsed time, measured in microseconds, since the last call to the
' <code>StartTiming</code> method.
'
' @remarks
'   Note that if the thread is pre-empted or completes its processor time quantum
'   during the measurement period, the duration until it is rescheduled will be
'   included in the measured time.
'
' @return
'   The number of microseconds since the most recent call to the
'   <code>StartTiming</code> method.
'@/
Public Function ElapsedTimeMicroseconds() As Single
Attribute ElapsedTimeMicroseconds.VB_Description = "Returns the elapsed time in microseconds since the most recent call to StartTiming."
Dim perfCounter As Currency
Dim diff As Currency
QueryPerformanceCounter perfCounter
diff = perfCounter - mStartPerfCounter
ElapsedTimeMicroseconds = CSng((1000000# * CDbl(diff)) / mPerfFreq)
End Function

''
' Starts a new measurement interval.
'@/
Public Sub StartTiming()
Attribute StartTiming.VB_Description = "A call to this method defines the point in time from which the ElapsedTimeMicroseconds method will measure."
QueryPerformanceCounter mStartPerfCounter
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


