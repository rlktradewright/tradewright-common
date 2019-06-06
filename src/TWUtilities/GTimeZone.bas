Attribute VB_Name = "GTimeZone"
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


Private Const ModuleName                    As String = "GTimeZone"

Private Const MaxNameLen                    As Long = MAX_PATH + 1

Private Const RegSubKeyTimezones            As String = "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Time Zones"

'@================================================================================
' Member variables
'@================================================================================

Private mTimeZoneNames()                    As String
Private mTimeZones                          As Collection
Private mLocalTimeZone                      As TimeZone

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

Public Function gGetAvailableTimeZoneNames() As String()
Const ProcName As String = "gGetAvailableTimeZoneNames"
On Error GoTo Err

ReDim mTimeZoneNames(200) As String

Dim hKey As Long
If RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
                RegSubKeyTimezones, _
                0&, _
                KEY_READ, _
                hKey) <> ERROR_SUCCESS Then
    Err.Raise ErrorCodes.ErrRuntimeException, , "Can't find timezones registry key"
End If

mTimeZoneNames(0) = "UTC"

Dim i As Long
Do
    i = i + 1
    Dim Name As String * MaxNameLen
    If RegEnumKey(hKey, i, Name, Len(Name)) = ERROR_NO_MORE_ITEMS Then Exit Do
    mTimeZoneNames(i) = gTrimNull(Name)
Loop

ReDim Preserve mTimeZoneNames(i - 1) As String

gGetAvailableTimeZoneNames = mTimeZoneNames

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetCurrentTimeZone() As TimeZone
Set gGetCurrentTimeZone = mLocalTimeZone
End Function

Public Function gGetCurrentTimeZoneName() As String
Dim tzi As TIME_ZONE_INFORMATION

Const ProcName As String = "gGetCurrentTimeZoneName"

On Error GoTo Err

GetTimeZoneInformation tzi
gGetCurrentTimeZoneName = gTrimNull(CStr(tzi.StandardName))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetTimeZone( _
                ByVal pTimezonename As String) As TimeZone

Const ProcName As String = "gGetTimeZone"

On Error GoTo Err

If pTimezonename = "" Then
    Set gGetTimeZone = mLocalTimeZone
    Exit Function
End If

On Error Resume Next
Set gGetTimeZone = mTimeZones(pTimezonename)
On Error GoTo Err

If gGetTimeZone Is Nothing Then
    Set gGetTimeZone = New TimeZone
    gGetTimeZone.Initialise pTimezonename
    mTimeZones.Add gGetTimeZone, pTimezonename
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetTimeZoneInformation( _
                ByVal pTimezonename As String, _
                ByRef pDisplayName As String, _
                ByRef pFirstYear As Long, _
                ByRef pLastyear As Long, _
                ByRef dynamicTZI() As TIME_ZONE_INFORMATION) As TIME_ZONE_INFORMATION
Dim hKey As Long
Dim dispName(2 * MaxNameLen) As Byte
Dim rtzi As REG_TIME_ZONE_INFORMATION
Dim tzi As TIME_ZONE_INFORMATION
Dim dynTzi() As TIME_ZONE_INFORMATION
Dim i As Long

Const ProcName As String = "gGetTimeZoneInformation"

On Error GoTo Err

If pTimezonename = "" Then
    GetTimeZoneInformation gGetTimeZoneInformation
    Exit Function
Else
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
                    RegSubKeyTimezones & "\" & pTimezonename, _
                    0&, _
                    KEY_READ, _
                    hKey) <> ERROR_SUCCESS Then
        Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid timezone Name: " & pTimezonename
    End If
    
    If RegQueryValueEx(hKey, _
                        "TZI", _
                        0&, _
                        0&, _
                        VarPtr(rtzi), _
                        Len(rtzi)) <> ERROR_SUCCESS Then
        Err.Raise ErrorCodes.ErrRuntimeException, , "Can't get TimeZoneInfo for: " & pTimezonename
    End If
    
    If RegQueryValueEx(hKey, _
                        "Display", _
                        0&, _
                        0&, _
                        VarPtr(dispName(0)), _
                        UBound(dispName) + 1) <> ERROR_SUCCESS Then
        Err.Raise ErrorCodes.ErrRuntimeException, , "Can't get TimeZoneInfo for: " & pTimezonename
    End If
    
    pDisplayName = gTrimNull(StrConv(dispName, vbUnicode))
    
    If RegQueryValueEx(hKey, _
                        "Std", _
                        0&, _
                        0&, _
                        VarPtr(tzi.StandardName(0)), _
                        UBound(tzi.StandardName) + 1) <> ERROR_SUCCESS Then
        Err.Raise ErrorCodes.ErrRuntimeException, , "Can't get TimeZoneInfo for: " & pTimezonename
    End If
    
    If RegQueryValueEx(hKey, _
                        "Dlt", _
                        0&, _
                        0&, _
                        VarPtr(tzi.DaylightName(0)), _
                        UBound(tzi.DaylightName) + 1) <> ERROR_SUCCESS Then
        Err.Raise ErrorCodes.ErrRuntimeException, , "Can't get TimeZoneInfo for: " & pTimezonename
    End If
    
    tzi.Bias = rtzi.Bias
    tzi.daylightBias = rtzi.daylightBias
    tzi.daylightDate = rtzi.daylightDate
    tzi.standardBias = rtzi.standardBias
    tzi.standardDate = rtzi.standardDate
    
    gGetTimeZoneInformation = tzi
    
    RegCloseKey hKey

    ' now get the dynamic timezone info
    
    If RegOpenKeyEx(HKEY_LOCAL_MACHINE, _
                    RegSubKeyTimezones & "\" & pTimezonename & "\Dynamic DST", _
                    0&, _
                    KEY_READ, _
                    hKey) <> ERROR_SUCCESS Then
        Exit Function
    End If
    
    If RegQueryValueEx(hKey, _
                        "FirstEntry", _
                        0&, _
                        0&, _
                        VarPtr(pFirstYear), _
                        4) <> ERROR_SUCCESS Then
        Err.Raise ErrorCodes.ErrRuntimeException, , "Can't get FirstEntry for: " & pTimezonename
    End If
    
    If RegQueryValueEx(hKey, _
                        "LastEntry", _
                        0&, _
                        0&, _
                        VarPtr(pLastyear), _
                        4) <> ERROR_SUCCESS Then
        Err.Raise ErrorCodes.ErrRuntimeException, , "Can't get LastEntry for: " & pTimezonename
    End If
    
    ReDim dynTzi(pLastyear - pFirstYear) As TIME_ZONE_INFORMATION
    
    For i = pFirstYear To pLastyear
        If RegQueryValueEx(hKey, _
                            CStr(i), _
                            0&, _
                            0&, _
                            VarPtr(rtzi), _
                            Len(rtzi)) <> ERROR_SUCCESS Then
            Err.Raise ErrorCodes.ErrRuntimeException, , "Can't get dynamic TimeZoneInfo for year " & i & " for: " & pTimezonename
        
        End If
        
        tzi.Bias = rtzi.Bias
        tzi.daylightBias = rtzi.daylightBias
        tzi.daylightDate = rtzi.daylightDate
        tzi.standardBias = rtzi.standardBias
        tzi.standardDate = rtzi.standardDate
        
        dynTzi(i - pFirstYear) = tzi
    Next
    
    dynamicTZI = dynTzi

    RegCloseKey hKey

End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gInit()
Const ProcName As String = "gInit"
On Error GoTo Err

Dim Name As String

Set mTimeZones = New Collection
Name = gGetCurrentTimeZoneName
Set mLocalTimeZone = New TimeZone
mLocalTimeZone.Initialise Name
mTimeZones.Add mLocalTimeZone, Name

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




