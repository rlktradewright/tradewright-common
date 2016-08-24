Attribute VB_Name = "Globals"
Option Explicit

'@================================================================================
' Description
'@================================================================================
'
'

'@================================================================================
' Interfaces
'@================================================================================

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                    As String = "TWUtilities40"
Private Const ModuleName                    As String = "Globals"

Public Const MinDateValue                   As Date = -657434 ' 1 Jan 100
Public Const MaxDateValue                   As Date = 2958465 ' 31 Dec 9999

Public Const MaxDoubleValue                 As Double = (2 - 2 ^ -52) * 2 ^ 1023
Public Const MinDoubleValue                 As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const MinPositiveDoubleValue         As Double = 2 ^ -1074

Public Const MaxLongValue                   As Long = &H7FFFFFFF
Public Const MinLongValue                   As Long = &H80000000

'@================================================================================
' Enums
'@================================================================================

' Red/Black tree node colors
Public Enum NodeColors
    BLACK
    Red
End Enum

Public Enum UserMessages
    UserMessageTimer = WM_USER + 1234
    UserMessageScheduleTasks
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mInitialised As Boolean
Private mTerminated As Boolean

Private mFSO As FileSystemObject

Private mPrintableKeyCodes(255) As Byte

Private mHexChars(15) As Byte   ' ASCII

Private mPostMessageForm As PostMessageForm

Private mSysInfo As SYSTEM_INFO

Private mLogTokens(9) As String

Private mPrevWndProc As Long

Private mMainWindowHandle As Long

Private mUnhandledErrorHandler As New UnhandledErrorHandler

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

Public Property Get gFileSystemObject() As FileSystemObject
If mFSO Is Nothing Then Set mFSO = New FileSystemObject
Set gFileSystemObject = mFSO
End Property

Public Property Get gInitialised() As Boolean
gInitialised = mInitialised
End Property

Public Property Get gRegExp() As RegExp
Static lRegexp As RegExp
If lRegexp Is Nothing Then Set lRegexp = New RegExp
Set gRegExp = lRegexp
End Property

Public Property Get gUnhandledErrorHandler() As UnhandledErrorHandler
Set gUnhandledErrorHandler = mUnhandledErrorHandler
End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub gAssert( _
                ByVal pCondition As Boolean, _
                Optional ByVal pMessage As String, _
                Optional ByVal pException As Long = ErrorCodes.ErrIllegalStateException)
If Not pCondition Then Err.Raise pException, , pMessage
End Sub

Public Sub gAssertArgument( _
                ByVal pCondition As Boolean, _
                Optional ByVal pMessage As String)
gAssert pCondition, pMessage, ErrorCodes.ErrIllegalArgumentException
End Sub

Public Function gBytesToHexString(inAr() As Byte) As String
Const ProcName As String = "gBytesToHexString"
On Error GoTo Err

ReDim outAr(4 * (UBound(inAr()) + 1) - 1) As Byte

Dim i As Long
For i = 0 To UBound(inAr)
    outAr(4 * i) = mHexChars(Int(inAr(i) / 16))
    outAr(4 * i + 2) = mHexChars(inAr(i) And &HF)
Next

gBytesToHexString = outAr

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gCopyFromObjects( _
                ByRef objAr() As Object, _
                ByVal destArPtr As Long)
CopyMemory destArPtr, VarPtr(objAr(0)), (UBound(objAr) + 1) * 4
ZeroMemory VarPtr(objAr(0)), (UBound(objAr) + 1) * 4
End Sub

Public Sub gCopyToObjects( _
                ByVal sourceArPtr As Long, _
                ByRef objAr() As Object)
CopyMemory VarPtr(objAr(0)), sourceArPtr, (UBound(objAr) + 1) * 4
ZeroMemory sourceArPtr, (UBound(objAr) + 1) * 4
End Sub

Public Function gCpuSpeedMhz() As Long
Const ProcName As String = "gCpuSpeedMhz"
On Error GoTo Err

Const RegKeyCpuSpeed As String = "HARDWARE\DESCRIPTION\System\CentralProcessor\0"

Static cpuSpeed As Long
If cpuSpeed = 0 Then
    Dim hKey As Long
    RegOpenKeyEx HKEY_LOCAL_MACHINE, RegKeyCpuSpeed, 0, KEY_READ, hKey
                     
    RegQueryValueEx hKey, "~MHz", 0, 0, VarPtr(cpuSpeed), Len(cpuSpeed)
    RegCloseKey hKey
End If

gCpuSpeedMhz = cpuSpeed

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateWriteableTextFile( _
                ByVal pFilename As String, _
                Optional ByVal pOverwrite As Boolean, _
                Optional ByVal pCreateBackup As Boolean, _
                Optional ByVal pUnicode As Boolean, _
                Optional ByVal pIncrementFilenameIfInUse As Boolean) As TextStream
Const ProcName As String = "gCreateWriteableTextFile"
On Error GoTo Err

Dim i As Long
Dim ext As String
Dim fname As String
Dim Path As String
Dim cantRename As Boolean

If pOverwrite And pCreateBackup Then
    If gFileSystemObject.FileExists(pFilename) Then
        gParseFilename pFilename, Path, fname, ext
        
        Dim f As File
        Set f = gFileSystemObject.GetFile(pFilename)
        If gFileSystemObject.FileExists(Path & fname & ".bak." & ext) Then
            i = 1
            Do While gFileSystemObject.FileExists(Path & fname & ".bak" & i & "." & ext)
                i = i + 1
            Loop
            
            On Error Resume Next    ' in case file is in use or not accessible
            f.Name = fname & ".bak" & i & "." & ext
            If Err.number = VBErrorCodes.VbErrPermissionDenied Then cantRename = True
            On Error GoTo Err
        Else
            On Error Resume Next    ' in case file is in use or not accessible
            f.Name = fname & ".bak." & ext
            If Err.number = VBErrorCodes.VbErrPermissionDenied Then cantRename = True
            On Error GoTo Err
        End If
    End If
End If

If Not cantRename Then
    Set gCreateWriteableTextFile = gFileSystemObject.OpenTextFile(pFilename, _
                                            IIf(pOverwrite, IOMode.ForWriting, IOMode.ForAppending), _
                                            True, _
                                            IIf(pUnicode, Tristate.TristateTrue, Tristate.TristateFalse))
Else
    ' can't rename the existing file: this may be either because it's in use by another
    ' process, or we don't have access permission
    
    gAssert pIncrementFilenameIfInUse, "File already in use or access denied", VBErrorCodes.VbErrPermissionDenied
    
    ' now increment the pFilename to find one we can use
    i = 1
    Do While gFileSystemObject.FileExists(Path & fname & "-" & i & "." & ext)
        i = i + 1
    Loop
    
    ' if we fail on this it must be because we don't have access
    Set gCreateWriteableTextFile = gFileSystemObject.OpenTextFile(Path & fname & "-" & i & "." & ext, _
                                            IIf(pOverwrite, IOMode.ForWriting, IOMode.ForAppending), _
                                            True, _
                                            IIf(pUnicode, Tristate.TristateTrue, Tristate.TristateFalse))
    
    
    
End If
Exit Function

Err:
If Err.number = VBErrorCodes.VbErrPathNotFound Then
    gCreateFolder Left$(pFilename, InStrRev(pFilename, "\") - 1)
    Resume
ElseIf Err.number = VBErrorCodes.VbErrPermissionDenied Then
    Err.Raise ErrorCodes.ErrSecurityException, , Err.Description
End If
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gCreateFolder( _
                ByVal folderPath As String)
Const ProcName As String = "gCreateFolder"
On Error GoTo Err

gFileSystemObject.CreateFolder folderPath

Exit Sub

Err:
If Err.number = VBErrorCodes.VbErrPathNotFound Then
    gCreateFolder Left$(folderPath, InStrRev(folderPath, "\") - 1)
    Resume
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gCreateIntervalTimer( _
                ByVal firstExpiryTime As Variant, _
                Optional ByVal firstExpiryUnits As ExpiryTimeUnits = ExpiryTimeUnits.ExpiryTimeUnitMilliseconds, _
                Optional ByVal repeatIntervalMillisecs As Long, _
                Optional ByVal useRandomIntervals As Boolean, _
                Optional ByVal pData As Variant) As IntervalTimer
Const ProcName As String = "gCreateIntervalTimer"
On Error GoTo Err

If IsMissing(pData) Then pData = Empty

Dim firstExpiryInterval As Long
If firstExpiryUnits = ExpiryTimeUnits.ExpiryTimeUnitDateTime Then
    gAssertArgument IsDate(firstExpiryTime), "firstExpiryTime is not a valid date"
    
    Dim ExpiryDate As Date
    ExpiryDate = gLocalToUtc(CDate(firstExpiryTime))
    
    On Error Resume Next
    firstExpiryInterval = (ExpiryDate - gGetTimestampUtc) * 86400 * 1000
    If Err.number = VBErrorCodes.VbErrOverflow Then firstExpiryInterval = -1
    On Error GoTo Err
    
    gAssertArgument firstExpiryInterval >= 0, "FirstExpiryTime has already passed"

Else
    gAssertArgument IsNumeric(firstExpiryTime), "ExpiryTime is not a valid number"
    
    If firstExpiryUnits = ExpiryTimeUnits.ExpiryTimeUnitMilliseconds Then
        firstExpiryInterval = CDbl(firstExpiryTime)
    ElseIf firstExpiryUnits = ExpiryTimeUnits.ExpiryTimeUnitSeconds Then
        firstExpiryInterval = CDbl(firstExpiryTime) * 1000
    ElseIf firstExpiryUnits = ExpiryTimeUnits.ExpiryTimeUnitMinutes Then
        firstExpiryInterval = CDbl(firstExpiryTime) * 60 * 1000
    ElseIf firstExpiryUnits = ExpiryTimeUnits.ExpiryTimeUnitHours Then
        firstExpiryInterval = CDbl(firstExpiryTime) * 60 * 60 * 1000
    ElseIf firstExpiryUnits = ExpiryTimeUnits.ExpiryTimeUnitDays Then
        firstExpiryInterval = CDbl(firstExpiryTime) * 24 * 60 * 60 * 1000
    Else
        gAssertArgument False, "FirstExpiryUnits argument invalid"
    End If
    
    gAssertArgument firstExpiryInterval >= 0, "FirstExpiryTime cannot be negative"
End If

gAssertArgument repeatIntervalMillisecs >= 0, "repeatIntervalMillisecs cannot be negative"

On Error GoTo Err

Set gCreateIntervalTimer = New IntervalTimer

gCreateIntervalTimer.Initialise firstExpiryInterval, _
                                repeatIntervalMillisecs, _
                                useRandomIntervals, _
                                pData

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateParameterStringParser( _
                ByVal Value As String, _
                Optional ByVal nameDelimiter As String = "=", _
                Optional ByVal parameterSeparator As String = ";", _
                Optional ByVal escapeCharacter As String = "\") As ParameterStringParser
Const ProcName As String = "gCreateParameterStringParser"
On Error GoTo Err

gAssertArgument Len(nameDelimiter) = 1 And _
                    Len(parameterSeparator) = 1 And _
                    Len(escapeCharacter) = 1, _
                "Delimiters and escape characters must be a single Character"

Set gCreateParameterStringParser = New ParameterStringParser
gCreateParameterStringParser.Initialise Value, _
                                    nameDelimiter, _
                                    parameterSeparator, _
                                    escapeCharacter

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateWeakReference( _
                ByVal target As Object) As WeakReference
Const ProcName As String = "gCreateWeakReference"
On Error GoTo Err

Set gCreateWeakReference = New WeakReference
gCreateWeakReference.Initialise target

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gEndProcess(ByVal pExitCode As Long)
TWWin32API.TerminateProcess GetCurrentProcess, pExitCode
End Sub

Public Function gGenerateGUID() As GUIDStruct
Const ProcName As String = "gGenerateGUID"
On Error GoTo Err

Dim lReturn As Long
lReturn = CoCreateGuid(gGenerateGUID)

gAssert lReturn = S_OK, "Can't create GUID", ErrorCodes.ErrRuntimeException

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function gGenerateGUIDString() As String
Const ProcName As String = "gGenerateGUIDString"
On Error GoTo Err

gGenerateGUIDString = gGUIDToString(gGenerateGUID)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGenerateID() As Long
gGenerateID = Int(Rnd() * &H7FFFFFFF) + 1
End Function

Public Function gGenerateTextID() As String
gGenerateTextID = Hex(gGenerateID)
End Function

Public Function gGetCommandLine() As String
Static lCommandLine As String

If lCommandLine = "" Then
    Dim lAddr As Long
    lAddr = GetCommandLine
    
    ReDim lBuf(getCommandLineLength(lAddr) - 1) As Byte
    CopyMemory VarPtr(lBuf(0)), lAddr, UBound(lBuf) + 1
    lCommandLine = lBuf
    
End If

gGetCommandLine = lCommandLine
End Function

Public Function gGetObjectKey(ByVal pObject As Object) As String
gGetObjectKey = Hex$(ObjPtr(pObject))
End Function

Public Function gGUIDToString(ByRef pGUID As GUIDStruct) As String
Const ProcName As String = "gGUIDToString"
On Error GoTo Err

Dim GUIDString As GUIDString
Dim iChars As Integer

iChars = StringFromGUID2(pGUID, GUIDString, Len(GUIDString))
' convert string to ANSI
gGUIDToString = GUIDString.GUIDProper

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gHandleUnexpectedError( _
                ByVal pProcedureName As String, _
                ByVal pModuleName As String, _
                Optional ByVal pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pNumber As Long, _
                Optional ByVal pDescription As String, _
                Optional ByVal pSource As String, _
                Optional ByVal pProjectName As String = ProjectName)
Const ProcName As String = "gHandleUnexpectedError"

Dim errSource As String: errSource = IIf(pSource <> "", pSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pDescription <> "", pDescription, Err.Description)
Dim errNum As Long: errNum = IIf(pNumber <> 0, pNumber, Err.number)

errSource = IIf(errSource <> "", errSource & vbCrLf, "") & _
            IIf(pProjectName <> "", pProjectName, ProjectName) & "." & _
            IIf(pModuleName <> "", pModuleName & ":", "") & _
            pProcedureName & _
            IIf(pFailpoint <> "", " At: " & pFailpoint, "")
If pLog Then gErrorLogger.Log LogLevels.LogLevelSevere, "Error " & errNum & ": " & errDesc & vbCrLf & errSource

If pReRaise Then
    If errNum = 0 Then
        Err.Raise ErrorCodes.ErrIllegalStateException, _
                errSource, _
                "gHandleUnexpectedError called in non-error context"
    Else
        Err.Raise errNum, errSource, errDesc
    End If
End If
End Sub

Public Function gNormalizeColor(ByVal pColor As Long) As Long
If pColor < 0 Then
    gNormalizeColor = GetSysColor(pColor And &HFF&)
Else
    gNormalizeColor = pColor
End If
End Function

Public Sub gNotifyUnhandledError( _
                ByVal pProcedureName As String, _
                ByVal pModuleName As String, _
                Optional ByVal pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByVal pErrorDesc As String, _
                Optional ByVal pErrorSource As String)
mUnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, pErrorNumber, pErrorDesc, pErrorSource
End Sub

Public Sub gHandleWin32Error()
Err.Raise ErrorCodes.ErrRuntimeException, , "Windows error " & GetLastError
End Sub

Public Function gHexStringToBytes(inString As String) As Byte()
Const ProcName As String = "gHexStringToBytes"
On Error GoTo Err

ReDim outAr(Len(inString) / 2 - 1) As Byte

Dim inAr() As Byte
inAr = inString

Dim i As Long
For i = 0 To UBound(outAr)
    outAr(i) = convertHexToByte(inAr(4 * i), inAr(4 * i + 2))
Next
gHexStringToBytes = outAr

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gInitialise()
Const ProcName As String = "gInitialise"
On Error GoTo Err

If mInitialised Then Exit Sub
If mTerminated Then Err.Raise ErrorCodes.ErrIllegalStateException, , "TWUtilities has been terminated"

mInitialised = True

Randomize

GLogging.gInit

GetSystemInfo mSysInfo
gCpuSpeedMhz

GTracer.gInit
GTimeZone.gInit
TimestampGlobals.gInit
GIntervalTimer.gInit
GClock.gInit
GTimerList.gInit

initPrintableKeyCodes

initHexChars

hook

gLogger.Log "TWUtilities initialisation completed", ProcName, ModuleName, LogLevelDetail

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gIntegersToHexString( _
                inAr() As Integer) As String
Const ProcName As String = "gIntegersToHexString"
On Error GoTo Err

ReDim outAr(8 * (UBound(inAr()) + 1) - 1) As Byte

Dim i As Long
For i = 0 To UBound(inAr)
    Dim val As Long
    If inAr(i) >= 0 Then
        val = inAr(i)
    Else
        ' convert to long ignoring sign
        Dim temp As Integer
        temp = inAr(i) And &H7FFF
        val = temp
        val = val Or &H8000&
    End If
    
    ' remember - little endian!!
    outAr(8 * i + 4) = mHexChars((val And &HF000) / &H1000)
    outAr(8 * i + 6) = mHexChars((val And &HF00) / &H100)
    
    outAr(8 * i) = mHexChars((val And &HF0) / &H10)
    outAr(8 * i + 2) = mHexChars(val And &HF)
Next

gIntegersToHexString = outAr

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function gIsInteger( _
                ByVal Value As Variant, _
                Optional ByVal minValue As Long = &H80000000, _
                Optional ByVal maxValue As Long = &H7FFFFFFF) As Boolean
Const ProcName As String = "gIsInteger"
On Error GoTo Err

gAssertArgument minValue <= maxValue, "minValue must not be greater than maxValue"
On Error GoTo Err

If IsNumeric(Value) Then
    Dim quantity As Long
    quantity = CLng(Value)
    If CDbl(Value) - quantity = 0 Then
        If quantity >= minValue And quantity <= maxValue Then
            gIsInteger = True
        End If
    End If
End If
                
Exit Function

Err:
If Err.number = VBErrorCodes.VbErrOverflow Then Exit Function
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gIsPrintableKey( _
                ByVal keyCode As KeyCodeConstants) As Boolean
Const ProcName As String = "gIsPrintableKey"
On Error GoTo Err

gIsPrintableKey = mPrintableKeyCodes(keyCode)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gIswhitespace( _
                ByVal Value As String) As Boolean
Const ProcName As String = "gIswhitespace"
On Error GoTo Err

Dim i As Long
For i = 1 To Len(Value)
    Select Case Mid$(Value, i, 1)
    Case " "
    Case vbCrLf
    Case vbNewLine
    Case vbTab
    Case vbFormFeed
    Case vbLf
    Case vbVerticalTab
    Case Else
        Exit Function
    End Select
Next
gIswhitespace = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gLongsToHexString( _
                inAr() As Integer) As String
Const ProcName As String = "gLongsToHexString"
On Error GoTo Err

ReDim outAr(18 * (UBound(inAr()) + 1) - 1) As Byte

Dim i As Long
For i = 0 To UBound(inAr)
    Dim val As Long
    val = inAr(i)
    
    ' remember - little endian!!
    If val >= 0 Then
        outAr(8 * i + 12) = mHexChars(val / &H10000000)
    Else
        outAr(8 * i + 12) = mHexChars(((val And &H70000000) / &H10000000) Or &H80000000)
    End If
    outAr(8 * i + 14) = mHexChars((val And &HF000000) / &H1000000)
    
    outAr(8 * i + 8) = mHexChars((val And &HF00000) / &H100000)
    outAr(8 * i + 10) = mHexChars((val And &HF0000) / &H10000)
    
    outAr(8 * i + 4) = mHexChars((val And &HF000) / &H1000)
    outAr(8 * i + 6) = mHexChars((val And &HF00) / &H100)
    
    outAr(8 * i) = mHexChars((val And &HF0) / &H10)
    outAr(8 * i + 2) = mHexChars(val And &HF)
Next

gLongsToHexString = outAr

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gMax( _
                ParamArray values() As Variant) As Variant
Const ProcName As String = "gMax"
On Error GoTo Err

Dim Value As Variant
For Each Value In values
    If IsNumeric(Value) Then
        If IsEmpty(gMax) Then
            gMax = Value
        ElseIf Value > gMax Then
            gMax = Value
        End If
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gMin( _
                ParamArray values() As Variant) As Variant
Const ProcName As String = "gMin"
On Error GoTo Err

Dim Value As Variant
For Each Value In values
    If IsNumeric(Value) Then
        If IsEmpty(gMin) Then
            gMin = Value
        ElseIf Value < gMin Then
            gMin = Value
        End If
    End If
Next

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gNumberOfProcessors()
Const ProcName As String = "gNumberOfProcessors"
On Error GoTo Err

gNumberOfProcessors = mSysInfo.dwNumberOfProcessors

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gParseFilename( _
                fullFilename As String, _
                Path As String, _
                pFilename As String, _
                extension As String)
Const ProcName As String = "gParseFilename"
On Error GoTo Err

Dim lastBackslashPosn As Long
lastBackslashPosn = InStrRev(fullFilename, "\")

Dim periodPosn As Long
periodPosn = InStrRev(fullFilename, ".")
If periodPosn < lastBackslashPosn Then periodPosn = 0

Dim filenameStub As String
If periodPosn = 0 Then
    extension = ""
    filenameStub = fullFilename
Else
    extension = Right$(fullFilename, Len(fullFilename) - periodPosn)
    filenameStub = Left$(fullFilename, periodPosn - 1)
End If

If lastBackslashPosn = 0 Then
    pFilename = filenameStub
    Path = ""
Else
    pFilename = Right$(filenameStub, Len(filenameStub) - lastBackslashPosn)
    Path = Left$(filenameStub, lastBackslashPosn)   ' include the backslash
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gPostUserMessage( _
                ByVal pMessageId As Long, _
                ByVal pUserData1 As Long, _
                ByVal pUserData2 As Long)
PostMessage mMainWindowHandle, pMessageId, pUserData1, pUserData2
End Sub

Public Function gRedirectMethod( _
                ByVal pObjectPointer As Long, _
                ByVal pvTableIndex As Long, _
                ByVal pNewAddress As Long) As Long
Dim vTableAddress As Long
CopyMemory VarPtr(vTableAddress), pObjectPointer, 4

Dim vTableEntryAddress As Long
vTableEntryAddress = vTableAddress + 4 * pvTableIndex

Dim oldAddress As Long
CopyMemory VarPtr(oldAddress), vTableEntryAddress, 4

Dim oldProtect As Long
If VirtualProtect(vTableEntryAddress, 4, PAGE_EXECUTE_READWRITE, oldProtect) = 0 Then
    Dim ErrorCode As Long
    ErrorCode = GetLastError
    Stop
End If

CopyMemory vTableEntryAddress, VarPtr(pNewAddress), 4

VirtualProtect vTableEntryAddress, 4, oldProtect, oldProtect

gRedirectMethod = oldAddress
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
Const ProcName As String = "gSetVariant"
On Error GoTo Err

If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortDoublesAsc( _
                ByRef Data() As Double, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Const ProcName As String = "gSortDoublesAsc"
On Error GoTo Err

If endIndex = 0 Then endIndex = UBound(Data)

If endIndex <= startIndex Then Exit Sub

Dim lowIndex As Long
lowIndex = startIndex

Dim highIndex As Long
highIndex = endIndex
  
Dim midIndex As Long
midIndex = (startIndex + endIndex) \ 2
    
Dim tempValue As Double
tempValue = Data(midIndex)
    
Do While (lowIndex <= highIndex)
    
    Do While (Data(lowIndex) < tempValue And lowIndex < endIndex)
        lowIndex = lowIndex + 1
    Loop
     
    Do While (tempValue < Data(highIndex) And highIndex > startIndex)
        highIndex = highIndex - 1
    Loop
           
    If (lowIndex <= highIndex) Then
        Dim tempHold As Double
        tempHold = Data(lowIndex)
        Data(lowIndex) = Data(highIndex)
        Data(highIndex) = tempHold
        lowIndex = lowIndex + 1
        highIndex = highIndex - 1
    End If
    
Loop
        
If (startIndex < highIndex) Then
    gSortDoublesAsc Data, startIndex, highIndex
End If
        
If (lowIndex < endIndex) Then
    gSortDoublesAsc Data, lowIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortDoublesDesc( _
                ByRef Data() As Double, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Const ProcName As String = "gSortDoublesDesc"
On Error GoTo Err

If endIndex = 0 Then endIndex = UBound(Data)

If endIndex <= startIndex Then Exit Sub

Dim lowIndex As Long
lowIndex = startIndex

Dim highIndex As Long
highIndex = endIndex
  
Dim midIndex As Long
midIndex = (startIndex + endIndex) \ 2
    
Dim tempValue As Double
tempValue = Data(midIndex)
    
Do While (lowIndex <= highIndex)
    
    Do While (Data(lowIndex) > tempValue And lowIndex < endIndex)
        lowIndex = lowIndex + 1
    Loop
     
    Do While (tempValue > Data(highIndex) And highIndex > startIndex)
        highIndex = highIndex - 1
    Loop
           
    If (lowIndex <= highIndex) Then
        Dim tempHold As Double
        tempHold = Data(lowIndex)
        Data(lowIndex) = Data(highIndex)
        Data(highIndex) = tempHold
        lowIndex = lowIndex + 1
        highIndex = highIndex - 1
    End If
    
Loop
        
If (startIndex < highIndex) Then
    gSortDoublesDesc Data, startIndex, highIndex
End If
        
If (lowIndex < endIndex) Then
    gSortDoublesDesc Data, lowIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortLongsAsc( _
                ByRef Data() As Long, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Dim lowIndex As Long
Dim highIndex As Long
Dim midIndex As Long
Dim tempValue As Long
Dim tempHold As Long

Const ProcName As String = "gSortLongsAsc"

On Error GoTo Err

If endIndex = 0 Then endIndex = UBound(Data)

If endIndex <= startIndex Then Exit Sub

lowIndex = startIndex
highIndex = endIndex
  
midIndex = (startIndex + endIndex) \ 2
    
tempValue = Data(midIndex)
    
Do While (lowIndex <= highIndex)
    
    Do While (Data(lowIndex) < tempValue And lowIndex < endIndex)
        lowIndex = lowIndex + 1
    Loop
     
    Do While (tempValue < Data(highIndex) And highIndex > startIndex)
        highIndex = highIndex - 1
    Loop
           
    If (lowIndex <= highIndex) Then
        tempHold = Data(lowIndex)
        Data(lowIndex) = Data(highIndex)
        Data(highIndex) = tempHold
        lowIndex = lowIndex + 1
        highIndex = highIndex - 1
    End If
    
Loop
        
If (startIndex < highIndex) Then
    gSortLongsAsc Data, startIndex, highIndex
End If
        
If (lowIndex < endIndex) Then
    gSortLongsAsc Data, lowIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortLongsDesc( _
                ByRef Data() As Long, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Dim lowIndex As Long
Dim highIndex As Long
Dim midIndex As Long
Dim tempValue As Long
Dim tempHold As Long

Const ProcName As String = "gSortLongsDesc"

On Error GoTo Err

If endIndex = 0 Then endIndex = UBound(Data)

If endIndex <= startIndex Then Exit Sub

lowIndex = startIndex
highIndex = endIndex
  
midIndex = (startIndex + endIndex) \ 2
    
tempValue = Data(midIndex)
    
Do While (lowIndex <= highIndex)
    
    Do While (Data(lowIndex) > tempValue And lowIndex < endIndex)
        lowIndex = lowIndex + 1
    Loop
     
    Do While (tempValue > Data(highIndex) And highIndex > startIndex)
        highIndex = highIndex - 1
    Loop
           
    If (lowIndex <= highIndex) Then
        tempHold = Data(lowIndex)
        Data(lowIndex) = Data(highIndex)
        Data(highIndex) = tempHold
        lowIndex = lowIndex + 1
        highIndex = highIndex - 1
    End If
    
Loop
        
If (startIndex < highIndex) Then
    gSortLongsDesc Data, startIndex, highIndex
End If
        
If (lowIndex < endIndex) Then
    gSortLongsDesc Data, lowIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortObjectsAsc( _
                ByRef Data() As Object, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
  
Dim lowIndex As Long
Dim highIndex As Long
Dim midIndex As Long
Dim obj As IComparable

' holds the address pointer for one object when switching object references. It
' is necessary to do this by copying memory rather than setting references to
' unsure that the correct interface pointer is set back in the Data array after
' the switch
Dim tempHold As Long

Const ProcName As String = "gSortObjectsAsc"

On Error GoTo Err

If endIndex = 0 Then endIndex = UBound(Data)

If endIndex <= startIndex Then Exit Sub

lowIndex = startIndex
highIndex = endIndex
  
midIndex = (startIndex + endIndex) \ 2
      
Set obj = Data(midIndex)
      
Do While (lowIndex <= highIndex)

    Do While (obj.CompareTo(Data(lowIndex)) > 0 And lowIndex < endIndex)
        lowIndex = lowIndex + 1
    Loop
     
    Do While (obj.CompareTo(Data(highIndex)) < 0 And highIndex > startIndex)
        highIndex = highIndex - 1
    Loop
           
    If (lowIndex <= highIndex) Then
        CopyMemory VarPtr(tempHold), VarPtr(Data(lowIndex)), 4
        CopyMemory VarPtr(Data(lowIndex)), VarPtr(Data(highIndex)), 4
        CopyMemory VarPtr(Data(highIndex)), VarPtr(tempHold), 4
        lowIndex = lowIndex + 1
        highIndex = highIndex - 1
    End If
    
Loop
        
If (startIndex < highIndex) Then
    gSortObjectsAsc Data, startIndex, highIndex
End If
          
If (lowIndex < endIndex) Then
    gSortObjectsAsc Data, lowIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortObjectsDesc( _
                ByRef Data() As Object, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
  
Dim lowIndex As Long
Dim highIndex As Long
Dim midIndex As Long
Dim obj As IComparable

' holds the address pointer for one object when switching object references. It
' is necessary to do this by copying memory rather than setting references to
' unsure that the correct interface pointer is set back in the Data array after
' the switch
Dim tempHold As Long

Const ProcName As String = "gSortObjectsDesc"

On Error GoTo Err

If endIndex = 0 Then endIndex = UBound(Data)

If endIndex <= startIndex Then Exit Sub

lowIndex = startIndex
highIndex = endIndex
  
midIndex = (startIndex + endIndex) \ 2
      
Set obj = Data(midIndex)
      
Do While (lowIndex <= highIndex)

    Do While (obj.CompareTo(Data(lowIndex)) < 0 And lowIndex < endIndex)
        lowIndex = lowIndex + 1
    Loop
     
    Do While (obj.CompareTo(Data(highIndex)) > 0 And highIndex > startIndex)
        highIndex = highIndex - 1
    Loop
           
    If (lowIndex <= highIndex) Then
        CopyMemory VarPtr(tempHold), VarPtr(Data(lowIndex)), 4
        CopyMemory VarPtr(Data(lowIndex)), VarPtr(Data(highIndex)), 4
        CopyMemory VarPtr(Data(highIndex)), VarPtr(tempHold), 4
        lowIndex = lowIndex + 1
        highIndex = highIndex - 1
    End If
    
Loop
        
If (startIndex < highIndex) Then
    gSortObjectsDesc Data, startIndex, highIndex
End If
          
If (lowIndex < endIndex) Then
    gSortObjectsDesc Data, lowIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortSinglesAsc( _
                ByRef Data() As Single, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Dim lowIndex As Long
Dim highIndex As Long
Dim midIndex As Long
Dim tempValue As Single
Dim tempHold As Single

Const ProcName As String = "gSortSinglesAsc"

On Error GoTo Err

If endIndex = 0 Then endIndex = UBound(Data)

If endIndex <= startIndex Then Exit Sub

lowIndex = startIndex
highIndex = endIndex
  
midIndex = (startIndex + endIndex) \ 2
    
tempValue = Data(midIndex)
    
Do While (lowIndex <= highIndex)
    
    Do While (Data(lowIndex) < tempValue And lowIndex < endIndex)
        lowIndex = lowIndex + 1
    Loop
     
    Do While (tempValue < Data(highIndex) And highIndex > startIndex)
        highIndex = highIndex - 1
    Loop
           
    If (lowIndex <= highIndex) Then
        tempHold = Data(lowIndex)
        Data(lowIndex) = Data(highIndex)
        Data(highIndex) = tempHold
        lowIndex = lowIndex + 1
        highIndex = highIndex - 1
    End If
    
Loop
        
If (startIndex < highIndex) Then
    gSortSinglesAsc Data, startIndex, highIndex
End If
        
If (lowIndex < endIndex) Then
    gSortSinglesAsc Data, lowIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortSinglesDesc( _
                ByRef Data() As Single, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Dim lowIndex As Long
Dim highIndex As Long
Dim midIndex As Long
Dim tempValue As Single
Dim tempHold As Single

Const ProcName As String = "gSortSinglesDesc"

On Error GoTo Err

If endIndex = 0 Then endIndex = UBound(Data)

If endIndex <= startIndex Then Exit Sub

lowIndex = startIndex
highIndex = endIndex
  
midIndex = (startIndex + endIndex) \ 2
    
tempValue = Data(midIndex)
    
Do While (lowIndex <= highIndex)
    
    Do While (Data(lowIndex) > tempValue And lowIndex < endIndex)
        lowIndex = lowIndex + 1
    Loop
     
    Do While (tempValue > Data(highIndex) And highIndex > startIndex)
        highIndex = highIndex - 1
    Loop
           
    If (lowIndex <= highIndex) Then
        tempHold = Data(lowIndex)
        Data(lowIndex) = Data(highIndex)
        Data(highIndex) = tempHold
        lowIndex = lowIndex + 1
        highIndex = highIndex - 1
    End If
    
Loop
        
If (startIndex < highIndex) Then
    gSortSinglesDesc Data, startIndex, highIndex
End If
        
If (lowIndex < endIndex) Then
    gSortSinglesDesc Data, lowIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortStringsAsc( _
                ByRef Data() As String, _
                Optional ByVal compareMethod As VBA.VbCompareMethod = vbTextCompare, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Dim lowIndex As Long
Dim highIndex As Long
Dim midIndex As Long
Dim tempValue As String
Dim tempHold As String

Const ProcName As String = "gSortStringsAsc"

On Error GoTo Err

Select Case compareMethod
Case vbBinaryCompare
Case vbTextCompare
Case Else
    gAssert False, "Only binary comparison and text comparison are permitted"
End Select

If endIndex = 0 Then endIndex = UBound(Data)

If endIndex <= startIndex Then Exit Sub

lowIndex = startIndex
highIndex = endIndex

midIndex = (startIndex + endIndex) \ 2
    
tempValue = Data(midIndex)
    
Do While (lowIndex <= highIndex)
    
    Do While (StrComp(Data(lowIndex), tempValue, compareMethod) = -1 And lowIndex < endIndex)
        lowIndex = lowIndex + 1
    Loop
     
    Do While (StrComp(tempValue, Data(highIndex), compareMethod) = -1 And highIndex > startIndex)
        highIndex = highIndex - 1
    Loop
           
    If (lowIndex <= highIndex) Then
        tempHold = Data(lowIndex)
        Data(lowIndex) = Data(highIndex)
        Data(highIndex) = tempHold
        lowIndex = lowIndex + 1
        highIndex = highIndex - 1
    End If
    
Loop
        
If (startIndex < highIndex) Then
    gSortStringsAsc Data, compareMethod, startIndex, highIndex
End If
        
If (lowIndex < endIndex) Then
    gSortStringsAsc Data, compareMethod, lowIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortStringsDesc( _
                ByRef Data() As String, _
                Optional ByVal compareMethod As VBA.VbCompareMethod = vbTextCompare, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
Dim lowIndex As Long
Dim highIndex As Long
Dim midIndex As Long
Dim tempValue As String
Dim tempHold As String

Const ProcName As String = "gSortStringsDesc"

On Error GoTo Err

Select Case compareMethod
Case vbBinaryCompare
Case vbTextCompare
Case Else
    gAssert False, "Only binary comparison and text comparison are permitted"
End Select

If endIndex = 0 Then endIndex = UBound(Data)

If endIndex <= startIndex Then Exit Sub

lowIndex = startIndex
highIndex = endIndex

midIndex = (startIndex + endIndex) \ 2
    
tempValue = Data(midIndex)
    
Do While (lowIndex <= highIndex)
    
    Do While (StrComp(Data(lowIndex), tempValue, compareMethod) = 1 And lowIndex < endIndex)
        lowIndex = lowIndex + 1
    Loop
     
    Do While (StrComp(tempValue, Data(highIndex), compareMethod) = 1 And highIndex > startIndex)
        highIndex = highIndex - 1
    Loop
           
    If (lowIndex <= highIndex) Then
        tempHold = Data(lowIndex)
        Data(lowIndex) = Data(highIndex)
        Data(highIndex) = tempHold
        lowIndex = lowIndex + 1
        highIndex = highIndex - 1
    End If
    
Loop
        
If (startIndex < highIndex) Then
    gSortStringsDesc Data, compareMethod, startIndex, highIndex
End If
        
If (lowIndex < endIndex) Then
    gSortStringsDesc Data, compareMethod, lowIndex, endIndex
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gSortTypedObjects( _
                ByVal arPtr As Long, _
                ByVal number As Long, _
                ByVal descending As Boolean, _
                Optional ByVal startIndex As Long, _
                Optional ByVal endIndex As Long)
  
Const ProcName As String = "gSortTypedObjects"

On Error GoTo Err

ReDim objs(number - 1) As Object
gCopyToObjects arPtr, objs

If descending Then
    gSortObjectsDesc objs, startIndex, endIndex
Else
    gSortObjectsAsc objs, startIndex, endIndex
End If

gCopyFromObjects objs, arPtr

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function gStringToHexString( _
                ByVal Value As String) As String
Const ProcName As String = "gStringToHexString"

On Error GoTo Err

ReDim ar(2 * Len(Value) - 1) As Byte

CopyMemory VarPtr(ar(0)), StrPtr(Value), 2 * Len(Value)
gStringToHexString = gBytesToHexString(ar)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub gTerminate()
Const ProcName As String = "gTerminate"
On Error Resume Next

If Not mInitialised Then Exit Sub

mTerminated = True
mInitialised = False

GIntervalTimer.gTerm
TimestampGlobals.gTerm
GTimerList.gTerm

unhook

gLogger.Log "TWUtilities uninitialisation completed", ProcName, ModuleName, LogLevelDetail

gLogManager.Finish
End Sub

Public Function gAdjustColorIntensity(ByVal pColor As Long, ByVal pFactor As Double) As Long
gAssertArgument pFactor < 1#, "pFactor must be < 1.0"

If (pColor And &H80000000) Then pColor = GetSysColor(pColor And &HFFFFFF)

gAdjustColorIntensity = (((pColor And &HFF0000) * pFactor) And &HFF0000) + _
            (((pColor And &HFF00&) * pFactor) And &HFF00) + _
            ((pColor And &HFF&) * pFactor)
End Function

Public Function gTrimNull(s As String) As String
Const ProcName As String = "gTrimNull"
On Error GoTo Err

gTrimNull = Left$(s, lstrlenW(StrPtr(s)))

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gUnsignedAdd( _
                ByVal num1 As Long, _
                ByVal num2 As Long) As Long
If num1 And &H80000000 Then
    gUnsignedAdd = ((num1 And &H7FFFFFFF) + num2) Or &H80000000
ElseIf &H7FFFFFFF - num1 >= num2 Then
    gUnsignedAdd = num1 + num2
Else
    gUnsignedAdd = (num2 - (&H7FFFFFFF - num1) - 1) Or &H80000000
End If
End Function

Public Function gVariantEquals(ByVal p1 As Variant, ByVal p2 As Variant) As Boolean
If IsMissing(p2) Or IsEmpty(p2) Then
    gVariantEquals = False
ElseIf IsNumeric(p1) And IsNumeric(p2) Then
    gVariantEquals = (p1 = p2)
ElseIf IsArray(p1) Then
    gVariantEquals = False
ElseIf IsObject(p1) And IsObject(p2) Then
    gVariantEquals = (p1 Is p2)
Else
    gVariantEquals = (p1 = p2)
End If
End Function

Public Function gVariantToString( _
                ByRef Value As Variant) As String
Const ProcName As String = "gVariantToString"
On Error GoTo Err

Dim baseType As VbVarType
baseType = VarType(Value) And (Not VbVarType.vbArray)

Dim s As String
If IsArray(Value) Then
    Dim sb As New StringBuilder
    Dim Var As Variant
    For Each Var In Value
        If sb.Length = 0 Then
            sb.Append "Array["
        Else
            sb.Append ", "
        End If
        If baseType = VbVarType.vbString Then
            sb.Append """"
            sb.Append gVariantToString(Var)
            sb.Append """"
        Else
            sb.Append gVariantToString(Var)
            sb.Append s
        End If
    Next
    sb.Append "]"
    gVariantToString = sb.ToString
    
ElseIf IsObject(Value) Then
    If TypeOf Value Is IStringable Then
        Dim obj As IStringable
        Set obj = Value
        s = obj.ToString
    ElseIf TypeOf Value Is IJSONable Then
        Dim objJ As IJSONable
        Set objJ = Value
        s = objJ.ToJSON
    ElseIf baseType <> VbVarType.vbObject Then
        ' this means the Value has a default property - we'll use
        ' that Value instead
        s = "DefaultProp(" & varToString(Value) & ")"
    Else
        ' nothing we can find out about this object
        s = "?"
    End If
    gVariantToString = "Obj(" & TypeName(Value) & "){" & s & "}"
Else
    gVariantToString = varToString(Value)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gXMLEncode(ByVal Value As String) As String
Const ProcName As String = "gXMLEncode"
On Error GoTo Err

gXMLEncode = Replace(Replace(Replace(Replace(Replace(Value, "&", "&amp;"), "'", "&apos;"), """", "&quot;"), "<", "&lt;"), ">", "&gt;")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gXMLDecode(ByVal Value As String) As String
Const ProcName As String = "gXMLDecode"

On Error GoTo Err

gXMLDecode = Replace(Replace(Replace(Replace(Value, "&amp;", "&"), "&apos;", "'"), "&lt;", "<"), "&gt;", ">")

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub Main()
gInitialise
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function convertHexToByte( _
                ByVal asc1 As Integer, _
                ByVal asc2 As Integer) As Byte
Const ProcName As String = "convertHexToByte"
On Error GoTo Err

convertHexToByte = 16 * hexCharToBin(asc1) + hexCharToBin(asc2)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function findMainWindowHandle() As Long
Const ProcName As String = "findMainWindowHandle"
On Error GoTo Err

mMainWindowHandle = mPostMessageForm.hwnd
Do
    Dim hwnd As Long
    hwnd = GetWindow(mMainWindowHandle, GW_OWNER)
    If hwnd = 0 Then Exit Do
    mMainWindowHandle = hwnd
Loop

findMainWindowHandle = mMainWindowHandle
Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getbyte(ByVal addr As Long) As Byte
Dim b As Byte
CopyMemory VarPtr(b), addr, 1
getbyte = b
End Function

Private Function getCommandLineLength(ByVal pBuffAddress) As Long
Dim i As Long
Do While True
    If getbyte(pBuffAddress + i) = 0 And getbyte(pBuffAddress + i + 1) = 0 Then
        getCommandLineLength = i
        Exit Do
    End If
    i = i + 2
Loop
End Function

Private Function hexCharToBin( _
                ByVal hexChar As Integer) As Integer
Const ProcName As String = "hexCharToBin"
On Error GoTo Err

gAssertArgument hexChar >= vbKey0, "Invalid hex Character"
If hexChar <= vbKey9 Then hexCharToBin = hexChar - vbKey0: Exit Function

gAssertArgument hexChar >= vbKeyA, "Invalid hex Character"
If hexChar <= vbKeyF Then hexCharToBin = 10 + hexChar - vbKeyA: Exit Function

hexChar = hexChar - 32
gAssertArgument hexChar >= vbKeyA, "Invalid hex Character"
If hexChar <= vbKeyF Then hexCharToBin = 10 + hexChar - vbKeyA: Exit Function

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub hook()
Const ProcName As String = "hook"
On Error GoTo Err

If mPrevWndProc <> 0 Then Exit Sub

Set mPostMessageForm = New PostMessageForm

mPrevWndProc = SetWindowLong(findMainWindowHandle, GWL_WNDPROC, AddressOf WindowProc)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub initHexChars()
mHexChars(0) = AscB("0")
mHexChars(1) = AscB("1")
mHexChars(2) = AscB("2")
mHexChars(3) = AscB("3")
mHexChars(4) = AscB("4")
mHexChars(5) = AscB("5")
mHexChars(6) = AscB("6")
mHexChars(7) = AscB("7")
mHexChars(8) = AscB("8")
mHexChars(9) = AscB("9")
mHexChars(10) = AscB("A")
mHexChars(11) = AscB("B")
mHexChars(12) = AscB("C")
mHexChars(13) = AscB("D")
mHexChars(14) = AscB("E")
mHexChars(15) = AscB("F")
End Sub

Private Sub initPrintableKeyCodes()
Dim i As Long
For i = 1 To UBound(mPrintableKeyCodes)
    Select Case i
    Case KeyCodeConstants.vbKeyBack, _
        KeyCodeConstants.vbKeyCancel, _
        KeyCodeConstants.vbKeyCapital, _
        KeyCodeConstants.vbKeyClear, _
        KeyCodeConstants.vbKeyControl, _
        KeyCodeConstants.vbKeyDelete, _
        KeyCodeConstants.vbKeyDown, _
        KeyCodeConstants.vbKeyEnd, _
        KeyCodeConstants.vbKeyEscape, _
        KeyCodeConstants.vbKeyExecute
        
    Case KeyCodeConstants.vbKeyF1, _
        KeyCodeConstants.vbKeyF2, _
        KeyCodeConstants.vbKeyF3, _
        KeyCodeConstants.vbKeyF4, _
        KeyCodeConstants.vbKeyF5, _
        KeyCodeConstants.vbKeyF6, _
        KeyCodeConstants.vbKeyF7, _
        KeyCodeConstants.vbKeyF8, _
        KeyCodeConstants.vbKeyF9, _
        KeyCodeConstants.vbKeyF10, _
        KeyCodeConstants.vbKeyF11, _
        KeyCodeConstants.vbKeyF12, _
        KeyCodeConstants.vbKeyF13, _
        KeyCodeConstants.vbKeyF14, _
        KeyCodeConstants.vbKeyF15, _
        KeyCodeConstants.vbKeyF16
    
    Case KeyCodeConstants.vbKeyHelp, _
        KeyCodeConstants.vbKeyHome, _
        KeyCodeConstants.vbKeyInsert, _
        KeyCodeConstants.vbKeyLButton, _
        KeyCodeConstants.vbKeyLeft, _
        KeyCodeConstants.vbKeyMButton, _
        KeyCodeConstants.vbKeyMenu, _
        KeyCodeConstants.vbKeyNumlock, _
        KeyCodeConstants.vbKeyPageDown, _
        KeyCodeConstants.vbKeyPause, _
        KeyCodeConstants.vbKeyPrint, _
        KeyCodeConstants.vbKeyRButton, _
        KeyCodeConstants.vbKeyReturn, _
        KeyCodeConstants.vbKeyRight, _
        KeyCodeConstants.vbKeyScrollLock, _
        KeyCodeConstants.vbKeySelect, _
        KeyCodeConstants.vbKeySeparator, _
        KeyCodeConstants.vbKeyShift, _
        KeyCodeConstants.vbKeySnapshot, _
        KeyCodeConstants.vbKeyTab, _
        KeyCodeConstants.vbKeyUp
    Case Else
        mPrintableKeyCodes(i) = True
    End Select

Next
End Sub

Private Sub unhook()
Const ProcName As String = "unhook"
On Error GoTo Err

If mPrevWndProc = 0 Then Exit Sub
SetWindowLong findMainWindowHandle, GWL_WNDPROC, mPrevWndProc
mPrevWndProc = 0

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function varToString( _
                ByRef Value As Variant) As String
Const ProcName As String = "varToString"
On Error GoTo Err

gAssertArgument Not IsArray(Value), "Argument must not be an array variant"

Select Case VarType(Value)
Case VbVarType.vbBoolean, _
        VbVarType.vbCurrency, _
        VbVarType.vbDecimal, _
        VbVarType.vbDouble, _
        VbVarType.vbError, _
        VbVarType.vbInteger, _
        VbVarType.vbLong, _
        VbVarType.vbSingle, _
        VbVarType.vbString
    varToString = CStr(Value)
Case VbVarType.vbByte
    varToString = Hex$(Value)
Case VbVarType.vbDataObject
    Dim dataObj As DataObject
    Set dataObj = Value
    
    Dim s As String
    Dim fn As Variant
    For Each fn In dataObj
        If Len(s) <> 0 Then s = s & ", "
        s = s & fn
    Next
    varToString = "DataObject{" & s & "}"
Case VbVarType.vbDate
    If Int(Value) = Value Then
        varToString = gFormatTimestamp(Value, TimestampDateOnlyISO8601)
    ElseIf Int(Value) = 0 Then
        varToString = gFormatTimestamp(Value, TimestampTimeOnlyISO8601)
    Else
        varToString = gFormatTimestamp(Value, TimestampDateAndTimeISO8601)
    End If
Case VbVarType.vbEmpty
    varToString = "EMPTY"
Case VbVarType.vbNull
    varToString = "NULL"
Case VbVarType.vbObject
    gAssertArgument False, "Argument must not be an object with no default property"
Case VbVarType.vbUserDefinedType
    varToString = "UDT{?}"

End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function WindowProc( _
                ByVal hwnd As Long, _
                ByVal uMsg As Long, _
                ByVal wParam As Long, _
                ByVal lParam As Long) As Long
Const ProcName As String = "WindowProc"
On Error GoTo Err

If uMsg = WM_TIMECHANGE Then
    WindowProc = CallWindowProc(mPrevWndProc, hwnd, uMsg, wParam, lParam)
    gSetBaseTimes
    gResetClocks
    Debug.Print "Time changed"
    gLogger.Log "Time changed", ProcName, ModuleName
ElseIf uMsg = WM_NCDESTROY Or uMsg = WM_CLOSE Then
    Globals.gTerminate
    WindowProc = CallWindowProc(mPrevWndProc, hwnd, uMsg, wParam, lParam)
ElseIf uMsg = UserMessages.UserMessageScheduleTasks Then
    DoEvents
    gTaskManager.ScheduleTasks
ElseIf uMsg = UserMessages.UserMessageTimer Then
    GIntervalTimer.gProcessUserTimerMsg wParam
Else
    WindowProc = CallWindowProc(mPrevWndProc, hwnd, uMsg, wParam, lParam)
End If

Exit Function

Err:
gNotifyUnhandledError ProcName, ModuleName
End Function


