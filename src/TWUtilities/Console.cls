VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' The <code>Console</code> object gives access to the program's console.
'
' This object cannot be created with the <code>New</code> operator: call the
' global <code>GetConsole</code> method.
'
' Note that programs created with Visual Basic are GUI programs unless special
' Action is taken. GUI programs are not given a console when they are Run. Calling
' <code>GetConsole</code> within a GUI program causes a console to be allocated
' which can then be used for input and output via the <code>Console</code> object.
' The console that is allocated to a GUI program that calls <code>GetConsole</code>
' is separate from any console (eg a command prompt) that was used to Run the program.
'
' To convert a GUI program into a console program (which is always given a console
' when Run), the .exe file created by Visual Basic needs to be modified. One way
' to do this is with a command like the following:
' <pre>
' <code>
'   link /EDIT /SUBSYSTEM:CONSOLE myprog.exe
' </code>
' </pre>
' where myprog.exe is the path and filename of your program. The link tool is
' to be found in C:\Program Files\Microsoft Visual Studio\VB98, so you
' can either Run it from there or add that folder to your PATH environment
' variable, and then you can Run the command in the folder that contains your
' program.
'
' If you make a program that uses the <code>Console</code> class into a console
' program in this way, then you can Run it from the command line and use the
' <code>&lt;</code>, <code>&gt;</code>, <code>&gt;&gt;</code> and
' <code>&brvbar;</code> command line operators in the usual way.
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


Private Const ModuleName                    As String = "Console"

'@================================================================================
' Member variables
'@================================================================================

Private mHwnd                               As Long

Private mStdInHandle                        As Long
Private mStdOutHandle                       As Long
Private mStdErrHandle                       As Long

' A handle to the active screen buffer, even if STDOUT and STDERR have been redirected
Private mConOutHandle                       As Long

' A handle to the active console input buffer, even if STDIN has been redirected
Private mConInHandle                       As Long

Private mEoFString                          As String

Private mBuf(8191)                          As Byte

Private mStrBld                             As StringBuilder
' the position at which the next input character will be placed in mStrBld
Private mCurrPosition                       As Long

Private mAtEof                              As Boolean

Private mOrigConsoleModeStdIn               As Long

Private mOrigConsoleCursorInfo              As CONSOLE_CURSOR_INFO
Private mCurrConsoleCursorInfo              As CONSOLE_CURSOR_INFO

Private mOverwriteMode                      As Boolean

Private mOrigTitle                          As String
Private mTitle                              As String

Private mRecbuf()                           As Integer

Private mForegroundColor                    As ConsoleColors
Private mBackgroundColor                    As ConsoleColors

' The cursor position immediately prior to starting to receive input
Private mPrevCursorPosition                 As COORD

Private mReadingFromConsole                 As Boolean
Private mPrompt                             As String

Private mScreenBufferSize                   As COORD

Private mInputHistory()                     As String
Private mInputHistoryIndex                  As Long
Private mCurrHistoryIndex                   As Long

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Const ProcName As String = "Class_Initialize"
On Error GoTo Err

' if we don't already have a console, get one
AllocConsole

mHwnd = GetConsoleWindow

mStdInHandle = GetStdHandle(STD_INPUT_HANDLE)
mStdOutHandle = GetStdHandle(STD_OUTPUT_HANDLE)
mStdErrHandle = GetStdHandle(STD_ERROR_HANDLE)

mConOutHandle = CreateFile("CONOUT$", _
                            GENERIC_READ Or GENERIC_WRITE, _
                            FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                            0, _
                            OPEN_EXISTING, _
                            0, _
                            0)

mConInHandle = CreateFile("CONIn$", _
                            GENERIC_READ Or GENERIC_WRITE, _
                            FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                            0, _
                            OPEN_EXISTING, _
                            0, _
                            0)

mEoFString = Chr$(26) ' set default EoF string to Ctrl-Z

GetConsoleMode mConInHandle, mOrigConsoleModeStdIn

SetConsoleMode mConInHandle, _
                ENABLE_WINDOW_INPUT Or _
                ENABLE_MOUSE_INPUT

    GetConsoleCursorInfo mConOutHandle, mOrigConsoleCursorInfo
    
    mCurrConsoleCursorInfo.bVisible = 1
    mCurrConsoleCursorInfo.dwSize = 10
    SetConsoleCursorInfo mConOutHandle, mCurrConsoleCursorInfo

mOrigTitle = Title

Dim screenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO
GetConsoleScreenBufferInfo mConOutHandle, screenBufferInfo
If Err.LastDllError = 0 Then
    mForegroundColor = screenBufferInfo.wAttributes And &HF
    mBackgroundColor = (screenBufferInfo.wAttributes And &HF0) / &H10
    
    mScreenBufferSize = screenBufferInfo.dwSize
End If

Set mStrBld = New StringBuilder
mStrBld.Initialise 80

ReDim mInputHistory(19) As String

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Class_Terminate()
Const ProcName As String = "Class_Terminate"
On Error GoTo Err

Title = mOrigTitle

If StdInType = FileTypeChar Then
    SetConsoleMode mConInHandle, mOrigConsoleModeStdIn
End If
    
SetConsoleCursorInfo mConOutHandle, mOrigConsoleCursorInfo

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

''
' Sets the background console colour.
'
' @remarks
'   The value supplied must be a member of the <code>ConsoleColors</code> enum, not
'   an RGB value.
' @param Value
'   The new background console colour.
'@/
Public Property Let BackgroundColor( _
                ByVal Value As ConsoleColors)
Const ProcName As String = "BackgroundColor"
On Error GoTo Err

If Not isValidConsoleColor(Value) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid ConsoleColor"

mBackgroundColor = Value

If Err.LastDllError = 0 Then
    Dim attrsWritten As Long
    FillConsoleOutputAttribute mConOutHandle, _
                                mForegroundColor Or (&H10 * mBackgroundColor), _
                                &H7FFFFFFF, _
                                0, _
                                VarPtr(attrsWritten)
    SetConsoleTextAttribute mConOutHandle, mForegroundColor Or (&H10 * mBackgroundColor)
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the current background console colour.
'
' @remarks
'   The value returned is a member of the <code>ConsoleColors</code> enum, not
'   an RGB value.
' @return
'   The current background console colour.
'@/
Public Property Get BackgroundColor() As ConsoleColors
BackgroundColor = mBackgroundColor
End Property

''
' Sets the string that is returned by the <code>ReadLine</code> method when the
' end of the input file is reached.
'
'   The default setting of this property is Chr(26), which is Ctrl-Z.
'
' @remarks
'   If this is set to an empty string, then an error <code>ErrorCodes.ErrEndOfFile</code>
'   is raised when end of file is reached.
'
'   Note that if this string is received in the console input data stream, it is passed
'   through to the application, but it does not prevent the application from reading
'   further input if there is any.
'
' @param Value
'   The string to be set as the end-of-file marker.
'@/
Public Property Let EofString( _
                ByVal Value As String)
mEoFString = Value
End Property

''
' Returns the string that is returned by the <code>ReadLine</code> method when the
' end of the input file is reached.
'
' @return
'   The string that is returned when end-of-file is reached.
'@/
Public Property Get EofString() As String
EofString = mEoFString
End Property

''
' Sets the foreground console colour.
'
' @remarks
'   The value supplied must be a member of the <code>ConsoleColors</code> enum, not
'   an RGB value.
' @param Value
'   The new foreground console colour.
'@/
Public Property Let ForegroundColor( _
                ByVal Value As ConsoleColors)
Const ProcName As String = "ForegroundColor"
On Error GoTo Err

If Not isValidConsoleColor(Value) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Value is not a valid ConsoleColor"

mForegroundColor = Value

Dim attrsWritten As Long
FillConsoleOutputAttribute mConOutHandle, _
                            mForegroundColor Or (&H10 * mBackgroundColor), _
                            &H7FFFFFFF, _
                            0, _
                            VarPtr(attrsWritten)
SetConsoleTextAttribute mConOutHandle, mForegroundColor Or (&H10 * mBackgroundColor)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the current foreground console colour.
'
' @remarks
'   The value returned is a member of the <code>ConsoleColors</code> enum, not
'   an RGB value.
' @return
'   The current foreground console colour.
'@/
Public Property Get ForegroundColor() As ConsoleColors
ForegroundColor = mForegroundColor
End Property

''
' Returns the handle to the console window.
'
' @return
'   The handle to the console window.
'@/
Public Property Get hWnd() As Long
hWnd = mHwnd
End Property

''
' Returns the file type for the StdErr stream.
'
' @return
'   The file type for the StdErr stream.
'@/
Public Property Get StdErrType() As FileTypes
Const ProcName As String = "StdErrType"
On Error GoTo Err

StdErrType = GetFileType(mStdErrHandle)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the file type for the StdIn stream.
'
' @return
'   The file type for the StdIn stream.
'@/
Public Property Get StdInType() As FileTypes
Const ProcName As String = "StdInType"
On Error GoTo Err

StdInType = GetFileType(mStdInHandle)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the file type for the StdOut stream.
'
' @return
'   The file type for the StdOut stream.
'@/
Public Property Get StdOutType() As FileTypes
Const ProcName As String = "StdOutType"
On Error GoTo Err

StdOutType = GetFileType(mStdOutHandle)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Sets the console window's Title.
'
' @param Value
'   The Title to be set for the console window.
'@/
Public Property Let Title( _
                ByVal Value As String)
Const ProcName As String = "Title"
On Error GoTo Err

SetConsoleTitle StrPtr(Value)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the console window's Title.
'
' @return
'   The console window's Title.
'@/
Public Property Get Title() As String
Const ProcName As String = "Title"
On Error GoTo Err

Dim titleLen As Long
Dim ar(2 * MAX_PATH - 1) As Byte
titleLen = GetConsoleTitle(VarPtr(ar(0)), MAX_PATH)

ReDim ar1(2 * titleLen - 1) As Byte
CopyMemory VarPtr(ar1(0)), VarPtr(ar(0)), 2 * titleLen
Title = ar1

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Reads the next line of input from the StdIn stream, excluding the carriage return
' and line feed at the end.
'
' @remarks
'   Keys typed by the user are echoed to the console, except that if the
'   <code>echoChar</code> argument is supplied they are replaced by this character.
'
'   If the StdIn stream is the console window, output to the console may occur
'   while the user is typing (eg calls to <code>writeLine</code> in an event handler).
'   Such output appears before the line that the user
'   is typing on (the previous content is scrolled up as required to make room for
'   the output).
' @return
'   The next line of input. This may be an empty string, indicating that there are
'   system messages waiting to be processed: the application should call DoEvents
'   and then call this method again. Note that this case is indistinguishable from the
'   user simply entering a blank line.
' @param prompt
'   This is output to the console on a new line and the cursor is positioned
'   immediately after it.
' @param echoChar
'   Specifies a single character that is echoed to the console instead of each character
'   the user actually types. If this argument is omitted or is an empty string, the
'   actual characters typed by the user are echoed. If it has the special value Chr(0),
'   then nothing the user types is echoed.
'@/
Public Function ReadLine( _
                Optional ByVal prompt As String, _
                Optional ByVal echoChar As String) As String
Const ProcName As String = "ReadLine"
On Error GoTo Err

If Len(echoChar) > 1 Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "echoChar must be a single character"

If StdInType = FileTypeChar Then
    ReadLine = ReadLineFromConsole(prompt, echoChar)
Else
    ReadLine = readLineFromFile
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Reads a line of input from the console, excluding the carriage return
' and line feed at the end. Note that this will accept input from the console
' even if the STDIN stream has been redirected.
'
' @remarks
'   Keys typed by the user are echoed to the console, except that if the
'   <code>echoChar</code> argument is supplied they are replaced by this character.
'
'   Output to the console may occur
'   while the user is typing (eg calls to <code>writeLine</code> in an event handler).
'   Such output appears before the line that the user
'   is typing on (the previous content is scrolled up as required to make room for
'   the output).
' @return
'   The next line of input. This may be an empty string, indicating that there are
'   system messages waiting to be processed: the application should call DoEvents
'   and then call this method again. Note that this case is indistinguishable from the
'   user simply entering a blank line.
' @param prompt
'   This is output to the console on a new line and the cursor is positioned
'   immediately after it.
' @param echoChar
'   Specifies a single character that is echoed to the console instead of each character
'   the user actually types. If this argument is omitted or is an empty string, the
'   actual characters typed by the user are echoed. If it has the special value Chr(0),
'   then nothing the user types is echoed.
'@/
Public Function ReadLineFromConsole( _
                ByVal prompt As String, _
                ByVal echoChar As String) As String
Const ProcName As String = "ReadLineFromConsole"
On Error GoTo Err

Static sNoInputOnPreviousCall As Boolean

mReadingFromConsole = True
If Not sNoInputOnPreviousCall Then
    mPrompt = prompt
    clearUserInput
    writePromptAndUserInput
ElseIf prompt <> mPrompt Then
    mPrompt = prompt
    writePromptAndUserInput
Else
    ' don't do anything to the screen
End If

Do
    Dim s As String
    Do
        s = ""
        If Not readInputFromConsole(echoChar, s) Then
            sNoInputOnPreviousCall = True
            mReadingFromConsole = False
            ReadLineFromConsole = ""
            Exit Function
        End If
        sNoInputOnPreviousCall = False
    Loop Until s <> ""
    
    If Right$(s, 2) = vbCrLf Then Exit Do
    
    If mOverwriteMode Then
        mStrBld.Overwrite s, mCurrPosition
    Else
        mStrBld.Insert s, mCurrPosition
    End If
    
    mCurrPosition = mCurrPosition + 1
Loop

ReadLineFromConsole = mStrBld.ToString
mReadingFromConsole = False

If ReadLineFromConsole <> "" Then updateInputHistory ReadLineFromConsole

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

''
' Writes a line to the StdErr stream. A carriage return and line feed are appended to
' the supplied string.
'
' @param msg
'   The line to be written to StdErr.
'@/
Public Sub WriteErrorLine(msg As String)
Const ProcName As String = "WriteErrorLine"
On Error GoTo Err

WriteErrorString msg
WriteErrorString vbCrLf

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Writes a string to the StdErr stream. No carriage return or line feed are appended.
'
' @param msg
'   The string to be written to StdErr.
'@/
Public Sub WriteErrorString( _
                msg As String)
Const ProcName As String = "WriteErrorString"
On Error GoTo Err

If StdErrType = FileTypeChar Then
    If StdOutType = FileTypeChar Then
        If mReadingFromConsole Then
            ' we need to move the prompt and anything the user has already typed out of the
            ' way. We do this by clearing it all out, then writing the msg argument,
            ' then rewriting the prompt and the input so far
            clearPromptAndUserInput
            writeErrorStringToConsole msg
            writePromptAndUserInput
        Else
            writeErrorStringToConsole msg
        End If
    Else
        writeErrorStringToConsole msg
    End If
Else
    Dim bytesWritten As Long
    WriteFile mStdErrHandle, StrPtr(StrConv(msg, vbFromUnicode)), Len(msg), bytesWritten, 0&
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Writes a line to the StdOut stream. A carriage return and line feed are appended to
' the supplied string.
'
' @param msg
'   The line to be written to StdOut.
'@/
Public Sub WriteLine( _
                msg As String)
Const ProcName As String = "WriteLine"
On Error GoTo Err

WriteString msg & vbCrLf

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Writes a line to the console. A carriage return and line feed are appended to
' the supplied string.
'
' @param msg
'   The line to be written to the console.
'@/
Public Sub WriteLineToConsole( _
                msg As String)
Const ProcName As String = "WriteLineToConsole"
On Error GoTo Err

WriteStringToConsole msg & vbCrLf

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Writes a string to the StdOut stream. No carriage return or line feed are appended.
'
' @param msg
'   The string to be written to StdOut.
'@/
Public Sub WriteString( _
                msg As String)
Const ProcName As String = "WriteString"
On Error GoTo Err

If StdOutType = FileTypeChar Then
    WriteStringToConsole msg
Else
    Dim bytesWritten As Long
    WriteFile mStdOutHandle, StrPtr(StrConv(msg, vbFromUnicode)), Len(msg), bytesWritten, 0&
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

''
' Writes a string to the console. No carriage return or line feed are appended.
'
' @param msg
'   The string to be written to the console.
'@/
Public Sub WriteStringToConsole( _
                msg As String)
Const ProcName As String = "WriteStringToConsole"
On Error GoTo Err

If mReadingFromConsole Then
    ' we need to move the prompt and anything the user has already typed out of the
    ' way. We do this by clearing it all out, then writing the msg argument,
    ' then rewriting the prompt and the input so far
    Dim s() As String: s = Split(msg, vbCrLf)
    Dim i As Long
    For i = 0 To UBound(s)
        clearPromptAndUserInput
        writeToConsole msg & vbCrLf
        writePromptAndUserInput
    Next
Else
    writeToConsole msg
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub clearPromptAndUserInput()
Const ProcName As String = "clearPromptAndUserInput"
On Error GoTo Err

Dim currCursorPos As COORD
currCursorPos = getCursorPosition

Dim newCursorPos As COORD
newCursorPos.y = currCursorPos.y - Int((Len(mPrompt) + mCurrPosition) / mScreenBufferSize.x)
'don't need to calculate x because a prompt always starts at the beginning of a line
newCursorPos.x = 0

Dim numCharsWritten As Long
FillConsoleOutputCharacter mConOutHandle, _
                AscW(" "), _
                Len(mPrompt) + mStrBld.Length, _
                newCursorPos.y * &H10000, _
                numCharsWritten
                
If mPrevCursorPosition.x <> 0 Then
    ' Reset the cursor position to where it was immediately before the user
    ' started typing
    newCursorPos.y = newCursorPos.y - 1
    newCursorPos.x = mPrevCursorPosition.x
End If
setCursorPosition newCursorPos.x, newCursorPos.y

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub clearUserInput()
Const ProcName As String = "clearUserInput"
On Error GoTo Err

mStrBld.Clear
mCurrPosition = 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deleteCharAtCursor()
Const ProcName As String = "deleteCharAtCursor"
On Error GoTo Err

If mCurrPosition <> mStrBld.Length + 1 Then
    mStrBld.Remove mCurrPosition, 1
    
    Dim cursorPosn As COORD
    cursorPosn = getCursorPosition
    writeToConsole mStrBld.SubString(mCurrPosition)
    writeToConsole " "
    setCursorPosition cursorPosn.x, cursorPosn.y
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub echoToConsole( _
                ByVal Value As String, _
                ByVal echoChar As String, _
                Optional ByVal pOverwrite As Boolean)
Const ProcName As String = "echoToConsole"
On Error GoTo Err

If Not mReadingFromConsole Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Not reading from console - can't echo"

Dim str As String
If echoChar = "" Then
    writeToConsole Value
ElseIf echoChar <> Chr$(0) Then
    If Len(Value) > 0 Then
        str = String$(Len(Value), echoChar)
        writeToConsole str
    End If
End If
If Not pOverwrite Then
    If mCurrPosition < mStrBld.Length + 1 Then
        Dim cursorPosn As COORD
        cursorPosn = getCursorPosition
        
        str = mStrBld.SubString(mCurrPosition)
        writeToConsole str
        
        setCursorPosition cursorPosn.x, cursorPosn.y
        
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function getCursorPosition() As COORD
Const ProcName As String = "getCursorPosition"
On Error GoTo Err

Dim csbi As CONSOLE_SCREEN_BUFFER_INFO
GetConsoleScreenBufferInfo mConOutHandle, csbi
getCursorPosition = csbi.dwCursorPosition

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function isValidConsoleColor( _
                ByVal Value As ConsoleColors) As Boolean
Const ProcName As String = "isValidConsoleColor"
On Error GoTo Err

Select Case Value
Case ConsoleColors.ConsoleBlack, _
        ConsoleColors.ConsoleBlue, _
        ConsoleColors.ConsoleCyan, _
        ConsoleColors.ConsoleDarkBlue, _
        ConsoleColors.ConsoleDarkCyan, _
        ConsoleColors.ConsoleDarkGray, _
        ConsoleColors.ConsoleDarkGreen, _
        ConsoleColors.ConsoleDarkMagenta, _
        ConsoleColors.ConsoleDarkRed, _
        ConsoleColors.ConsoleDarkYellow, _
        ConsoleColors.ConsoleGray, _
        ConsoleColors.ConsoleGreen, _
        ConsoleColors.ConsoleMagenta, _
        ConsoleColors.ConsoleRed, _
        ConsoleColors.ConsoleWhite, _
        ConsoleColors.ConsoleYellow
    isValidConsoleColor = True
End Select

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub moveCursorEnd()
Const ProcName As String = "moveCursorEnd"
On Error GoTo Err

Dim cursorPosn As COORD
cursorPosn = getCursorPosition

cursorPosn.x = (Len(mPrompt) + mStrBld.Length) Mod mScreenBufferSize.x
cursorPosn.y = cursorPosn.y - _
                Int((Len(mPrompt) + mCurrPosition) / mScreenBufferSize.x) + _
                Int((Len(mPrompt) + mStrBld.Length) / mScreenBufferSize.x)

setCursorPosition cursorPosn.x, cursorPosn.y
mCurrPosition = mStrBld.Length + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub moveCursorHome()
Const ProcName As String = "moveCursorHome"
On Error GoTo Err

Dim cursorPosn As COORD
cursorPosn = getCursorPosition

cursorPosn.x = Len(mPrompt) Mod mScreenBufferSize.x
cursorPosn.y = cursorPosn.y - _
                Int((Len(mPrompt) + mCurrPosition) / mScreenBufferSize.x) + _
                Int(Len(mPrompt) / mScreenBufferSize.x)

setCursorPosition cursorPosn.x, cursorPosn.y
mCurrPosition = 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub moveCursorLeft()
Const ProcName As String = "moveCursorLeft"
On Error GoTo Err

If mCurrPosition = 1 Then Exit Sub

mCurrPosition = mCurrPosition - 1

Dim cursorPosn As COORD
cursorPosn = getCursorPosition
If cursorPosn.x = 0 Then
    cursorPosn.x = mScreenBufferSize.x - 1
    cursorPosn.y = cursorPosn.y - 1
Else
    cursorPosn.x = cursorPosn.x - 1
End If
setCursorPosition cursorPosn.x, cursorPosn.y

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub moveCursorRight()
Const ProcName As String = "moveCursorRight"
On Error GoTo Err

If mCurrPosition = mStrBld.Length + 1 Then Exit Sub

mCurrPosition = mCurrPosition + 1

Dim cursorPosn As COORD
cursorPosn = getCursorPosition
If cursorPosn.x = mScreenBufferSize.x - 1 Then
    cursorPosn.x = 0
    cursorPosn.y = cursorPosn.y + 1
Else
    cursorPosn.x = cursorPosn.x + 1
End If
setCursorPosition cursorPosn.x, cursorPosn.y

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

' moves to the specified position in the user input
Private Sub moveCursorToPosition( _
                ByVal position As Long)
Const ProcName As String = "moveCursorToPosition"
On Error GoTo Err

Dim cursorPosn As COORD
cursorPosn = getCursorPosition

cursorPosn.x = (Len(mPrompt) + position - 1) Mod mScreenBufferSize.x
cursorPosn.y = cursorPosn.y - _
                Int((Len(mPrompt) + mCurrPosition) / mScreenBufferSize.x) + _
                Int((Len(mPrompt) + position) / mScreenBufferSize.x)

setCursorPosition cursorPosn.x, cursorPosn.y
mCurrPosition = position

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function readInputFromConsole( _
                ByVal pEchoChar As String, _
                ByRef pValue As String) As Boolean
Const ProcName As String = "readInputFromConsole"
On Error GoTo Err

Dim failPoint As String
Static recordLength As Long
Dim recsRead As Long
Dim keyEventRec As KEY_EVENT_RECORD
Dim mouseEventRec As MOUSE_EVENT_RECORD
Dim windowBufferSizeRec As WINDOW_BUFFER_SIZE_RECORD
Dim menuEventRec As MENU_EVENT_RECORD
Dim focusEventRec As FOCUS_EVENT_RECORD
Dim Key As String
Dim errNum As Long
Dim i As Long

If recordLength = 0 Then
    recordLength = gMax(Len(focusEventRec), _
                        Len(keyEventRec), _
                        Len(menuEventRec), _
                        Len(mouseEventRec), _
                        Len(windowBufferSizeRec))
    ReDim mRecbuf(Int(((recordLength + 4) + 1) / 2 - 1)) As Integer ' extra 4 for record type
                                                                    ' and alignment padding
End If


If Not waitForConsoleInput Then
    pValue = ""
    readInputFromConsole = False
    Exit Function
End If

failPoint = 100

ReadConsoleInput mConInHandle, VarPtr(mRecbuf(0)), 1, recsRead
errNum = Err.LastDllError
If errNum <> 0 Then Err.Raise ErrorCodes.ErrRuntimeException, , "Error " & errNum & " when reading record from STDIN"

failPoint = 150

Select Case mRecbuf(0)
Case KEY_EVENT

    CopyMemory VarPtr(keyEventRec), VarPtr(mRecbuf(2)), Len(keyEventRec)
    
    failPoint = 200

    Key = ChrW(keyEventRec.uChar)
    
    If gIsPrintableKey(keyEventRec.wVirtualKeyCode) Then
    
        failPoint = 300
        
        If keyEventRec.bKeyDown Then
            pValue = Key
            For i = 2 To gMin(keyEventRec.wRepeatCount, 10)
                pValue = pValue & Key
            Next
            echoToConsole pValue, pEchoChar, mOverwriteMode
        End If
    Else
        failPoint = 400
    
        Select Case keyEventRec.wVirtualKeyCode
        Case vbKeyReturn, vbKeySeparator
            If keyEventRec.bKeyDown Then
                pValue = vbCrLf
                ' we'll ignore the repeat count because it's difficult to handle
                echoToConsole pValue, ""
            End If
        Case vbKeyBack
            If keyEventRec.bKeyDown Then
                If mCurrPosition <> 1 Then
                    For i = 1 To keyEventRec.wRepeatCount
                        moveCursorLeft
                        deleteCharAtCursor
                    Next
                End If
            End If
        Case vbKeyDelete
            If keyEventRec.bKeyDown Then
                For i = 1 To keyEventRec.wRepeatCount
                    deleteCharAtCursor
                Next
            End If
        Case vbKeyDown
            If keyEventRec.bKeyDown Then
                If mCurrHistoryIndex < mInputHistoryIndex Then
                    mCurrHistoryIndex = mCurrHistoryIndex + 1
                    clearPromptAndUserInput
                    setUserInputFromHistory
                    writePromptAndUserInput
                 End If
            End If
        Case vbKeyEnd
            If keyEventRec.bKeyDown Then
                moveCursorEnd
            End If
        Case vbKeyEscape
            If keyEventRec.bKeyDown Then
                clearPromptAndUserInput
                clearUserInput
                writePromptAndUserInput
            End If
        Case vbKeyHome
            If keyEventRec.bKeyDown Then
                moveCursorHome
            End If
        Case vbKeyInsert
            If keyEventRec.bKeyDown Then
                If mOverwriteMode Then
                    mCurrConsoleCursorInfo.dwSize = 50
                    mOverwriteMode = False
                Else
                    mCurrConsoleCursorInfo.dwSize = 50
                    mOverwriteMode = True
                    SetConsoleCursorInfo mConOutHandle, mCurrConsoleCursorInfo
                End If
            End If
        Case vbKeyLeft
            If keyEventRec.bKeyDown Then
                If keyEventRec.dwControlKeyState And LEFT_CTRL_PRESSED Then
                    moveCursorToPosition mStrBld.WordPositionLeft(mCurrPosition)
                Else
                    For i = 1 To keyEventRec.wRepeatCount
                        moveCursorLeft
                    Next
                End If
            End If
        Case vbKeyRight
            If keyEventRec.bKeyDown Then
                If keyEventRec.dwControlKeyState And LEFT_CTRL_PRESSED Then
                    moveCursorToPosition mStrBld.WordPositionRight(mCurrPosition)
                Else
                    For i = 1 To keyEventRec.wRepeatCount
                        moveCursorRight
                    Next
                End If
            End If
        Case vbKeyUp
            If keyEventRec.bKeyDown Then
                If mCurrHistoryIndex > 0 Then
                    mCurrHistoryIndex = mCurrHistoryIndex - 1
                    clearPromptAndUserInput
                    setUserInputFromHistory
                    writePromptAndUserInput
                 End If
            End If
        End Select
    End If
Case MOUSE_EVENT
    
    failPoint = 500
    
    CopyMemory VarPtr(mouseEventRec), VarPtr(mRecbuf(2)), Len(mouseEventRec)
Case WINDOW_BUFFER_SIZE_EVENT

    failPoint = 600

    CopyMemory VarPtr(windowBufferSizeRec), VarPtr(mRecbuf(2)), Len(windowBufferSizeRec)
Case MENU_EVENT
    
    failPoint = 700
    
    CopyMemory VarPtr(menuEventRec), VarPtr(mRecbuf(2)), Len(menuEventRec)
Case FOCUS_EVENT
    
    failPoint = 800
    
    CopyMemory VarPtr(focusEventRec), VarPtr(mRecbuf(2)), Len(focusEventRec)
Case Else
    WriteErrorLine "Unknown record type " & mRecbuf(0)
    pValue = "****"
End Select

readInputFromConsole = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName, failPoint
End Function

Private Function readLineFromFile() As String
Const ProcName As String = "readLineFromFile"
On Error GoTo Err

If mAtEof Then
    If mEoFString <> "" Then
        readLineFromFile = mEoFString
        Exit Function
    End If
    
    Err.Raise ErrorCodes.ErrEndOfFile, , "End of file reached"
End If

Dim i As Long
i = -1
Do
    i = i + 1
    
    Dim bytesRead As Long
    ReadFile mStdInHandle, VarPtr(mBuf(i)), 1, bytesRead, 0&
    
    Dim errNum As Long
    errNum = Err.LastDllError
    If errNum <> 0 Then
        If errNum = ERROR_BROKEN_PIPE Then
            mAtEof = True
            
            ' need to adjust it so it looks like we've got the CR & LF
            If i = 0 Then
                ' there's nothing on this line so end of file
                If mEoFString <> "" Then
                    readLineFromFile = mEoFString
                    Exit Function
                End If
                
                Err.Raise ErrorCodes.ErrEndOfFile, , "End of file reached"
            
            ElseIf mBuf(i - 1) = AscB(vbCr) Then
                ' we've had the CR but not the LF, so this would have been
                ' the LF, so no need to adjust i
            Else
                ' we've had neither, so this would have been the CR, so
                ' need to increment i
                i = i + 1
            End If
            Exit Do
        End If
        
        Err.Raise ErrorCodes.ErrRuntimeException, , "Error " & errNum & " when reading from STDIN"
    End If
    
    If bytesRead = 0 Then
        mAtEof = True
        If mEoFString <> "" Then
            readLineFromFile = mEoFString
            Exit Function
        Else
            Err.Raise ErrorCodes.ErrEndOfFile, , "End of file reached"
        End If
    End If
    
    If i >= 1 Then
        If mBuf(i - 1) = AscB(vbCr) And mBuf(i) = AscB(vbLf) Then Exit Do
    End If
Loop

If i > 1 Then
    ReDim ar(i - 2) As Byte
    CopyMemory VarPtr(ar(0)), VarPtr(mBuf(0)), i - 1
    readLineFromFile = StrConv(ar, vbUnicode)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Private Sub setCursorPosition( _
                ByVal x As Integer, _
                ByVal y As Integer)
Const ProcName As String = "setCursorPosition"
On Error GoTo Err

Dim csbi As CONSOLE_SCREEN_BUFFER_INFO
GetConsoleScreenBufferInfo mConOutHandle, csbi

If x < 0 Or y < 0 Or x > (csbi.dwSize.x - 1) Or y > (csbi.dwSize.y - 1) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid cursor position: " & x & "," & y

SetConsoleCursorPosition mConOutHandle, y * &H10000 + x

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub setUserInputFromHistory()
Const ProcName As String = "setUserInputFromHistory"
On Error GoTo Err

Set mStrBld = New StringBuilder
mStrBld.Initialise 80
mStrBld.Append mInputHistory(mCurrHistoryIndex)
mCurrPosition = mStrBld.Length + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function translateKeyEventRecord( _
                ByRef keyEventRec As KEY_EVENT_RECORD) As String
Const ProcName As String = "translateKeyEventRecord"
On Error GoTo Err

ReDim buf(15) As Byte
Dim l As Long

Dim keyboardState(255) As Byte
GetKeyboardState VarPtr(keyboardState(0))
If Err.LastDllError <> 0 Then Err.Raise ErrorCodes.ErrRuntimeException, , "Error " & Err.LastDllError & " reading keyboard State"

' set the State for the control Keys - see GetKeyboardState API documentation
' for details
If keyEventRec.dwControlKeyState And RIGHT_ALT_PRESSED Then
    keyboardState(VK_RMENU) = keyboardState(VK_RMENU) Or &H80
Else
    keyboardState(VK_RMENU) = keyboardState(VK_RMENU) And &H7F
End If

If keyEventRec.dwControlKeyState And LEFT_ALT_PRESSED Then
    keyboardState(VK_LMENU) = keyboardState(VK_LMENU) Or &H80
Else
    keyboardState(VK_LMENU) = keyboardState(VK_LMENU) And &H7F
End If

If keyEventRec.dwControlKeyState And RIGHT_CTRL_PRESSED Then
    keyboardState(VK_RCONTROL) = keyboardState(VK_RCONTROL) Or &H80
Else
    keyboardState(VK_RCONTROL) = keyboardState(VK_RCONTROL) And &H7F
End If

If keyEventRec.dwControlKeyState And LEFT_CTRL_PRESSED Then
    keyboardState(VK_LCONTROL) = keyboardState(VK_LCONTROL) Or &H80
Else
    keyboardState(VK_LCONTROL) = keyboardState(VK_LCONTROL) And &H7F
End If

If keyEventRec.dwControlKeyState And SHIFT_PRESSED Then
    keyboardState(vbKeyShift) = keyboardState(vbKeyShift) Or &H80
Else
    keyboardState(vbKeyShift) = keyboardState(vbKeyShift) And &H7F
End If

If keyEventRec.dwControlKeyState And NUMLOCK_ON Then
    keyboardState(vbKeyNumlock) = keyboardState(vbKeyNumlock) Or &H1
Else
    keyboardState(vbKeyNumlock) = keyboardState(vbKeyNumlock) And &HF7
End If

If keyEventRec.dwControlKeyState And SCROLLLOCK_ON Then
    keyboardState(vbKeyScrollLock) = keyboardState(vbKeyScrollLock) Or &H1
Else
    keyboardState(vbKeyScrollLock) = keyboardState(vbKeyScrollLock) And &HF7
End If

If keyEventRec.dwControlKeyState And CAPSLOCK_ON Then
    keyboardState(vbKeyCapital) = keyboardState(vbKeyCapital) Or &H1
Else
    keyboardState(vbKeyCapital) = keyboardState(vbKeyCapital) And &HF7
End If

l = toUnicode(keyEventRec.wVirtualKeyCode, _
            keyEventRec.wVirtualScanCode, _
            VarPtr(keyboardState(0)), _
            VarPtr(buf(0)), _
            (UBound(buf) + 1) / 2, _
            0&)
If l > 0 Then
    ReDim Preserve buf(2 * l - 1) As Byte
    translateKeyEventRecord = buf
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub updateInputHistory( _
                ByVal Value As String)
Const ProcName As String = "updateInputHistory"
On Error GoTo Err

Dim newInput As Boolean
newInput = True

Dim i As Long
For i = 0 To mInputHistoryIndex
    If mInputHistory(i) = Value Then
        newInput = False
        Exit For
    End If
Next

If Value <> mInputHistory(mInputHistoryIndex) Then
    mInputHistoryIndex = mInputHistoryIndex + 1
    If mInputHistoryIndex > UBound(mInputHistory) Then
        ReDim Preserve mInputHistory((UBound(mInputHistory) + 1) * 2 - 1) As String
    End If
    mInputHistory(mInputHistoryIndex) = Value
End If

If newInput Then mCurrHistoryIndex = mInputHistoryIndex + 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function waitForConsoleInput() As Boolean
Const ProcName As String = "waitForConsoleInput"
On Error GoTo Err

Dim ret As Long
ret = MsgWaitForMultipleObjects(1, _
                                mConInHandle, _
                                False, _
                                INFINITE, _
                                QS_ALLEVENTS)
If ret = WAIT_FAILED Then gHandleWin32Error

waitForConsoleInput = (ret = WAIT_OBJECT_0)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub writeErrorStringToConsole( _
                msg As String)
Const ProcName As String = "writeErrorStringToConsole"
On Error GoTo Err

writeToConsole msg

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub writePromptAndUserInput()
Const ProcName As String = "writePromptAndUserInput"
On Error GoTo Err

mPrevCursorPosition = getCursorPosition

If mPrevCursorPosition.x <> 0 Then
    writeToConsole vbCrLf    ' move to start of next line
End If

Dim cursorPosn As COORD
cursorPosn = getCursorPosition
writeToConsole mPrompt
writeToConsole mStrBld.ToString

' now reposition the cursor
cursorPosn.x = (mCurrPosition + Len(mPrompt) - 1) Mod mScreenBufferSize.x
cursorPosn.y = cursorPosn.y + Int((mCurrPosition + Len(mPrompt) - 1) / mScreenBufferSize.x)
setCursorPosition cursorPosn.x, cursorPosn.y

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub writeToConsole( _
                msg As String)
Const ProcName As String = "writeToConsole"
On Error GoTo Err

Dim bytesWritten As Long
If WriteConsole(mConOutHandle, _
            StrPtr(msg), _
            Len(msg), _
            bytesWritten, _
            0&) Then
Else
    WriteErrorLine "Error " & Err.LastDllError & " writing to console"
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub





