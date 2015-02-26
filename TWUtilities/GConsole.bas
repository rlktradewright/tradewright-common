Attribute VB_Name = "GConsole"

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

Public Type CONSOLE_CURSOR_INFO
    dwSize As Long
    bVisible As Long
End Type

Public Type COORD
    x                   As Integer
    y                   As Integer
End Type

Public Type SMALL_RECT
    Left                As Integer
    Top                 As Integer
    Right               As Integer
    Bottom              As Integer
End Type

Public Type CONSOLE_SCREEN_BUFFER_INFO
    dwSize              As COORD
    dwCursorPosition    As COORD
    wAttributes         As Integer
    srWindow            As SMALL_RECT
    dwMaximumWindowSize As COORD
End Type

Public Type FOCUS_EVENT_RECORD
    bSetFocus           As Long
End Type

Public Type KEY_EVENT_RECORD
    bKeyDown            As Long
    wRepeatCount        As Integer
    wVirtualKeyCode     As Integer
    wVirtualScanCode    As Integer
    uChar               As Integer  ' unicode
    dwControlKeyState   As Long
End Type

Public Type MENU_EVENT_RECORD
    dwCommandId         As Long
End Type

Public Type MOUSE_EVENT_RECORD
    dwMousePosition     As COORD
    dwButtonState       As Long
    dwControlKeyState   As Long
    dwEventFlags        As Long
End Type

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Type WINDOW_BUFFER_SIZE_RECORD
    dwSize              As COORD
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                     As String = "GConsole"

Public Const BACKGROUND_BLUE                As Long = &H10   ' background Color contains blue.
Public Const BACKGROUND_GREEN               As Long = &H20   ' background Color contains green.
Public Const BACKGROUND_RED                 As Long = &H40   ' background Color contains red.
Public Const BACKGROUND_INTENSITY           As Long = &H80   ' background Color is intensified.

Public Const ENABLE_ECHO_INPUT              As Long = &H4
Public Const ENABLE_LINE_INPUT              As Long = &H2
Public Const ENABLE_MOUSE_INPUT             As Long = &H10
Public Const ENABLE_PROCESSED_INPUT         As Long = &H1
Public Const ENABLE_PROCESSED_OUTPUT        As Long = &H1
Public Const ENABLE_WINDOW_INPUT            As Long = &H8
Public Const ENABLE_WRAP_AT_EOL_OUTPUT      As Long = &H2

Public Const ERROR_BROKEN_PIPE              As Long = &H6D

Public Const FILE_SHARE_READ                As Long = &H1
Public Const FILE_SHARE_WRITE               As Long = &H2

Public Const FILE_TYPE_CHAR                 As Long = &H2
Public Const FILE_TYPE_DISK                 As Long = &H1
Public Const FILE_TYPE_PIPE                 As Long = &H3
Public Const FILE_TYPE_UNKNOWN              As Long = &H0

Public Const FOREGROUND_BLUE                As Long = &H1    ' text Color contains blue.
Public Const FOREGROUND_GREEN               As Long = &H2    ' text Color contains green.
Public Const FOREGROUND_RED                 As Long = &H4    ' text Color contains red.
Public Const FOREGROUND_INTENSITY           As Long = &H8    ' text Color is intensified.

Public Const GENERIC_READ                   As Long = &H80000000
Public Const GENERIC_WRITE                  As Long = &H40000000

Public Const INVALID_HANDLE_VALUE           As Long = -1

Public Const OPEN_EXISTING                  As Long = 3

Public Const STD_INPUT_HANDLE               As Long = -10&
Public Const STD_OUTPUT_HANDLE              As Long = -11&
Public Const STD_ERROR_HANDLE               As Long = -12&

Public Const KEY_EVENT                      As Integer = &H1
Public Const MOUSE_EVENT                    As Integer = &H2
Public Const WINDOW_BUFFER_SIZE_EVENT       As Integer = &H4
Public Const MENU_EVENT                     As Integer = &H8
Public Const FOCUS_EVENT                    As Integer = &H10

Public Const RIGHT_ALT_PRESSED              As Long = &H1   ' the right alt key is pressed.
Public Const LEFT_ALT_PRESSED               As Long = &H2   ' the left alt key is pressed.
Public Const RIGHT_CTRL_PRESSED             As Long = &H4   ' the right ctrl key is pressed.
Public Const LEFT_CTRL_PRESSED              As Long = &H8   ' the left ctrl key is pressed.
Public Const SHIFT_PRESSED                  As Long = &H10  ' the shift key is pressed.
Public Const NUMLOCK_ON                     As Long = &H20  ' the numlock light is on.
Public Const SCROLLLOCK_ON                  As Long = &H40  ' the scrolllock light is on.
Public Const CAPSLOCK_ON                    As Long = &H80  ' the capslock light is on.
Public Const ENHANCED_KEY                   As Long = &H100 ' the key is enhanced.

Public Const VK_LSHIFT                      As Long = &HA0
Public Const VK_RSHIFT                      As Long = &HA1
Public Const VK_LCONTROL                    As Long = &HA2
Public Const VK_RCONTROL                    As Long = &HA3
Public Const VK_LMENU                       As Long = &HA4
Public Const VK_RMENU                       As Long = &HA5

'@================================================================================
' External function declarations
'@================================================================================

Public Declare Function AllocConsole Lib "Kernel32" () As Long

Public Declare Function CreateFile Lib "Kernel32" Alias "CreateFileA" ( _
                ByVal lpFileName As String, _
                ByVal dwDesiredAccess As Long, _
                ByVal dwShareMode As Long, _
                ByVal lpSecurityAttributes As Long, _
                ByVal dwCreationDisposition As Long, _
                ByVal dwFlagsAndAttributes As Long, _
                ByVal hTemplateFile As Long) As Long
                
Public Declare Function FillConsoleOutputAttribute Lib "Kernel32" ( _
                ByVal hConsoleOutput As Long, _
                ByVal wAttribute As Long, _
                ByVal nLength As Long, _
                ByVal dwWriteCoord As Long, _
                ByVal lpNumberOfAttrsWritten As Long) As Long

Public Declare Function FillConsoleOutputCharacter Lib "Kernel32" Alias "FillConsoleOutputCharacterW" ( _
                ByVal hConsoleOutput As Long, _
                ByVal cCharacter As Integer, _
                ByVal nLength As Long, _
                ByVal dwWriteCoord As Long, _
                lpNumberOfCharsWritten As Long) As Long

Public Declare Function GetConsoleCursorInfo Lib "Kernel32" ( _
                ByVal hConsoleOutput As Long, _
                lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long
                
Public Declare Function GetConsoleMode Lib "Kernel32" ( _
                ByVal hConsoleHandle As Long, _
                lpMode As Long) As Long
                
Public Declare Function GetConsoleScreenBufferInfo Lib "Kernel32" ( _
                ByVal hConsoleOutput As Long, _
                lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long

Public Declare Function GetConsoleTitle Lib "Kernel32" Alias "GetConsoleTitleW" ( _
                ByVal lpConsoleTitle As Long, _
                ByVal nSize As Long) As Long

Public Declare Function GetFileType Lib "Kernel32" ( _
                ByVal hFile As Long) As Long

Public Declare Function GetConsoleWindow Lib "Kernel32" () As Long

Public Declare Function GetKeyboardState Lib "user32" ( _
                ByVal pbKeyState As Long) As Long

Public Declare Function GetStdHandle Lib "Kernel32" ( _
                ByVal nStdHandle As Long) As Long

Public Declare Function GetQueueStatus Lib "user32" ( _
                ByVal fuFlags As Long) As Long

Public Declare Function ReadConsole Lib "Kernel32" Alias "ReadConsoleW" ( _
                ByVal hFile As Long, _
                ByVal IpBuffer As Long, _
                ByVal nNumberOfBytesToRead As Long, _
                IpNumberOfBytesRead As Long, _
                ByVal lpReserved As Any) As Long

Public Declare Function ReadConsoleInput Lib "Kernel32" Alias "ReadConsoleInputW" ( _
                ByVal hFile As Long, _
                ByVal IpBuffer As Long, _
                ByVal nNumberOfRecordsToRead As Long, _
                IpNumberOfRecordsRead As Long) As Long

Public Declare Function ReadFile Lib "Kernel32" ( _
                ByVal hFile As Long, _
                ByVal IpBuffer As Long, _
                ByVal nNumberOfBytesToRead As Long, _
                IpNumberOfBytesRead As Long, _
                ByVal IpOverlapped As Any) As Long

Public Declare Function SetConsoleCursorInfo Lib "Kernel32" ( _
                ByVal hConsoleOutput As Long, _
                lpConsoleCursorInfo As CONSOLE_CURSOR_INFO) As Long

Public Declare Function SetConsoleCursorPosition Lib "Kernel32" ( _
                ByVal hConsoleOutput As Long, _
                ByVal dwCursorPosition As Long) As Long
                
Public Declare Function SetConsoleCtrlHandler Lib "Kernel32" ( _
                ByVal HandlerRoutine As Long, _
                ByVal Add As Long) As Long
                
Public Declare Function SetConsoleMode Lib "Kernel32" ( _
                ByVal hConsoleHandle As Long, _
                ByVal dwMode As Long) As Long
                
Public Declare Function SetConsoleScreenBufferSize Lib "Kernel32" ( _
                ByVal hConsoleOutput As Long, _
                dwSize As COORD) As Long

Public Declare Function SetConsoleTextAttribute Lib "Kernel32" ( _
                ByVal hConsoleOutput As Long, _
                ByVal wAttributes As Long) As Long

Public Declare Function SetConsoleTitle Lib "Kernel32" Alias "SetConsoleTitleW" ( _
                ByVal lpConsoleTitle As Long) As Long

Public Declare Function toUnicode Lib "user32" Alias _
                "ToUnicode" (ByVal wVirtKey As _
                Long, ByVal wScanCode As _
                Long, ByVal lpKeyState As _
                Long, ByVal pwszBuff As _
                Long, ByVal cchBuff As _
                Long, ByVal wFlags As Long) As Long

Public Declare Function WriteConsole Lib "Kernel32" Alias "WriteConsoleW" ( _
                ByVal hConsoleOutput As Long, _
                ByVal lpBuffer As Long, _
                ByVal nNumberOfCharsToWrite As Long, _
                lpNumberOfCharsWritten As Long, _
                lpReserved As Any) As Long
                
Public Declare Function WriteFile Lib "Kernel32" ( _
                ByVal hFile As Long, _
                ByVal IpBuffer As Long, _
                ByVal nNumberOfBytesToWrite As Long, _
                IpNumberOfBytesWritten As Long, _
                ByVal IpOverlapped As Any) As Long

'@================================================================================
' Member variables
'@================================================================================

Private mConsole                            As Console

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

Public Function gGetConsole() As Console
If mConsole Is Nothing Then
    Set mConsole = New Console
End If

Set gGetConsole = mConsole
End Function

'@================================================================================
' Helper Functions
'@================================================================================


