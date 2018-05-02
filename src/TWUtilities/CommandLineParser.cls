VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CommandLineParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Provides facilities for an application to determine the number and values of
' arguments and switches in a string (normally the arguments part of the command
' used to start the application).
'
' To create a <code>CommandLineParser</code> object, call the global
' <code>CreateCommandLineParser</code> method.
'
' The format of the argument string passed to the <code>CreateCommandLineParser</code> method
' is as follows:
'
' <pre>
'   [&lt;argument&gt; | &lt;switch&gt;] [&lt;sep&gt; (&lt;argument&gt; | &lt;switch&gt;)]...
' </pre>
'
' ie, there is a sequence of arguments or switches, separated by separator characters. The
' separator character is specified in the call to <code>CreateCommandLineParser</code>.
'
' Arguments that contain the separator character must be enclosed in double quotes. Double quotes
' appearing within an argument must be repeated.
'
' Switches have the following format:
'
' <pre>
'   ( "/" | "-")&lt;identifier&gt; [":"&lt;switchValue&gt;]
' </pre>
'
' ie the switch starts with a forward slash or a hyphen followed by an identifier, and
' optionally followed by a colon and the switch value. Switch identifiers are not
' case-sensitive. Switch values that contain the separator character must be enclosed in
' double quotes. Double quotes appearing within a switch value must be repeated.
'
' If a switch consisting only of "//" or "--" is supplied, then everything else in the
' input is interpreted as arguments.
'
' Examples (these examples use a space as the separator character):
' <pre>
'   anArgument -sw1 anotherArg -sw2:42
' </pre>
' <pre>
'   "C:\Program Files\MyApp\myapp.ini" -out:C:\MyLogs\myapp.Log
' </pre>
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
' Constants
'@================================================================================


Private Const ModuleName                    As String = "CommandLineParser"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mArgs() As String
Private mSwitches() As SwitchEntry
Private mSep As String
Private mCommandLine As String

Private mNumberOfSwitches As Long
Private mNumberOfArgs As Long

Private mTreatRemainingArgsAsArgs As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
ReDim mArgs(15) As String
ReDim mSwitches(15) As SwitchEntry
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
' Returns the nth argument, where n is the value supplied to this method.
'
'@remarks
'   If the requested argument has not been supplied, an empty string is returned.
' @return
'   A String value containing the nth argument, where n is the value of the i parameter.
' @param i
'   The number of the argument to be returned. The first argument is number 0.
'@/
Public Property Get Arg( _
                ByVal i As Integer) As String
If mNumberOfArgs = 0 Then Exit Property

gAssertArgument gIsInteger(i, 0, mNumberOfArgs - 1), "Invalid argument index"
Arg = mArgs(i)
End Property

''
' Returns an array of strings containing the arguments.
'
' @return
'   A String array containing the arguments.
'
'@/
Public Property Get Args() As String()
Args = mArgs
End Property

''
' Returns the number of arguments.
'
' @return
'   The number of arguments.
'
'@/
Public Property Get NumberOfArgs() As Integer
NumberOfArgs = mNumberOfArgs
End Property

''
' Returns the number of switches.
'
' @return
'   The number of switches.
'
'@/
Public Property Get NumberOfSwitches() As Integer
NumberOfSwitches = mNumberOfSwitches
End Property

''
' Indicates whether the specified switch was included.
'
' @return
'   If the specified switch was included, <code>True</code> is
'   returned. Otherwise <code>False</code> is returned.
'
' @param s
'   The identifier of the switch whose inclusion is to be indicated.
'
'@/
Public Property Get Switch(ByVal s As String) As Boolean
If mNumberOfSwitches = 0 Then Exit Property

Dim i As Long
For i = 0 To mNumberOfSwitches - 1
    If CStr(mSwitches(i).Switch) = UCase$(s) Then
        Switch = True
        Exit For
    End If
Next
Exit Property

Err:
End Property

''
' Returns an array of <code>SwitchEntry</code>s containing the
' switch identifiers and values.
'
' @return
'   An array of <code>SwitchEntry</code>s containing the
'   switch identifiers and values.
'
'@/
Public Property Get Switches() As SwitchEntry()
Switches = mSwitches
End Property

''
' Returns the value of the specified switch.
'
'@remarks
'   If the requested switch has not been supplied, or no value
'   was supplied for the switch, an empty string is returned.
' @return
'   A String containing the value for the specified switch.
' @param s
'   The identifier of the switch whose value is to be returned.
'
'@/
Public Property Get SwitchValue( _
                ByVal s As String) As String
If mNumberOfSwitches = 0 Then Exit Property

Dim i As Long
For i = 0 To mNumberOfSwitches - 1
    If CStr(mSwitches(i).Switch) = UCase$(s) Then
        SwitchValue = mSwitches(i).Value
        Exit For
    End If
Next
Exit Property

Err:
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise( _
                ByVal commandLine As String, _
                ByVal separator As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

mCommandLine = Trim$(commandLine)
mSep = separator
getArgs

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


Private Function ContainsUnbalancedQuotes( _
                inString As String) As Boolean
Const ProcName As String = "ContainsUnbalancedQuotes"
On Error GoTo Err

Dim pos As Long
pos = InStrRev(inString, """")
If pos = 0 Then Exit Function
Do While pos <> 0
    ContainsUnbalancedQuotes = Not ContainsUnbalancedQuotes
    If pos = 1 Then Exit Do
    pos = InStrRev(inString, """", pos - 1)
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub getArgs()
Const ProcName As String = "getArgs"
On Error GoTo Err

If mCommandLine = "" Then Exit Sub

Dim partialArg As String
Dim Var As Variant
For Each Var In Split(mCommandLine, mSep)
    Dim argument As String: argument = Var
    If partialArg = "" And argument = "" And mSep = " " Then
        ' discard spaces when the separator is a space and we don't have unbalanced quotes
    Else
        If partialArg <> "" Then partialArg = partialArg & mSep
        partialArg = partialArg & argument
        If Not ContainsUnbalancedQuotes(partialArg) Then
            setSwitchOrArg Trim$(partialArg)
            partialArg = ""
        End If
    End If
Next

If partialArg <> "" Then
    setSwitchOrArg Trim$(partialArg)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function isAlphaChar(val As String) As Boolean
isAlphaChar = True
If val >= "A" And val <= "Z" Then Exit Function
If val >= "a" And val <= "z" Then Exit Function
isAlphaChar = False
End Function

Private Sub setArg(ByVal val As String)
If mNumberOfArgs > UBound(mArgs) Then
    ReDim Preserve mArgs(2 * (UBound(mArgs) + 1) - 1) As String
End If
mArgs(mNumberOfArgs) = val
mNumberOfArgs = mNumberOfArgs + 1
End Sub

Private Sub setSwitch(ByVal val As String)
If mNumberOfSwitches > UBound(mSwitches) Then
    ReDim Preserve mSwitches(2 * (UBound(mSwitches) + 1) - 1) As SwitchEntry
End If

Dim i As Long
i = InStr(val, ":")

If i > 0 Then
    mSwitches(mNumberOfSwitches).Switch = UCase$(Left$(val, i - 1))
    mSwitches(mNumberOfSwitches).Value = Right$(val, Len(val) - i)
Else
    mSwitches(mNumberOfSwitches).Switch = UCase$(val)
End If
mNumberOfSwitches = mNumberOfSwitches + 1
End Sub

Private Sub setSwitchOrArg(val As String)
If val = "--" Or val = "//" Then
    mTreatRemainingArgsAsArgs = True
    Exit Sub
End If

If mTreatRemainingArgsAsArgs Then
    setArg val
ElseIf (Left$(val, 1) = "/" Or Left$(val, 1) = "-") And _
        Len(val) > 1 And _
        isAlphaChar(Mid$(val, 2, 1)) _
Then
    setSwitch Mid$(val, 2)
Else
    setArg val
End If
End Sub


