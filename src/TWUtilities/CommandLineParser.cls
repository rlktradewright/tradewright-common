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
' separator Character is specified in the call to <code>CreateCommandLineParser</code>.
'
' Arguments that contain the separator Character must be enclosed in double quotes. Double quotes
' appearing within an argument must be repeated.
'
' Switches have the following format:
'
' <pre>
'   ( "/" | "-")&lt;identifier&gt; [":"&lt;switchValue&gt;]
' </pre>
'
' ie the switch starts with a forward slash or a hyphen followed by an identifier, and
' optionally followed by a colon and the switch Value. Switch identifiers are not
' case-sensitive. Switch values that contain the separator Character must be enclosed in
' double quotes. Double quotes appearing within a switch Value must be repeated.
'
' Examples (these examples use a space as the separator Character):
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

''
' Returns the nth argument, where n is the Value supplied to this method.
'
'@remarks
'   If the requested argument has not been supplied, an empty string is returned.
' @return
'   A String Value containing the nth argument, where n is the Value of the i parameter.
' @param i
'   The number of the argument to be returned. The first argument is number 0.
'@/
Public Property Get Arg( _
                ByVal i As Integer) As String
On Error Resume Next
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
On Error GoTo Err
NumberOfArgs = UBound(mArgs) + 1
Exit Property

Err:
NumberOfArgs = 0
End Property

''
' Returns the number of switches.
'
' @return
'   The number of switches.
'
'@/
Public Property Get NumberOfSwitches() As Integer
On Error GoTo Err
NumberOfSwitches = UBound(mSwitches) + 1
Exit Property

Err:
NumberOfSwitches = 0
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
Dim i As Long

On Error GoTo Err

For i = 0 To UBound(mSwitches)
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
' Returns the Value of the specified switch.
'
'@remarks
'   If the requested switch has not been supplied, or no Value
'   was supplied for the switch, an empty string is returned.
' @return
'   A String containing the Value for the specified switch.
' @param s
'   The identifier of the switch whose Value is to be returned.
'
'@/
Public Property Get SwitchValue( _
                ByVal s As String) As String
Dim i As Long

On Error GoTo Err

For i = 0 To UBound(mSwitches)
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
Dim pos As Long
Const ProcName As String = "ContainsUnbalancedQuotes"

On Error GoTo Err

pos = Len(inString) - 1 ' exclude the last char
If pos = 0 Then Exit Function
pos = InStrRev(inString, """", pos)
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
Dim rawArgs() As String
Dim Arg As String
Dim v As Variant
Dim quotedArg As String
Dim inQuotedArg As Boolean
Dim iA As Integer
Dim iSw As Integer
Dim unbalancedQuotes As Boolean

Const ProcName As String = "getArgs"

On Error GoTo Err

If mCommandLine = "" Then Exit Sub

rawArgs = Split(mCommandLine, mSep)

For Each v In rawArgs
    Arg = CStr(v)
    Arg = Replace(Arg, """""", """")
    If Not inQuotedArg Then
        If Arg = "" And mSep = "" Then
        Else
            If Left$(Arg, 1) = """" Or _
                ((Left$(Arg, 1) = "/" Or Left$(Arg, 1) = "-") And InStr(1, Arg, ":""")) <> 0 _
            Then
                inQuotedArg = True
                quotedArg = Replace(Arg, """", "", , 1) ' Right$(arg, Len(arg) - 1)
                If ContainsUnbalancedQuotes(quotedArg) Then
                    unbalancedQuotes = Not unbalancedQuotes
                End If
                If Right$(quotedArg, 1) = """" And Not unbalancedQuotes Then
                    ' the ending quote is also in this arg
                    inQuotedArg = False
                    quotedArg = Left$(quotedArg, Len(quotedArg) - 1)
                    If Left$(quotedArg, 1) = "/" Or _
                        Left$(quotedArg, 1) = "-" _
                    Then
                        setSwitch Right$(quotedArg, Len(quotedArg) - 1), iSw
                        iSw = iSw + 1
                    Else
                        setArg quotedArg, iA
                        iA = iA + 1
                    End If
                Else
                    If Right$(Arg, 1) = """" Then
                        unbalancedQuotes = Not unbalancedQuotes
                    End If
                End If
            ElseIf (Left$(Arg, 1) = "/" Or _
                Left$(Arg, 1) = "-") And _
                Len(Arg) >= 2 _
            Then
                setSwitch Right$(Arg, Len(Arg) - 1), iSw
                iSw = iSw + 1
            Else
                setArg Arg, iA
                iA = iA + 1
            End If
        End If
    Else
        If ContainsUnbalancedQuotes(Arg) Then
            unbalancedQuotes = Not unbalancedQuotes
        End If
        If Right$(Arg, 1) = """" And Not unbalancedQuotes Then
            inQuotedArg = False
            quotedArg = quotedArg & mSep & Left$(Arg, Len(Arg) - 1)
            If Left$(quotedArg, 1) = "/" Or _
                Left$(quotedArg, 1) = "-" _
            Then
                setSwitch Right$(quotedArg, Len(quotedArg) - 1), iSw
                iSw = iSw + 1
            Else
                setArg quotedArg, iA
                iA = iA + 1
            End If
        Else
            If Right$(Arg, 1) = """" Then
                unbalancedQuotes = Not unbalancedQuotes
            End If
            quotedArg = quotedArg & mSep & Arg
        End If
    End If
Next

If iA <> 0 Then ReDim Preserve mArgs(iA - 1) As String
If iSw <> 0 Then ReDim Preserve mSwitches(iSw - 1) As SwitchEntry

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub


Private Sub setArg( _
                ByVal val As String, _
                ByVal pIndex As Long)
On Error Resume Next
If pIndex > UBound(mArgs) Then
    ReDim Preserve mArgs(UBound(mArgs) + 10) As String
    If Err.number = 9 Then
         ReDim mArgs(10) As String
    End If
End If
mArgs(pIndex) = val
End Sub

Private Sub setSwitch( _
                ByVal val As String, _
                ByVal pIndex As Long)
Dim i As Long
Const ProcName As String = "setSwitch"

On Error GoTo Err

On Error Resume Next
If pIndex > UBound(mSwitches) Then
    ReDim Preserve mSwitches(UBound(mSwitches) + 10) As SwitchEntry
    If Err.number = 9 Then
         ReDim mSwitches(10) As SwitchEntry
    End If
End If

i = InStr(val, ":")

If i > 0 Then
    mSwitches(pIndex).Switch = UCase$(Left$(val, i - 1))
    mSwitches(pIndex).Value = Right$(val, Len(val) - i)
Else
    mSwitches(pIndex).Switch = UCase$(val)
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub



