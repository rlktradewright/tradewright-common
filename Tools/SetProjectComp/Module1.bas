Attribute VB_Name = "Module1"
yOption Explicit

Public Sub Main()

Dim filename As String
Dim clp As CommandLineParser
Dim mode As String
Dim con As Console
Dim ts As TextStream
Dim fs As FileSystemObject
Dim lines() As String
Dim linesIndex As Long
Dim i As Long

On Error GoTo Err

InitialiseTWUtilities

Set con = GetConsole

Set clp = CreateCommandLineParser(Command)

filename = clp.Arg(0)

If filename = "" Then
    con.WriteErrorLine "Project filename must be supplied as first argument"
    Exit Sub
End If

If Not clp.Switch("mode") Then
    con.WriteErrorLine "Must supply /mode switch"
    Exit Sub
End If

mode = UCase$(clp.SwitchValue("mode"))
If mode = "P" Or mode = "B" Then
Else
    con.WriteErrorLine "Mode must be 'P' or 'B'"
    Exit Sub
End If

Set fs = New FileSystemObject
Set ts = fs.OpenTextFile(filename, ForReading)

ReDim lines(1) As String
linesIndex = -1

Do While Not ts.AtEndOfStream
    linesIndex = linesIndex + 1
    If linesIndex > UBound(lines) Then ReDim Preserve lines(2 * (UBound(lines) + 1) - 1) As String
    lines(linesIndex) = ts.ReadLine
Loop

ts.Close

Dim alreadyInRequiredMode As Boolean
For i = 0 To linesIndex
    If mode = "P" And lines(i) = "CompatibleMode=""1""" Then
        con.WriteLine "Already in Project Compatibility mode"
        alreadyInRequiredMode = True
        Exit For
    ElseIf mode = "B" And lines(i) = "CompatibleMode=""2""" Then
        con.WriteLine "Already in Binary Compatibility mode"
        alreadyInRequiredMode = True
        Exit For
    End If
Next

If Not alreadyInRequiredMode Then
    Set ts = fs.OpenTextFile(filename, ForWriting)
    
    For i = 0 To linesIndex
        If mode = "P" Then
            If Left(lines(i), Len("CompatibleMode=")) = "CompatibleMode=" Then
                ts.WriteLine "CompatibleMode=""1"""
            ElseIf lines(i) = "VersionCompatible32=""1""" Then
                ' this line is to be removed so don't write anything
            Else
                ts.WriteLine lines(i)
            End If
        Else
            If lines(i) = "CompatibleMode=""1""" Then
                ts.WriteLine "CompatibleMode=""2"""
                ts.WriteLine "VersionCompatible32=""1"""
            Else
                ts.WriteLine lines(i)
            End If
        End If
    Next
    
    ts.Close
    con.WriteLine "File adjusted to " & IIf(mode = "P", "Project", "Binary") & " Compatibility"
End If

TerminateTWUtilities
Exit Sub

Err:
con.WriteErrorLine "Error " & Err.Number & ": " & Err.Description
TerminateTWUtilities
EndProcess 2
End Sub
