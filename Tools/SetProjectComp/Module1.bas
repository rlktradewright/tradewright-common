Attribute VB_Name = "Module1"
Option Explicit

Public Sub Main()
On Error GoTo Err

Dim lCon As Console: Set lCon = GetConsole

Dim lClp As CommandLineParser: Set lClp = CreateCommandLineParser(Command)

Dim lFilename As String: lFilename = getFileName(lClp)

Dim lMode As String: lMode = getMode(lClp)

Dim lRevisionNumber As Long: lRevisionNumber = getRevision(lClp)

Dim lFs As New FileSystemObject

Dim lLines As Collection
Set lLines = getLines(lFilename, lFs)

Dim lModeUpdateNeeded As Boolean
Dim lVersionUpdateNeeded As Boolean
checkIfChangesNeeded lLines, lMode, lRevisionNumber, lCon, lModeUpdateNeeded, lVersionUpdateNeeded

If Not lModeUpdateNeeded And Not lVersionUpdateNeeded Then Exit Sub

If lModeUpdateNeeded Then adjustMode lLines, lMode, lCon
If lVersionUpdateNeeded Then adjustVersion lLines, lRevisionNumber, lCon

writeNewFile lFs, lFilename, lLines

Exit Sub

Err:
lCon.WriteErrorLine "Error " & Err.Number & ": " & Err.Description
EndProcess 2
End Sub

Private Sub adjustMode( _
                ByRef pLines As Collection, _
                ByVal pMode As String, _
                ByVal pCon As Console)
Dim i As Long
Dim v As Variant
For Each v In pLines
    Dim s As String: s = CStr(v)
    i = i + 1
    
    If s = "VersionCompatible32=""1""" Then
        ' this line will be rewritten only if needed, so delete it
        pLines.Remove i
        i = i - 1
    ElseIf startsWith(s, "CompatibleMode=") Then
        If pMode = "P" Then
            pLines.Remove i
            pLines.Add "CompatibleMode=""1""", , i
            pCon.WriteLine "File adjusted to Project Compatibility"
        ElseIf pMode = "B" Then
            pLines.Remove i
            pLines.Add "CompatibleMode=""2""", , i
            pLines.Add "VersionCompatible32=""1""", , i + 1
            i = i + 1
            pCon.WriteLine "File adjusted to Binary Compatibility"
        ElseIf pMode = "N" Then
            pLines.Remove i
            pLines.Add "CompatibleMode=""0""", , i
            pCon.WriteLine "File adjusted to No Compatibility"
        End If
    End If
Next
End Sub

Private Sub adjustVersion( _
                ByRef pLines As Collection, _
                ByVal pRevisionNumber As Long, _
                ByVal pCon As Console)
Dim i As Long
Dim v As Variant
For Each v In pLines
    Dim s As String: s = CStr(v)
    i = i + 1
    
    If startsWith(s, "RevisionVer=") Then
        pLines.Remove i
        pLines.Add "RevisionVer=" & CStr(pRevisionNumber), , i
        pCon.WriteLine "Revision version set to " & CStr(pRevisionNumber)
    End If
Next
End Sub

Private Sub checkIfChangesNeeded( _
                ByRef pLines As Collection, _
                ByVal pMode As String, ByVal pRevisionNumber As Long, _
                ByVal pCon As Console, _
                ByRef pModeUpdateNeeded As Boolean, _
                ByRef pVersionUpdateNeeded As Boolean)
Dim v As Variant
For Each v In pLines
    Dim s As String: s = CStr(v)
    
    If startsWith(s, "CompatibleMode=") Then
        If pMode = "P" And s = "CompatibleMode=""1""" Then
            pCon.WriteLine "Already in Project Compatibility mode"
        ElseIf pMode = "B" And s = "CompatibleMode=""2""" Then
            pCon.WriteLine "Already in Binary Compatibility mode"
        Else
            pModeUpdateNeeded = True
        End If
    ElseIf startsWith(s, "RevisionVer=") Then
        If s <> ("RevisionVer=" & CLng(pRevisionNumber)) Then pVersionUpdateNeeded = True
    End If
Next
End Sub

Private Function getFileName(ByVal pClp As CommandLineParser) As String
getFileName = pClp.Arg(0)
If getFileName = "" Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Project filename must be supplied as first argument"
End Function

Private Function getLines(ByVal pFilename As String, ByVal pFs As FileSystemObject) As Collection
Dim lTs As TextStream
Set lTs = pFs.OpenTextFile(pFilename, ForReading)

Dim lLines As New Collection

Do While Not lTs.AtEndOfStream
    lLines.Add lTs.ReadLine
Loop

lTs.Close

Set getLines = lLines
End Function

Private Function getMode(ByVal pClp As CommandLineParser) As String
If Not pClp.Switch("mode") Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Must supply /mode switch"

getMode = UCase$(pClp.SwitchValue("mode"))
If getMode <> "P" And getMode <> "B" And getMode <> "N" Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Mode must be 'P' or 'B' or 'N'"
End Function

Private Function getRevision(pClp) As Long
Dim lRevString As String
lRevString = pClp.Arg(1)
If lRevString = "" Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Product revision number must be supplied as second argument"
If Not IsInteger(lRevString, 0, 9999) Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Product revision number must be an integer 0-9999"
getRevision = CLng(lRevString)
End Function

Private Function startsWith(ByVal s As String, ByVal pSubStr As String) As Boolean
startsWith = (Left$(UCase$(s), Len(pSubStr)) = UCase$(pSubStr))
End Function

Private Sub writeNewFile( _
                ByVal pFso As FileSystemObject, _
                ByVal pFilename As String, _
                ByRef pLines As Collection)
Dim lTs As TextStream
Set lTs = pFso.OpenTextFile(pFilename, ForWriting)

Dim v As Variant
For Each v In pLines
    Dim s As String: s = CStr(v)
    lTs.WriteLine s
Next

lTs.Close
End Sub



