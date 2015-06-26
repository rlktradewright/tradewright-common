VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigSectionCache"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ConfigSectionCache"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigSections                             As New Collection

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

Friend Function Add( _
                ByVal pConfigSection As ConfigurationSection) As ConfigurationSection
Const ProcName As String = "Add"

On Error GoTo Err

mConfigSections.Add pConfigSection, pConfigSection.Path
Set Add = pConfigSection

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function Item( _
                ByVal Path As String) As ConfigurationSection
On Error Resume Next
Set Item = mConfigSections.Item(Path)
End Function

Friend Sub Remove( _
                ByVal pConfigSection As ConfigurationSection)
Dim i As Long
Dim section As ConfigurationSection

Const ProcName As String = "Remove"

On Error GoTo Err

For i = mConfigSections.Count To 1 Step -1
    Set section = mConfigSections(i)
    If InStr(1, section.Path, pConfigSection.Path) = 1 Then
        mConfigSections.Remove i
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub rename( _
                ByVal oldPath As String, _
                ByVal newPath As String)
Dim i As Long
Dim section As ConfigurationSection
Dim s As String

Const ProcName As String = "rename"

On Error GoTo Err

For i = mConfigSections.Count To 1 Step -1
    Set section = mConfigSections(i)
    If InStr(1, section.Path, oldPath) = 1 Then
        mConfigSections.Remove i
        s = section.Path
        s = Replace(s, oldPath, newPath, 1, 1, vbTextCompare)
        mConfigSections.Add section, s
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveChildren( _
                ByVal pConfigSection As ConfigurationSection)
Dim i As Long
Dim section As ConfigurationSection

Const ProcName As String = "RemoveChildren"

On Error GoTo Err

For i = mConfigSections.Count To 1 Step -1
    Set section = mConfigSections(i)
    If InStr(1, section.Path, pConfigSection.Path) = 1 And Not section Is pConfigSection Then
        mConfigSections.Remove i
    End If
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================


