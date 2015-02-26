Attribute VB_Name = "GConfigurationStore"
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

Private Const ModuleName                            As String = "GConfigurationStore"

Public Const AttributeNameName                      As String = "__Name"
Public Const AttributeNamePrivate                   As String = "__Private"
Public Const AttributeNameRenderer                  As String = "__Renderer"
Public Const AttributeNameType                      As String = "__Type"

Public Const AttributeValueFalse                    As String = "False"
Public Const AttributeValueTrue                     As String = "True"
Public Const AttributeValueTypeBoolean              As String = "Boolean"
Public Const AttributeValueTypeSelection            As String = "Selection"

Public Const ConfigNameSelection                    As String = "__Selection"
Public Const ConfigNameSelections                   As String = "__Selections"

Public Const ConfigSectionPathSeparator             As String = "/"
Public Const AttributePathNameSeparator             As String = "&"
Public Const ValuePathNameSeparator                 As String = "."
Public Const RootSectionName                        As String = "Configuration"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigPaths                                As New Collection

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

Public Function gGetConfigPath( _
                ByVal Path As String) As ConfigurationPath
Const ProcName As String = "gGetConfigPath"

On Error GoTo Err

On Error Resume Next
Set gGetConfigPath = mConfigPaths(Path)
On Error GoTo Err

If gGetConfigPath Is Nothing Then
    Set gGetConfigPath = New ConfigurationPath
    gGetConfigPath.Initialise Path
    mConfigPaths.Add gGetConfigPath, Path
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================


