Attribute VB_Name = "GTWUtilities"
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

Private Const ModuleName                            As String = "GTWUtilities"

Private Const DefaultApplicationName                As String = "Application"

Private Const SHGFP_TYPE_CURRENT = &H0
Private Const SHGFP_TYPE_DEFAULT = &H1

Private Const SwitchConfig                          As String = "Config"
Private Const SwitchSettings                        As String = "Settings"

'@================================================================================
' Member variables
'@================================================================================

Private mApplicationName As String
Private mApplicationGroupName As String
Private mDefaultLogListener As ILogListener

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

Public Property Let gApplicationName(ByVal Value As String)
mApplicationName = Value
End Property

Public Property Get gApplicationName() As String
If mApplicationName = "" Then mApplicationName = DefaultApplicationName
gApplicationName = mApplicationName
End Property

Public Property Let gApplicationGroupName(ByVal Value As String)
mApplicationGroupName = Value
End Property

Public Property Get gApplicationGroupName() As String
gApplicationGroupName = mApplicationGroupName
End Property

Public Property Get gApplicationSettingsFolder() As String
Const ProcName As String = "gApplicationSettingsFolder"

On Error GoTo Err

gApplicationSettingsFolder = gGetSpecialFolderPath(FolderIdAppdata) & _
                    IIf(gApplicationGroupName <> "", "\" & gApplicationGroupName & "\", "\") & _
                    gApplicationName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let gDefaultLogListener(ByVal Value As ILogListener)
Set mDefaultLogListener = Value
End Property

Public Property Get gDefaultLogListener() As ILogListener
Set gDefaultLogListener = mDefaultLogListener
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateConfigurationStore( _
                ByVal pConfigStoreProvider As IConfigStoreProvider, _
                ByVal pFilename As String) As ConfigurationStore
Const ProcName As String = "gCreateConfigurationStore"

On Error GoTo Err

Set gCreateConfigurationStore = New ConfigurationStore
gCreateConfigurationStore.Initialise pConfigStoreProvider, _
                                    pFilename

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreateXMLConfigurationProvider( _
                Optional ByVal pApplicationName As String, _
                Optional ByVal pApplicationVersion As String) As IConfigStoreProvider
                
Dim xmlConfigProvider As xmlConfigProvider

Const ProcName As String = "gCreateXMLConfigurationProvider"

On Error GoTo Err

Set xmlConfigProvider = New xmlConfigProvider
xmlConfigProvider.Initialise pApplicationName, pApplicationVersion
Set gCreateXMLConfigurationProvider = xmlConfigProvider

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetCommandLineParser(ByVal pCommandLine As String) As CommandLineParser
Const ProcName As String = "gGetCommandLineParser"
Static clp As CommandLineParser

On Error GoTo Err

If clp Is Nothing Then
    Set clp = New CommandLineParser
    clp.Initialise pCommandLine, " "
End If
Set gGetCommandLineParser = clp

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetDefaultConfigurationStore( _
                ByVal pCommandLine As String, _
                ByVal pConfigFileVersion As String, _
                ByVal pCreate As Boolean, _
                ByVal pIgnoreInvalid As Boolean, _
                ByVal pOptions As ConfigFileOptions) As ConfigurationStore
Const ProcName As String = "gGetDefaultConfigurationStore"
Dim baseConfigStoreProvider As IConfigStoreProvider

Select Case pOptions
    Case ConfigFileOptionNone
    Case ConfigFileOptionFirstArg
    Case ConfigFileOptionConfigSwitch
    Case ConfigFileOptionSettingsSwitch
    Case Else
        Err.Raise ErrorCodes.ErrIllegalArgumentException, , "pOptions must be a member of the ConfigFileOptions enum"
End Select

On Error Resume Next
Set baseConfigStoreProvider = gLoadConfigProviderFromXMLFile(getConfigFilename(pCommandLine, pOptions))
On Error GoTo Err

If pCreate Then
    Set baseConfigStoreProvider = gCreateXMLConfigurationProvider(gApplicationName, pConfigFileVersion)
    Set gGetDefaultConfigurationStore = gCreateConfigurationStore(baseConfigStoreProvider, getConfigFilename(pCommandLine, pOptions))
ElseIf baseConfigStoreProvider Is Nothing Then
Else
    Set gGetDefaultConfigurationStore = gCreateConfigurationStore(baseConfigStoreProvider, _
                                                       getConfigFilename(pCommandLine, pOptions))
    If gGetDefaultConfigurationStore.ApplicationName <> gApplicationName Or _
        gGetDefaultConfigurationStore.fileVersion <> pConfigFileVersion _
    Then
        Dim lErrMsg As String
        lErrMsg = "The configuration store is not the correct format for this program" & vbCrLf & _
                "Current app name is " & gApplicationName & "; config store app name is " & gGetDefaultConfigurationStore.ApplicationName & vbCrLf & _
                "Required file version is " & pConfigFileVersion & "; config store file version is " & gGetDefaultConfigurationStore.fileVersion

        Set gGetDefaultConfigurationStore = Nothing
        
        If pIgnoreInvalid Then
            gLogger.Log lErrMsg, ProcName, ModuleName, LogLevelSevere
        Else
            Err.Raise ErrorCodes.ErrIllegalStateException, , lErrMsg
        End If
    End If

End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gGetSpecialFolderPath( _
                ByVal folderId As FolderIdentifiers) As String
Const ProcName As String = "gGetSpecialFolderPath"

On Error GoTo Err

gGetSpecialFolderPath = getFolderPath(folderId, SHGFP_TYPE_CURRENT)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gLoadConfigProviderFromXMLFile( _
                ByVal filePath As String) As IConfigStoreProvider
Dim xmlConfigProvider As xmlConfigProvider

Const ProcName As String = "gLoadConfigProviderFromXMLFile"

On Error GoTo Err

Set xmlConfigProvider = New xmlConfigProvider
xmlConfigProvider.load filePath

Set gLoadConfigProviderFromXMLFile = xmlConfigProvider

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function getConfigFilename( _
                ByVal pCommandLine As String, _
                ByVal pOptions As ConfigFileOptions) As String
Const ProcName As String = "getConfigFilename"


On Error GoTo Err

Select Case pOptions
Case ConfigFileOptionNone

Case ConfigFileOptionFirstArg
    getConfigFilename = gGetCommandLineParser(pCommandLine).Arg(0)
Case ConfigFileOptionConfigSwitch
    getConfigFilename = gGetCommandLineParser(pCommandLine).SwitchValue(SwitchConfig)
Case ConfigFileOptionSettingsSwitch
    getConfigFilename = gGetCommandLineParser(pCommandLine).SwitchValue(SwitchSettings)
End Select

If getConfigFilename = "" Then
    getConfigFilename = gApplicationSettingsFolder & "\settings.xml"
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getFolderPath( _
                folderId As FolderIdentifiers, _
                SHGFP_TYPE As Long) As String

Dim lBuff As String
Dim lResult As Long

Const ProcName As String = "getFolderPath"

On Error GoTo Err

lBuff = Space$(MAX_PATH)

lResult = SHGetFolderPath(-1, _
                   folderId, _
                   -1, _
                   SHGFP_TYPE, _
                   lBuff)

If lResult = S_OK Then
    getFolderPath = gTrimNull(lBuff)
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
   
End Function


