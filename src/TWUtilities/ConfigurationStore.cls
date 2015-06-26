VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigurationStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class provides a means to easily store and retrieve application State
' data in a configuration dataset. It is typically used to remember settings that
' can be acted on the next time the application starts.
'
' The actual storage and retrieval facilities are provided by an underlying object
' of <code>ConfigStoreProvider</code> type. This object must be created first, and then
' a <code>ConfigurationStore</code> object can be obtained using the global
' <code>CreateConfigurationStore</code> method.
'
' A <code>ConfigurationStore</code> represents the contents of a configuration dataset
' as a hierarchy of named <em>Configuration Sections</em>. The hierarchy may have any depth
' required by the application. Each configuration section contains one or more named
' <em>settings</em>, and may also have one or more named <em>attributes</em>.
' Attributes give further information about the configuration section: for example,
' a configuration section may have an attribute called Disabled, which indicates
' whether any of the settings within the configuration section (and any child
' configuration sections) are to be actioned.
'
' A configuration section is referenced using a <em>path</em>, which consists of the
' sequence of configuration section names down the hierarchy from the root to the
' target configuration section, separated using '/' characters. The root name
' itself is not specified, but the '/' separating the root name from the
' next name in the Path must be supplied.
'
'

'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IChangeListener

'@================================================================================
' Events
'@================================================================================

Event Change(ev As ChangeEventData)

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "ConfigurationStore"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigStoreProvider                        As ConfigStoreProvider

Private mFilename                                   As String

Private mRootConfigSection                          As ConfigurationSection
Private mAppConfigsSection                          As ConfigurationSection

Private mCache                                      As New ConfigSectionCache

Private mChangeListeners                            As Listeners

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mChangeListeners = New Listeners
End Sub

'@================================================================================
' ChangeListener Interface Members
'@================================================================================

Private Sub IChangeListener_Change(ev As ChangeEventData)
Const ProcName As String = "ChangeListener_Change"
On Error GoTo Err

Set ev.Source = Me

mChangeListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mChangeListeners.Count
    Dim lListener As IChangeListener
    Set lListener = mChangeListeners.GetListener(i)
    lListener.Change ev
Next

RaiseEvent Change(ev)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ApplicationName() As String
Const ProcName As String = "ApplicationName"
On Error GoTo Err

ApplicationName = mConfigStoreProvider.ApplicationName

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Dirty() As Boolean
Const ProcName As String = "Dirty"
On Error GoTo Err

Dirty = mConfigStoreProvider.Dirty

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get fileVersion() As String
Const ProcName As String = "FileVersion"
On Error GoTo Err

fileVersion = mConfigStoreProvider.ApplicationVersion

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Registers an object to act as a listener for changes to the configuration Data.
'
' @remarks
'   The listener object must implement the <code>ChangeListener</code>
'   interface.
'
'   When a change is notified, the <code>change</code> field in the <code>ChangeEventData</code>
'   is a Value from the <code>ConfigChangeTypes</code> enum.
' @param listener
'   The object which is to act as a listener.
'@/
Public Sub AddChangeListener( _
                ByVal pListener As IChangeListener)
Const ProcName As String = "addChangeListener"
On Error GoTo Err

mChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function AddConfigurationSection( _
                ByVal configSectionPath As String, _
                Optional ByVal Comment As String, _
                Optional ByVal RendererProgId As String) As ConfigurationSection
Const ProcName As String = "AddConfigurationSection"
On Error GoTo Err

Set AddConfigurationSection = AddConfigSection(configSectionPath, False, Comment, RendererProgId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function AddPrivateConfigurationSection( _
                ByVal configSectionPath As String, _
                Optional ByVal Comment As String, _
                Optional ByVal RendererProgId As String) As ConfigurationSection
Const ProcName As String = "AddPrivateConfigurationSection"
On Error GoTo Err

Set AddPrivateConfigurationSection = AddConfigSection(configSectionPath, True, Comment, RendererProgId)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetConfigurationSection( _
                ByVal configSectionPath As String) As ConfigurationSection
Const ProcName As String = "GetConfigurationSection"
On Error GoTo Err

If configSectionPath = "" Then
    Set GetConfigurationSection = mRootConfigSection
    Exit Function
End If

Set GetConfigurationSection = mCache.Item(configSectionPath)

If GetConfigurationSection Is Nothing Then
    Dim sectionPath As ConfigurationPath
    Set sectionPath = gGetConfigPath(configSectionPath)
    If Not sectionPath.IsValidSectionPath Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Not a valid configuration section Path"
    
    If sectionPath.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Relative Path not permitted"
    
    Dim parentSection As ConfigurationSection
    Set parentSection = GetConfigurationSection(sectionPath.ParentPath)
    If Not parentSection Is Nothing Then
        Set GetConfigurationSection = parentSection.GetConfigurationSection(sectionPath.TailPath)
    Else
        Set GetConfigurationSection = mRootConfigSection.GetConfigurationSection(sectionPath.TailPath)
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetSetting( _
                ByVal settingPath As String, _
                Optional ByVal defaultValue As String) As String
Const ProcName As String = "GetSetting"
On Error GoTo Err

Dim cp As ConfigurationPath
Set cp = gGetConfigPath(settingPath)
If cp.IsValidAttributePath Then
ElseIf cp.IsValidValuePath Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Not a valid attribute or Value Path"
End If

If cp.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Relative attribute or Value Path not permitted"

Dim section As ConfigurationSection
Set section = GetConfigurationSection(cp.ParentPath)

If section Is Nothing Then
    GetSetting = defaultValue
ElseIf cp.IsValidAttributePath Then
    GetSetting = section.GetAttribute(cp.TailName, defaultValue)
Else
    GetSetting = section.GetConfigurationSection(Right$(cp.TailPath, Len(cp.TailPath) - 1)).Value
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Sub Initialise( _
                ByVal pConfigStoreProvider As ConfigStoreProvider, _
                ByVal pFilename As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

mFilename = pFilename

Set mConfigStoreProvider = pConfigStoreProvider
mConfigStoreProvider.AddChangeListener Me
Set mRootConfigSection = New ConfigurationSection
mRootConfigSection.Initialise Me, "", mConfigStoreProvider.RootItem, Nothing, mCache, False

Set mChangeListeners = New Listeners

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

''
' Unregisters a change listener object.
'
' @param listener
'   The object to be unregistered.
'@/
Public Sub RemoveChangeListener( _
                ByVal pListener As IChangeListener)
Const ProcName As String = "removeChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveConfigurationSection( _
                ByVal configSectionPath As String)
Const ProcName As String = "RemoveConfigurationSection"
On Error GoTo Err

Dim sectionPath As ConfigurationPath
Set sectionPath = gGetConfigPath(configSectionPath)
If Not sectionPath.IsValidSectionPath Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Not a valid configuration section Path"

Dim parentSection As ConfigurationSection
Set parentSection = GetConfigurationSection(sectionPath.ParentPath)
If Not parentSection Is Nothing Then
    parentSection.RemoveConfigurationSection sectionPath.TailPath
Else
    mRootConfigSection.RemoveConfigurationSection sectionPath.TailPath
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveSetting( _
                ByVal settingPath As String)
Const ProcName As String = "RemoveSetting"
On Error GoTo Err

Dim cp As ConfigurationPath
Set cp = gGetConfigPath(settingPath)
If cp.IsValidAttributePath Then
ElseIf cp.IsValidValuePath Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Not a valid attribute or Value Path"
End If

If cp.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Not relative attribute or Value Path not permitted"

Dim section As ConfigurationSection
Set section = GetConfigurationSection(cp.ParentPath)
If section Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid Path - config section does not exist"

section.RemoveSetting cp.TailPath

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub Save( _
                Optional ByVal pFilename As String)
Const ProcName As String = "save"
On Error GoTo Err

If pFilename <> "" Then
    mConfigStoreProvider.Save pFilename
Else
    mConfigStoreProvider.Save mFilename
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetBooleanSetting( _
                ByVal settingPath As String, _
                ByVal Value As Boolean)
Const ProcName As String = "SetBooleanSetting"
On Error GoTo Err

Dim section As ConfigurationSection
Set section = SetSettingEx(settingPath, Value)
section.SetAttribute AttributeNameType, AttributeValueTypeBoolean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
                
Public Sub SetSelectionSetting( _
                ByVal settingPath As String, _
                ByVal Value As String, _
                ByRef permittedValues() As String)
Const ProcName As String = "SetSelectionSetting"
On Error GoTo Err

Dim section As ConfigurationSection
Set section = SetSettingEx(settingPath, CStr(Value))
section.SetAttribute AttributeNameType, AttributeValueTypeSelection

Dim selsItem As ConfigItem
On Error Resume Next
Set selsItem = section.Config.ChildItems.Item(ConfigNameSelections)
On Error GoTo Err

If selsItem Is Nothing Then
    Set selsItem = section.Config.ChildItems.AddItem(ConfigNameSelections)
Else
    selsItem.ChildItems.Clear
End If

Dim permittedValue As Variant
For Each permittedValue In permittedValues
    selsItem.ChildItems.AddItem ConfigNameSelection, permittedValue
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetSetting( _
                ByVal settingPath As String, _
                ByVal Value As String)
Const ProcName As String = "SetSetting"
On Error GoTo Err

SetSettingEx settingPath, Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AddConfigSection( _
                ByVal configSectionPath As String, _
                ByVal isPrivate As Boolean, _
                ByVal Comment As String, _
                ByVal RendererProgId As String) As ConfigurationSection
Const ProcName As String = "AddConfigSection"
On Error GoTo Err

Dim sectionPath As ConfigurationPath
Set sectionPath = gGetConfigPath(configSectionPath)

If Not sectionPath.IsValidSectionPath Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid section Path"

If sectionPath.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Relative section Path not permitted"

Set AddConfigSection = GetConfigurationSection(configSectionPath)

If AddConfigSection Is Nothing Then
    If sectionPath.ParentPath = "" Then
        If isPrivate Then
            Set AddConfigSection = mRootConfigSection.AddPrivateConfigurationSection(sectionPath.TailPath, Comment, RendererProgId)
        Else
            Set AddConfigSection = mRootConfigSection.AddConfigurationSection(sectionPath.TailPath, Comment, RendererProgId)
        End If
    Else
        Dim parentSection As ConfigurationSection
        Set parentSection = GetConfigurationSection(sectionPath.ParentPath)
        If parentSection Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid Path - config section does not exist"
        
        If isPrivate Then
            Set AddConfigSection = parentSection.AddPrivateConfigurationSection(sectionPath.TailPath, Comment, RendererProgId)
        Else
            Set AddConfigSection = parentSection.AddConfigurationSection(sectionPath.TailPath, Comment, RendererProgId)
        End If
        
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function SetSettingEx( _
                ByVal settingPath As String, _
                ByVal Value As String) As ConfigurationSection
Const ProcName As String = "SetSettingEx"
On Error GoTo Err

Dim cp As ConfigurationPath
Set cp = gGetConfigPath(settingPath)
If cp.IsValidAttributePath Then
ElseIf cp.IsValidValuePath Then
Else
    Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Not a valid attribute or Value Path"
End If

If cp.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Not relative attribute or Value Path not permitted"

Dim section As ConfigurationSection
Set section = GetConfigurationSection(cp.ParentPath)
If section Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid Path - config section does not exist"

Set SetSettingEx = section.SetSettingEx(gGetConfigPath(cp.TailPath), Value)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function



