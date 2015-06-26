VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ConfigurationSection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements Enumerable

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

Private Const ModuleName                            As String = "ConfigurationSection"

'@================================================================================
' Member variables
'@================================================================================

Private mConfigFile                                 As ConfigurationStore
Private mConfig                                     As IConfigItem
Private mConfigEnumble                              As Enumerable
Private mRelativePath                               As String
Private mCache                                      As ConfigSectionCache
Private mIsPrivate                                  As Boolean

Private mParentRef                                  As WeakReference

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Enumerable Interface Members
'@================================================================================

Private Property Get Enumerable_Enumerator() As Enumerator
Const ProcName As String = "Enumerable_Enumerator"

On Error GoTo Err

Set Enumerable_Enumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Function Enumerable_NextItem(Item As Variant, pCookie As Variant) As Boolean
Dim en As Enumerator
Dim ci As IConfigItem
Dim section As ConfigurationSection
Dim relPath As String

Const ProcName As String = "Enumerable_NextItem"

On Error GoTo Err

If IsEmpty(pCookie) Then Set pCookie = mConfig.ChildItems.Enumerator

Set en = pCookie
If en.MoveNext Then
    Set ci = en.Current
    If ci.Name = ConfigNameSelections Then
        ' we mustn't return this!
        If Not en.MoveNext Then
            Item = Empty
            Exit Function
        End If
        Set ci = en.Current
    End If
    relPath = ci.Name
    If ci.GetDefaultableAttribute(AttributeNameName, "") <> "" Then relPath = relPath & "(" & ci.GetAttribute(AttributeNameName) & ")"
    Set section = GetConfigurationSection(relPath)
    
    If section Is Nothing Then
        Set section = New ConfigurationSection
        section.Initialise mConfigFile, _
                            relPath, _
                            ci, _
                            Me, _
                            mCache, _
                            CBool(ci.GetDefaultableAttribute(AttributeNamePrivate, AttributeValueFalse))
        mCache.Add section
    End If
    Set Item = section
    Enumerable_NextItem = True
Else
    Item = Empty
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Sub Enumerable_Remove(pCookie As Variant)
Dim en As Enumerator
Dim ci As IConfigItem
Dim Path As String

Const ProcName As String = "Enumerable_Remove"

On Error GoTo Err

Set en = pCookie
Set ci = en.Current
Path = ci.Name
If ci.GetAttribute(AttributeNameName) Then Path = Path & "(" & ci.GetAttribute(AttributeNameName) & ")"
RemoveConfigurationSection Path

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Function Enumerable_Skip(ByVal numToSkip As Long, pCookie As Variant) As Boolean
Dim en As Enumerator
Dim i As Long

Const ProcName As String = "Enumerable_Skip"

On Error GoTo Err

Set en = pCookie

For i = 1 To numToSkip
    If Not en.MoveNext Then Exit Function
Next

Enumerable_Skip = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Property Get Enumerable_UpdateSequence() As Long
Enumerable_UpdateSequence = mConfigEnumble.UpdateSequence
End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get ChildEnumerator() As Enumerator
Const ProcName As String = "ChildEnumerator"

On Error GoTo Err

Set ChildEnumerator = getEnumerator

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Comment() As String
Const ProcName As String = "Comment"

On Error GoTo Err

Comment = mConfig.Comment

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Comment( _
                ByVal Value As String)
Const ProcName As String = "Comment"

On Error GoTo Err

mConfig.Comment = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ConfigurationStore() As ConfigurationStore
Const ProcName As String = "ConfigurationStore"

On Error GoTo Err

Set ConfigurationStore = mConfigFile

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get Config() As IConfigItem
Const ProcName As String = "Config"

On Error GoTo Err

Set Config = mConfig

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let InstanceQualifier( _
                ByVal Value As String)
Dim newRelativePath As String
Dim configPath As ConfigurationPath

Const ProcName As String = "InstanceQualifier"

On Error GoTo Err

Set configPath = gGetConfigPath(mRelativePath)
If configPath.InstanceQualifier = "" Then Err.Raise ErrorCodes.ErrIllegalStateException, , "Config section doesn't currently have an instance qualifier"


newRelativePath = Replace(mRelativePath, _
                            "(" & configPath.InstanceQualifier & ")", _
                            "(" & Value & ")", _
                            1, _
                            1, _
                            vbTextCompare)
mCache.rename Path, Parent.Path & ConfigSectionPathSeparator & newRelativePath
mRelativePath = newRelativePath
SetAttribute AttributeNameName, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get InstanceQualifier() As String
Const ProcName As String = "InstanceQualifier"

On Error GoTo Err

InstanceQualifier = GetAttribute(AttributeNameName)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Name() As String
Const ProcName As String = "Name"

On Error GoTo Err

Name = mConfig.Name

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Parent() As ConfigurationSection
Const ProcName As String = "Parent"

On Error GoTo Err

If Not mParentRef Is Nothing Then Set Parent = mParentRef.target

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Path() As String
Const ProcName As String = "Path"

On Error GoTo Err

If Parent Is Nothing Then
    Path = mRelativePath
Else
    Path = Parent.Path & ConfigSectionPathSeparator & mRelativePath
End If

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get RendererProgId() As String
Const ProcName As String = "RendererProgId"

On Error GoTo Err

RendererProgId = GetAttribute(AttributeNameRenderer, "")

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let RendererProgId( _
                ByVal Value As String)
Const ProcName As String = "RendererProgId"

On Error GoTo Err

SetAttribute AttributeNameRenderer, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Let Value(ByVal Value As String)
Const ProcName As String = "Value"

On Error GoTo Err

mConfig.Value = Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get Value() As String
Const ProcName As String = "Value"

On Error GoTo Err

Value = mConfig.Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub AddChildSections(ByVal pConfigItem As IConfigItem)
Const ProcName As String = "AddChildSections"

On Error GoTo Err

Dim lItem As IConfigItem
Dim en As Enumerator

Set en = pConfigItem.ChildItems.Enumerator
Do While en.MoveNext
    Set lItem = en.Current
    
    Dim lPath As String
    lPath = Path & "/" & lItem.Name
    Dim lInstanceQual As String
    lInstanceQual = lItem.GetDefaultableAttribute(AttributeNameName, "")
    If lInstanceQual <> "" Then lPath = lPath & "(" & lInstanceQual & ")"
    
    Dim lConfigSection As New ConfigurationSection
    lConfigSection.Initialise mConfigFile, _
                            lPath, _
                            lItem, _
                            Me, _
                            mCache, _
                            CBool(lItem.GetDefaultableAttribute(AttributeNamePrivate, "False"))
    
    lConfigSection.AddChildSections lItem
Loop

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

Set AddConfigurationSection = AddConfigSection(configSectionPath, mIsPrivate, Comment, RendererProgId)

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

Public Function CloneConfigSection( _
                ByVal pConfigSection As ConfigurationSection) As ConfigurationSection
Const ProcName As String = "CloneConfigSection"

On Error GoTo Err

Dim lRelativePath As String
lRelativePath = pConfigSection.Name
If pConfigSection.Config.GetDefaultableAttribute(AttributeNameName, "") <> "" Then lRelativePath = lRelativePath & "(" & pConfigSection.Config.GetAttribute(AttributeNameName) & ")"

Dim lNewConfigSection As ConfigurationSection
Set lNewConfigSection = mConfigFile.GetConfigurationSection(Path & ConfigSectionPathSeparator & lRelativePath)

If Not lNewConfigSection Is Nothing Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Target config section already exists"
    
Set lNewConfigSection = New ConfigurationSection
Dim lClonedConfigItem As IConfigItem
Set lClonedConfigItem = mConfig.ChildItems.AddClone(pConfigSection.Config, True)
lNewConfigSection.Initialise mConfigFile, _
                                lRelativePath, _
                                lClonedConfigItem, _
                                Me, _
                                mCache, _
                                CBool(pConfigSection.Config.GetDefaultableAttribute(AttributeNamePrivate, "False"))
mCache.Add lNewConfigSection

' now we need to clone the subtree of ConfigurationSections from the cloned ConfigItem
AddChildSections lClonedConfigItem

Set CloneConfigSection = lNewConfigSection

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetAttribute( _
                ByVal Name As String, _
                Optional ByVal defaultValue As String) As String
Const ProcName As String = "GetAttribute"

On Error GoTo Err

GetAttribute = mConfig.GetDefaultableAttribute(Name, defaultValue)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetConfigurationSection( _
                ByVal configSectionPath As String) As ConfigurationSection
Dim relativeSectionPath As ConfigurationPath
Dim childItem As IConfigItem

Dim absPath As String

Const ProcName As String = "GetConfigurationSection"

On Error GoTo Err

Set relativeSectionPath = gGetConfigPath(configSectionPath)

If Not relativeSectionPath.IsValidSectionPath Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid section Path"
If Not relativeSectionPath.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "A relative config section Path is required"

absPath = Path & ConfigSectionPathSeparator & configSectionPath
Set GetConfigurationSection = mCache.Item(absPath)

If GetConfigurationSection Is Nothing Then
    If relativeSectionPath.ParentPath = "" Then
        Set childItem = getChildItem(relativeSectionPath.TailName, relativeSectionPath.InstanceQualifier)
        If Not childItem Is Nothing Then
            Set GetConfigurationSection = New ConfigurationSection
            GetConfigurationSection.Initialise mConfigFile, _
                                            relativeSectionPath.Path, _
                                            childItem, _
                                            Me, _
                                            mCache, _
                                            childItem.GetDefaultableAttribute(AttributeNamePrivate, AttributeValueFalse) = AttributeValueTrue
            mCache.Add GetConfigurationSection
        End If
    Else
        ' the required config section is not a direct child of this one
        Set GetConfigurationSection = mConfigFile.GetConfigurationSection(absPath)
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function getNextSection() As ConfigurationSection
Dim section As ConfigurationSection
Dim en As Enumerator

Const ProcName As String = "getNextSection"

On Error GoTo Err

If Parent Is Nothing Then Exit Function

Set en = Parent.ChildEnumerator
        
Do While en.MoveNext
    Set section = en.Current
    If section Is Me Then Exit Do
Loop

If en.MoveNext Then Set getNextSection = en.Current

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function getPreviousSection() As ConfigurationSection
Dim section As ConfigurationSection
Dim en As Enumerator

Const ProcName As String = "getPreviousSection"

On Error GoTo Err

If Parent Is Nothing Then Exit Function

Set en = Parent.ChildEnumerator
        
Do While en.MoveNext
    Set getPreviousSection = section
    Set section = en.Current
    If section Is Me Then Exit Do
Loop

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function GetSetting( _
                ByVal relativeSettingPath As String, _
                Optional ByVal defaultValue As String) As String
Dim relSettingPath As ConfigurationPath
Dim section As ConfigurationSection

Const ProcName As String = "GetSetting"

On Error GoTo Err

Set relSettingPath = gGetConfigPath(relativeSettingPath)

If Not relSettingPath.IsValidAttributePath And Not relSettingPath.IsValidValuePath Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid setting Path"

If Not relSettingPath.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "A relative setting Path is required"

If relSettingPath.IsValidAttributePath Then
    If relSettingPath.ParentPath = "" Then
        GetSetting = mConfig.GetDefaultableAttribute(relSettingPath.TailName, defaultValue)
    Else
        GetSetting = mConfigFile.GetSetting(Path & ConfigSectionPathSeparator & relSettingPath.Path, defaultValue)
    End If
Else
    If relSettingPath.ParentPath = "" Then
        Set section = GetConfigurationSection(relSettingPath.TailName)
    Else
        Set section = GetConfigurationSection(relSettingPath.ParentPath & ConfigSectionPathSeparator & relSettingPath.TailName)
    End If
    If section Is Nothing Then
        GetSetting = defaultValue
    Else
        GetSetting = section.Value
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Friend Function Initialise( _
                ByVal pConfigFile As ConfigurationStore, _
                ByVal relativePath As String, _
                ByVal Config As IConfigItem, _
                ByVal parentConfigSection As ConfigurationSection, _
                ByVal cache As ConfigSectionCache, _
                ByVal isPrivate As Boolean)
Const ProcName As String = "Initialise"

On Error GoTo Err

Set mConfigFile = pConfigFile
Set mConfig = Config
If Not parentConfigSection Is Nothing Then Set mParentRef = gCreateWeakReference(parentConfigSection)
Set mConfigEnumble = mConfig.ChildItems
mRelativePath = relativePath
Set mCache = cache
mIsPrivate = isPrivate

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function MoveDown() As Boolean
Dim targetSection As ConfigurationSection

Const ProcName As String = "MoveDown"

On Error GoTo Err

Set targetSection = getNextSection
If Not targetSection Is Nothing Then
    Set targetSection = targetSection.getNextSection
    If Not targetSection Is Nothing Then
        Parent.Config.ChildItems.MoveItemBefore Me.Config, targetSection.Config
    Else
        Parent.Config.ChildItems.MoveItemBefore Me.Config, Nothing
    End If
    MoveDown = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function MoveUp() As Boolean
Dim targetSection As ConfigurationSection

Const ProcName As String = "MoveUp"

On Error GoTo Err

Set targetSection = getPreviousSection
If Not targetSection Is Nothing Then
    Parent.Config.ChildItems.MoveItemBefore Me.Config, targetSection.Config
    MoveUp = True
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
Set NewEnum = getEnumerator
End Function

Public Sub Remove()
Const ProcName As String = "Remove"

On Error GoTo Err

mConfigFile.RemoveConfigurationSection Path

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveAttribute( _
                ByVal Name As String)
Const ProcName As String = "RemoveAttribute"

On Error GoTo Err

mConfig.RemoveAttribute Name

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveAllChildren()
Dim en As Enumerator
Dim sections As New Collection
Dim section As ConfigurationSection

Const ProcName As String = "RemoveAllChildren"

On Error GoTo Err

Set en = ChildEnumerator
Do While en.MoveNext
    Set section = en.Current
    sections.Add section
Loop

For Each section In sections
    mConfig.ChildItems.Remove section.Config
    mCache.Remove section
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveConfigurationSection( _
                ByVal configSectionPath As String)
Dim relativeSectionPath As ConfigurationPath
Dim section As ConfigurationSection

Const ProcName As String = "RemoveConfigurationSection"

On Error GoTo Err

Set relativeSectionPath = gGetConfigPath(configSectionPath)

If Not relativeSectionPath.IsValidSectionPath Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid section Path"
If Not relativeSectionPath.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "A relative config section Path is required"

If relativeSectionPath.ParentPath = "" Then
    Set section = GetConfigurationSection(relativeSectionPath.Path)
    If Not section Is Nothing Then
        mConfig.ChildItems.Remove section.Config
        mCache.Remove section
    End If
Else
    ' the config section to be removed is not a direct child of this one
    mConfigFile.RemoveConfigurationSection Path & ConfigSectionPathSeparator & configSectionPath
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveSetting( _
                ByVal relativeSettingPath As String)
Dim relSettingPath As ConfigurationPath
Dim section As ConfigurationSection

Const ProcName As String = "RemoveSetting"

On Error GoTo Err

Set relSettingPath = gGetConfigPath(relativeSettingPath)

If Not relSettingPath.IsValidAttributePath And Not relSettingPath.IsValidValuePath Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid setting Path"
If Not relSettingPath.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "A relative setting Path is required"

If relSettingPath.IsValidAttributePath Then
    If relSettingPath.ParentPath = "" Then
        mConfig.RemoveAttribute relSettingPath.TailName
    Else
        mConfigFile.RemoveSetting Path & ConfigSectionPathSeparator & relSettingPath.Path
    End If
Else
    Set section = GetConfigurationSection(Right$(relSettingPath.Path, Len(relSettingPath.Path) - 1))
    If Not section Is Nothing Then
        mConfig.ChildItems.Remove section.Config
        mCache.Remove section
    End If
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Function SetAttribute( _
                ByVal Name As String, _
                ByVal Value As String) As String
Const ProcName As String = "SetAttribute"

On Error GoTo Err

mConfig.SetAttribute Name, Value

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Sub SetBooleanSetting( _
                ByVal relativeSettingPath As String, _
                ByVal Value As Boolean)
Const ProcName As String = "SetBooleanSetting"

On Error GoTo Err

Dim section As ConfigurationSection
Dim relSettingPath As ConfigurationPath

Set relSettingPath = gGetConfigPath(relativeSettingPath)

Set section = SetSettingEx(relSettingPath, CStr(Value))
If Not relSettingPath.IsValidAttributePath Then section.SetAttribute AttributeNameType, AttributeValueTypeBoolean

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetSelectionSetting( _
                ByVal relativeSettingPath As String, _
                ByVal Value As String, _
                ByRef permittedValues() As String)
Const ProcName As String = "SetSelectionSetting"

On Error GoTo Err

Dim section As ConfigurationSection
Dim permittedValue As Variant
Dim selsItem As IConfigItem

Dim relSettingPath As ConfigurationPath

Set relSettingPath = gGetConfigPath(relativeSettingPath)

Set section = SetSettingEx(relSettingPath, CStr(Value))
If Not relSettingPath.IsValidAttributePath Then section.SetAttribute AttributeNameType, AttributeValueTypeSelection

On Error Resume Next
Set selsItem = section.Config.ChildItems.Item(ConfigNameSelections)
On Error GoTo Err

If selsItem Is Nothing Then
    Set selsItem = section.Config.ChildItems.AddItem(ConfigNameSelections)
Else
    selsItem.ChildItems.Clear
End If

For Each permittedValue In permittedValues
    selsItem.ChildItems.AddItem ConfigNameSelection, permittedValue
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub SetSetting( _
                ByVal relativeSettingPath As String, _
                ByVal Value As String)
Const ProcName As String = "SetSetting"

On Error GoTo Err

Dim relSettingPath As ConfigurationPath

Set relSettingPath = gGetConfigPath(relativeSettingPath)

SetSettingEx relSettingPath, Value

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Function SetSettingEx( _
                ByVal relSettingPath As ConfigurationPath, _
                ByVal Value As String) As ConfigurationSection
Dim section As ConfigurationSection

Const ProcName As String = "SetSettingEx"

On Error GoTo Err

If Not relSettingPath.IsValidAttributePath And Not relSettingPath.IsValidValuePath Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid setting Path"
If Not relSettingPath.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "A relative setting Path is required"

If relSettingPath.IsValidAttributePath Then
    If relSettingPath.ParentPath = "" Then
        Set section = Me
        section.SetAttribute relSettingPath.TailName, Value
    Else
        Set section = mConfigFile.GetConfigurationSection(Path & ConfigSectionPathSeparator & relSettingPath.ParentPath)
        section.SetAttribute relSettingPath.TailName, Value
    End If
ElseIf relSettingPath.ParentPath = "" Then
        Set section = AddConfigurationSection(relSettingPath.TailName)
        section.Value = Value
Else
    Set section = mConfigFile.AddConfigurationSection(Path & ConfigSectionPathSeparator & relSettingPath.ParentPath & ConfigSectionPathSeparator & relSettingPath.TailName)
    section.Value = Value
End If

Set SetSettingEx = section

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================

Private Function AddConfigSection( _
                ByVal pNewConfigSectionPath As String, _
                ByVal isPrivate As Boolean, _
                ByVal Comment As String, _
                ByVal RendererProgId As String) As ConfigurationSection
Dim relativeSectionPath As ConfigurationPath

Const ProcName As String = "AddConfigSection"

On Error GoTo Err

Set relativeSectionPath = gGetConfigPath(pNewConfigSectionPath)

If Not relativeSectionPath.IsValidSectionPath Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "Invalid section Path"
If Not relativeSectionPath.IsRelative Then Err.Raise ErrorCodes.ErrIllegalArgumentException, , "A relative config section Path is required"

Set AddConfigSection = mConfigFile.GetConfigurationSection(Path & ConfigSectionPathSeparator & pNewConfigSectionPath)

If AddConfigSection Is Nothing Then
    If relativeSectionPath.ParentPath = "" Then
        Set AddConfigSection = New ConfigurationSection
        AddConfigSection.Initialise mConfigFile, _
                                        relativeSectionPath.Path, _
                                        mConfig.ChildItems.AddItem(relativeSectionPath.TailName), _
                                        Me, _
                                        mCache, _
                                        isPrivate
        If Comment <> "" Then AddConfigSection.Comment = Comment
        If isPrivate And Not mIsPrivate Then AddConfigSection.SetAttribute AttributeNamePrivate, _
                                                        AttributeValueTrue
        If relativeSectionPath.InstanceQualifier <> "" Then
            AddConfigSection.SetAttribute AttributeNameName, _
                                                relativeSectionPath.InstanceQualifier
        End If
        If RendererProgId <> "" Then AddConfigSection.SetAttribute AttributeNameRenderer, _
                                                        RendererProgId
        mCache.Add AddConfigSection
    Else
        ' the required new config section is not a direct child of this one
        If isPrivate Then
            Set AddConfigSection = mConfigFile.AddPrivateConfigurationSection(Path & ConfigSectionPathSeparator & pNewConfigSectionPath, Comment, RendererProgId)
        Else
            Set AddConfigSection = mConfigFile.AddConfigurationSection(Path & ConfigSectionPathSeparator & pNewConfigSectionPath, Comment, RendererProgId)
        End If
    End If
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getChildItem( _
                ByVal Name As String, _
                ByVal qualifier) As IConfigItem
Dim childEnum As Enumerator

Const ProcName As String = "getChildItem"

On Error GoTo Err

Set childEnum = mConfig.ChildItems.Enumerator

qualifier = UCase$(qualifier)

If qualifier = "" Then
    On Error Resume Next
    Set getChildItem = mConfig.ChildItems.Item(Name)
    On Error GoTo Err
Else
    Do While childEnum.MoveNext
        Set getChildItem = childEnum.Current
        If getChildItem.Name = Name And _
            UCase$(getChildItem.GetDefaultableAttribute(AttributeNameName, "")) = qualifier _
        Then
            Exit Function
        End If
    Loop
    Set getChildItem = Nothing
End If

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Private Function getEnumerator() As Enumerator
Const ProcName As String = "getEnumerator"

On Error GoTo Err

Set getEnumerator = New Enumerator
getEnumerator.Initialise Me

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function


