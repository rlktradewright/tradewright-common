VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLConfigProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class is an implementation of the <code>ConfigStoreProvider</code> interface that
' stores the configuration Data in an XML file.
'
' Objects of this class cannot be created using the <code>New</code> operator. Use
' the <code>CreateXMLConfigurationProvider</code> or <code>LoadConfigProviderFromXMLFile</code>
' global methods.
'
' To access the configuration Data, set an object variable of type <code>ConfigStoreProvider</code>
' to this object, for example:
'
' <pre>
' <tt>
' Dim cf As ConfigStoreProvider<br />
' Dim root As IConfigItem<br />
' Dim outputFilename As String<br />
' <br />
' Set cf = LoadConfigProviderFromXMLFile("C:\MyConfig.xml")<br />
' Set root = cf.rootItem    ' this is actually an XMLConfigItem object<br />
' outputFilename = root.childItems.Item("outputFilename").value<br />
' </tt>
' </pre>
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements ConfigStoreProvider

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


Private Const ModuleName                    As String = "XMLConfigProvider"

Private Const AttributeNameAppName          As String = "__AppName"
Private Const AttributeNameAppVersion       As String = "__AppVersion"
Private Const AttributeNameFileVersion      As String = "__FileVersion"

Private Const ElementNameRoot               As String = "Configuration"

'@================================================================================
' Member variables
'@================================================================================

Private mXMLdoc                             As DOMDocument60
Attribute mXMLdoc.VB_VarHelpID = -1

Private mRootItem                           As XMLConfigItem

Private mFilePath                           As String

Private mChangeListeners                    As New Listeners

Private mDirty                              As Boolean

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Initialize()
Set mChangeListeners = New Listeners
End Sub

'@================================================================================
' ConfigStoreProvider Interface Members
'@================================================================================

Private Sub ConfigStoreProvider_addChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "ConfigStoreProvider_addChangeListener"
On Error GoTo Err

mChangeListeners.Add pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ConfigStoreProvider_applicationName() As String
ConfigStoreProvider_applicationName = ApplicationName
End Property

Private Property Let ConfigStoreProvider_applicationName(ByVal RHS As String)
ApplicationName = RHS
End Property

Private Property Let ConfigStoreProvider_applicationVersion(ByVal RHS As String)
ApplicationVersion = RHS
End Property

Private Property Get ConfigStoreProvider_applicationVersion() As String
ConfigStoreProvider_applicationVersion = ApplicationVersion
End Property

Private Property Get ConfigStoreProvider_dirty() As Boolean
ConfigStoreProvider_dirty = mDirty
End Property

Private Sub ConfigStoreProvider_removeChangeListener(ByVal pListener As IChangeListener)
Const ProcName As String = "ConfigStoreProvider_removeChangeListener"
On Error GoTo Err

mChangeListeners.Remove pListener

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Property Get ConfigStoreProvider_rootItem() As IConfigItem
Set ConfigStoreProvider_rootItem = RootItem
End Property

Private Sub ConfigStoreProvider_save(Optional ByVal filePath As String)
Const ProcName As String = "ConfigStoreProvider_save"
On Error GoTo Err

Save filePath

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Properties
'@================================================================================

Friend Property Let ApplicationName( _
                ByVal Value As String)
Const ProcName As String = "ApplicationName"
On Error GoTo Err

mRootItem.SetAttribute AttributeNameAppName, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ApplicationName() As String
Const ProcName As String = "ApplicationName"
On Error GoTo Err

ApplicationName = mRootItem.GetAttribute(AttributeNameAppName)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Let ApplicationVersion( _
                ByVal Value As String)
Const ProcName As String = "applicationVersion"
On Error GoTo Err

mRootItem.SetAttribute AttributeNameAppVersion, Value

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get ApplicationVersion() As String
Const ProcName As String = "applicationVersion"
On Error GoTo Err

ApplicationVersion = mRootItem.GetAttribute(AttributeNameAppVersion)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get RootItem() As XMLConfigItem
Set RootItem = mRootItem
End Property

Friend Property Get fileVersion() As String
Const ProcName As String = "fileVersion"
On Error GoTo Err

fileVersion = mRootItem.GetAttribute(AttributeNameFileVersion)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Friend Property Get XMLDoc() As DOMDocument60
Set XMLDoc = mXMLdoc
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Dirty()
Const ProcName As String = "Dirty"
On Error GoTo Err

If Not mDirty Then
    mDirty = True
    fireChange ConfigDirty
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Initialise( _
                ByVal pApplicationName As String, _
                ByVal pApplicationVersion As String)
Const ProcName As String = "Initialise"
On Error GoTo Err

Set mXMLdoc = New DOMDocument60
mXMLdoc.preserveWhiteSpace = True
createDocumentRoot pApplicationName, pApplicationVersion

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub load( _
                ByVal filePath As String)
Const ProcName As String = "load"
On Error GoTo Err

Set mChangeListeners = New Listeners
mFilePath = filePath
Set mXMLdoc = New DOMDocument60
mXMLdoc.preserveWhiteSpace = True
loadDocument mFilePath

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub Save( _
                Optional ByVal filePath As String)
Const ProcName As String = "save"
On Error GoTo Err

If filePath = "" Then
    filePath = mFilePath
End If
mXMLdoc.Save filePath

mDirty = False
fireChange ConfigClean

Exit Sub

Err:
If Err.number = -2147024893 Then
    gCreateFolder Left$(filePath, InStrRev(filePath, "\") - 1)
    Resume
End If
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub createDocumentRoot( _
                ByVal pApplicationName As String, _
                ByVal pApplicationVersion As String)
Const ProcName As String = "createDocumentRoot"
On Error GoTo Err

mXMLdoc.appendChild mXMLdoc.createProcessingInstruction("xml", "version='1.0'")

Dim element As IXMLDOMElement
Set element = mXMLdoc.createElement(ElementNameRoot)

Set mRootItem = New XMLConfigItem
mRootItem.Initialise Me, element, "", "", 1
Set mXMLdoc.documentElement = element
If pApplicationName <> "" Then mRootItem.SetAttribute AttributeNameAppName, pApplicationName
If pApplicationVersion <> "" Then mRootItem.SetAttribute AttributeNameAppVersion, pApplicationVersion
mRootItem.SetAttribute AttributeNameFileVersion, "1.1"

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub fireChange( _
                ByVal ChangeType As ConfigChangeTypes)
Const ProcName As String = "fireChange"
On Error GoTo Err

Dim ev As ChangeEventData
Set ev.Source = Me
ev.ChangeType = ChangeType

mChangeListeners.SetCurrentListeners
Dim i As Long
For i = 1 To mChangeListeners.Count
    Dim lListener As IChangeListener
    Set lListener = mChangeListeners.GetListener(i)
    lListener.Change ev
Next

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub loadDocument( _
                ByVal xmlSource As Variant)
Const ProcName As String = "loadDocument"
On Error GoTo Err

mXMLdoc.load xmlSource

Dim element As IXMLDOMElement
Set element = mXMLdoc.documentElement
Set mRootItem = New XMLConfigItem
mRootItem.InitialiseFromNode Me, element, Nothing, 1

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

