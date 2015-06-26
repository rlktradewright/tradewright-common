VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IConfigStoreProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This interface is implemented by classes that give access to files of configuration
' data (called <em>configuration datasets</em>), typically used by applications to record
' State information for use the Next time the application is Run.
'
' A configuration dataset conceptually contains a set of <em>configuration items</em>,
' organised in a tree structure. There is a single configuration Item at the root
' of the tree, called the <em>root configuration item</em>.
'
' Each configuration item has a name, an optional text Value, an optional set of
' attributes, and an optional comment. Each attribute consists of a name and a
' text Value.
'
' An example of a configuration Item might be:
'   <pre>
'   ServiceProvider=Realtime Data Source{Enabled="True" ProgId="IBTWSSP26.RealtimeDataServiceProvider" LogLevel="1"} // configure the tick Data SP
'   </pre>'
' Here, the configuration item name is &quot;ServiceProvider&quot;, its Value
' is &quot;Realtime Data Source&quot;, and it has three attributes: Enabled,
' ProgId and LogLevel. Note that the format of this example is purely illustrative: the
' actual form in which a configuration Item is stored, and the representation of the
' hierarchical relationship between configuration items, depends on the nature of
' the storage file type.
'
' Currently the only class that implements this interface is
' <code>XMLConfigProvider</code>, which is used to create and access configuration
' Data in an XML format. To create an empty instance of <code>XMLConfigProvider</code>,
' use the <code>CreateXMLConfigurationProvider</code> global method. To create an instance
' of <code>XMLConfigProvider</code> containing configuration Data stored in an
' existing file, use the <code>LoadConfigProviderFromXMLFile</code> global method.
'
' Note that the facilities provided by the <code>ConfigStoreProvider</code> class are quite
' low-level, and not particularly easy to use. It is recommended that applications
' use the <code>ConfigurationFile</code> class instead, which presents a more
' intuitive interface. Use the <code>CreateConfigurationStore</code> global method to
' obtain a <code>ConfigurationStore</code> object from a <code>ConfigStoreProvider</code>
' object.
' @see ConfigItem
' @see ConfigItems
' @see TWUtilities.CreateXMLConfigurationProvider
' @see TWUtilities.LoadConfigProviderFromXMLFile
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


Private Const ModuleName                    As String = "IConfigStoreProvider"

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

''
' Stores the specified application Name in the configuration dataset.
'
' @remarks
' @param Value
'   The application Name to be stored.
'
'@/
Public Property Let ApplicationName( _
                ByVal Value As String)
End Property

''
' Returns the application Name stored in the configuration dataset, if any.
'
' @remarks
' @return
'   The application Name stored in the configuration dataset, if any.
' @see
'
'@/
Public Property Get ApplicationName() As String
End Property

''
' Stores the specified application version in the configuration dataset.
'
' @remarks
' @param Value
'   The application version to be stored.
' @see
'
'@/
Public Property Let ApplicationVersion( _
                ByVal Value As String)
End Property

''
' Returns the application version stored in the configuration dataset, if any.
'
' @remarks
' @return
'   The application version stored in the configuration dataset, if any.
' @see
'
'@/
Public Property Get ApplicationVersion() As String
End Property

''
' Returns <code>True</code> if the configuration Data contains unsaved changes.
'
' @remarks
' @return
'   <code>True</code> if there are unsaved changes, otherwise <code>False</code>.
' @see
'
'@/
Public Property Get Dirty() As Boolean
End Property

''
' Returns the root<code>ConfigItem</code> object.
'
' @remarks
' @return
'   Returns a <code>ConfigItem</code> object containing the root configuration Item.
' @see
'
'@/
Public Property Get RootItem() As IConfigItem
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
                
End Sub

''
' Unregisters a change listener object.
'
' @param listener
'   The object to be unregistered.
'@/
Public Sub RemoveChangeListener( _
                ByVal pListener As IChangeListener)
                
End Sub

''
' Saves the configuration Data to the specified file.
'
' @param filePath
'   The Path and filename in which the configuration Data is to be stored.
' @see
'
'@/
Public Sub Save( _
                Optional ByVal filePath As String)
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

