VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IConfigurable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This interface enables a class to declare that its objects can be configured using
' information from a <code>ConfigurationStore</code>.
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

Private Const ModuleName                            As String = "IConfigurable"

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
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

''
' Specifies the <code>ConfigurationSection</code> where configuration information
' about this object is to be stored.
'
' @param pConfig
'   the <code>ConfigurationSection</code> where configuration information
'   is to be stored.
'
'@/
Public Property Let ConfigurationSection(ByVal pConfig As ConfigurationSection)

End Property

'@================================================================================
' Methods
'@================================================================================

''
' Specifies the <code>ConfigurationSection</code> that contains configuration information
' to initialise this object.
'
' @param pConfig
'   the <code>ConfigurationSection</code> that contains configuration information
'   to initialise this object.
'
'@/
Public Sub LoadFromConfig(ByVal pConfig As ConfigurationSection)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================

''
' Instructs the object to remove its configuration information.
'@/
Public Sub RemoveFromConfig()

End Sub




