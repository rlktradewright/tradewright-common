VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Bitmap"
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

Private Const ModuleName                            As String = "Bitmap"

Private Const ConfigSettingFilename                 As String = "&Filename"

'@================================================================================
' Member variables
'@================================================================================

Private mHandle                             As Long

Private mFilename                           As String
Private mConfig                             As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
If mHandle <> 0 Then DeleteObject mHandle
End Sub

'@================================================================================
' XXXX Interface Members
'@================================================================================

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)

Const ProcName As String = "ConfigurationSection"
Dim failpoint As String
On Error GoTo Err

If Value Is Nothing Then
    RemoveFromConfig
    Exit Property
End If

Set mConfig = Value

mConfig.SetSetting ConfigSettingFilename, mFilename

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

Friend Property Get Handle() As Long
Handle = mHandle
End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pFilename As String)

Const ProcName As String = "Initialise"
On Error GoTo Err

mFilename = pFilename
mHandle = LoadImage(0, StrPtr(mFilename), IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE)
If mHandle = 0 Then gHandleWin32Error GetLastError
Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub

Friend Sub LoadFromConfig( _
                ByVal pConfig As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"
Dim failpoint As String
On Error GoTo Err

Set mConfig = pConfig
Initialise pConfig.GetSetting(ConfigSettingFilename)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"
Dim failpoint As String
On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove
Set mConfig = Nothing

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================




