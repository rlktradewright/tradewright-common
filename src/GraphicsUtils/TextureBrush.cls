VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextureBrush"
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

Implements IBrush
Implements INativeBrush

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

Private Const ModuleName                            As String = "TextureBrush"

Private Const ConfigSectionBitmap                   As String = "Bitmap"

'@================================================================================
' Member variables
'@================================================================================

Private mBitmap                             As Bitmap

Private mHandle                             As Long

Private mConfig                             As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
deleteBrushObject
End Sub

'@================================================================================
' IBrush Interface Members
'@================================================================================

Private Property Let IBrush_ConfigurationSection(ByVal RHS As ConfigurationSection)
Const ProcName As String = "IBrush_ConfigurationSection"
On Error GoTo Err

ConfigurationSection = RHS

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub IBrush_RemoveFromConfig()
Const ProcName As String = "IBrush_RemoveFromConfig"
On Error GoTo Err

RemoveFromConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' INativeBrush Interface Members
'@================================================================================

Private Property Get INativeBrush_Handle() As Long
Const ProcName As String = "INativeBrush_Handle"
On Error GoTo Err

If mHandle = 0 Then createTheBrush
INativeBrush_Handle = mHandle

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

Public Property Get Bitmap() As Bitmap
Set Bitmap = mBitmap
End Property

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

gSetBrushClassInConfig Me, mConfig
mBitmap.ConfigurationSection = mConfig.AddConfigurationSection(ConfigSectionBitmap)

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName

End Property

'@================================================================================
' Methods
'@================================================================================

Friend Sub Initialise(ByVal pBitmap As Bitmap)
Const ProcName As String = "Initialise"
On Error GoTo Err

AssertArgument Not pBitmap Is Nothing, "pBitmap must be a Bitmp object"

Set mBitmap = pBitmap

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
Set mBitmap = gLoadBitmapFromConfig(mConfig.GetConfigurationSection(ConfigSectionBitmap))

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

Private Sub createTheBrush()
Const ProcName As String = "createTheBrush"
Dim failpoint As String
On Error GoTo Err

mHandle = CreatePatternBrush(mBitmap.Handle)
If mHandle = 0 Then gHandleWin32Error GetLastError

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub deleteBrushObject()
Const ProcName As String = "deleteBrushObject"
On Error GoTo Err

If mHandle <> 0 Then
    If DeleteObject(mHandle) = 0 Then gHandleWin32Error GetLastError
    mHandle = 0
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName

End Sub


