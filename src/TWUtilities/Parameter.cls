VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Parameter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

''
' This class represents a text key/Value pair.
'
' Instances of this class cannot be created directly using the <code>New</code> operator.
' Instead you must use the <code>getParameterValue</code>, <code>setParameterValue</code>
' or <code>setParameterValues</code>methods of the <code>Parameters</code> class.
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements Comparable
Implements Configurable

'@================================================================================
' Events
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                As String = "Parameter"

Private Const ConfigSettingName                         As String = "&Name"
Private Const ConfigSettingValue                        As String = "&Value"

'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Member variables
'@================================================================================

Private mName As String
Private mValue As String

Private mConfig                                         As ConfigurationSection

'@================================================================================
' Class Event Handlers
'@================================================================================

'@================================================================================
' Comparable Interface Members
'@================================================================================

Private Function Comparable_compareTo(ByVal obj As Comparable) As Long
Dim param As Parameter

Const ProcName As String = "Comparable_compareTo"

On Error GoTo Err

If obj Is Nothing Then Err.Raise ErrorCodes.ErrNullPointerException, , "obj argument must not be Nothing"

Set param = obj
Comparable_compareTo = StrComp(mName, param.Name, vbTextCompare)

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Configurable Interface Members
'@================================================================================

Private Property Let Configurable_ConfigurationSection(ByVal RHS As ConfigurationSection)
Const ProcName As String = "Configurable_ConfigurationSection"
On Error GoTo Err

ConfigurationSection = RHS

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Private Sub Configurable_LoadFromConfig(ByVal pConfig As ConfigurationSection)
Const ProcName As String = "Configurable_LoadFromConfig"
On Error GoTo Err

LoadFromConfig pConfig

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Private Sub Configurable_RemoveFromConfig()
Const ProcName As String = "Configurable_RemoveFromConfig"
On Error GoTo Err

RemoveFromConfig

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

Friend Property Let ConfigurationSection( _
                ByVal Value As ConfigurationSection)
Const ProcName As String = "ConfigurationSection"

On Error GoTo Err

Set mConfig = Value

setSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Sets the Name of this <code>Parameter</code> object.
'
' @param val
'   The new Name of this object.
'@/
Friend Property Let Name(ByVal val As String)
Const ProcName As String = "Name"

On Error GoTo Err

mName = val
setSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the Name of this <code>Parameter</code> object.
'
' @return
'   The current Name of this <code>Parameter</code> object.
'@/
Public Property Get Name() As String
Name = mName
End Property

''
' Sets the Value of this <code>Parameter</code> object.
'
' @param val
'   The new Value of this object.
'@/
Friend Property Let Value(ByVal val As String)
Const ProcName As String = "Value"

On Error GoTo Err

mValue = val
setSettings

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

''
' Returns the Value of this <code>Parameter</code> object.
'
' @return
'   The current Value of this <code>Parameter</code> object.
'@/
Public Property Get Value() As String
Value = mValue
End Property

'@================================================================================
' Methods
'@================================================================================

''
' Indicates whether this object and the specified <code>Parameter</code> object
' are equal.
'
' @remarks
'   To be considered equal, the two objects must have the same Name and the same
'   Value, using a case-insensitive comparison.
' @return
'   <code>True</code> if the objects are equal, otherwise <code>False</code>.
' @param param
'   The <code>Parameter</code> object to which this object is to be compared.
'@/
Public Function Equals(param As Parameter) As Boolean
Const ProcName As String = "Equals"

On Error GoTo Err

If param Is Nothing Then Exit Function

If param Is Me Then
    Equals = True
    Exit Function
End If

If UCase$(mName) <> UCase$(param.Name) Then Exit Function

If UCase$(mValue) <> UCase$(param.Value) Then Exit Function

Equals = True

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName

End Function

Friend Sub LoadFromConfig( _
                ByVal Config As ConfigurationSection)
Const ProcName As String = "LoadFromConfig"

On Error GoTo Err

Set mConfig = Config
mName = mConfig.GetSetting(ConfigSettingName)
mValue = mConfig.GetSetting(ConfigSettingValue)

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Friend Sub RemoveFromConfig()
Const ProcName As String = "RemoveFromConfig"

On Error GoTo Err

If Not mConfig Is Nothing Then mConfig.Remove

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Helper Functions
'@================================================================================

Private Sub setSettings()
Const ProcName As String = "setSettings"

On Error GoTo Err

If Not mConfig Is Nothing Then
    mConfig.SetSetting ConfigSettingName, mName
    mConfig.SetSetting ConfigSettingValue, mValue
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub
