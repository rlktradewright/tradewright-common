Attribute VB_Name = "GPen"
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

Private Const ModuleName                            As String = "GPen"

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

'@================================================================================
' Methods
'@================================================================================

Public Function gCreateLogicalPen( _
                ByVal pColor As Long, _
                ByVal pWidth As Double, _
                ByVal pLineStyle As LineStyles, _
                ByVal pHatchStyle As HatchStyles, _
                ByVal pGraphics As Graphics) As Pen
Const ProcName As String = "gCreateLogicalPen"
On Error GoTo Err

Set gCreateLogicalPen = New Pen
gCreateLogicalPen.Initialise pColor, pWidth, pLineStyle, pHatchStyle, False, pGraphics

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gCreatePixelPen( _
                ByVal pColor As Long, _
                ByVal pWidth As Double, _
                ByVal pLineStyle As LineStyles, _
                ByVal pHatchStyle As HatchStyles) As Pen
Const ProcName As String = "gCreatePixelPen"
On Error GoTo Err

Set gCreatePixelPen = New Pen
gCreatePixelPen.Initialise pColor, pWidth, pLineStyle, pHatchStyle, True, Nothing

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gLoadPenFromConfig( _
                ByVal pConfig As ConfigurationSection) As Pen
Const ProcName As String = "gLoadPenFromConfig"
On Error GoTo Err

Set gLoadPenFromConfig = New Pen

gLoadPenFromConfig.LoadFromConfig pConfig


Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

'@================================================================================
' Helper Functions
'@================================================================================




