Attribute VB_Name = "Globals"
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

Public Enum RegionSelectionModes
    RegionSelectionModeAnd = RGN_AND
    RegionSelectionModeCopy = RGN_COPY
    RegionSelectionModeDiff = RGN_DIFF
    RegionSelectionModeOr = RGN_OR
    RegionSelectionModeXOr = RGN_XOR
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Public Const ProjectName                            As String = "GraphicsUtils40"

Private Const ModuleName                            As String = "Globals"

Public Const Pi As Double = 3.14159265358979

Public Const MinDouble As Double = -(2 - 2 ^ -52) * 2 ^ 1023
Public Const MaxDouble As Double = (2 - 2 ^ -52) * 2 ^ 1023

Public Const MinSingle As Single = -(2 - 2 ^ -23) * 2 ^ 127
Public Const MaxSingle As Single = (2 - 2 ^ -23) * 2 ^ 127

Public Const MaxLong As Long = &H7FFFFFFF
Public Const MinLong As Long = &H80000000

Public Const MaxSystemColor As Long = &H80000018

Public Const LogicalUnitsPerPixel      As Long = 1000
Public Const TwipsPerPoint             As Long = 24

Public Const TwipsPerCm As Double = 1440 / 2.54

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

Public Property Get gLogger() As FormattingLogger
Static lLogger As FormattingLogger
Const ProcName As String = "gLogger"

On Error GoTo Err

If lLogger Is Nothing Then
    Set lLogger = CreateFormattingLogger("graphicsutils.log", ProjectName)
End If
Set gLogger = lLogger

Exit Property

Err:
gHandleUnexpectedError ProcName, ModuleName
End Property

Public Property Get ghNullBrush() As Long
Static shNullBrush As Long
If shNullBrush = 0 Then shNullBrush = GetStockObject(NULL_BRUSH)
ghNullBrush = shNullBrush
End Property

Public Property Get ghNullPen() As Long
Static shNullPen As Long
If shNullPen = 0 Then shNullPen = GetStockObject(NULL_PEN)
ghNullPen = shNullPen
End Property

Public Property Get ghPathPen() As Long
Static shPathPen As Long
If shPathPen = 0 Then shPathPen = GetStockObject(BLACK_PEN)
ghPathPen = shPathPen
End Property

Public Property Get gPathBrush() As IBrush
Static sPathBrush As IBrush
If sPathBrush Is Nothing Then Set sPathBrush = gCreateBrush(vbBlack)
Set gPathBrush = sPathBrush
End Property

Public Property Get gPathPen() As Pen
Static sPathPen As Pen
If sPathPen Is Nothing Then Set sPathPen = gCreatePixelPen(vbBlack, 1, LineSolid, HatchNone)
Set gPathPen = sPathPen
End Property

'@================================================================================
' Methods
'@================================================================================

Public Function gDegreesToRadians( _
                ByVal degrees As Double) As Double
gDegreesToRadians = degrees * Pi / 180
End Function

Public Function gDoubleArrayFromString( _
                ByRef pInput As String) As Double()
Dim ar() As Variant
Dim result() As Double
Dim i As Long
Const ProcName As String = "gDoubleArrayFromString"
On Error GoTo Err

ar = Split(pInput, ",")
ReDim result(UBound(ar)) As Double
For i = 0 To UBound(ar)
    result(i) = CDbl(ar(i))
Next

gDoubleArrayFromString = result

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gDoubleArrayToString( _
                ByRef pArray() As Double) As String
On Error Resume Next

Dim s As String
Dim i As Double

For i = 0 To UBound(pArray)
    If i <> 0 Then s = s & ","
    s = s & CStr(pArray(i))
Next

gDoubleArrayToString = s

End Function

Public Sub gHandleWin32Error(ByVal pErrorCode As Long)
Const ProcName As String = "gHandleWin32Error"
On Error GoTo Err

Err.Raise ErrorCodes.ErrRuntimeException, , "Windows error " & pErrorCode

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

Public Sub gHandleUnexpectedError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pReRaise As Boolean = True, _
                Optional ByVal pLog As Boolean = False, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

HandleUnexpectedError pProcedureName, ProjectName, pModuleName, pFailpoint, pReRaise, pLog, errNum, errDesc, errSource
End Sub

Public Sub gNotifyUnhandledError( _
                ByRef pProcedureName As String, _
                ByRef pModuleName As String, _
                Optional ByRef pFailpoint As String, _
                Optional ByVal pErrorNumber As Long, _
                Optional ByRef pErrorDesc As String, _
                Optional ByRef pErrorSource As String)
Dim errSource As String: errSource = IIf(pErrorSource <> "", pErrorSource, Err.Source)
Dim errDesc As String: errDesc = IIf(pErrorDesc <> "", pErrorDesc, Err.Description)
Dim errNum As Long: errNum = IIf(pErrorNumber <> 0, pErrorNumber, Err.Number)

UnhandledErrorHandler.Notify pProcedureName, pModuleName, ProjectName, pFailpoint, errNum, errDesc, errSource
End Sub

Public Function gIsValidColor( _
                ByVal Value As Long) As Boolean
                
If Value > &HFFFFFF Then Exit Function
If Value < 0 And Value > MaxSystemColor Then Exit Function
gIsValidColor = True
End Function

Public Function gLongArrayFromString( _
                ByRef pInput As String) As Long()
Const ProcName As String = "gLongArrayFromString"
On Error GoTo Err

Dim ar() As Variant
Dim result() As Long
Dim i As Long

ar = Split(pInput, ",")
ReDim result(UBound(ar)) As Long
For i = 0 To UBound(ar)
    result(i) = CLng(ar(i))
Next

gLongArrayFromString = result

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gLongArrayToString( _
                ByRef pArray() As Long) As String
On Error Resume Next

Dim s As String
Dim i As Long

For i = 0 To UBound(pArray)
    If i <> 0 Then s = s & ","
    s = s & CStr(pArray(i))
Next

gLongArrayToString = s
End Function

Public Function gMaxDouble( _
                ByVal pVal1 As Double, _
                ByVal pVal2 As Double) As Double
If pVal1 >= pVal2 Then
    gMaxDouble = pVal1
Else
    gMaxDouble = pVal2
End If
End Function
               
Public Function gMaxLong( _
                ByVal pVal1 As Long, _
                ByVal pVal2 As Long) As Long
If pVal1 >= pVal2 Then
    gMaxLong = pVal1
Else
    gMaxLong = pVal2
End If
End Function
               
Public Function gMinDouble( _
                ByVal pVal1 As Double, _
                ByVal pVal2 As Double) As Double
If pVal1 <= pVal2 Then
    gMinDouble = pVal1
Else
    gMinDouble = pVal2
End If
End Function
               
Public Function gMinLong( _
                ByVal pVal1 As Long, _
                ByVal pVal2 As Long) As Long
If pVal1 <= pVal2 Then
    gMinLong = pVal1
Else
    gMinLong = pVal2
End If
End Function
               
Public Function gNormalizeColor(ByVal pColor As Long) As Long
If pColor < 0 Then
    gNormalizeColor = GetSysColor(pColor And &HFF&)
Else
    gNormalizeColor = pColor
End If
End Function

Public Sub gPrintGdiDiagnostics(ByVal phDC As Long)
Dim lSize As GDI_SIZE
Dim lPoint As GDI_POINT
Debug.Print "Graphics mode: " & GetGraphicsMode(phDC)
Debug.Print "Mapping mode : " & GetMapMode(phDC)
GetWindowExtEx phDC, lSize
Debug.Print "WindowExt    : " & lSize.cx & ", " & lSize.cy
GetWindowOrgEx phDC, lPoint
Debug.Print "WindowOrg    : " & lPoint.X & "," & lPoint.Y
GetViewportExtEx phDC, lSize
Debug.Print "ViewportExt  : " & lSize.cx & ", " & lSize.cy
End Sub

Public Function gRadiansToDegrees( _
                ByVal radians As Double) As Double
gRadiansToDegrees = radians * 180 / Pi
End Function

Public Sub gSetVariant(ByRef pTarget As Variant, ByRef pSource As Variant)
If IsObject(pSource) Then
    Set pTarget = pSource
Else
    pTarget = pSource
End If
End Sub

Public Function gSingleArrayFromString( _
                ByRef pInput As String) As Single()
Dim ar() As Variant
Dim result() As Single
Dim i As Long
Const ProcName As String = "gSingleArrayFromString"
On Error GoTo Err

ar = Split(pInput, ",")
ReDim result(UBound(ar)) As Single
For i = 0 To UBound(ar)
    result(i) = CDbl(ar(i))
Next

gSingleArrayFromString = result

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function gSingleArrayToString( _
                ByRef pArray() As Single) As String
On Error Resume Next

Dim s As String
Dim i As Single

For i = 0 To UBound(pArray)
    If i <> 0 Then s = s & ","
    s = s & CStr(pArray(i))
Next

gSingleArrayToString = s

End Function

'@================================================================================
' Helper Functions
'@================================================================================


