VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LayeredGraphics"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
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

Public Enum LayerNumbers
    LayerMin = LayerNumberRange.MinLayer
    LayerInvisible = LayerMin
    LayerMinVisible = LayerInvisible + 1
    LayerBackground = LayerMinVisible
    LayerGrid = LayerMinVisible + 5
    LayerGridText = LayerNumbers.LayerGrid + 3
    LayerLowestUser = LayerGridText + 10
    LayerHighestUser = Globals.LayerNumberRange.MaxLayer - 10
    LayerTitle = LayerNumberRange.MaxLayer - 5
    LayerPointer = LayerNumberRange.MaxLayer - 2
    LayerMax = LayerNumberRange.MaxLayer
End Enum

'@================================================================================
' Types
'@================================================================================

Public Type GraphObjListCookie
    From                As Long
    To                  As Long
    Current             As Long
End Type

Public Type LayeredListCookie
    StartLayer          As LayerNumbers
    StopLayer           As LayerNumbers
    CurrLayer           As LayerNumbers
    En                  As Enumerator
    UpLayers            As Boolean
End Type

Public Type ListMemberships
    MinStrip            As Long
    MaxStrip            As Long
    StripListMemberships    As LinkedList
End Type

'@================================================================================
' Constants
'@================================================================================

Private Const ModuleName                            As String = "LayeredGraphics"

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

Public Function CreateLayeredGraphicsEngine( _
                ByVal pGraphics As Graphics, _
                Optional ByVal pStripWidth As Double = 1#, _
                Optional ByVal pStripOrigin As Double = 0#, _
                Optional ByVal pUseDeferredPainting As Boolean = True, _
                Optional ByVal pDeferredPaintDelay As Long = 20) As Controller
Const ProcName As String = "CreateLayeredGraphicsEngine"
On Error GoTo Err

AssertArgument Not pGraphics Is Nothing, "pGraphics cannot be Nothing"
AssertArgument IsInteger(pDeferredPaintDelay, 5, 1000), "pDeferredPaintDelay must be between 5 and 1000"

Set CreateLayeredGraphicsEngine = New Controller
CreateLayeredGraphicsEngine.Initialise pGraphics, pUseDeferredPainting, pDeferredPaintDelay, pStripWidth, pStripOrigin

Exit Function

Err:
gHandleUnexpectedError ProcName, ModuleName
End Function

Public Function IsValidLayerNumber(ByVal pValue As Long) As Boolean
IsValidLayerNumber = gIsValidLayerNumber(pValue)
End Function

'@================================================================================
' Helper Functions
'@================================================================================




