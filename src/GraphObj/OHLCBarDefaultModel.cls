VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OHLCBarDefaultModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''
' Description here
'
'@/

'@================================================================================
' Interfaces
'@================================================================================

Implements IOHLCBarModel

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

Private Const ModuleName                            As String = "OHLCBarDefaultModel"
Private mCloseValue                                 As Double
Private mHighValue                                  As Double
Private mLowValue                                   As Double
Private mOpenValue                                  As Double
Private mX                                          As Double

'@================================================================================
' Member variables
'@================================================================================

'@================================================================================
' Class Event Handlers
'@================================================================================

Private Sub Class_Terminate()
gRemoveAllChangeListeners Me
End Sub

'@================================================================================
' IOHLCBarModel Interface Members
'@================================================================================

Public Property Let IOHLCBarModel_CloseValue(ByVal Value As Double)
mCloseValue = Value
gFireChange Me, OHLCBarCloseValueChanged
End Property

Public Property Get IOHLCBarModel_CloseValue() As Double
IOHLCBarModel_CloseValue = mCloseValue
End Property

Public Property Let IOHLCBarModel_LowValue(ByVal Value As Double)
mLowValue = Value
gFireChange Me, OHLCBarLowValueChanged
End Property

Public Property Get IOHLCBarModel_LowValue() As Double
IOHLCBarModel_LowValue = mLowValue
End Property

Public Property Let IOHLCBarModel_HighValue(ByVal Value As Double)
mHighValue = Value
gFireChange Me, OHLCBarHighValueChanged
End Property

Public Property Get IOHLCBarModel_HighValue() As Double
IOHLCBarModel_HighValue = mHighValue
End Property

Public Property Let IOHLCBarModel_OpenValue(ByVal Value As Double)
mOpenValue = Value
gFireChange Me, OHLCBarOpenValueChanged
End Property

Public Property Get IOHLCBarModel_OpenValue() As Double
IOHLCBarModel_OpenValue = mOpenValue
End Property

Public Property Let IOHLCBarModel_X(ByVal Value As Double)
mX = Value
gFireChange Me, OHLCBarXChanged
End Property

Public Property Get IOHLCBarModel_X() As Double
IOHLCBarModel_X = mX
End Property

Private Sub IOHLCBarModel_AddChangeListener(ByVal pListener As IChangeListener)
gAddChangeListener Me, pListener
End Sub

Private Sub IOHLCBarModel_RemoveChangeListener(ByVal pListener As IChangeListener)
gRemoveChangeListener Me, pListener
End Sub

'@================================================================================
' XXXX Event Handlers
'@================================================================================

'@================================================================================
' Properties
'@================================================================================

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================






