VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IGraphicObjectSeries"
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


Private Const ModuleName                    As String = "IGraphicObjectSeries"

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

Public Property Get Controller() As Controller

End Property

Public Property Get Count() As Long

End Property

Public Property Get ExtendedPropertyHost() As ExtendedPropertyHost

End Property

Public Property Get GraphicsUpdateNumber() As Long

End Property

Public Property Get Id() As Long

End Property

Public Property Get IsSelectable() As Boolean

End Property

Public Property Let IsSelected(ByVal Value As Boolean)

End Property

Public Property Get IsSelected() As Boolean

End Property

Public Property Get LastGraphicsChange() As GraphicsChangeTypes

End Property

Public Property Get LayeredGraphicsModel() As LayeredGraphicsModel

End Property

Public Property Get Name() As String

End Property

Public Property Get Renderer() As IRenderer

End Property

Public Property Let Renderer(ByVal Value As IRenderer)

End Property

Public Property Get SeriesID() As Long

End Property

Public Property Get SeriesUpdateNumber() As Long

End Property

'@================================================================================
' Methods
'@================================================================================

Public Sub Click()

End Sub

Public Sub DblCLick()

End Sub

Public Sub Finish()

End Sub

Public Sub Initialise( _
                ByVal pName As String, _
                ByVal pLayer As LayerNumbers, _
                ByVal pSeriesID As Long, _
                ByVal pController As Controller)

End Sub

Public Function Item(ByVal pIndex As Long) As IGraphicObject

End Function

Public Sub Remove(ByVal pIndex As Long)

End Sub

Public Sub RemoveObject(ByVal Value As IGraphicObject)

End Sub

'@================================================================================
' Helper Functions
'@================================================================================




