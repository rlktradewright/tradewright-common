VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TWControls"
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

Enum AlignmentSettings
    TwGridAlignLeftTop = 0
    TwGridAlignLeftCenter = 1
    TwGridAlignLeftBottom = 2
    TwGridAlignCenterTop = 3
    TwGridAlignCenterCenter = 4
    TwGridAlignCenterBottom = 5
    TwGridAlignRightTop = 6
    TwGridAlignRightCenter = 7
    TwGridAlignRightBottom = 8
    TwGridAlignGeneral = 9
End Enum

Enum AllowUserResizeSettings
    TwGridResizeNone = 0
    TwGridResizeColumns = 1
    TwGridResizeRows = 2
    TwGridResizeBoth = 3
End Enum

Enum AllowUserReorderSettings
    TwGridReorderNone = 0
    TwGridReorderColumns = 1
    TwGridReorderRows = 2
    TwGridReorderBoth = 3
End Enum

Enum FillStyleSettings
    TwGridFillSingle = 0
    TwGridFillRepeat = 1
End Enum

Enum FocusRectSettings
    TwGridFocusNone = 0
    TwGridFocusLight = 1
    TwGridFocusHeavy = 2
    TwGridFocusBroken = 3
End Enum

Enum GridLineSettings
    TwGridGridNone = 0
    TwGridGridFlat = 1
    TwGridGridInset = 2
    TwGridGridRaised = 3
End Enum

Enum HighLightSettings
    TwGridHighlightNever = 0
    TwGridHighlightAlways = 1
    TwGridHighlightWithFocus = 2
End Enum

Enum MsgBoxResults
    MsgBoxAbort = vbAbort
    MsgBoxCancel = vbCancel
    MsgBoxIgnore = vbIgnore
    MsgBoxNo = vbNo
    MsgBoxOK = vbOK
    MsgBoxRetry = vbRetry
    MsgBoxYes = vbYes
End Enum

Enum MsgBoxStyles
    MsgBoxOKOnly = vbOKOnly
    MsgBoxOKCancel = vbOKCancel
    MsgBoxAbortRetryIgnore = vbAbortRetryIgnore
    MsgBoxYesNoCancel = vbYesNoCancel
    MsgBoxYesNo = vbYesNo
    MsgBoxRetryCancel = vbRetryCancel
    MsgBoxDefaultButton1 = vbDefaultButton1
    MsgBoxDefaultButton2 = vbDefaultButton2
    MsgBoxDefaultButton3 = vbDefaultButton3
    MsgBoxDefaultButton4 = vbDefaultButton4
    MsgBoxCritical = vbCritical
    MsgBoxQuestion = vbQuestion
    MsgBoxExclamation = vbExclamation
    MsgBoxInformation = vbInformation
    MsgBoxApplicationModal = vbApplicationModal
    MsgBoxSystemModal = vbSystemModal
    MsgBoxMsgBoxHelpButton = vbMsgBoxHelpButton
    MsgBoxMsgBoxRight = vbMsgBoxRight
    MsgBoxMsgBoxRtlReading = vbMsgBoxRtlReading
    MsgBoxMsgBoxSetForeground = vbMsgBoxSetForeground
End Enum

Enum RowSizingSettings
    TwGridRowSizeIndividual = 0
    TwGridRowSizeAll = 1
End Enum

Enum ScrollBarsSettings
    TwGridScrollBarNone = 0
    TwGridScrollBarHorizontal = 1
    TwGridScrollBarVertical = 2
    TwGridScrollBarBoth = 3
End Enum

Enum SelectionModeSettings
    TwGridSelectionFree = 0
    TwGridSelectionByRow = 1
    TwGridSelectionByColumn = 2
End Enum

Enum TextStyleSettings
    TwGridTextFlat = 0
    TwGridTextRaised = 1
    TwGridTextInset = 2
    TwGridTextRaisedLight = 3
    TwGridTextInsetLight = 4
End Enum

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================


Private Const ModuleName                    As String = "TWControls"

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

Public Function ChoosePath(ByVal pInitialPath As String, Optional ByVal pTheme As ITheme) As String
Const ProcName As String = "ChoosePath"
On Error GoTo Err

Dim f As New fPathChooser
f.Path = pInitialPath
If Not pTheme Is Nothing Then f.Theme = pTheme
f.show vbModal
If Not f.cancelled Then ChoosePath = f.Path
Unload f

Exit Function

Err:
gNotifyUnhandledError ProcName, ModuleName
End Function

Public Sub ModelessMsgBox( _
                ByVal pPrompt As String, _
                ByVal pButtons As MsgBoxStyles, _
                Optional ByVal pTitle As String, _
                Optional pOwner As Variant, _
                Optional ByVal pTheme As ITheme)
Const ProcName As String = "ModelessMsgBox"
On Error GoTo Err

Dim lMsgBox As New fMsgBox
lMsgBox.Initialise pPrompt, pButtons, pTitle
If Not pTheme Is Nothing Then lMsgBox.Theme = pTheme
If Not pOwner Is Nothing Then
    lMsgBox.show vbModeless, pOwner
Else
    lMsgBox.show vbModeless
End If

Exit Sub

Err:
gHandleUnexpectedError ProcName, ModuleName
End Sub

'@================================================================================
' Methods
'@================================================================================

'@================================================================================
' Helper Functions
'@================================================================================


