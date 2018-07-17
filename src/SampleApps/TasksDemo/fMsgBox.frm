VERSION 5.00
Object = "{99CC0176-59AF-4A52-B7C0-192026D3FE5D}#32.0#0"; "TWControls40.ocx"
Begin VB.Form fMsgBox 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin TWControls40.TWModelessMsgBox TWModelessMsgBox1 
      Height          =   1500
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   2646
   End
End
Attribute VB_Name = "fMsgBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Event Result( _
                ByVal Value As MsgBoxResults)
                
'@================================================================================
' Enums
'@================================================================================

'@================================================================================
' Types
'@================================================================================

'@================================================================================
' Constants
'@================================================================================

Private Const ProjectName                   As String = "ModelessMsgBoxTester"
Private Const ModuleName                    As String = "fMsgBox"

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
' Control Event Handlers
'@================================================================================

Private Sub TWModelessMsgBox1_Result(ByVal Value As MsgBoxResults)
RaiseEvent Result(Value)
Unload Me
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

Public Sub initialise( _
                ByVal prompt As String, _
                ByVal buttons As MsgBoxStyles, _
                Optional ByVal title As String)
TWModelessMsgBox1.initialise prompt, buttons, title
End Sub

'@================================================================================
' Helper Functions
'@================================================================================






