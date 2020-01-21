VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   12795
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11940
   LinkTopic       =   "Form2"
   ScaleHeight     =   12795
   ScaleWidth      =   11940
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11655
      ExtentX         =   20558
      ExtentY         =   13150
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
WebBrowser1.Navigate "E:\projects\tradewright-common\src\Tools\RegExpTest\RegExpHelp.htm"
End Sub

Private Sub Form_Resize()
WebBrowser1.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
