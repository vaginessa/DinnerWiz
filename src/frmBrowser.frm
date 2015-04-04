VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmBrowser 
   Caption         =   "frmBrowser"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   Icon            =   "frmBrowser.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   7935
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   5880
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      ExtentX         =   13573
      ExtentY         =   8070
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
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    'XP Themes
    If IsThemed() Then
        FixThemeSupport Controls
    End If

    SetIcon Me.hWnd, "AAA", True

    '''''''''''''

    frmBrowser.Caption = "How to Use Dinner Wizard"
    WebBrowser.Navigate (App.Path & devDir & "\Help.html")

End Sub

Private Sub cmdClose_Click()

    Unload frmBrowser

End Sub

Private Sub Form_Resize()

    WebBrowser.Width = frmBrowser.ScaleWidth - (WebBrowser.Left * 2)
    WebBrowser.height = frmBrowser.ScaleHeight - (WebBrowser.Top * 2) - cmdClose.height - WebBrowser.Top
    
    cmdClose.Top = frmBrowser.ScaleHeight - cmdClose.height - WebBrowser.Top
    cmdClose.Left = WebBrowser.Left + (WebBrowser.Width / 2) - (cmdClose.Width / 2)

End Sub

