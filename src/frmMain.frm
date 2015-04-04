VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dinner Wizard 2.11 by RL Vision"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":CCA7
   ScaleHeight     =   509
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   421
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   1320
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picPrinter 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6735
      Left            =   6600
      ScaleHeight     =   6675
      ScaleWidth      =   5115
      TabIndex        =   23
      Top             =   360
      Width           =   5175
   End
   Begin VB.PictureBox picButtons 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   975
      Index           =   5
      Left            =   1740
      Picture         =   "frmMain.frx":1A00B
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   184
      TabIndex        =   22
      Top             =   870
      Width           =   2760
   End
   Begin VB.PictureBox picButtons 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   270
      Index           =   4
      Left            =   4365
      Picture         =   "frmMain.frx":1BD2D
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   71
      TabIndex        =   21
      Top             =   6720
      Width           =   1065
   End
   Begin VB.PictureBox picOverlay 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   5760
      ScaleHeight     =   2355
      ScaleWidth      =   6315
      TabIndex        =   19
      Top             =   120
      Width           =   6315
   End
   Begin VB.Timer tmrOverlay 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picButtons 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   510
      Index           =   3
      Left            =   4635
      Picture         =   "frmMain.frx":1C689
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   112
      TabIndex        =   18
      Top             =   7125
      Width           =   1680
   End
   Begin VB.PictureBox picButtons 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   510
      Index           =   2
      Left            =   3060
      Picture         =   "frmMain.frx":1D10A
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   17
      Top             =   7125
      Width           =   1755
   End
   Begin VB.PictureBox picButtons 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   510
      Index           =   1
      Left            =   1515
      Picture         =   "frmMain.frx":1DB00
      ScaleHeight     =   34
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   117
      TabIndex        =   16
      Top             =   7125
      Width           =   1755
   End
   Begin VB.PictureBox picButtons 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   0
      Picture         =   "frmMain.frx":1E547
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   114
      TabIndex        =   15
      Top             =   7140
      Width           =   1710
   End
   Begin VB.Timer TimerMouse 
      Interval        =   50
      Left            =   720
      Top             =   120
   End
   Begin VB.Label lblAutosize 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "lblAutosize"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   5760
      TabIndex        =   20
      Top             =   2760
      Width           =   1035
      Visible         =   0   'False
   End
   Begin VB.Label lblMondayLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mon:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00063669&
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   2145
      Width           =   1095
   End
   Begin VB.Label lblTuesdayLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Tue:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00063669&
      Height          =   255
      Left            =   1680
      TabIndex        =   11
      Top             =   2745
      Width           =   1095
   End
   Begin VB.Label lblSundayLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sun:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00063669&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   5745
      Width           =   1095
   End
   Begin VB.Label lblSaturdayLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Sat:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00063669&
      Height          =   255
      Left            =   1680
      TabIndex        =   9
      Top             =   5145
      Width           =   1095
   End
   Begin VB.Label lblFridayLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fri:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00063669&
      Height          =   255
      Left            =   1680
      TabIndex        =   8
      Top             =   4545
      Width           =   1095
   End
   Begin VB.Label lblThursdayLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Thu:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00063669&
      Height          =   255
      Left            =   1680
      TabIndex        =   7
      Top             =   3945
      Width           =   1095
   End
   Begin VB.Label lblWednesdayLabel 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Wed:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00063669&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   3345
      Width           =   1095
   End
   Begin VB.Label lblTuesday 
      BackStyle       =   0  'Transparent
      Caption         =   "lblTuesday"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   5
      Top             =   2745
      Width           =   2415
   End
   Begin VB.Label lblWednesday 
      BackStyle       =   0  'Transparent
      Caption         =   "lblWednesday"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   4
      Top             =   3345
      Width           =   2415
   End
   Begin VB.Label lblThursday 
      BackStyle       =   0  'Transparent
      Caption         =   "lblThursday"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   3
      Top             =   3945
      Width           =   2415
   End
   Begin VB.Label lblFriday 
      BackStyle       =   0  'Transparent
      Caption         =   "lblFriday"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   4545
      Width           =   2415
   End
   Begin VB.Label lblSaturday 
      BackStyle       =   0  'Transparent
      Caption         =   "lblSaturday"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   1
      Top             =   5145
      Width           =   2415
   End
   Begin VB.Label lblSunday 
      BackStyle       =   0  'Transparent
      Caption         =   "lblSunday"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   0
      Top             =   5745
      Width           =   2415
   End
   Begin VB.Label lblMonday 
      BackStyle       =   0  'Transparent
      Caption         =   "lblMonday"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2880
      TabIndex        =   14
      Top             =   2145
      Width           =   2415
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Click on 'menu' above to create a weekly menu. Each click makes a new menu!"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   3015
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private chkMonday_LastValue
Private chkTuesday_LastValue
Private chkWednesday_LastValue
Private chkThursday_LastValue
Private chkFriday_LastValue
Private chkSaturday_LastValue
Private chkSunday_LastValue

Private currentArrayIndex As Integer
Private editNameArrayIndex As Integer
Private editNameListIndex As Integer
Private allowUpdate As Boolean
Private useGrayChckboxes As Boolean
Private dataChanged As Boolean

Dim myTT() As CTooltip

Private Sub Form_Load()

    allowUpdate = False
    useGrayChckboxes = False
    dataChanged = False

    picOverlay.Left = 0
    picOverlay.Top = 0
    picOverlay.Visible = False
    picButtons(0).Visible = False
    picButtons(1).Visible = False
    picButtons(2).Visible = False
    picButtons(3).Visible = False

    iBrowserDoc = -1
    
    
    picOverlay.Picture = frmMain.Picture
    
    'change font on win98
    If Not IsWin2000Plus Then
        'Book Antiqua fonten fanns i 98, men inte xp, den fonten som är inställd i gui är en ersättare som ser likadan ut...
        sFont = "Book Antiqua"
        lblInfo.Font = sFont
        lblMonday.Font = sFont
        lblMondayLabel.Font = sFont
        lblTuesday.Font = sFont
        lblTuesdayLabel.Font = sFont
        lblWednesday.Font = sFont
        lblWednesdayLabel.Font = sFont
        lblThursday.Font = sFont
        lblThursdayLabel.Font = sFont
        lblFriday.Font = sFont
        lblFridayLabel.Font = sFont
        lblSaturday.Font = sFont
        lblSaturdayLabel.Font = sFont
        lblSunday.Font = sFont
        lblSundayLabel.Font = sFont
    End If
    
    'increase spacing a bit
    iPlus = 1
    lblTuesday.Top = lblTuesday.Top + iPlus
    lblTuesdayLabel.Top = lblTuesdayLabel.Top + iPlus
    lblWednesday.Top = lblWednesday.Top + iPlus * 2
    lblWednesdayLabel.Top = lblWednesdayLabel.Top + iPlus * 2
    lblThursday.Top = lblThursday.Top + iPlus * 3
    lblThursdayLabel.Top = lblThursdayLabel.Top + iPlus * 3
    lblFriday.Top = lblFriday.Top + iPlus * 4
    lblFridayLabel.Top = lblFridayLabel.Top + iPlus * 4
    lblSaturday.Top = lblSaturday.Top + iPlus * 5
    lblSaturdayLabel.Top = lblSaturdayLabel.Top + iPlus * 5
    lblSunday.Top = lblSunday.Top + iPlus * 6
    lblSundayLabel.Top = lblSundayLabel.Top + iPlus * 6


    'move text a bit
    iPlus = 3
    lblMonday.Top = lblMonday.Top + iPlus
    lblTuesday.Top = lblTuesday.Top + iPlus
    lblWednesday.Top = lblWednesday.Top + iPlus
    lblThursday.Top = lblThursday.Top + iPlus
    lblFriday.Top = lblFriday.Top + iPlus
    lblSaturday.Top = lblSaturday.Top + iPlus
    lblSunday.Top = lblSunday.Top + iPlus

    'move some controlles
    lblInfo.Left = 128
    lblInfo.Top = 168
    
    'XP Themes
    If IsThemed() Then
        FixThemeSupport Controls
    End If

    SetIcon Me.hWnd, "AAA", True

    '''''''''''''
    
    'load food data
    
    ReDim myFood(0) ' entry 0 is not used
    
    ver = 0
    bFirstTimeUsing = False
    bFirstTimeClicking = False
    bFirstTimeHolding = False
    
    Call getSystemFolders(sysFolders)
    
    LoadSettings
    
    'recall menu
    If sCurrentMenu <> "" Then
        tmp = Split(sCurrentMenu, ",")
        If UBound(myFood) >= tmp(0) Then lblMonday = myFood(tmp(0)).Name Else lblMonday = ""
        If UBound(myFood) >= tmp(1) Then lblTuesday = myFood(tmp(1)).Name Else lblTuesday = ""
        If UBound(myFood) >= tmp(2) Then lblWednesday = myFood(tmp(2)).Name Else lblWednesday = ""
        If UBound(myFood) >= tmp(3) Then lblThursday = myFood(tmp(3)).Name Else lblThursday = ""
        If UBound(myFood) >= tmp(4) Then lblFriday = myFood(tmp(4)).Name Else lblFriday = ""
        If UBound(myFood) >= tmp(5) Then lblSaturday = myFood(tmp(5)).Name Else lblSaturday = ""
        If UBound(myFood) >= tmp(6) Then lblSunday = myFood(tmp(6)).Name Else lblSunday = ""
        
        Call FixLabel(lblMonday)
        Call FixLabel(lblTuesday)
        Call FixLabel(lblWednesday)
        Call FixLabel(lblThursday)
        Call FixLabel(lblFriday)
        Call FixLabel(lblSaturday)
        Call FixLabel(lblSunday)
        
        
        lblInfo.Visible = False
    Else
        lblInfo.Visible = True
        lblMonday.Visible = False
        lblMondayLabel.Visible = False
        lblTuesday.Visible = False
        lblTuesdayLabel.Visible = False
        lblWednesday.Visible = False
        lblWednesdayLabel.Visible = False
        lblThursday.Visible = False
        lblThursdayLabel.Visible = False
        lblFriday.Visible = False
        lblFridayLabel.Visible = False
        lblSaturday.Visible = False
        lblSaturdayLabel.Visible = False
        lblSunday.Visible = False
        lblSundayLabel.Visible = False
    End If
    
    
    Randomize

    allowUpdate = True


    Call SetTooltips

End Sub

Private Sub Form_Activate()
    
    If bFirstTimeUsing = True Then
        Call MsgBox("Welcome to Dinner Wizard! Please take some time to read the help text to make sure that you understand how it works!", vbInformation, "Welcome!")
        iBrowserDoc = 1
        frmBrowser.Show vbModal
        bFirstTimeUsing = False
        iBrowserDoc = -1
    End If

End Sub


Private Sub imgClick_Click()
    If bFirstTimeClicking = True Then
        Call MsgBox("Since this is the first time you are creating a menu, I have entered a few common dishes for you. These are just to get you started, and you can safely remove them as you please! To add and remove meals, please click on 'Setup' at the bottom.", vbInformation)
        bFirstTimeClicking = False
        SaveSettings
    End If
    tmrOverlay.Enabled = True
End Sub


Private Sub TabStrip_Click()

    frmGenerate.Visible = False
    frmFood.Visible = False
    wbHelp.Visible = False
    wbAbout.Visible = False

    If TabStrip.SelectedItem.Index = 1 Then
        frmGenerate.Visible = True
    ElseIf TabStrip.SelectedItem.Index = 2 Then
        frmFood.Visible = True
    ElseIf TabStrip.SelectedItem.Index = 3 Then
        wbHelp.Visible = True
    ElseIf TabStrip.SelectedItem.Index = 4 Then
        wbAbout.Visible = True
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)

    SaveSettings

    Call UnloadXpApp

End Sub

Private Sub GenerateMenu()

    lblInfo.Visible = False
    lblMonday.Visible = True
    lblMondayLabel.Visible = True
    lblTuesday.Visible = True
    lblTuesdayLabel.Visible = True
    lblWednesday.Visible = True
    lblWednesdayLabel.Visible = True
    lblThursday.Visible = True
    lblThursdayLabel.Visible = True
    lblFriday.Visible = True
    lblFridayLabel.Visible = True
    lblSaturday.Visible = True
    lblSaturdayLabel.Visible = True
    lblSunday.Visible = True
    lblSundayLabel.Visible = True

    oldMenu = Split(sCurrentMenu, ",")
    For n = 1 To UBound(myFood())
        myFood(n).alreadyUsedInCurrentMenu = False

        If UBound(myFood) >= 14 Then
            For i = LBound(oldMenu) To UBound(oldMenu)
                If n = Val(oldMenu(i)) Then
                    myFood(n).alreadyUsedInCurrentMenu = True
                End If
            Next
        End If
    Next

    'create menu
    sCurrentMenu = ""
    bGeneratErrors = False
    
    If lblMonday.ForeColor = vbButtonText Then GenerateDay lblMonday, 0 Else sCurrentMenu = sCurrentMenu & oldMenu(0) & ","
    If lblTuesday.ForeColor = vbButtonText Then GenerateDay lblTuesday, 1 Else sCurrentMenu = sCurrentMenu & oldMenu(1) & ","
    If lblWednesday.ForeColor = vbButtonText Then GenerateDay lblWednesday, 2 Else sCurrentMenu = sCurrentMenu & oldMenu(2) & ","
    If lblThursday.ForeColor = vbButtonText Then GenerateDay lblThursday, 3 Else sCurrentMenu = sCurrentMenu & oldMenu(3) & ","
    If lblFriday.ForeColor = vbButtonText Then GenerateDay lblFriday, 4 Else sCurrentMenu = sCurrentMenu & oldMenu(4) & ","
    If lblSaturday.ForeColor = vbButtonText Then GenerateDay lblSaturday, 5 Else sCurrentMenu = sCurrentMenu & oldMenu(5) & ","
    If lblSunday.ForeColor = vbButtonText Then GenerateDay lblSunday, 6 Else sCurrentMenu = sCurrentMenu & oldMenu(6) & ","

    lblMonday.ForeColor = vbButtonText
    lblTuesday.ForeColor = vbButtonText
    lblWednesday.ForeColor = vbButtonText
    lblThursday.ForeColor = vbButtonText
    lblFriday.ForeColor = vbButtonText
    lblSaturday.ForeColor = vbButtonText
    lblSunday.ForeColor = vbButtonText

End Sub

Private Sub lblFriday_Click()

    If lblFriday.ForeColor = vbButtonText Then
        lblFriday.ForeColor = &H8000&
    Else
        lblFriday.ForeColor = vbButtonText
    End If

    DoEvents
    Call TestFirstTimeHolding

End Sub

Private Sub lblMonday_Click()

    If lblMonday.ForeColor = vbButtonText Then
        lblMonday.ForeColor = &H8000&
    Else
        lblMonday.ForeColor = vbButtonText
    End If

    DoEvents
    Call TestFirstTimeHolding

End Sub

Private Sub lblSaturday_Click()

    If lblSaturday.ForeColor = vbButtonText Then
        lblSaturday.ForeColor = &H8000&
    Else
        lblSaturday.ForeColor = vbButtonText
    End If

    DoEvents
    Call TestFirstTimeHolding

End Sub

Private Sub lblSunday_Click()

    If lblSunday.ForeColor = vbButtonText Then
        lblSunday.ForeColor = &H8000&
    Else
        lblSunday.ForeColor = vbButtonText
    End If

    DoEvents
    Call TestFirstTimeHolding

End Sub

Private Sub lblThursday_Click()

    If lblThursday.ForeColor = vbButtonText Then
        lblThursday.ForeColor = &H8000&
    Else
        lblThursday.ForeColor = vbButtonText
    End If

    DoEvents
    Call TestFirstTimeHolding

End Sub

Private Sub lblTuesday_Click()

    If lblTuesday.ForeColor = vbButtonText Then
        lblTuesday.ForeColor = &H8000&
    Else
        lblTuesday.ForeColor = vbButtonText
    End If

    DoEvents
    Call TestFirstTimeHolding

End Sub

Private Sub lblWednesday_Click()

    If lblWednesday.ForeColor = vbButtonText Then
        lblWednesday.ForeColor = &H8000&
    Else
        lblWednesday.ForeColor = vbButtonText
    End If

    DoEvents
    Call TestFirstTimeHolding

End Sub


Private Sub TimerMouse_Timer()

    If iBrowserDoc <> -1 Then Exit Sub
    If Screen.MousePointer <> 0 Then Exit Sub

    For n = 0 To picButtons.Count - 1
        picButtons(n).Visible = False
    Next

    For n = 0 To picButtons.Count - 1

        If MouseY(frmMain.hWnd) > picButtons(n).Top And MouseY(frmMain.hWnd) < picButtons(n).Top + picButtons(n).height Then
            If MouseX(frmMain.hWnd) > picButtons(n).Left And MouseX(frmMain.hWnd) < picButtons(n).Left + picButtons(n).Width Then
                picButtons(n).Visible = True
                SetCursor LoadCursor(0, IDC_HAND)   'set hand cursor

                Exit For
            End If
        End If

    Next

End Sub

Private Sub picButtons_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

    If iBrowserDoc <> -1 Then Exit Sub
    If Screen.MousePointer <> 0 Then Exit Sub
    
    Select Case Index
    Case 0:
        SetCursor LoadCursor(0, IDC_HAND)   'set hand cursor
        
        If bFirstTimeClicking = True Then
            Call MsgBox("Since this is the first time you are using the program, I have entered a few common dishes for you. These are just to get you started, and you can safely remove them as you please!", vbInformation)
            bFirstTimeClicking = False
            SaveSettings
        End If
        
        iBrowserDoc = 99
        frmManage.Show vbModal
        iBrowserDoc = -1
    
    Case 1:
        iBrowserDoc = 1
        frmBrowser.Show vbModal
        iBrowserDoc = -1

    Case 2:
        Call PrintMenu
    
    Case 3:
        ret = ShellExecute(Me.hWnd, vbNullString, "http://www.rlvision.com/script/redirect.asp?app=dinnerwiz_donate", vbNullString, "c:\", SW_SHOWNORMAL)
    
    Case 4: 'web link
        ret = ShellExecute(Me.hWnd, vbNullString, "http://www.rlvision.com", vbNullString, "c:\", SW_SHOWNORMAL)
    
    Case 5: 'generate menu
        If bFirstTimeClicking = True Then
            Call MsgBox("Since this is the first time you are creating a menu, I have entered a few common dishes for you. These are just to get you started, and you can safely remove them as you please! To add and remove meals, please click on 'Setup' at the bottom.", vbInformation)
            bFirstTimeClicking = False
            SaveSettings
        End If
        
        tmrOverlay.Enabled = True
        
    End Select
    
    If Index <> 5 Then Call SetTooltips    'they disappear if I don't do this...
    
End Sub

Private Sub picButtons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If iBrowserDoc <> -1 Then Exit Sub
    If Screen.MousePointer <> 0 Then Exit Sub
    SetCursor LoadCursor(0, IDC_HAND)   'set hand cursor
End Sub

Private Sub tmrOverlay_Timer()

    Static height
    Static mode

    If height = 0 Then height = 80 'init
    

    If mode = 0 Then
        Screen.MousePointer = 11
        If height < 440 Then
            picOverlay.Visible = True
            height = height + 10
        Else
            mode = 1
            Call GenerateMenu
        End If
    ElseIf mode = 1 Then
        If height > 130 Then
            height = height - 10
        Else
            mode = 0
            tmrOverlay.Enabled = False
            picOverlay.Visible = False
            Screen.MousePointer = 0
            If bGeneratErrors = True And bGeneratErrors_ShowWarning = True Then
                Call MsgBox("Some days are empty because the program could not find enough entries to create" & vbNewLine & "a unique menu. Fix this by adding more dishes or ticking more days for the current ones.", vbInformation)
                bGeneratErrors_ShowWarning = False
                SaveSettings
            End If
        End If
    End If

    picOverlay.height = height

End Sub

Private Sub GenerateDay(myLabel As Label, myDay As Integer)

    ReDim tmp(0)
    For n = 1 To UBound(myFood())
        If myFood(n).CheckedDays(myDay) = 1 And myFood(n).alreadyUsedInCurrentMenu = False Then
            ReDim Preserve tmp(UBound(tmp) + 2)
            tmp(UBound(tmp()) - 1) = myFood(n).Name
            tmp(UBound(tmp())) = myFood(n).Name
        End If
        If myFood(n).CheckedDays(myDay) = 2 And myFood(n).alreadyUsedInCurrentMenu = False Then
            ReDim Preserve tmp(UBound(tmp) + 1)
            tmp(UBound(tmp)) = myFood(n).Name
        End If
    Next
    If UBound(tmp()) > 0 Then
        n = Int((UBound(tmp()) - 1 + 1) * Rnd + 1)
        myLabel = tmp(n)
        For nn = 1 To UBound(myFood())
            If myFood(nn).Name = tmp(n) Then
                myFood(nn).alreadyUsedInCurrentMenu = True
                sCurrentMenu = sCurrentMenu & nn & ","
            End If
        Next
    Else
        sCurrentMenu = sCurrentMenu & "0,"
        myLabel = ""
        bGeneratErrors = True
    End If


    Call FixLabel(myLabel)

End Sub

Private Sub FixLabel(myLabel As Label)

    myLabel = Replace(myLabel, "&", "&&")
    
    lblAutosize = " ..."
    minusWidth = lblAutosize.Width
    
    'line 1
    line1 = myLabel
    lblAutosize = line1
    bCut = False
    While lblAutosize.Width > 160
        line1 = Left(line1, Len(line1) - 1)
        lblAutosize = line1
        bCut = True
        'DoEvents
    Wend
    If bCut = True Then
        While Right(line1, 1) <> " " And Len(line1) > 0
            line1 = Left(line1, Len(line1) - 1)
            'DoEvents
        Wend
    End If
    
    'line 2
    line2 = ""
    If bCut = True And Len(line1) > 0 Then
        line2 = Mid(myLabel, Len(line1))
        lblAutosize = Trim(line2)
        bCut = False
        While lblAutosize.Width > 160 - minusWidth
            line2 = Left(line2, Len(line2) - 1)
            lblAutosize = line2
            bCut = True
            'DoEvents
        Wend
        If bCut = True Then
            While Right(line2, 1) <> " " And Len(line2) > 0
                line2 = Left(line2, Len(line2) - 1)
                'DoEvents
            Wend
            line2 = line2 & " ..."
        End If
    End If
    ''
    If Len(line1) = 0 Then line1 = myLabel  'fix for really long strings
    
    myLabel = line1 & vbNewLine & line2

End Sub

Private Sub SetTooltips()

    'tooltips won't work inside IDE for some reason

    Dim n As Integer
    
    ReDim myTT(5)
    For n = LBound(myTT) To UBound(myTT)
        Set myTT(n) = New CTooltip
        myTT(n).Style = TTBalloon
        myTT(n).Icon = TTIconInfo
        myTT(n).VisibleTime = 10000
        myTT(n).WrapText = True

        Select Case n
        Case 0:
            myTT(n).Title = "Setup"
            myTT(n).TipText = "This is where you add and remove dishes that you want to have on your menu."
        Case 1:
            myTT(n).Title = "Help"
            myTT(n).TipText = "Learn how to use Dinner Wizard and show additional information about this program."
        Case 2:
            myTT(n).Title = "Print"
            myTT(n).TipText = "Print this menu! A print dialog will show where you can choose printer and other settings."
        Case 3:
            myTT(n).Title = "Donate"
            myTT(n).TipText = "This program is freeware, but if you like it you can show your appreciation and make me happy by donating a small sum."
        Case 4:
            myTT(n).Title = "Go to Homepage"
            myTT(n).TipText = "Visit my homepage for other neat applications!"
        Case 5:
            myTT(n).Title = "Generate Menu"
            myTT(n).TipText = "Click here to generate a new weekly menu!"
        End Select
        myTT(n).Create picButtons(n).hWnd
    Next

End Sub

Private Sub PrintMenu()

    'sorry about the Swedish comments in here...

    Dim P As Object
    
    On Error GoTo errHandler


    bPrintToPicture = False
    
    'settings
    marginBase = 30     '-del av pappret
    scaleFactorMenu = 0.55   '% av papprets bredd
    scaleFactorDec1 = 0.32  '%av papprets bredd
    scaleFactorDec2 = 0.27  '% av papprets bredd
    lineYOffset = 0.02  '% av papprets höjd
    menuHeader = 0.12    '% av papprets höjd
    
    'text settings
    textYOffset = 0.09  '% av papprets höjd
    leftOffset = 0.07   'offset i sidled från mitten av pappret
    textYStart = 0.04   'mellanrum från linjen till först raden
    
    If bPrintToPicture = False Then
        CommonDialog.CancelError = True
        CommonDialog.PrinterDefault = True  'detta sätter default printer (i systemet?) så att printer objektet pekar rätt. men det funkar bara en gång per körning. fattar inte varför...
        CommonDialog.Flags = cdlPDHidePrintToFile + cdlPDNoPageNums + cdlPDNoSelection
        On Error Resume Next
        ' display the dialog
        CommonDialog.ShowPrinter
        If Err.Number = 32755 Then Exit Sub ' user cancelled
        On Error GoTo errHandler
        
        Set P = Printer
        Printer.Orientation = vbPRORPortrait
        Printer.Copies = CommonDialog.Copies
        Debug.Print Printer.DeviceName
    Else
        picPrinter.Cls
        picPrinter.Width = picPrinter.height * 0.707  'A4
        'picPrinter.Width = picPrinter.height * 0.77  'Letter
        Set P = picPrinter
        frmMain.Width = 12000
    End If


    sFont = "Palatino Linotype"
    If Not IsWin2000Plus Then sFont = "Book Antiqua"
    P.FontName = sFont

    'beräkna marginaler
    ratio = P.ScaleWidth / P.ScaleHeight
    marginX = (P.ScaleWidth / marginBase)
    marginY = (P.ScaleHeight / marginBase * ratio)
    
    If bPrintToPicture = True Then
        P.Line (marginX, marginY)-(P.ScaleWidth - marginX, P.ScaleHeight - marginY), vbRed, B
    End If
    
    'bilder
    '(för att printa transparenta bilder måste man använda en imagelist och köpoera bilden därifrån istället för via loadpicture/paintpicture metoden)
    
    'menu text = 757x247
    desiredWidth = (P.ScaleWidth - (marginX * 2)) * scaleFactorMenu
    origWidth = P.ScaleX(757, vbPixels, P.ScaleMode)
    ratio = desiredWidth / origWidth
    origHeight = P.ScaleY(247, vbPixels, P.ScaleMode)
    desiredHeight = origHeight * ratio
    X = (P.ScaleWidth / 2) - (desiredWidth / 2)
    Y = (P.ScaleHeight - (marginY / 2)) * menuHeader
    P.PaintPicture LoadPicture(App.Path & "\print_menu.gif"), X, Y, desiredWidth, desiredHeight
    
    'linje under meny text
    textY = Y
    textY = textY + desiredHeight + (P.ScaleHeight - (marginY / 2)) * lineYOffset
    P.Line (marginX + (marginX / 2), textY)-(P.ScaleWidth - marginX - (marginX / 2), textY), &H63669 ', BF
    textY = textY + 1
    P.Line (marginX + (marginX / 2), textY)-(P.ScaleWidth - marginX - (marginX / 2), textY), &H63669 ', BF

    'dekorationer ska vara ca 25% av bredden (exkl. marginaler)
        
    'decoration 1 (top/right) = 299x187
    desiredWidth = (P.ScaleWidth - (marginX * 2)) * scaleFactorDec1
    origWidth = P.ScaleX(299, vbPixels, P.ScaleMode)
    ratio = desiredWidth / origWidth
    origHeight = P.ScaleY(187, vbPixels, P.ScaleMode)
    desiredHeight = origHeight * ratio
    P.PaintPicture LoadPicture(App.Path & "\print_dec1.gif"), P.ScaleWidth - marginX - desiredWidth, marginY, desiredWidth, desiredHeight
    
    'decoration 1 (bottom/left) = 297x481
    desiredWidth = (P.ScaleWidth - (marginX * 2)) * scaleFactorDec2
    origWidth = P.ScaleX(297, vbPixels, P.ScaleMode)
    ratio = desiredWidth / origWidth
    origHeight = P.ScaleY(481, vbPixels, P.ScaleMode)
    desiredHeight = origHeight * ratio
    P.PaintPicture LoadPicture(App.Path & "\print_dec2.gif"), marginX, P.ScaleHeight - marginY - desiredHeight, desiredWidth, desiredHeight
    
    
    'mat & veckodagar
    Dim days(7): Dim food(7)
    days(0) = "Mon:": days(1) = "Tue:": days(2) = "Wed:": days(3) = "Thu:": days(4) = "Fri:": days(5) = "Sat:": days(6) = "Sun:"
    food(0) = lblMonday: food(1) = lblTuesday: food(2) = lblWednesday: food(3) = lblThursday: food(4) = lblFriday: food(5) = lblSaturday: food(6) = lblSunday
    
    'räkna ut storleken på fonterna för att det ska stämma oavsett upplösning & papper osv
    P.FontSize = 1
    P.FontBold = True
    P.FontItalic = True
    While P.TextWidth("Mon") < (P.ScaleWidth / 11)
        P.FontSize = P.FontSize + 1
        DoEvents
    Wend
    daySize = P.FontSize
    dayHeight = P.TextHeight("Mon")
        
    P.FontSize = 1
    P.FontBold = False
    P.FontItalic = True
    While P.TextWidth("Mon") < (P.ScaleWidth / 14)
        P.FontSize = P.FontSize + 1
        DoEvents
    Wend
    foodSize = P.FontSize
    foodPlusY = ((dayHeight - P.TextHeight("Mon")) / 3) * 2

    textY = textY + (P.ScaleHeight - (marginY / 2)) * textYStart
    
    For n = 0 To 6

        'day
        sMsg = days(n)
        P.FontSize = daySize
        P.ForeColor = &H63669
        P.FontBold = True
        P.FontItalic = True
        
        P.CurrentX = (P.ScaleWidth / 2) - P.TextWidth(sMsg) - (P.ScaleWidth * leftOffset)
        P.CurrentY = textY
        P.Print sMsg
    
        'food
        food(n) = Replace(food(n), "&&", "&")
        food(n) = Replace(food(n), " " & vbNewLine & " ", vbNewLine)
        P.FontSize = foodSize
        P.ForeColor = vbBlack
        P.FontBold = False
        P.FontItalic = True
        
        tmp = Split(food(n), vbNewLine)
        For nn = LBound(tmp) To UBound(tmp)
            sMsg = tmp(nn)
            P.CurrentX = (P.ScaleWidth / 2) - ((P.ScaleWidth * leftOffset) / 2)
            P.CurrentY = textY + foodPlusY + nn * (P.TextHeight(sMsg) * 0.9)
            P.Print sMsg
        Next
        
        'increase y
        textY = textY + (P.ScaleHeight - (marginY / 2)) * textYOffset
    
    Next
    
    
    sMsg = "made with DinnerWiz by RL Vision"
    P.FontSize = 1
    P.FontBold = False
    P.FontItalic = True
    P.ForeColor = &H555555
    While P.TextWidth(sMsg) < (P.ScaleWidth / 5)
        P.FontSize = P.FontSize + 1
        DoEvents
    Wend
    P.CurrentX = P.ScaleWidth - P.TextWidth(sMsg) - marginX
    P.CurrentY = P.ScaleHeight - P.TextHeight(sMsg) - marginY
    P.Print sMsg


    If bPrintToPicture = False Then
        Printer.EndDoc
    End If
    

    Exit Sub
errHandler:

    Call MsgBox("Error " & Err.Number & " " & Err.Description, vbCritical)

End Sub

Private Sub TestFirstTimeHolding()
    
    If bFirstTimeHolding = True Then
        Call MsgBox("Clicking on a dish marks it green. The next time you generate a menu, marked items will be kept.", vbInformation, "Information")
        bFirstTimeHolding = False
    End If

End Sub

