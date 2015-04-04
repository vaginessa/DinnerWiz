VERSION 5.00
Begin VB.Form frmManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Dishes"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   Icon            =   "frmManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      ToolTipText     =   "Removes ALL items in the list!"
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Del"
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      ToolTipText     =   "Removes the currently selected item in the food list"
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "Add Dish"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   10
      ToolTipText     =   "Adds a new dish to your food list"
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CheckBox chkSunday 
      Caption         =   "Sundays"
      Height          =   255
      Left            =   3840
      TabIndex        =   9
      Top             =   3720
      Width           =   1695
   End
   Begin VB.CheckBox chkSaturday 
      Caption         =   "Saturdays"
      Height          =   255
      Left            =   3840
      TabIndex        =   8
      Top             =   3360
      Width           =   1695
   End
   Begin VB.CheckBox chkFriday 
      Caption         =   "Fridays"
      Height          =   255
      Left            =   3840
      TabIndex        =   7
      Top             =   3000
      Width           =   1695
   End
   Begin VB.CheckBox chkThursday 
      Caption         =   "Thursdays"
      Height          =   255
      Left            =   3840
      TabIndex        =   6
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CheckBox chkWednesday 
      Caption         =   "Wednesdays"
      Height          =   255
      Left            =   3840
      TabIndex        =   5
      Top             =   2280
      Width           =   1695
   End
   Begin VB.CheckBox chkTuesday 
      Caption         =   "Tuesdays"
      Height          =   255
      Left            =   3840
      TabIndex        =   4
      Top             =   1920
      Width           =   1695
   End
   Begin VB.CheckBox chkMonday 
      Caption         =   "Mondays"
      Height          =   255
      Left            =   3840
      TabIndex        =   3
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox txtName 
      Height          =   285
      Left            =   3480
      MaxLength       =   64
      TabIndex        =   2
      Top             =   600
      Width           =   2055
   End
   Begin VB.ListBox lstFood 
      Height          =   5325
      Left            =   240
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      TabIndex        =   0
      ToolTipText     =   "Save changes and close this window"
      Top             =   5520
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Suitable to eat on:"
      Height          =   255
      Left            =   3480
      TabIndex        =   13
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   3480
      TabIndex        =   12
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private currentArrayIndex As Integer
Private editNameArrayIndex As Integer
Private editNameListIndex As Integer
Private allowUpdate As Boolean
Private useGrayChckboxes As Boolean
Private dataChanged As Boolean

Private Sub Form_Load()

    allowUpdate = False
    useGrayChckboxes = False
    dataChanged = False

    'fill food list
    For n = 1 To UBound(myFood())
        lstFood.AddItem myFood(n).Name
    Next

    If lstFood.ListCount > 0 Then lstFood.Selected(0) = True

    allowUpdate = True


    'XP Themes
    If IsThemed() Then
        FixThemeSupport Controls
    End If

    SetIcon Me.hWnd, "AAA", True

    '''''''''''''

End Sub

Private Sub Form_Activate()

    frmManage.cmdNew.SetFocus

End Sub

Private Sub lstFood_Click()

    allowUpdate = False

    'get entry in food list

    For n = 1 To UBound(myFood())
        If lstFood.List(lstFood.ListIndex) = myFood(n).Name Then Exit For
    Next
    currentArrayIndex = n

    txtName = myFood(n).Name

    chkMonday_LastValue = 0
    chkTuesday_LastValue = 0
    chkWednesday_LastValue = 0
    chkThursday_LastValue = 0
    chkFriday_LastValue = 0
    chkSaturday_LastValue = 0
    chkSunday_LastValue = 0

    chkMonday.Value = myFood(n).CheckedDays(0)
    chkTuesday.Value = myFood(n).CheckedDays(1)
    chkWednesday.Value = myFood(n).CheckedDays(2)
    chkThursday.Value = myFood(n).CheckedDays(3)
    chkFriday.Value = myFood(n).CheckedDays(4)
    chkSaturday.Value = myFood(n).CheckedDays(5)
    chkSunday.Value = myFood(n).CheckedDays(6)

    chkMonday_LastValue = myFood(n).CheckedDays(0)
    chkTuesday_LastValue = myFood(n).CheckedDays(1)
    chkWednesday_LastValue = myFood(n).CheckedDays(2)
    chkThursday_LastValue = myFood(n).CheckedDays(3)
    chkFriday_LastValue = myFood(n).CheckedDays(4)
    chkSaturday_LastValue = myFood(n).CheckedDays(5)
    chkSunday_LastValue = myFood(n).CheckedDays(6)

    allowUpdate = True

End Sub


Private Sub cmdNew_Click()

    newFood = InputBox("Please enter a name for the new dish:", "New Dish")
retry:
    If newFood <> "" Then

        'validate entry, may not be duplicate
        dupFound = False
        For n = 1 To UBound(myFood())
            If myFood(n).Name = newFood Then
                newFood = InputBox("Duplicate name. Please enter another name:", "New Dish", newFood)
                GoTo retry
            End If
        Next
        '''''''''''''''''''''''''''''''''''''


        n = UBound(myFood()) + 1

        ReDim Preserve myFood(n)

        myFood(n).Name = newFood

        For i = 0 To 6
            myFood(n).CheckedDays(i) = 1
        Next

        lstFood.AddItem newFood
        DoEvents

        'select item just added
        For n = 0 To lstFood.ListCount
            If lstFood.List(n) = newFood Then Exit For
        Next
        lstFood.Selected(n) = True

        dataChanged = True

    End If

End Sub

Private Sub chkMonday_Click()
    If useGrayChckboxes = True Then If chkMonday_LastValue = 1 Then chkMonday.Value = 2
    chkMonday_LastValue = chkMonday.Value
    Call Update_Values
End Sub
Private Sub chkTuesday_Click()
    If useGrayChckboxes = True Then If chkTuesday_LastValue = 1 Then chkTuesday.Value = 2
    chkTuesday_LastValue = chkTuesday.Value
    Call Update_Values
End Sub
Private Sub chkwednesday_Click()
    If useGrayChckboxes = True Then If chkWednesday_LastValue = 1 Then chkWednesday.Value = 2
    chkWednesday_LastValue = chkWednesday.Value
    Call Update_Values
End Sub
Private Sub chkthursday_Click()
    If useGrayChckboxes = True Then If chkThursday_LastValue = 1 Then chkThursday.Value = 2
    chkThursday_LastValue = chkThursday.Value
    Call Update_Values
End Sub
Private Sub chkfriday_Click()
    If useGrayChckboxes = True Then If chkFriday_LastValue = 1 Then chkFriday.Value = 2
    chkFriday_LastValue = chkFriday.Value
    Call Update_Values
End Sub
Private Sub chksaturday_Click()
    If useGrayChckboxes = True Then If chkSaturday_LastValue = 1 Then chkSaturday.Value = 2
    chkSaturday_LastValue = chkSaturday.Value
    Call Update_Values
End Sub
Private Sub chksunday_Click()
    If useGrayChckboxes = True Then If chkSunday_LastValue = 1 Then chkSunday.Value = 2
    chkSunday_LastValue = chkSunday.Value
    Call Update_Values
End Sub

Private Sub cmdDelete_Click()

    If lstFood.ListCount = 0 Then Exit Sub

    sel = lstFood.ListIndex

    ret = MsgBox("Remove '" & lstFood.List(lstFood.ListIndex) & "'?", vbYesNo + vbQuestion)
    If ret = vbYes Then
    
        txtName = ""
        chkMonday = 0
        chkTuesday = 0
        chkWednesday = 0
        chkThursday = 0
        chkFriday = 0
        chkSaturday = 0
        chkSunday = 0
    
        'remove from array (move all items one step up)
        For n = currentArrayIndex To UBound(myFood()) - 1
            myFood(n).Name = myFood(n + 1).Name
            For i = 0 To 6
                myFood(n).CheckedDays(i) = myFood(n + 1).CheckedDays(i)
                myFood(n + 1).CheckedDays(i) = 0
            Next
        Next
        ReDim Preserve myFood(UBound(myFood()) - 1)
    
        'remove from listbox
        lstFood.RemoveItem (lstFood.ListIndex)

        If lstFood.ListCount > sel Then
            lstFood.Selected(sel) = True
        ElseIf lstFood.ListCount > 0 Then
            lstFood.Selected(lstFood.ListCount - 1) = True
        End If

        sCurrentMenu = ""

        dataChanged = True

    End If

End Sub

Private Sub txtName_GotFocus()
    
    editNameArrayIndex = currentArrayIndex
    editNameListIndex = lstFood.ListIndex

End Sub

Private Sub txtName_LostFocus()

retry:

    'validate entry, may not be duplicate
    dupFound = False
    For n = 1 To UBound(myFood())
        If myFood(n).Name = txtName And n <> editNameArrayIndex Then
            dupFound = True
            Exit For
        End If
    Next
    
    If dupFound = True Then
        newName = InputBox("The name already exist! Please enter another name:", "Duplicate Name", txtName)
        If newName <> "" Then
            allowUpdate = False
            txtName = newName
            lstFood.List(editNameListIndex) = newName
            myFood(editNameArrayIndex).Name = newName
            allowUpdate = True
        End If
        GoTo retry
    End If

End Sub

Private Sub txtName_Change()

    If allowUpdate = False Or lstFood.ListCount = 0 Then Exit Sub

    Call Update_Values

End Sub

Private Sub Update_Values()

    If allowUpdate = False Then Exit Sub

    If lstFood.ListCount = 0 Then Exit Sub

    If txtName <> myFood(currentArrayIndex).Name Then
        myFood(currentArrayIndex).Name = txtName
        lstFood.List(lstFood.ListIndex) = txtName
    End If

    myFood(currentArrayIndex).CheckedDays(0) = chkMonday.Value
    myFood(currentArrayIndex).CheckedDays(1) = chkTuesday.Value
    myFood(currentArrayIndex).CheckedDays(2) = chkWednesday.Value
    myFood(currentArrayIndex).CheckedDays(3) = chkThursday.Value
    myFood(currentArrayIndex).CheckedDays(4) = chkFriday.Value
    myFood(currentArrayIndex).CheckedDays(5) = chkSaturday.Value
    myFood(currentArrayIndex).CheckedDays(6) = chkSunday.Value

    dataChanged = True

End Sub

Private Sub cmdClear_Click()


    ret = MsgBox("Are you sure you want to remove all dishes from the list?", vbYesNo + vbQuestion)
    If ret = vbYes Then
    
        ReDim myFood(0) ' entry 0 is not used
    
        'remove from listbox
        lstFood.Clear
        txtName = ""
        chkMonday = 0
        chkTuesday = 0
        chkWednesday = 0
        chkThursday = 0
        chkFriday = 0
        chkSaturday = 0
        chkSunday = 0
        
        
        sCurrentMenu = ""
    
        dataChanged = True

    End If


End Sub

Private Sub cmdClose_Click()
    
    'save settings file
    If dataChanged = True Then
    
        SaveSettings
        
    End If
    '''''''''''''''
    
    Unload Me
End Sub

