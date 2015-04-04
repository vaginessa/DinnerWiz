Attribute VB_Name = "Module"
Public bFirstTimeUsing As Boolean
Public bFirstTimeHolding As Boolean
Public bFirstTimeClicking As Boolean
Public bGeneratErrors As Boolean
Public bGeneratErrors_ShowWarning As Boolean

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'mouse x/y
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

'hand mouse pointer
Public Const IDC_HAND = 32649&
Public Const IDC_ARROW = 32512&
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long


Type food
    Name As String
    CheckedDays(7) As Integer
    alreadyUsedInCurrentMenu As Boolean
End Type

Public myFood() As food
Public sCurrentMenu As String
Public iBrowserDoc As Integer


' Get mouse X coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen
Function MouseX(Optional ByVal hWnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hWnd Then ScreenToClient hWnd, lpPoint
    MouseX = lpPoint.X
End Function

' Get mouse Y coordinates in pixels
'
' If a window handle is passed, the result is relative to the client area
' of that window, otherwise the result is relative to the screen
Function MouseY(Optional ByVal hWnd As Long) As Long
    Dim lpPoint As POINTAPI
    GetCursorPos lpPoint
    If hWnd Then ScreenToClient hWnd, lpPoint
    MouseY = lpPoint.Y
End Function

Public Sub SaveSettings()

    myFile = sysFolders.AppData & "\RL Vision\DinnerWiz\DinnerWiz.dat"

    If Dir(App.Path & "\DinnerWiz.dat") <> "" Then
        myFile = App.Path & "\DinnerWiz.dat"
    Else

        If Dir(sysFolders.AppData & "\RL Vision\*.*", vbArchive + vbReadOnly + vbSystem + vbDirectory) = "" Then
            MkDir (sysFolders.AppData & "\RL Vision")
        End If
        If Dir(sysFolders.AppData & "\RL Vision\DinnerWiz\*.*", vbArchive + vbReadOnly + vbSystem + vbDirectory) = "" Then
            MkDir (sysFolders.AppData & "\RL Vision\DinnerWiz")
        End If

    End If

    Open myFile For Output As #1

        Print #1, 3 'version

        If bFirstTimeUsing = True Then Print #1, 1 Else Print #1, 0
        If bFirstTimeClicking = True Then Print #1, 1 Else Print #1, 0
        If bFirstTimeHolding = True Then Print #1, 1 Else Print #1, 0
        If bGeneratErrors_ShowWarning = True Then Print #1, 1 Else Print #1, 0

        Print #1, sCurrentMenu

        For n = 1 To UBound(myFood())
            
            myLine = myFood(n).Name & "¤" & _
                        myFood(n).CheckedDays(0) & "¤" & _
                        myFood(n).CheckedDays(1) & "¤" & _
                        myFood(n).CheckedDays(2) & "¤" & _
                        myFood(n).CheckedDays(3) & "¤" & _
                        myFood(n).CheckedDays(4) & "¤" & _
                        myFood(n).CheckedDays(5) & "¤" & _
                        myFood(n).CheckedDays(6) & "¤"


            Print #1, myLine
            
        Next
        
        dataChanged = False

    Close #1

End Sub

Public Sub LoadSettings()

    On Error Resume Next
    
    myFile = sysFolders.AppData & "RL Vision\DinnerWiz\DinnerWiz.dat"
    If Dir(App.Path & "\DinnerWiz.dat") <> "" Then myFile = App.Path & "\DinnerWiz.dat"
    If Dir(myFile, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) = "" Then
        myFile = App.Path & "\DefaultSettings.dat"
    Else
        If FileLen(myFile) = 0 Then myFile = App.Path & "\DefaultSettings.dat"
    End If

     If Dir(myFile, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) <> "" Then
        Open myFile For Input As #1
        
            Line Input #1, ver  'version
            
            Line Input #1, tmp
            If tmp = 1 Then
                bFirstTimeUsing = True
            End If
    
            Line Input #1, tmp
            If tmp = 1 Then
                bFirstTimeClicking = True
            End If

            If ver >= 3 Then
                Line Input #1, tmp
                If tmp = 1 Then
                    bFirstTimeHolding = True
                End If
            Else
                bFirstTimeHolding = True
            End If
            
            Line Input #1, tmp
            If tmp = 1 Then
                bGeneratErrors_ShowWarning = True
            End If

            Line Input #1, sCurrentMenu
            sCurrentMenu = Trim(sCurrentMenu)
    
            While Not EOF(1)
        
                Line Input #1, tmp
                tmp = Split(tmp, "¤")
                
                n = UBound(myFood()) + 1
                
                ReDim Preserve myFood(n)
                
                myFood(n).Name = tmp(0)
                
                For i = 0 To 6
                    myFood(n).CheckedDays(i) = tmp(i + 1)
                Next
                    
                DoEvents
    
            Wend
        
        Close #1
    End If


    Exit Sub


End Sub
