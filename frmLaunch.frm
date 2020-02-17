VERSION 5.00
Begin VB.Form frmLaunch 
   BackColor       =   &H000F0F79&
   BorderStyle     =   0  'None
   Caption         =   " Auto Close Internet Explorer"
   ClientHeight    =   3375
   ClientLeft      =   1080
   ClientTop       =   0
   ClientWidth     =   4350
   Icon            =   "frmLaunch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkRandomize 
      BackColor       =   &H000F0F79&
      Caption         =   "randomize relaunch within 30 secconds"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   1200
      Width           =   3135
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   1560
      TabIndex        =   6
      Text            =   "http://www.ventaja.co.nr"
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Text            =   "1"
      Top             =   840
      Width           =   735
   End
   Begin VB.Timer tmrOpen 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2160
      Top             =   4680
   End
   Begin VB.Label lblAddress 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000F0F79&
      Caption         =   "Web Address:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   480
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   45
      Left            =   0
      Top             =   3325
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3075
      Index           =   1
      Left            =   4300
      Top             =   340
      Width           =   45
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   3075
      Index           =   0
      Left            =   0
      Top             =   340
      Width           =   45
   End
   Begin VB.Image imgPauseDown 
      Height          =   345
      Left            =   2400
      Picture         =   "frmLaunch.frx":164A
      Top             =   4200
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgPauseUp 
      Height          =   345
      Left            =   120
      Picture         =   "frmLaunch.frx":1838
      Top             =   4200
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgStartDown 
      Height          =   345
      Left            =   2400
      Picture         =   "frmLaunch.frx":1A27
      Top             =   3840
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image imgStartUp 
      Height          =   345
      Left            =   120
      Picture         =   "frmLaunch.frx":1C16
      Top             =   3840
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image butPause 
      Height          =   345
      Left            =   960
      Picture         =   "frmLaunch.frx":1E06
      Top             =   2880
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image butStart 
      Height          =   345
      Left            =   960
      Picture         =   "frmLaunch.frx":1FF5
      Top             =   1680
      Width           =   2250
   End
   Begin VB.Image imgExit 
      Height          =   225
      Left            =   3600
      Picture         =   "frmLaunch.frx":21E5
      Top             =   53
      Width           =   675
   End
   Begin VB.Image imgButtonDown 
      Height          =   225
      Left            =   2400
      Picture         =   "frmLaunch.frx":2307
      Top             =   3600
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Image imgButtonUp 
      Height          =   225
      Left            =   1680
      Picture         =   "frmLaunch.frx":2429
      Top             =   3600
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Label lblSecondsToOpen 
      Alignment       =   2  'Center
      BackColor       =   &H000F0F79&
      Caption         =   "0s until next action"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   3135
   End
   Begin VB.Label lblOpened 
      Alignment       =   2  'Center
      BackColor       =   &H000F0F79&
      Caption         =   "Opened 0 windows"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label lblMinutes 
      BackColor       =   &H000F0F79&
      Caption         =   "Minutes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   1
      Top             =   840
      Width           =   615
   End
   Begin VB.Label lblLaunchClose 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000F0F79&
      Caption         =   "Launch/Close every:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   840
      Width           =   1575
   End
   Begin VB.Image imgTitleBar 
      Height          =   345
      Left            =   0
      Picture         =   "frmLaunch.frx":254B
      Top             =   0
      Width           =   4350
   End
End
Attribute VB_Name = "frmLaunch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Adam Knox

Option Explicit

Private sAppPath As String, intTime As Integer, dblTotal As Double, intStarted As Integer, intTime2 As Integer
Private intCount As Integer, intSeconds As Integer, intRandom As Integer
Private CurrX, CurrY As Single

Private Sub butPause_Click()
'Disable timer
    tmrOpen.Enabled = False
'Disable the pause button
    butPause.Visible = False
'Enable the start button
    butStart.Visible = True
'Disable the textboxes
    txtTime.Enabled = True
    txtAddress.Enabled = True
'Enable Check Button
    chkRandomize.Enabled = True
End Sub

Private Sub butPause_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the exit picture
    butPause.Picture = imgPauseDown.Picture
End Sub

Private Sub butPause_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the exit picture
    butPause.Picture = imgPauseUp.Picture
End Sub

Private Sub butStart_Click()
    If Val(txtTime.Text) > 0 Then
    
        If intStarted = 0 Then
        'Open program
            Shell "C:\Program Files\Internet Explorer\iexplore.exe " & txtAddress
        'Increase the total windows opened
            dblTotal = dblTotal + 1
        'Display the total number of windows opened
            lblOpened.Caption = "Opened " & dblTotal & " windows"
        ElseIf intStarted = 1 Then

        'Reset the seconds if necessary
            If intTime2 <> intTime Then
                intSeconds = 0
                intCount = 0
            End If
        End If
        
    'Setup the timer
        tmrOpen.Enabled = True
    'Set the time to a variable
        intTime = txtTime.Text
    'Enable the pause button
        butPause.Visible = True
    'Disable the start button
        butStart.Visible = False
    'Disable the textboxes
        txtAddress.Enabled = False
        txtTime.Enabled = False
    'Disable Check Button
        chkRandomize.Enabled = False
            
        If chkRandomize.Value = 1 Then
            Randomize Timer
            intRandom = Int(Rnd * 30) + 30
        ElseIf chkRandomize.Value = 0 Then
            intRandom = 60
        End If
    
    
    Else
        MsgBox "Please select how often the program should open"
    End If
End Sub

Private Sub butStart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the exit picture
    butStart.Picture = imgStartDown.Picture
End Sub

Private Sub butStart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the exit picture
    butStart.Picture = imgStartUp.Picture
End Sub

Private Sub imgExit_Click()

'Exit the program
    Unload Me
End Sub

Private Sub imgExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the exit picture
    imgExit.Picture = imgButtonDown.Picture
End Sub

Private Sub imgExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Change the exit picture
    imgExit.Picture = imgButtonUp.Picture
End Sub

Private Sub imgTitlebar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Me.Left = Me.Left + (X - CurrX)
    Me.Top = Me.Top + (Y - CurrY)
End If
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    Me.Left = Me.Left + (X - CurrX)
    Me.Top = Me.Top + (Y - CurrY)
End If
End Sub

Private Sub imgTitlebar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    CurrX = X
    CurrY = Y
End Sub

Private Sub tmrOpen_Timer()
'Dimension Variable

    
    intCount = intCount + 1
    intSeconds = intSeconds + 1
    intTime2 = intTime
    
    If intStarted = 0 Then intStarted = 1
    
    lblSecondsToOpen.Caption = (intTime * intRandom) - intSeconds & "s before the window is reopened"
    
    If intCount = intTime * intRandom Then
    'Randomize if necessary
        If chkRandomize.Value = 1 Then
            Randomize Timer
            intRandom = Int(Rnd * 30) + 30
        ElseIf chkRandomize.Value = 0 Then
            intRandom = 60
        End If
    'Reset the number of seconds till the window opens
        intSeconds = 0
    'Increase the total windows opened
        dblTotal = dblTotal + 1
    'Display the total number of windows opened
        lblOpened.Caption = "Opened " & dblTotal & " windows"
    'Close explorer
        Dim shWin As New ShellWindows
        Dim IE As InternetExplorer

        For Each IE In shWin
            IE.Quit
        Next
    
        'Call EndTask("Ventaja Designs - The way of the future - Windows Internet Explorer")
    'Open program
        Shell "C:\Program Files\Internet Explorer\iexplore.exe " & txtAddress
    
    'Reset the timer
        intCount = 0
    End If
End Sub
