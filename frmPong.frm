VERSION 5.00
Begin VB.Form frmPong 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   Caption         =   "Pong"
   ClientHeight    =   5535
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6600
      Top             =   5040
   End
   Begin VB.Label lblScore2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   6600
      TabIndex        =   4
      Top             =   120
      Width           =   105
   End
   Begin VB.Label lblPaused 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "Start!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   555
      Left            =   3000
      TabIndex        =   0
      Top             =   2520
      Width           =   1260
   End
   Begin VB.Shape shpWallBottom 
      Height          =   15
      Left            =   0
      Top             =   5520
      Width           =   7335
   End
   Begin VB.Shape shpWallTop 
      Height          =   15
      Left            =   0
      Top             =   0
      Width           =   7335
   End
   Begin VB.Image shpBall 
      Height          =   450
      Left            =   3360
      Picture         =   "frmPong.frx":0000
      Top             =   1920
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image shpPlayer1 
      Height          =   1140
      Left            =   120
      Picture         =   "frmPong.frx":0B82
      Top             =   1680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Image shpPlayer2 
      Height          =   1140
      Left            =   6720
      Picture         =   "frmPong.frx":1FF4
      Top             =   1680
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "Computer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   5520
      TabIndex        =   2
      Top             =   120
      Width           =   930
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackColor       =   &H80000007&
      Caption         =   "Player"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblScore1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000012&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   105
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileStart 
         Caption         =   "&Start"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuFileP 
         Caption         =   "&Pause"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuFilePaused 
         Caption         =   "&Paused"
         Shortcut        =   {F3}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileBar 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About Pong"
      End
   End
End
Attribute VB_Name = "frmPong"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Vmom As Integer
Dim Hmom As Integer
Dim PadSpeed As Integer
Dim OrigPaddleLoc As Integer
Dim OrigBallLocY As Integer
Dim OrigBallLocX As Integer
Dim Score1 As Integer, Score2 As Integer
Dim Tmp As Integer
Dim Missed As Integer



Private Sub BallMissed()
Missed = Missed + 1
End Sub


Private Sub Update_Score(Player As Integer)
Select Case Player
Case 1
Score1 = Score1 + 1
lblScore1.Caption = Format(Score1, "#0")
lblScore1.Refresh
Case 2
Score2 = Score2 + 1
lblScore2.Caption = Format(Score2, "#0")
lblScore2.Refresh
End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Then    'the up key
    PadSpeed = -200
ElseIf KeyCode = 40 Then    'the down key
    PadSpeed = 200
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
PadSpeed = 0    'stop the paddle from moving
End Sub


Private Sub Form_Load()
Hmom = -150
Vmom = 0
'record the origional starting locations for everything
OrigPaddleLoc = shpPlayer1.Top
OrigBallLocX = shpBall.Left
OrigBallLocY = shpBall.Top
End Sub


Private Sub lblPaused_Click()
lblPaused.Caption = "Start!"
If lblPaused.Caption = "Start!" Then
Timer1.Enabled = True
shpBall.Visible = True
shpPlayer1.Visible = True
shpPlayer2.Visible = True
End If
lblPaused.Visible = False
mnuFilePaused.Visible = False
mnuFileP.Visible = True
mnuFileStart.Visible = False
End Sub

Private Sub mnuFileExit_Click()
Dim Message As String
If Timer1.Enabled = False Then
End
Else
Message = Message + "The " + lblName.Caption + " scored " + lblScore1.Caption + " and the Computer scored " + lblScore2.Caption + vbCr
MsgBox Message, vbOKOnly + vbInformation, "Your Score"
End
End If
End Sub

Private Sub mnuFileP_Click()
Timer1.Enabled = False
mnuFilePaused.Visible = True
mnuFileP.Visible = False
lblPaused.Visible = True
lblPaused.Caption = "Paused!"
End Sub

Private Sub mnuFilePaused_Click()
Timer1.Enabled = True
mnuFilePaused.Visible = False
mnuFileP.Visible = True
lblPaused.Visible = False
End Sub


Private Sub mnuFileStart_Click()
Timer1.Enabled = True
shpBall.Visible = True
shpPlayer1.Visible = True
shpPlayer2.Visible = True
lblPaused.Visible = False
mnuFileStart.Visible = False
End Sub

Private Sub mnuHelpAbout_Click()
frmAbout.Show
End Sub

Private Sub Timer1_Timer()
'move the ball
shpBall.Top = shpBall.Top + Vmom
shpBall.Left = shpBall.Left + Hmom
'check to see if the ball's hit a wall
If shpBall.Top + shpBall.Height >= shpWallBottom.Top Then
shpBall.Top = shpWallBottom.Top - shpBall.Height
Vmom = -Vmom
'Wav file goes here
ElseIf shpBall.Top <= shpWallTop.Top + shpWallTop.Height Then
shpBall.Top = shpWallTop.Top + shpWallTop.Height
Vmom = -Vmom
'Wav file goes here
End If
'move the paddle
If PadSpeed <> 0 Then
shpPlayer1.Top = shpPlayer1.Top + PadSpeed
End If
'check to see if the paddle's hit a wall
If shpPlayer1.Top <= shpWallTop.Top + shpWallTop.Height Then
shpPlayer1.Top = shpWallTop.Top + shpWallTop.Height
ElseIf shpPlayer1.Top + shpPlayer1.Height >= shpWallBottom.Top Then
shpPlayer1.Top = shpWallBottom.Top - shpPlayer1.Height
End If
If shpPlayer2.Top <= shpWallTop.Top + shpWallTop.Height Then
shpPlayer2.Top = shpWallTop.Top + shpWallTop.Height
ElseIf shpPlayer2.Top + shpPlayer2.Height >= shpWallBottom.Top Then
shpPlayer2.Top = shpWallBottom.Top - shpPlayer2.Height
End If
'move the computer player's paddle
If shpBall.Top < shpPlayer2.Top Then
shpPlayer2.Top = shpPlayer2.Top - 250
ElseIf shpBall.Top > shpPlayer2.Top + shpPlayer2.Height Then
shpPlayer2.Top = shpPlayer2.Top + 250
End If
'if the ball has hit player 1's paddle
If shpBall.Left <= shpPlayer1.Left + shpPlayer1.Width And shpBall.Left >= shpPlayer1.Left - shpPlayer1.Width Then
If shpBall.Top + shpBall.Height >= shpPlayer1.Top And shpBall.Top <= shpPlayer1.Top + shpPlayer1.Height Then
'calculate the angle it's deflecting off at
Tmp = ((shpPlayer1.Top + (shpPlayer1.Height / 2)) - (shpBall.Top + (shpBall.Height / 2))) * 0.55
Vmom = Vmom + -Tmp
'Wav file goes here
shpBall.Left = shpPlayer1.Left + shpPlayer1.Width
'deflect the ball
Hmom = -Hmom
End If
End If
'if the ball has hit player 2's paddle
If shpBall.Left + shpBall.Width >= shpPlayer2.Left And shpBall.Left <= shpPlayer2.Left + shpPlayer2.Width Then
If shpBall.Top + shpBall.Height >= shpPlayer2.Top And shpBall.Top <= shpPlayer2.Top + shpPlayer2.Height Then
'calculate the angle it's deflecting off at
Tmp = ((shpPlayer2.Top + (shpPlayer2.Height / 2)) - (shpBall.Top + (shpBall.Height / 2))) * 0.55
Vmom = Vmom + -Tmp
'Wav files goes here
shpBall.Left = shpPlayer2.Left - shpBall.Width
'deflect the ball
Hmom = -Hmom
End If
End If
'Scores
If (shpBall.Left) > frmPong.ScaleWidth Then
Call Update_Score(1)
End If
If (shpBall.Left + shpBall.Width) < 0 Then
Call Update_Score(2)
End If
'reset the ball and paddles to their origional location
If shpBall.Left + shpBall.Width < 0 Then
shpBall.Left = OrigBallLocX
shpBall.Top = OrigBallLocY
shpPlayer1.Top = OrigPaddleLoc
shpPlayer2.Top = OrigPaddleLoc
Hmom = -150
Vmom = 0
ElseIf shpBall.Left > frmPong.Width Then
'reset the ball and paddles to their origional location
shpBall.Left = OrigBallLocX
shpBall.Top = OrigBallLocY
shpPlayer1.Top = OrigPaddleLoc
shpPlayer2.Top = OrigPaddleLoc
Hmom = 150
Vmom = 0
End If
End Sub


