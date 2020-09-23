VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Pong"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   3630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
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
      Height          =   1695
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOk_Click()
Unload frmAbout
End Sub


Private Sub Form_Load()
Label1.Caption = "Welcome to my game of Pong, F1 = Start, F2 = Pause, F3 = Unpause, F4 = Exit, Use the Text in the game itself to Start and Unpause the game and use the up and down arrow keys to play the game, Enjoy, By Lee Foy"
End Sub


