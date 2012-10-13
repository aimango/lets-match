VERSION 5.00
Begin VB.Form frmScores 
   BackColor       =   &H00AEF0B7&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4560
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5565
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   4560
   ScaleWidth      =   5565
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picHard 
      BackColor       =   &H00AEF0B7&
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   2880
      ScaleHeight     =   3315
      ScaleWidth      =   2475
      TabIndex        =   2
      Top             =   1080
      Width           =   2535
   End
   Begin VB.PictureBox picEasy 
      BackColor       =   &H00AEF0B7&
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      ScaleHeight     =   3315
      ScaleWidth      =   2595
      TabIndex        =   1
      Top             =   1080
      Width           =   2655
   End
   Begin VB.CommandButton btnExit 
      BackColor       =   &H00AEF0B7&
      Caption         =   "&Exit"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmScores.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Image imgScore 
      Height          =   600
      Left            =   600
      Picture         =   "frmScores.frx":6852
      Top             =   240
      Width           =   3000
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nHScore As Integer
Dim sListedUser As String

Private Sub btnExit_Click()
    Unload Me
End Sub

Private Sub imgScore_Click()
    picHard.Cls
    picEasy.Cls
    picHard.Print "Username", "High Score"
    picEasy.Print "Username", "High Score"
    
    Open "hard_scores.txt" For Input As #1
        Do Until EOF(1)
            Input #1, sListedUser, nHScore
            picHard.Print sListedUser, nHScore
        Loop
    Close #1
    
    Open "easy_scores.txt" For Input As #2
        Do Until EOF(2)
            Input #2, sListedUser, nHScore
            picEasy.Print sListedUser, nHScore
        Loop
    Close #2
End Sub

