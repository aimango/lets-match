VERSION 5.00
Begin VB.Form frm1P_Easy 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6885
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7785
   ControlBox      =   0   'False
   FillColor       =   &H00400000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   6885
   ScaleWidth      =   7785
   Begin VB.CommandButton btnCont 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Continue"
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
      Left            =   6000
      Picture         =   "frm1P_Easy.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnPause 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Pause"
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
      Left            =   4800
      Picture         =   "frm1P_Easy.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnExit 
      BackColor       =   &H00FFC0C0&
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
      Left            =   6000
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frm1P_Easy.frx":D0A4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnStart 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Start"
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
      Left            =   4800
      Picture         =   "frm1P_Easy.frx":138F6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   840
      Top             =   360
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "High Scores to Beat"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   1455
      Left            =   4200
      TabIndex        =   6
      Top             =   5160
      Width           =   3015
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Best Score"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label lblBestScore 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   1095
      End
   End
   Begin VB.Timer tmrAscend 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   360
      Top             =   360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Your Stats"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000015&
      Height          =   2895
      Left            =   4200
      TabIndex        =   0
      Top             =   1920
      Width           =   3015
      Begin VB.Label lblTurns 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Turns Taken:"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Matches:"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label lblMatches 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lblUser 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblz 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Time Taken (sec):"
         BeginProperty Font 
            Name            =   "Tw Cen MT"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblCurrentTime 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   2040
         TabIndex        =   1
         Top             =   1080
         Width           =   735
      End
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   15
      Left            =   2760
      Top             =   5160
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   14
      Left            =   2040
      Top             =   5160
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   13
      Left            =   1320
      Top             =   5160
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   12
      Left            =   600
      Top             =   5160
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   11
      Left            =   2760
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   10
      Left            =   2040
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   9
      Left            =   1320
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   8
      Left            =   600
      Top             =   4080
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   7
      Left            =   2760
      Top             =   3000
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   6
      Left            =   2040
      Top             =   3000
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   5
      Left            =   1320
      Top             =   3000
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   4
      Left            =   600
      Top             =   3000
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   3
      Left            =   2760
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   2
      Left            =   2040
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   1
      Left            =   1320
      Top             =   1920
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   0
      Left            =   600
      Top             =   1920
      Width           =   750
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "1 Player - Easy Mode"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   12
      Top             =   840
      Width           =   3135
   End
   Begin VB.Label lblDelay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      ForeColor       =   &H80000015&
      Height          =   375
      Left            =   1680
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "frm1P_Easy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arnFlip(16), arnCardIndex(8) As Integer
Dim blnAlready(16) As Boolean

Private Sub Form_Load()
    CenterForm Me
    sUser = InputBox("Welcome to 1 Player Easy Mode. Please set your username")
    lblUser.Caption = sUser
    
    'shows Top Score to Beat
    Open "topscoreE.txt" For Input As #1
        Do Until EOF(1)
            Input #1, nBestScore
            lblBestScore.Caption = nBestScore
        Loop
    Close #1

End Sub

Private Sub btnStart_Click()
    Dim n, y, d As Integer
    btnPause.Enabled = True
    btnCont.Enabled = False
    tmrAscend.Enabled = True

    lblCurrentTime.Caption = "0"
    nClickCount = 0
    lblTurns.Caption = "0"
    lblDelay.Caption = "0"
    lblMatches.Caption = "0"
    

    For y = 0 To 7
        arnCardIndex(y) = 0
    Next
    Randomize 'Distinct values
    'Gets card face values ready
    For n = 0 To 15
        imgCard(n).Enabled = True
        imgCard(n).Picture = frmCards.imgFlip(100) 'If a user restarts
        Do
            arnFlip(n) = Random(0, 7)
        Loop Until arnCardIndex(arnFlip(n)) <= 1
        arnCardIndex(arnFlip(n)) = arnCardIndex(arnFlip(n)) + 1
    Next

    For d = 0 To 15
        blnAlready(d) = False
    Next
    btnStart.Caption = "Restart"
    
End Sub

Private Sub imgCard_Click(Index As Integer)
    nClickCount = nClickCount + 1
    If nClickCount = 1 Then
        n1st = Index
        imgCard(n1st).Enabled = False 'disable ability to click twice
        imgCard(n1st).Picture = frmCards.imgEasy(arnFlip(n1st))
        
    ElseIf nClickCount = 2 Then
        n2nd = Index
        lblTurns.Caption = lblTurns.Caption + 1
        imgCard(n2nd).Enabled = False
        imgCard(n2nd).Picture = frmCards.imgEasy(arnFlip(n2nd))
        
        If imgCard(n1st).Picture = imgCard(n2nd).Picture Then 'a match
            lblMatches.Caption = Val(lblMatches.Caption) + 1
            imgCard(n1st).Enabled = False
            imgCard(n2nd).Enabled = False
            blnAlready(n1st) = True
            blnAlready(n2nd) = True
           
        Else 'no match
            tmrDelay.Enabled = True
        End If
        nClickCount = 0
    End If
    
    If Val(lblMatches.Caption) = 8 Then 'when user gets all matches
        MsgBox ("Great Job!")
        tmrAscend.Enabled = False
        GoScores
    End If
End Sub

Private Sub tmrDelay_Timer()
     Dim c, t As Integer
    
    For c = 0 To 15
        imgCard(c).Enabled = False 'disable ability to click during delay
    Next
    If lblDelay.Caption = 3 Then
        lblDelay.Caption = 0
        imgCard(n1st).Picture = frmCards.imgFlip(100) 'flips back around
        imgCard(n2nd).Picture = frmCards.imgFlip(100)
        imgCard(n1st).Enabled = True 'able to click again
        imgCard(n2nd).Enabled = True
        For t = 0 To 15
            imgCard(t).Enabled = True
            If blnAlready(t) = True Then
                imgCard(t).Enabled = False 'if already a match, disable clickability
            End If
        Next
        tmrDelay.Enabled = False
        Exit Sub
    Else
        lblDelay.Caption = lblDelay.Caption + 1
    End If
End Sub

Private Sub tmrAscend_Timer()
    If lblCurrentTime.Caption = 80 Then 'when  time runs out
        tmrAscend.Enabled = False
        If MsgBox("Game Over. Would you like to restart?", vbYesNo) = vbNo Then
            GoScores
            btnStart.Caption = "Start"
        Else
            btnStart_Click
        End If
    Else
        lblCurrentTime.Caption = Val(lblCurrentTime.Caption) + 1
    End If
End Sub


Function Score(Time, Turns As Integer) As Integer
    Score = (100 - Turns) * (81 - Time)
End Function

Private Sub GoScores()
    Dim FF, nTurnsTaken As Integer
    FF = FreeFile
    nTime = Val(lblCurrentTime.Caption)
    nTurnsTaken = lblTurns.Caption
    nScore = Score(nTime, nTurnsTaken)
    MsgBox ("Your score is " & nScore)
    If nScore > nBestScore Then
        MsgBox ("Congrats, you've beaten the High Score for Easy Mode!")
        Open "topscoreE.txt" For Output As #1
            Write #1, nScore
        Close #1
    End If
    Open "easy_scores.txt" For Append As FF
        Print #FF, sUser
        Print #FF, nScore
    Close FF
    btnStart.Caption = "Start" 'if the form is loaded again
    frmScores.Visible = True
    frm1P_Easy.Visible = False
End Sub

Private Sub btnExit_Click()
    If MsgBox("You are now exiting the game.", vbInformation + vbOKCancel, "Close Game") = vbOK Then
        Unload Me
    End If
End Sub

Private Sub btnPause_Click()
    Dim u As Integer
    btnCont.Enabled = True
    btnPause.Enabled = False
    For u = 0 To 15
        imgCard(u).Enabled = False
    Next
    tmrAscend.Enabled = False
End Sub

Private Sub btnCont_Click()
    Dim v As Integer
    btnPause.Enabled = True
    btnCont.Enabled = False
    tmrAscend.Enabled = True
    For v = 0 To 15
        imgCard(v).Enabled = True
        If blnAlready(v) = True Then 'user can't click on matched cards
            imgCard(v).Enabled = False
        End If
    Next
End Sub
