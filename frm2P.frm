VERSION 5.00
Begin VB.Form frm2P 
   Appearance      =   0  'Flat
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7170
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   9060
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   7170
   ScaleWidth      =   9060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnStart 
      BackColor       =   &H0080C0FF&
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
      Left            =   6240
      Picture         =   "frm2P.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnExit 
      BackColor       =   &H0080C0FF&
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
      Left            =   7440
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frm2P.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton btnPause 
      BackColor       =   &H0080C0FF&
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
      Left            =   6240
      Picture         =   "frm2P.frx":D0A4
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton btnCont 
      BackColor       =   &H0080C0FF&
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
      Left            =   7440
      Picture         =   "frm2P.frx":138F6
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   960
      Width           =   1215
   End
   Begin VB.Frame frameP2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Player 2"
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
      Left            =   240
      TabIndex        =   9
      Top             =   3840
      Width           =   3015
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
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
         Index           =   2
         Left            =   240
         TabIndex        =   16
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
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
         Index           =   2
         Left            =   240
         TabIndex        =   15
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label lblUser 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
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
         Index           =   2
         Left            =   360
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label lblz 
         BackColor       =   &H0080C0FF&
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
         Index           =   2
         Left            =   240
         TabIndex        =   13
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblMatchesP2 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   2040
         TabIndex        =   12
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblCurrentTimeP2 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   2040
         TabIndex        =   11
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblTurnsP2 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   2040
         TabIndex        =   10
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Frame frameP1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Player 1"
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
      Height          =   2775
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   3015
      Begin VB.Label lblz 
         BackColor       =   &H0080C0FF&
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
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblUser 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
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
         Index           =   1
         Left            =   360
         TabIndex        =   7
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
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
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080C0FF&
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
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   1680
         Width           =   1575
      End
      Begin VB.Label lblCurrentTimeP1 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   2040
         TabIndex        =   4
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label lblMatchesP1 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblTurnsP1 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000015&
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   1680
         Width           =   855
      End
   End
   Begin VB.Timer tmrP2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   960
      Top             =   0
   End
   Begin VB.Timer tmrP1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer tmrDelay 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   0
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   0
      Left            =   3600
      Top             =   2160
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   1
      Left            =   4440
      Top             =   2160
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   3
      Left            =   6120
      Top             =   2160
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   4
      Left            =   3600
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   5
      Left            =   4440
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   6
      Left            =   5280
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   7
      Left            =   6120
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   8
      Left            =   3600
      Top             =   4560
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   9
      Left            =   4440
      Top             =   4560
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   10
      Left            =   5280
      Top             =   4560
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   11
      Left            =   6120
      Top             =   4560
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   12
      Left            =   3600
      Top             =   5760
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   13
      Left            =   4440
      Top             =   5760
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   14
      Left            =   5280
      Top             =   5760
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   15
      Left            =   6120
      Top             =   5760
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   2
      Left            =   5280
      Top             =   2160
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   16
      Left            =   6960
      Top             =   2160
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   17
      Left            =   6960
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   18
      Left            =   6960
      Top             =   4560
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   19
      Left            =   6960
      Top             =   5760
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   20
      Left            =   7800
      Top             =   2160
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   21
      Left            =   7800
      Top             =   3360
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   22
      Left            =   7800
      Top             =   4560
      Width           =   750
   End
   Begin VB.Image imgCard 
      BorderStyle     =   1  'Fixed Single
      Height          =   1125
      Index           =   23
      Left            =   7800
      Top             =   5760
      Width           =   750
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080C0FF&
      Caption         =   "2 Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   17
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label lblDelay 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Enabled         =   0   'False
      ForeColor       =   &H80000015&
      Height          =   255
      Left            =   3240
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frm2P"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim arnFlip(24), arnCardIndex(12) As Integer
Dim blnAlready(23) As Boolean
Dim nTimeP1, nTimeP2, nTurnsTakenP1, nTurnsTakenP2 As Integer
Dim nMatchesP1, nMatchesP2, nScoreP1, nScoreP2 As Integer
Dim nCountWho As Integer

Private Sub Form_Load()
    CenterForm Me
    lblUser(1).Caption = InputBox("Welcome to 2 Player Mode. Please set Player 1 username")
    lblUser(2).Caption = InputBox("Now please set Player 2 username")
End Sub

Private Sub btnStart_Click()
    Dim n, y, d As Integer
    lblCurrentTimeP1.Caption = "0"
    lblCurrentTimeP2.Caption = "0"

    nClickCount = 0
    nTurnsTakenP1 = 0
    nTurnsTakenP2 = 0
    lblDelay.Caption = 0
    nMatchesP1 = 0
    nMatchesP2 = 0
    nCountWho = 0
    lblUser(1).ForeColor = &HFFFFFF 'white, P1 goes first

    btnPause.Enabled = True
    btnCont.Enabled = False
    tmrP1.Enabled = True 'P1 goes first
    
    For y = 0 To 11
        arnCardIndex(y) = 0
    Next
    
    Randomize 'Distinct values
    'Gets card face values ready
    For n = 0 To 23
        imgCard(n).Enabled = True
        imgCard(n).Picture = frmCards.imgFlip(100) 'If a user restarts
        Do
            arnFlip(n) = Random(0, 11)
        Loop Until arnCardIndex(arnFlip(n)) <= 1
        arnCardIndex(arnFlip(n)) = arnCardIndex(arnFlip(n)) + 1
    Next

    For d = 0 To 23
        blnAlready(d) = False
    Next
    
    btnStart.Caption = "Restart"
    
End Sub

Private Sub imgCard_Click(Index As Integer)
    nClickCount = nClickCount + 1
    If nClickCount = 1 Then
        n1st = Index
        imgCard(n1st).Enabled = False 'disable ability to click twice
        imgCard(n1st).Picture = frmCards.imgHard(arnFlip(n1st))
        
    ElseIf nClickCount = 2 Then
        nCountWho = nCountWho + 1
        n2nd = Index
        If nCountWho Mod 2 = 0 Then
            nTurnsTakenP2 = nTurnsTakenP2 + 1
            tmrP2.Enabled = False
            tmrP1.Enabled = True
        Else
            nTurnsTakenP1 = nTurnsTakenP1 + 1
            tmrP1.Enabled = False
            tmrP2.Enabled = True
        End If
        lblTurnsP1.Caption = nTurnsTakenP1
        lblTurnsP2.Caption = nTurnsTakenP2
        imgCard(n2nd).Enabled = False
        imgCard(n2nd).Picture = frmCards.imgHard(arnFlip(n2nd))

        If imgCard(n1st).Picture = imgCard(n2nd).Picture Then ' a match
            If nCountWho Mod 2 = 0 Then
                nMatchesP2 = nMatchesP2 + 1
                lblMatchesP2.Caption = nMatchesP2
                tmrP2.Enabled = False
                tmrP1.Enabled = True
            Else
                nMatchesP1 = nMatchesP1 + 1
                lblMatchesP1.Caption = nMatchesP1
                tmrP1.Enabled = False
                tmrP2.Enabled = True
            End If
            imgCard(n1st).Enabled = False
            imgCard(n2nd).Enabled = False
            blnAlready(n1st) = True
            blnAlready(n2nd) = True
        Else 'no match
            tmrDelay.Enabled = True
        End If
        nClickCount = 0

    End If
    
    If nMatchesP1 + nMatchesP2 = 12 Then 'when users gets all matches
        tmrP1.Enabled = False
        tmrP2.Enabled = False
        MsgBox ("Great Job!")
        GoScores
        If MsgBox("Would you like to play 2 player again?", vbYesNo) = vbNo Then
            frm2P.Visible = False
        Else
            btnStart_Click
        End If
    End If
End Sub

Private Sub tmrDelay_Timer()
    Dim c, t As Integer
    For c = 0 To 23
        imgCard(c).Enabled = False 'disable ability to click during delay
    Next
    If lblDelay.Caption = 3 Then
        imgCard(n1st).Picture = frmCards.imgFlip(100)
        imgCard(n2nd).Picture = frmCards.imgFlip(100)
        imgCard(n1st).Enabled = True 'able to click again
        imgCard(n2nd).Enabled = True
        For t = 0 To 23
            imgCard(t).Enabled = True
            If blnAlready(t) = True Then
                imgCard(t).Enabled = False 'if already a match, disable clickability
            End If
        Next
        lblDelay.Caption = 0
        If nCountWho Mod 2 Then 'tells whose turn it is
            lblUser(1).ForeColor = &H80000015 'grey
            lblUser(2).ForeColor = &HFFFFFF 'white
        Else
            lblUser(1).ForeColor = &HFFFFFF
            lblUser(2).ForeColor = &H80000015
        End If
        tmrDelay.Enabled = False
        Exit Sub
    Else
        lblDelay.Caption = lblDelay.Caption + 1
    End If
End Sub

Private Sub GoScores()
    Dim nTurnsTakenP1, nTurnsTakenP2 As Integer
    nTimeP1 = Val(lblCurrentTimeP1.Caption)
    nTimeP2 = Val(lblCurrentTimeP2.Caption)
    nTurnsTakenP1 = Val(lblTurnsP1.Caption)
    nTurnsTakenP2 = Val(lblTurnsP2.Caption)
    nScoreP1 = nMatchesP1 * (70 - nTimeP1) * (50 - nTurnsTakenP1)
    nScoreP2 = nMatchesP2 * (70 - nTimeP2) * (50 - nTurnsTakenP2)
    MsgBox ("Player 1 score is " & nScoreP1 & ". Player 2 score is " & nScoreP2)
    If nScoreP1 > nScoreP2 Then
        MsgBox ("Player 1 is the winner!")
    ElseIf nScoreP1 = nScoreP2 Then
        MsgBox ("There's a tie!")
    Else
        MsgBox ("Player 2 is the winner!")
    End If
    btnStart.Caption = "Start" 'if the form is loaded again
End Sub

Private Sub tmrP1_Timer()
    lblCurrentTimeP1.Caption = lblCurrentTimeP1.Caption + 1
End Sub

Private Sub tmrP2_Timer()
    lblCurrentTimeP2.Caption = lblCurrentTimeP2.Caption + 1
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
    For u = 0 To 23
        imgCard(u).Enabled = False
    Next
    If lblUser(1).ForeColor = &HFFFFFF Then
        tmrP1.Enabled = False
    ElseIf lblUser(2).ForeColor = &HFFFFFF Then
        tmrP2.Enabled = False
    End If
End Sub
Private Sub btnCont_Click()
    Dim v As Integer
    btnPause.Enabled = True
    btnCont.Enabled = False
    If lblUser(1).ForeColor = &HFFFFFF Then
        tmrP1.Enabled = True
    ElseIf lblUser(2).ForeColor = &HFFFFFF Then
        tmrP2.Enabled = True
    End If
    For v = 0 To 23
        imgCard(v).Enabled = True
        If blnAlready(v) = True Then 'user can't click on matched cards
            imgCard(v).Enabled = False
        End If
    Next
        
End Sub
