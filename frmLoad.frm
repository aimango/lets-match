VERSION 5.00
Begin VB.Form frmLoad 
   BackColor       =   &H00FFCCBB&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3300
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5265
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   3300
   ScaleWidth      =   5265
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSelect 
      BackColor       =   &H00FFCCBB&
      Caption         =   "&Select"
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
      Left            =   3480
      Picture         =   "frmLoad.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton btnExit 
      BackColor       =   &H00FFCCBB&
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
      Left            =   3480
      MaskColor       =   &H00FFC0C0&
      Picture         =   "frmLoad.frx":6852
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ListBox lstNavigate 
      BackColor       =   &H00FFCCBB&
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFCCBB&
      Caption         =   "Let's Match"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   39.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CenterForm Me
    lstNavigate.AddItem "1 Player - Easy Mode"
    lstNavigate.AddItem "1 Player - Hard Mode"
    lstNavigate.AddItem "2 Player"
    lstNavigate.AddItem "High Scores"
    lstNavigate.AddItem "About Game"
End Sub

Private Sub btnExit_Click()
If MsgBox("You are now exiting the game.", vbInformation + vbOKCancel, "Close Game") = vbOK Then
        Unload Me
    End If
End Sub

Private Sub btnSelect_Click()

    If lstNavigate = "1 Player - Easy Mode" Then
        frm1P_Easy.Visible = True
    ElseIf lstNavigate = "1 Player - Hard Mode" Then
        frm1P_Hard.Visible = True
    ElseIf lstNavigate = "2 Player" Then
        frmPvsP.Visible = True
    ElseIf lstNavigate = "High Scores" Then
        frmScores.Visible = True
    ElseIf lstNavigate = "About Game" Then

        MsgBox "About Let's Match" & _
            vbCrLf & "Created by: Elisa Lou" & _
            vbCrLf & "Class TIK20I" & _
            vbCrLf & "Last updated January 24, 2008" & _
            vbCrLf & "2007-2008", vbInformation, "About Let's Match"
        
    End If
End Sub
