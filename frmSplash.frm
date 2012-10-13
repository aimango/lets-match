VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3585
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Moveable        =   0   'False
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3585
   ScaleWidth      =   6000
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrLoad 
      Interval        =   100
      Left            =   4560
      Top             =   120
   End
   Begin VB.Label lblLoadTime 
      BackColor       =   &H00FFFFFF&
      Caption         =   "0"
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CenterForm Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = 1 And lblLoadTime.Caption = 10 Then
        frmLoad.Visible = True
        frmSplash.Visible = False
    End If
End Sub

Private Sub tmrLoad_Timer()
    If lblLoadTime.Caption = "15" Then
        frmLoad.Visible = True
        frmSplash.Visible = False
        tmrLoad.Enabled = False
    Else
        lblLoadTime.Caption = Val(lblLoadTime.Caption) + 1
    End If
End Sub
