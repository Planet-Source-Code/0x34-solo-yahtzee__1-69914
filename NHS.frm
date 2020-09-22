VERSION 5.00
Begin VB.Form NHS 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Your Name Please....."
   ClientHeight    =   2295
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "NHS.frx":0000
   ScaleHeight     =   2295
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   400
      Left            =   3960
      Top             =   1800
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3480
      Top             =   1800
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   0
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "1st"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Name Please"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "1st"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   2780
      TabIndex        =   6
      Top             =   860
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New High Score!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "New High Score!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   10
      TabIndex        =   3
      Top             =   140
      Width           =   4455
   End
End
Attribute VB_Name = "NHS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Solo Yahtzee

'                                       Programming By
'                                  Ken Slater 2007/2008/2009
'                                            0x34

Private Sub Command1_Click()
    EnterIt
End Sub

Public Sub EnterIt()
    If Text1.Text <> "" Then
        Scores.HSHolder = Text1
        ATS(NHSplace).Name = Text1
    Else
        Scores.HSHolder = "(Unknown)"
        ATS(NHSplace).Name = "(Unknown)"
    End If
    ATS(NHSplace).Score = PreFOT
    ATS(NHSplace).Qnt = AmtYaht
    ATS(NHSplace).Date = Format$(Now, "m/d/yy")
    If iTEST = False Then
        Scores.HighScore = PreFOT
        frmMain.Caption = " Solo Yahtzee                          High Score = " & ATS(1).Score & _
        " by " & ATS(1).Name
        ATS(NHSplace).Score = PreFOT
        ATS(NHSplace).Name = Scores.HSHolder
        ATS(NHSplace).Qnt = AmtYaht
        ATS(NHSplace).Date = Format$(Now, "m/d/yy")
        SaveHighScore
    Else
        frmMain.Caption = " Solo Yahtzee                TEST MODE High Score = " & PreFOT & "  (Not Saved)"
    End If
    frmScores.Show vbModal
    Unload Me
End Sub

Private Sub Form_Load()
Dim Y As Integer
Dim T As Integer
Dim K As String
Dim TEMP(3) As String
    Timer1.Enabled = True
    If PreFOT > ATS(1).Score Then   'Score Sorter Bucket Burgade
        ATS(5).Date = ATS(4).Date: ATS(5).Name = ATS(4).Name: ATS(5).Score = ATS(4).Score: ATS(5).Qnt = ATS(4).Qnt
        ATS(4).Date = ATS(3).Date: ATS(4).Name = ATS(3).Name: ATS(4).Score = ATS(3).Score: ATS(4).Qnt = ATS(3).Qnt
        ATS(3).Date = ATS(2).Date: ATS(3).Name = ATS(2).Name: ATS(3).Score = ATS(2).Score: ATS(3).Qnt = ATS(2).Qnt
        ATS(2).Date = ATS(1).Date: ATS(2).Name = ATS(1).Name: ATS(2).Score = ATS(1).Score: ATS(2).Qnt = ATS(1).Qnt
        K = "1st Place": NHSplace = 1
    ElseIf PreFOT > ATS(2).Score Then
        ATS(5).Date = ATS(4).Date: ATS(5).Name = ATS(4).Name: ATS(5).Score = ATS(4).Score: ATS(5).Qnt = ATS(4).Qnt
        ATS(4).Date = ATS(3).Date: ATS(4).Name = ATS(3).Name: ATS(4).Score = ATS(3).Score: ATS(4).Qnt = ATS(3).Qnt
        ATS(3).Date = ATS(2).Date: ATS(3).Name = ATS(2).Name: ATS(3).Score = ATS(2).Score: ATS(3).Qnt = ATS(2).Qnt
        K = "2nd Place": NHSplace = 2
    ElseIf PreFOT > ATS(3).Score Then
        ATS(5).Date = ATS(4).Date: ATS(5).Name = ATS(4).Name: ATS(5).Score = ATS(4).Score: ATS(5).Qnt = ATS(4).Qnt
        ATS(4).Date = ATS(3).Date: ATS(4).Name = ATS(3).Name: ATS(4).Score = ATS(3).Score: ATS(4).Qnt = ATS(3).Qnt
        K = "3rd Place": NHSplace = 3
    ElseIf PreFOT > ATS(4).Score Then
        ATS(5).Date = ATS(4).Date: ATS(5).Name = ATS(4).Name: ATS(5).Score = ATS(4).Score: ATS(5).Qnt = ATS(4).Qnt
        K = "4th Place": NHSplace = 4
    Else
        K = "5th Place": NHSplace = 5
    End If
    Label5.Caption = K: Label4 = K
    ScorePlaced = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnterIt
    End If
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    On Error Resume Next
    Text1.SetFocus
End Sub

Private Sub Timer2_Timer()
    If Label4.Visible = True Then
        Label4.Visible = False
        Label5.Visible = False
    Else
        Label4.Visible = True
        Label5.Visible = True
    End If
End Sub
