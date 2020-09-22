VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmScores 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Scores This Session"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin RichTextLib.RichTextBox RT 
      Height          =   3615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   6376
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmScores.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblHighScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Highest Score EVER ="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   3700
      Width           =   5055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Highest Score EVER ="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   30
      TabIndex        =   2
      Top             =   3730
      Width           =   5055
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Solo Yahtzee

'                                       Programming By
'                                    Ken Slater 2007/2008
'                                            0x34



Private Sub Form_Load() ' My highest yet - 618 Points! 4 Yahtzee's!
Dim T As Integer
Dim Y As Integer
Dim NeedToSave As Boolean
Dim iScrs(100) As Integer
On Error GoTo ERROR
    NeedToSave = False
    RT = ""
    If GameScores(500) < 1 Then
        RT.Locked = False
        RT.SelStart = 0
        RT.Font.Bold = True
        RT.SelColor = vbRed
        RT.SelText = vbNewLine & vbNewLine & vbNewLine & vbNewLine & " No Games Played this Session" & vbNewLine & vbNewLine
        Me.Caption = "  Scores This Session "
        lblHighScore = "Highest Score EVER = " & Scores.HighScore
        Label1 = lblHighScore
        RT.Locked = True
        Exit Sub
    End If
    Me.Caption = "  Scores This Session"
    lblHighScore = "Highest Score EVER = " & Scores.HighScore
    Label1 = lblHighScore
    RT.Locked = False
    RT.SelStart = 0
    RT.Font.Bold = False
    For T = 1 To GameScores(500)
        RT.SelColor = vbGreen
        RT.SelText = " " & GameScores(T) & " Points" & vbNewLine
    Next
    RT.Locked = True
    Exit Sub
ERROR:
    MsgBox "An error occurred while looking over scores.", vbInformation, "Ken Goofed"
End Sub

Private Sub RT_Click()
    Unload Me
End Sub
