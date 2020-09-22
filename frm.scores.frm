VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmScores 
   BackColor       =   &H00800000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " "
   ClientHeight    =   8160
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8160
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox GrfButt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   600
      Picture         =   "frm.scores.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   50
      Top             =   1960
      Width           =   255
   End
   Begin VB.PictureBox GrfButt 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   240
      Picture         =   "frm.scores.frx":0102
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   49
      Top             =   1960
      Width           =   255
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6720
      ScaleHeight     =   315
      ScaleWidth      =   270
      TabIndex        =   48
      Top             =   4080
      Width           =   330
   End
   Begin VB.CommandButton GrfButton 
      Caption         =   "G"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1920
      TabIndex        =   46
      ToolTipText     =   "Large Graph"
      Top             =   5760
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   360
      Picture         =   "frm.scores.frx":0204
      ScaleHeight     =   330
      ScaleWidth      =   810
      TabIndex        =   43
      Top             =   0
      Width           =   810
   End
   Begin MSComDlg.CommonDialog ComD1 
      Left            =   120
      Top             =   7560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer CloseTIM 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2040
      Top             =   7080
   End
   Begin VB.Timer OpenTIM 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1560
      Top             =   7080
   End
   Begin VB.Timer SAVETimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1080
      Top             =   7080
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Height          =   2295
      Left            =   3720
      TabIndex        =   30
      Top             =   4560
      Width           =   3495
      Begin VB.CommandButton Command5 
         Height          =   195
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Populate 20 Games - TEST"
         Top             =   240
         Width           =   135
      End
      Begin VB.CommandButton ClearSelected 
         Caption         =   "Clear Selected HighScores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   775
         TabIndex        =   32
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Open YHS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   615
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Import HighScores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         TabIndex        =   40
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton MergeFile 
         Caption         =   "Merge HighScore File"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2280
         TabIndex        =   38
         Top             =   900
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00800000&
         Caption         =   "Check1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   4
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   37
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00800000&
         Caption         =   "Check1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   3
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   36
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00800000&
         Caption         =   "Check1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   2
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   35
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00800000&
         Caption         =   "Check1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   34
         Top             =   600
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00800000&
         Caption         =   "Check1"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   240
         MaskColor       =   &H00FFFFFF&
         TabIndex        =   33
         Top             =   360
         Width           =   1935
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Export HighScores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   31
         Top             =   1680
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   2295
      Left            =   0
      TabIndex        =   12
      Top             =   4560
      Width           =   3735
      Begin VB.CommandButton Command6 
         Caption         =   "Save Current Configuration"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   39
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Load Defaults"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         TabIndex        =   29
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Position"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Score"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Name"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Date"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   24
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Apply"
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
         Left            =   1200
         TabIndex        =   23
         Top             =   1680
         Width           =   975
      End
      Begin VB.HScrollBar HS 
         Height          =   255
         Index           =   0
         Left            =   1680
         Max             =   255
         TabIndex        =   15
         Top             =   240
         Width           =   1455
      End
      Begin VB.HScrollBar HS 
         Height          =   255
         Index           =   1
         Left            =   1680
         Max             =   255
         TabIndex        =   14
         Top             =   480
         Width           =   1455
      End
      Begin VB.HScrollBar HS 
         Height          =   255
         Index           =   2
         Left            =   1680
         Max             =   255
         TabIndex        =   13
         Top             =   720
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Yahtzee Count"
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   45
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   750
         Width           =   1215
      End
      Begin VB.Label LV 
         BackStyle       =   0  'Transparent
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   22
         Top             =   240
         Width           =   615
      End
      Begin VB.Label LV 
         BackStyle       =   0  'Transparent
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   21
         Top             =   480
         Width           =   615
      End
      Begin VB.Label LV 
         BackStyle       =   0  'Transparent
         Caption         =   "255"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   3240
         TabIndex        =   20
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Spot 
         BorderStyle     =   1  'Fixed Single
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "R"
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
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   18
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "G"
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
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   17
         Top             =   480
         Width           =   255
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "B"
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
         Height          =   255
         Index           =   2
         Left            =   1320
         TabIndex        =   16
         Top             =   720
         Width           =   255
      End
   End
   Begin VB.Timer SPLtim 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   600
      Top             =   7080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   120
      Top             =   7080
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Top             =   7680
      Width           =   255
   End
   Begin RichTextLib.RichTextBox RF 
      Height          =   1575
      Left            =   1920
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frm.scores.frx":105E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RD 
      Height          =   1575
      Left            =   600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   360
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frm.scores.frx":10E4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RS 
      Height          =   1575
      Left            =   0
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   360
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Appearance      =   0
      TextRTF         =   $"frm.scores.frx":116A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox RT 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2220
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   4154
      _Version        =   393217
      BackColor       =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frm.scores.frx":11E3
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
   Begin RichTextLib.RichTextBox RG 
      Height          =   1575
      Left            =   6120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frm.scores.frx":125E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   360
      Picture         =   "frm.scores.frx":12E4
      ScaleHeight     =   330
      ScaleWidth      =   810
      TabIndex        =   42
      Top             =   20
      Width           =   810
   End
   Begin RichTextLib.RichTextBox Rn 
      Height          =   1575
      Left            =   1440
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   360
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   2778
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"frm.scores.frx":213E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label SH 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   3000
      TabIndex        =   11
      Top             =   60
      Width           =   1095
   End
   Begin VB.Label Label4 
      Height          =   135
      Left            =   6240
      TabIndex        =   9
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scores This Session"
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
      TabIndex        =   7
      Top             =   1920
      Width           =   7215
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Scores This Session"
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
      TabIndex        =   8
      Top             =   1960
      Width           =   7215
   End
   Begin VB.Label lblHighScore 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "High Scores"
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
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "High Scores"
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
      Left            =   90
      TabIndex        =   2
      Top             =   50
      Width           =   7095
   End
End
Attribute VB_Name = "frmScores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Solo Yahtzee

'                                       Programming By
'                               Ken Slater 2007/2008/2009/2010
'                                            0x34

'                                 A Very ACTIVE Scoreboard!!
'              During game play, right click above the DATES to activate "MORE" button
'                        Graphing functions added in December of 2009

Option Explicit

Dim ColNS As Boolean
Dim ErrorCatcher As Boolean
Dim A As ColorConstants
Dim mIndex As Integer
Dim iOPEN As Boolean
Dim iExit As Boolean
Dim iEnding As Boolean
Dim MaxiScore As Integer
Dim PointCount As Integer
Dim ShowGraph As Boolean
Dim iAverage As Integer

Private Sub ClearSelected_Click()
Dim T As Integer
Dim uGo As Boolean
    uGo = False
    For T = 0 To 4
        If Check1(T).Value = 1 And ATS(T + 1).Score > 0 Then uGo = True
    Next
    If uGo Then
        If MsgBox("You are about to PERMANETLY DELETE these Scores from the HighScore File!   " & vbNewLine & _
        "Are you sure you want to Continue?", vbYesNo, " WARNING") = vbYes Then
            For T = 0 To 4
                If Check1(T).Value = 1 Then
                    ATS(T + 1).Date = ""
                    ATS(T + 1).Name = ""
                    ATS(T + 1).Score = 0
                End If
            Next
            SortEm
            FillATS
            SaveHighScore
            Scores.HighScore = GetHighScore
            SavePath = ComD1.FileName
            frmMain.Caption = " Solo Yahtzee                         High Score = " & Scores.HighScore & _
            " by " & Scores.HSHolder
            For T = 0 To 4
                Check1(T).Value = 0
            Next
        End If
    Else
        MsgBox "You've gotta select a good one first...     ", vbInformation, "...nice"
        For T = 0 To 4
            Check1(T).Value = 0
        Next
    End If
End Sub

Private Sub Command1_Click()
    If Option1(0).Value = True Then
        ColRTB(0) = A
    ElseIf Option1(1).Value = True Then
        ColRTB(1) = A
    ElseIf Option1(2).Value = True Then
        ColRTB(2) = A
    ElseIf Option1(3).Value = True Then
        ColRTB(3) = A
    ElseIf Option1(4).Value = True Then
        ColRTB(4) = A
    End If
    FillATS
End Sub

Private Sub Command2_Click()
    DefCOLs
    FillATS
    UpdateColors
End Sub

Public Sub DefCOLs()    'Default Colors - Use editor to customize
    ColRTB(0) = vbGreen
    ColRTB(1) = vbRed
    ColRTB(2) = &HA4FF01
    ColRTB(3) = vbWhite
    ColRTB(4) = RGB(255, 114, 0)
End Sub

Private Sub Command3_Click()    ' Export HighScore File
On Error GoTo Error
    MousePointer = vbDefault
        With ComD1
            .FileName = "YHS"
            If SavePath <> "" Then
                .InitDir = SavePath
            Else
                .InitDir = App.Path & "\"
            End If
            .Flags = &H1
            .Flags = &H2
            .Flags = &H4
            .DefaultExt = "bin"
            .Filter = "Dialogue Files (*.bin)|*.bin"
            .DialogTitle = "Export Yahtzee HighScore File"
            .ShowSave
        End With
        MousePointer = vbHourglass
        If ComD1.FileName = "" Then
            MousePointer = vbDefault
            Exit Sub
        End If
        If ExportYHS(ComD1.FileName) = False Then
            MousePointer = vbDefault
            MsgBox "Error Exporting HighScores File!", vbCritical, "ERROR"
            Exit Sub
        End If
        MousePointer = vbDefault
        Exit Sub
Error:
    MousePointer = vbDefault
    If Err.Number = 32755 Then Exit Sub
    MsgBox "Save Error: #" & Err.Number & " " & Error$(Err.Number), vbCritical, "Error"
End Sub

Private Sub Command4_Click()    ' Open the HighScore file in NotePad
On Error GoTo Error
    ShellExecute 0&, vbNullString, "c:\Windows\Notepad.exe", App.Path & "\YHS.bin", vbNullString, 10
    Exit Sub
Error:
    MsgBox "Error starting ""Notepad.exe"" in a Shell." & vbNewLine & _
    "Error #" & Err.Number & " - " & Error$(Err.Number), vbExclamation, "Sorry"
End Sub

Private Sub Command5_Click()    ' This TEST routine generates 20 played games.
Dim T As Integer
Dim K As Integer
    If CInt(GameScores(500, 0)) > 0 Then Exit Sub ' Exit if games played
    Randomize
    GameScores(500, 0) = 20
    For T = 1 To 20
        K = Int(Rnd * 400) + 100
        Totals(T) = K
        GameScores(T, 0) = "Test#" & T & " " & K
    Next
    For T = 1 To 20
        K = Int(Rnd * 5)
        GameScores(T, 1) = K
    Next
    Beep
    Unload frmScores
End Sub

Private Sub Command6_Click() ' Save current color configuration
    If WriteCustomColors(ColRTB(0), ColRTB(1), ColRTB(2), ColRTB(3), ColRTB(4)) = True Then
        CustomColor.Position = ColRTB(0)
        CustomColor.Score = ColRTB(1)
        CustomColor.Name = ColRTB(2)
        CustomColor.Date = ColRTB(3)
        CustomColor.YatCnt = ColRTB(4)
        CustFound = True
    Else
        MsgBox "Error saving Custom Colors!      ", vbCritical, "ERROR"
    End If
End Sub

Private Sub CallSub() 'Button Animation Control for Graph
    If GrfButt(0).Visible = True Then
        GrfButt(0).Visible = False
        GrfButt(1).Visible = True
    Else
        GrfButt(0).Visible = True
        GrfButt(1).Visible = False
    End If
    Timer1.Enabled = True
    OpenCloseGraph  'Calls Open or Close Graph
    If ShowGraph Then
        GrfButt(0).Visible = False
        GrfButt(1).Visible = True
    Else
        GrfButt(0).Visible = True
        GrfButt(1).Visible = False
    End If
End Sub

Private Sub Form_Load()
Dim T As Integer
Dim Y As Integer
Dim L As Integer
Dim NeedToSave As Boolean
Dim iScrs(100) As Integer
On Error GoTo Error
    Label4.BackStyle = vbTransparent    'Hide the MORE BUTTON Switch
    For T = 0 To 1
        GrfButt(T).Left = 240
        GrfButt(T).Top = 1960
    Next
    GrfButt(1).Visible = False
    P1.Visible = False
    P1.AutoRedraw = True
    PointCount = CInt(GameScores(500, 0))
    MaxiScore = GetMax(PointCount)
    P1.ScaleWidth = 1 + PointCount
    P1.ScaleHeight = 50 + MaxiScore
    iOPEN = False
    If PUB = False Then '      Make this FALSE to hide the MORE button - Set in Main.Load (Form1)
        Picture1.Visible = False
        Picture2.Visible = False
    Else
        Picture1.Visible = True
        Picture2.Visible = False
    End If
    If PointCount > 0 Then
        GrfButt(0).Visible = True
    Else
        GrfButt(0).Visible = False
    End If
    L = 50
    For T = 0 To 4
        Check1(T).Top = Check1(T).Top - L
        L = L - 10
    Next
    Me.Height = 4920
    If ScorePlaced = True Then
        ScorePlaced = False
        SPLtim.Enabled = True
        iPlay ("NewHighScore.wav")
    Else
        SPLtim.Enabled = False
    End If
    For T = 0 To 2
        LV(T).Caption = "0"
        HS(T).Value = 0
    Next
    Spot.BackColor = RGB(HS(0), HS(1), HS(2))
    A = RGB(HS(0), HS(1), HS(2)): Label5 = Hex(A)
    If CustFound = True Then
        ColRTB(0) = CustomColor.Position
        ColRTB(1) = CustomColor.Score
        ColRTB(2) = CustomColor.Name
        ColRTB(3) = CustomColor.Date
        ColRTB(4) = CustomColor.YatCnt
    Else
        DefCOLs
    End If
    iAverage = 0
    If GraphOpen Then
        GrfButt(0).Visible = False
        GrfButt(1).Visible = True
        iAverage = GetAVG
        P1.Left = 30
        P1.Top = 2250
        P1.Visible = True
        P1.Width = 7170
        P1.Height = 2310
        ShowGraph = True
    End If
    mIndex = 1
    UpdateColors
    NeedToSave = False  '                                       Label4 is just above date colom
    RT = ""
    If PointCount < 1 Then
        RT.Locked = False
        RT.SelStart = 0
        RT.Font.Bold = True
        RT.SelColor = vbRed
        RT.SelText = vbNewLine & vbNewLine & "              No Games Played this Session" & vbNewLine
        FillATS
        Label1 = lblHighScore
        RT.Locked = True
        Timer1.Enabled = True
        Exit Sub
    End If
    FillATS
    Label1 = lblHighScore
    RT.Locked = False
    RT.SelStart = 0
    RT.Font.Bold = False
    If Not PointCount = 1 Then
        Label2.Caption = "Scores This Session (" & GameScores(500, 0) & " games played)"
        Label3.Caption = Label2.Caption
    Else
        Label2.Caption = "Scores This Session (" & GameScores(500, 0) & " game played)"
        Label3.Caption = Label2.Caption
    End If
    For T = 1 To PointCount
        RT.SelColor = vbGreen
        If GameScores(T, 1) <> "1" Then
            RT.SelText = " " & GameScores(T, 0) & " Points" & "  (" & GameScores(T, 1) & " Yahtzees)" & vbNewLine
        Else
            RT.SelText = " " & GameScores(T, 0) & " Points" & "  (" & GameScores(T, 1) & " Yahtzee)" & vbNewLine
        End If
    Next
    If PointCount > 1 Then
        iAverage = GetAVG
        RT.SelFontSize = 4
        RT.SelText = " " & vbNewLine
        RT.SelColor = vbCyan
        RT.SelFontSize = 19
        RT.SelAlignment = 2
        RT.SelText = " Average Score = " & iAverage & " Points" & vbNewLine
    End If
    RT.SelStart = (Len(RT.Text))
    RT.Locked = True
    Timer1.Enabled = True
    Exit Sub
Error:
    If ErrorCatcher = False Then    'This prevents looping error windows
        ErrorCatcher = True
        mError = Err.Number
        MsgBox "An error occurred while looking over scores.", vbInformation, "Ken Goofed..."
    Else
        mError = Err.Number
        ErrorCatcher = False
        Unload Me
    End If
End Sub

Private Function GetAVG() As Integer
Dim T As Integer
    GetAVG = 0
    For T = 1 To PointCount
        GetAVG = GetAVG + Totals(T)
    Next
    GetAVG = (GetAVG / PointCount)
End Function

Private Sub OpenCloseGraph()
    P1.Left = 30
    P1.Top = 2250
    If P1.Visible = True Then
        iPlay ("ModClick7.wav")
        Do While P1.Width > 200     'Close
            P1.Width = P1.Width - 100
            If P1.Height > 10 Then
                P1.Height = P1.Height - 20
            Else
                P1.Height = 10
            End If
            DoEvents
        Loop
        P1.Width = 10
        P1.Height = 10
        P1.Visible = False
        ShowGraph = False
    Else
        P1.Width = 10
        P1.Height = 10
        P1.Visible = True
        iPlay ("ModClick5.wav")
        Do While P1.Width < 7169    'Open
            P1.Width = P1.Width + 65
            If P1.Height < 2310 Then
                P1.Height = P1.Height + 20
            Else
                P1.Height = 2310
            End If
            DoEvents
        Loop
        P1.Width = 7170
        P1.Height = 2310
        ShowGraph = True
    End If
End Sub

Private Sub GrfButt_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        CallSub
    Else
        iPlay ("ModClick5.wav")
        Graph.Show vbModal  'Open Big Graph Form
    End If
End Sub

Private Sub P1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        CallSub
    ElseIf Button = 2 Then
        iPlay ("ModClick5.wav")
        Graph.Show vbModal 'Open Big Graph Form
    End If
End Sub

Private Sub P1_Resize()
On Error Resume Next
    P1.ScaleWidth = 100 + (CInt(GameScores(500, 0)) * 100)
    P1.ScaleHeight = 50 + MaxiScore
    P1.Cls
    GridIt
    GraphIt
End Sub

Private Function GetMax(A As Integer) As Integer ' Get Highest Score this Session
Dim T As Integer
    GetMax = 0
    For T = 1 To A
        If Totals(T) > GetMax Then
            GetMax = Totals(T)
        End If
    Next
End Function

Private Sub GridIt() 'Draw the Background Grid of the Graph
Dim T As Integer
Dim E As Long
Dim K As Integer
Dim Divs As Integer
    Divs = Round(P1.ScaleHeight / 10, 0)
    E = Divs
    For T = 0 To 10
        P1.Line (0, E)-(P1.ScaleWidth, E), RGB(0, 0, 110)
        E = E + Divs
    Next
    P1.FontSize = 6
    P1.ForeColor = RGB(0, 255, 255)
    K = 1
    For T = 100 To (PointCount * 100) + 100 Step 100
        P1.Line (T, 0)-(T, P1.ScaleHeight), RGB(0, 0, 110)
        P1.CurrentX = (T - 20)
        If PointCount < 4 Then
            P1.CurrentX = (T - 10)
        ElseIf PointCount < 20 Then
            P1.CurrentX = (T - 20)
        Else
            P1.CurrentX = (T - 30)
        End If
        P1.CurrentY = P1.ScaleHeight - ((Divs / 4) * 3)
        If K - 1 < PointCount Then
            P1.Print K
            K = K + 1
        End If
    Next
End Sub

Public Sub GraphIt()    'Graph Gameplay for this Session
Dim T As Integer
Dim K As Integer
Dim CirSize As Integer
    K = 1
    If PointCount = 0 Then Exit Sub
    If PointCount > 1 Then
        For T = 100 To ((PointCount * 100) - 100) Step 100
            P1.Line (T, ((P1.ScaleHeight - 5) - Totals(K)))-((T + 100), ((P1.ScaleHeight - 5) - Totals(K + 1))), RGB(0, 255, 0)
            K = K + 1
        Next
    End If
    K = 90
    For T = 1 To PointCount
        With P1
            .CurrentX = K
            .CurrentY = (P1.ScaleHeight - 50) - Totals(T)
            .FontSize = 7
            .ForeColor = RGB(255, 255, 0)
        End With
        P1.Print Totals(T) & "-" & GameScores(T, 1)
        K = K + 100
    Next
    K = 1
    If PointCount < 20 Then
        CirSize = PointCount
    Else
        CirSize = PointCount - 5
    End If
    For T = 100 To (PointCount * 100) Step 100
        P1.Circle (T, ((P1.ScaleHeight - 5) - Totals(K))), CirSize, RGB(255, 0, 0)
        K = K + 1
    Next
    If PointCount > 1 Then
        With P1
            .ForeColor = RGB(255, 0, 0)
            .CurrentY = (((P1.ScaleHeight - 5) - iAverage) - ((P1.ScaleHeight / 100) * 5))
            .CurrentX = 10
            .FontSize = 6
                P1.Print "AVG"
            .CurrentY = P1.ScaleHeight - (iAverage - ((P1.ScaleHeight / 100) / 6))
            .CurrentX = 10
            .FontSize = 6
                P1.Print iAverage
            .DrawStyle = vbDot   '2
                P1.Line (50, ((P1.ScaleHeight - 5) - iAverage))-((P1.ScaleWidth - 50), ((P1.ScaleHeight - 5) - iAverage)), RGB(200, 0, 0)
            .DrawStyle = vbSolid  '0
        End With
    End If
End Sub

Public Sub FillATS()    ' Fill the HIGH SCOREs RichTextBoxes
Dim T As Integer
    RS.Locked = False
    RS.Text = ""
    RS.SelStart = 0
    RD.Locked = False
    RD.Text = ""
    RD.SelStart = 0
    RF.Locked = False
    RF.Text = ""
    RF.SelStart = 0
    Rn.Locked = False
    Rn.Text = ""
    Rn.SelStart = 0
    RG.Locked = False
    RG.Text = ""
    RG.SelStart = 0
    For T = 1 To 5
        RS.Font.Bold = True
        RS.SelColor = ColRTB(0)
        If T < 5 Then
            RS.SelText = " " & T & " :" & vbNewLine
        Else
            RS.SelText = " " & T & " :"
        End If
        RD.Font.Bold = True
        RD.SelColor = ColRTB(1)
        If ATS(T).Score > 0 Then
            RD.SelText = CStr(ATS(T).Score) & vbNewLine
        Else
            RD.SelText = " " & vbNewLine
        End If
        RF.Font.Bold = True
        RF.SelColor = ColRTB(2)
        RF.SelText = ATS(T).Name & vbNewLine
        Check1(T - 1).Caption = ATS(T).Name
        Rn.Font.Bold = True
        Rn.SelColor = ColRTB(4)
        If ATS(T).Name <> "" Then
            Rn.SelText = ATS(T).Qnt & vbNewLine
        Else
            Rn.SelText = vbNewLine
        End If
        RG.Font.Bold = False
        RG.SelColor = ColRTB(3)
        RG.SelText = ATS(T).Date & vbNewLine
    Next
    RS.Locked = True
    RD.Locked = True
    RF.Locked = True
    RG.Locked = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If iOPEN Then
        iExit = True
        Cancel = 1
        On Error Resume Next
        iPlay ("OpnCls.wav")
        CloseTIM.Enabled = True
        iEnding = True
        Exit Sub
    End If
    If iEnding = False Then
        iPlay ("ModClick1.wav")
    Else
        iEnding = False
    End If
    If P1.Visible = True Then
        GraphOpen = True
    Else
        GraphOpen = False
    End If
End Sub

Private Sub GrfButton_Click()
    Graph.Show vbModal
End Sub

Private Sub HS_Change(Index As Integer)
    LV(Index).Caption = Hex(HS(Index).Value)
    Spot.BackColor = RGB(HS(0), HS(1), HS(2))
    A = RGB(HS(0), HS(1), HS(2)): Label5 = Hex(A)
End Sub

Private Sub HS_Scroll(Index As Integer)
    LV(Index).Caption = Hex(HS(Index).Value)
    Spot.BackColor = RGB(HS(0), HS(1), HS(2))
    A = RGB(HS(0), HS(1), HS(2)): Label5 = Hex(A)
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) ' Show DATEs
    If Button = 2 Then
        iPlay ("ModClick2.wav")
        If PUB = False Then
            PUB = True
            If Me.Height > 4930 Then
                Picture2.Visible = True
                Picture1.Visible = False
            Else
                Picture1.Visible = True
                Picture2.Visible = False
            End If
            If GameScores(500, 0) > 2 Then
                GrfButton.Visible = True
            Else
                GrfButton.Visible = False
            End If
        Else
            PUB = False
            Picture1.Visible = False
            Picture2.Visible = False
            GrfButton.Visible = False
        End If
    End If
End Sub

Private Sub MergeFile_Click()
On Error GoTo Error
    MousePointer = vbDefault
    With ComD1
        .InitDir = App.Path
        .Flags = &H2
        .DefaultExt = "bin"
        .Filter = "Yahtzee HighScore Files (*.bin)|*.bin"
        .DialogTitle = "Merge a HighScore File into Mine"
        .Flags = &H4
        .Flags = &H1000
        .ShowOpen
    End With
    MousePointer = vbHourglass
    If ComD1.FileName = "" Then
        MousePointer = vbDefault
        Exit Sub
    End If
    If OpenYHSmrg(ComD1.FileName) = False Then
        MousePointer = vbDefault
        MsgBox "Error Importing Yahtzee HighScore Merge File. " & vbNewLine & _
        "ERROR #" & mError & " - " & Error$(mError) & vbNewLine & _
        " " & vbNewLine & "Data appears currupt.", vbCritical, "Merge File Error"
        Exit Sub
    End If
    MergeEm
    SortEm
    FillATS
    SaveHighScore
    Scores.HighScore = GetHighScore
    SavePath = ComD1.FileName
    frmMain.Caption = " Solo Yahtzee                         High Score = " & Scores.HighScore & _
    " by " & Scores.HSHolder
    MousePointer = vbDefault
    Exit Sub
Error:
    MousePointer = vbDefault
    If Err.Number = 32755 Then Exit Sub
    MousePointer = vbDefault
    MsgBox "Error Importing Yahtzee HighScore File. " & vbNewLine & _
        "ERROR #" & mError & " - " & Error$(mError) & vbNewLine & _
        " " & vbNewLine & "Data appears currupt.", vbCritical, "Merge File Error"
    On Error Resume Next
        Close #1
End Sub

Private Sub OpenTIM_Timer()
    If Me.Height < 7230 Then
        Me.Height = Me.Height + 75
        Me.Top = Me.Top - (75 / 2)
    Else
        OpenTIM.Enabled = False
        Me.Height = 7230
        iOPEN = True
    End If
End Sub

Private Sub CloseTIM_Timer()
    If Me.Height > 4920 Then
        Me.Height = Me.Height - 75
        Me.Top = Me.Top + (75 / 2)
    Else
        CloseTIM.Enabled = False
        Me.Height = 4920
        iOPEN = False
        If iExit Then
            iExit = False
            Unload Me
        End If
    End If
End Sub

Private Sub Option1_Click(Index As Integer)
    mIndex = Index
    UpdateColors
End Sub

Public Sub UpdateColors()
    GetColors (ColRTB(mIndex))
    HS(0).Value = CL.Red
    HS(1).Value = CL.Green
    HS(2).Value = CL.Blue
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SetTIMs
        Timer1.Enabled = True
    End If
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        SetTIMs
        Timer1.Enabled = True
    End If
End Sub

Private Sub RD_Click()
    Unload Me
End Sub

Private Sub RF_Click()
    Unload Me
End Sub

Private Sub RG_Click()
    Unload Me
End Sub

Private Sub RS_Click()
    Unload Me
End Sub

Private Sub RT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Unload Me
    Else
        If PointCount > 0 Then
            CallSub
        Else
            Unload Me
        End If
    End If
End Sub

Private Sub SH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) ' Expand Window
    If Button = 2 Then
        SetTIMs
        Timer1.Enabled = True
    End If
End Sub

Private Sub SetTIMs()
On Error Resume Next
    iPlay ("OpnCls.wav")
    If Me.Height < 7229 Then
        CloseTIM.Enabled = False
        OpenTIM.Enabled = True
        If PUB Then
            Picture2.Visible = True
            Picture1.Visible = False
        End If
    Else
        If PUB Then
            Picture1.Visible = True
            Picture2.Visible = False
        End If
        OpenTIM.Enabled = False
        CloseTIM.Enabled = True
    End If
    If GameScores(500, 0) > 2 And PUB = True Then
        GrfButton.Visible = True
    Else
        GrfButton.Visible = False
    End If
End Sub

Private Sub SPLtim_Timer()  'If a new High Score has been added, flash the position number
Dim T As Integer
    RS.Locked = False
    RS.Text = ""
    RS.SelStart = 0
    For T = 1 To 5
        If T = NHSplace Then
            RS.Font.Bold = True
            If ColNS = True Then
                ColNS = False
                RS.SelColor = vbYellow
            Else
                ColNS = True
                RS.SelColor = vbBlack
            End If
            RS.SelText = " " & T & " :" & vbNewLine
        Else
            RS.Font.Bold = True
            RS.SelColor = vbGreen
            RS.SelText = " " & T & " :" & vbNewLine
        End If
    Next
    RS.Locked = True
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Unload Me
    End If
End Sub

Private Sub Timer1_Timer()  'This timer moves the cursor out of the score window. Aesthetics!
    Timer1.Enabled = False
    On Error Resume Next
    Text1.SetFocus
End Sub

Private Sub Command7_Click()    'Import HighScore File
On Error GoTo Error
    MousePointer = vbDefault
    With ComD1
        .InitDir = App.Path
        .Flags = &H2
        .DefaultExt = "bin"
        .Filter = "Yahtzee HighScore Files (*.bin)|*.bin"
        .DialogTitle = "Import HighScore File"
        .Flags = &H4
        .Flags = &H1000
        .ShowOpen
    End With
    MousePointer = vbHourglass
    If ComD1.FileName = "" Then
        MousePointer = vbDefault
        Exit Sub
    End If
    If OpenYHS(ComD1.FileName) = False Then
        MousePointer = vbDefault
        MsgBox "Error Importing Yahtzee HighScore File. " & vbNewLine & _
        "ERROR #" & mError & " - " & Error$(mError) & vbNewLine & _
        " " & vbNewLine & "Data appears currupt.", vbCritical, "Import File Error"
        Exit Sub
    End If
    SortEm
    FillATS
    SaveHighScore
    Scores.HighScore = GetHighScore
    SavePath = ComD1.FileName
    frmMain.Caption = " Solo Yahtzee                         High Score = " & Scores.HighScore & _
    " by " & Scores.HSHolder
    MousePointer = vbDefault
    Exit Sub
Error:
    MousePointer = vbDefault
    If Err.Number = 32755 Then Exit Sub
    MousePointer = vbDefault
    MsgBox "Error Importing Yahtzee HighScore File. " & vbNewLine & _
        "ERROR #" & mError & " - " & Error$(mError) & vbNewLine & _
        " " & vbNewLine & "Data appears currupt.", vbCritical, "Import File Error"
    On Error Resume Next
        Close #1
End Sub
