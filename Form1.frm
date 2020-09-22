VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solo Yahtzee II    by 0x34"
   ClientHeight    =   3870
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   3870
   ScaleWidth      =   9120
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer BonSNDTim 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   6600
      Top             =   1080
   End
   Begin VB.Timer SHTim 
      Enabled         =   0   'False
      Interval        =   60
      Left            =   6600
      Top             =   600
   End
   Begin VB.Timer MenuNotifyTIM 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6600
      Top             =   120
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   12
      Left            =   5160
      Picture         =   "Form1.frx":9A58E
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   120
      Top             =   2160
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   11
      Left            =   4920
      Picture         =   "Form1.frx":9A690
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   119
      Top             =   2160
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   10
      Left            =   4680
      Picture         =   "Form1.frx":9A792
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   118
      Top             =   2160
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   9
      Left            =   4440
      Picture         =   "Form1.frx":9A894
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   117
      Top             =   2160
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   8
      Left            =   4200
      Picture         =   "Form1.frx":9A996
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   116
      Top             =   2160
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   7
      Left            =   5880
      Picture         =   "Form1.frx":9AA98
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   115
      Top             =   1920
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   6
      Left            =   5640
      Picture         =   "Form1.frx":9AB9A
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   114
      Top             =   1920
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   5
      Left            =   5400
      Picture         =   "Form1.frx":9AC9C
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   113
      Top             =   1920
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   4
      Left            =   5160
      Picture         =   "Form1.frx":9AD9E
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   112
      Top             =   1920
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   3
      Left            =   4920
      Picture         =   "Form1.frx":9AEA0
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   111
      Top             =   1920
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   2
      Left            =   4680
      Picture         =   "Form1.frx":9AFA2
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   110
      Top             =   1920
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   1
      Left            =   4440
      Picture         =   "Form1.frx":9B0A4
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   109
      Top             =   1920
      Width           =   200
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   200
      Index           =   0
      Left            =   4200
      Picture         =   "Form1.frx":9B1A6
      ScaleHeight     =   195
      ScaleWidth      =   195
      TabIndex        =   108
      Top             =   1920
      Width           =   200
   End
   Begin VB.Timer SpTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6120
      Top             =   120
   End
   Begin VB.PictureBox Chance 
      Height          =   495
      Index           =   2
      Left            =   3360
      Picture         =   "Form1.frx":9B2A8
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   105
      ToolTipText     =   "Chance"
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox Chance 
      Height          =   495
      Index           =   1
      Left            =   3360
      Picture         =   "Form1.frx":9BCE4
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   104
      ToolTipText     =   "Chance"
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox LgStr 
      Height          =   495
      Index           =   2
      Left            =   2760
      Picture         =   "Form1.frx":9C71E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   103
      ToolTipText     =   "Large Straight"
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox LgStr 
      Height          =   495
      Index           =   1
      Left            =   2760
      Picture         =   "Form1.frx":9D15A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   102
      ToolTipText     =   "Large Straight"
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox SmStr 
      Height          =   495
      Index           =   2
      Left            =   2160
      Picture         =   "Form1.frx":9DB96
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   101
      ToolTipText     =   "Small Straight"
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox SmStr 
      Height          =   495
      Index           =   1
      Left            =   2160
      Picture         =   "Form1.frx":9E5D2
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   100
      ToolTipText     =   "Small Straight"
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox FullH 
      Height          =   495
      Index           =   2
      Left            =   1560
      Picture         =   "Form1.frx":9F00E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   99
      ToolTipText     =   "Full House"
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox FullH 
      Height          =   495
      Index           =   1
      Left            =   1560
      Picture         =   "Form1.frx":9FA4A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   98
      ToolTipText     =   "Full House"
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox FourK 
      Height          =   495
      Index           =   2
      Left            =   960
      Picture         =   "Form1.frx":A0486
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   97
      ToolTipText     =   "Four of a Kind"
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox FourK 
      Height          =   495
      Index           =   1
      Left            =   960
      Picture         =   "Form1.frx":A0EC2
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   96
      ToolTipText     =   "Four of a Kind"
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox ThreeK 
      Height          =   495
      Index           =   2
      Left            =   360
      Picture         =   "Form1.frx":A18FE
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   95
      ToolTipText     =   "Three of a Kind"
      Top             =   1560
      Width           =   495
   End
   Begin VB.PictureBox ThreeK 
      Height          =   495
      Index           =   1
      Left            =   360
      Picture         =   "Form1.frx":A233A
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   94
      ToolTipText     =   "Three of a Kind"
      Top             =   1440
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   18
      Left            =   3360
      Picture         =   "Form1.frx":A2D76
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   93
      Top             =   600
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   17
      Left            =   2760
      Picture         =   "Form1.frx":A358A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   92
      Top             =   600
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   16
      Left            =   2160
      Picture         =   "Form1.frx":A3D9E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   91
      Top             =   600
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   15
      Left            =   1560
      Picture         =   "Form1.frx":A45B2
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   90
      Top             =   600
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   14
      Left            =   960
      Picture         =   "Form1.frx":A4DC6
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   89
      Top             =   600
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   13
      Left            =   360
      Picture         =   "Form1.frx":A55DA
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   88
      Top             =   600
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   12
      Left            =   3360
      Picture         =   "Form1.frx":A5DEE
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   87
      Top             =   480
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   11
      Left            =   2760
      Picture         =   "Form1.frx":A6602
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   86
      Top             =   480
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   10
      Left            =   2160
      Picture         =   "Form1.frx":A6E16
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   85
      Top             =   480
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   9
      Left            =   1560
      Picture         =   "Form1.frx":A762A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   84
      Top             =   480
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   8
      Left            =   960
      Picture         =   "Form1.frx":A7E3E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   83
      Top             =   480
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   7
      Left            =   360
      Picture         =   "Form1.frx":A8652
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   82
      Top             =   480
      Width           =   442
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   225
      Left            =   6240
      TabIndex        =   77
      Top             =   1440
      Visible         =   0   'False
      Width           =   200
   End
   Begin VB.Timer FlashTIM 
      Enabled         =   0   'False
      Interval        =   80
      Left            =   6120
      Top             =   600
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   360
      Picture         =   "Form1.frx":A8E66
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   70
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   1680
      Picture         =   "Form1.frx":AFBA4
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   69
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   5640
      Picture         =   "Form1.frx":B68E2
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   68
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   3000
      Picture         =   "Form1.frx":BD620
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   67
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   7
      Left            =   4320
      Picture         =   "Form1.frx":C435E
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   66
      Top             =   2640
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   360
      Picture         =   "Form1.frx":CB09C
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   65
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   360
      Picture         =   "Form1.frx":CDCB6
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   64
      Top             =   3840
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   360
      Picture         =   "Form1.frx":D08D0
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   63
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   360
      Picture         =   "Form1.frx":D34EA
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   62
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   360
      Picture         =   "Form1.frx":D6104
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   61
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000007&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   360
      Picture         =   "Form1.frx":D8D1E
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   60
      Top             =   2880
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   1680
      Picture         =   "Form1.frx":DB938
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   59
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   1680
      Picture         =   "Form1.frx":DE552
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   58
      Top             =   3840
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   1680
      Picture         =   "Form1.frx":E116C
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   57
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   1680
      Picture         =   "Form1.frx":E3D86
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   56
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   1680
      Picture         =   "Form1.frx":E69A0
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   55
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   1680
      Picture         =   "Form1.frx":E95BA
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   54
      Top             =   2880
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   3000
      Picture         =   "Form1.frx":EC1D4
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   51
      Top             =   4200
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   3000
      Picture         =   "Form1.frx":EEDEE
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   50
      Top             =   3960
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   3000
      Picture         =   "Form1.frx":F1A08
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   53
      Top             =   3720
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   3000
      Picture         =   "Form1.frx":F4622
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   52
      Top             =   3480
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   3000
      Picture         =   "Form1.frx":F723C
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   49
      Top             =   3240
      Width           =   975
   End
   Begin VB.PictureBox Picture3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   3000
      Picture         =   "Form1.frx":F9E56
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   48
      Top             =   3000
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   4320
      Picture         =   "Form1.frx":FCA70
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   47
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   4320
      Picture         =   "Form1.frx":FF68A
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   46
      Top             =   3840
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   4320
      Picture         =   "Form1.frx":1022A4
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   45
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   4320
      Picture         =   "Form1.frx":104EBE
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   44
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   4320
      Picture         =   "Form1.frx":107AD8
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   43
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   4320
      Picture         =   "Form1.frx":10A6F2
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   42
      Top             =   2880
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   6
      Left            =   5640
      Picture         =   "Form1.frx":10D30C
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   41
      Top             =   4080
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   5
      Left            =   5640
      Picture         =   "Form1.frx":10FF26
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   40
      Top             =   3840
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   5640
      Picture         =   "Form1.frx":112B40
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   39
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   3
      Left            =   5640
      Picture         =   "Form1.frx":11575A
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   38
      Top             =   3360
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   2
      Left            =   5640
      Picture         =   "Form1.frx":118374
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   37
      Top             =   3120
      Width           =   975
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   1
      Left            =   5640
      Picture         =   "Form1.frx":11AF8E
      ScaleHeight     =   915
      ScaleWidth      =   915
      TabIndex        =   36
      Top             =   2880
      Width           =   975
   End
   Begin VB.PictureBox Yatz 
      Height          =   460
      Index           =   2
      Left            =   4200
      Picture         =   "Form1.frx":11DBA8
      ScaleHeight     =   405
      ScaleWidth      =   1755
      TabIndex        =   35
      ToolTipText     =   "Yahtzee"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox Yatz 
      Height          =   460
      Index           =   1
      Left            =   4200
      Picture         =   "Form1.frx":1204B4
      ScaleHeight     =   405
      ScaleWidth      =   1755
      TabIndex        =   34
      ToolTipText     =   "Yahtzee"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.PictureBox Yatz 
      Height          =   460
      Index           =   0
      Left            =   4200
      Picture         =   "Form1.frx":122DC0
      ScaleHeight     =   405
      ScaleWidth      =   1755
      TabIndex        =   33
      ToolTipText     =   "Yahtzee"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Timer RollTIM 
      Enabled         =   0   'False
      Interval        =   40
      Left            =   4200
      Top             =   6360
   End
   Begin VB.Timer YatzTIM 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4680
      Top             =   6360
   End
   Begin VB.PictureBox ThreeK 
      Height          =   495
      Index           =   0
      Left            =   360
      Picture         =   "Form1.frx":1256CA
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   31
      ToolTipText     =   "Three of a Kind"
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox FourK 
      Height          =   495
      Index           =   0
      Left            =   960
      Picture         =   "Form1.frx":126104
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   29
      ToolTipText     =   "Four of a Kind"
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox FullH 
      Height          =   495
      Index           =   0
      Left            =   1560
      Picture         =   "Form1.frx":126B3E
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   27
      ToolTipText     =   "Full House"
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox SmStr 
      Height          =   495
      Index           =   0
      Left            =   2160
      Picture         =   "Form1.frx":127578
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   25
      ToolTipText     =   "Small Straight"
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox LgStr 
      Height          =   495
      Index           =   0
      Left            =   2760
      Picture         =   "Form1.frx":127FB2
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   23
      ToolTipText     =   "Large Straight"
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox Chance 
      Height          =   495
      Index           =   0
      Left            =   3360
      Picture         =   "Form1.frx":1289EC
      ScaleHeight     =   435
      ScaleWidth      =   435
      TabIndex        =   21
      ToolTipText     =   "Chance"
      Top             =   1320
      Width           =   495
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   1
      Left            =   360
      Picture         =   "Form1.frx":129426
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   16
      Top             =   360
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   2
      Left            =   960
      Picture         =   "Form1.frx":129C38
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   14
      Top             =   360
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   3
      Left            =   1560
      Picture         =   "Form1.frx":12A44A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   12
      Top             =   360
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   4
      Left            =   2160
      Picture         =   "Form1.frx":12AC5C
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   10
      Top             =   360
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   5
      Left            =   2760
      Picture         =   "Form1.frx":12B46E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   8
      Top             =   360
      Width           =   442
   End
   Begin VB.PictureBox Picture6 
      Height          =   440
      Index           =   6
      Left            =   3360
      Picture         =   "Form1.frx":12BC80
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   6
      Top             =   360
      Width           =   442
   End
   Begin VB.CommandButton RollDice 
      Caption         =   "Right Click for Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton AllScores 
      Height          =   225
      Left            =   8880
      TabIndex        =   2
      Top             =   720
      Width           =   135
   End
   Begin VB.Timer OverTIM 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   3240
      Top             =   6360
   End
   Begin VB.Timer JokerTIM 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3720
      Top             =   6360
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Menu"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   255
      Left            =   120
      TabIndex        =   122
      Top             =   30
      Width           =   615
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5880
      TabIndex        =   106
      Top             =   2040
      Width           =   615
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   3
      Left            =   8880
      Top             =   2400
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   2
      Left            =   8880
      Top             =   2520
      Width           =   135
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Opaque
      Height          =   135
      Index           =   1
      Left            =   8880
      Top             =   2640
      Width           =   135
   End
   Begin VB.Label TESTBox 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   76
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   5
      Left            =   5520
      TabIndex        =   75
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   4
      Left            =   4200
      TabIndex        =   74
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   3
      Left            =   2880
      TabIndex        =   73
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   2
      Left            =   1560
      TabIndex        =   72
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00004040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1215
      Index           =   1
      Left            =   240
      TabIndex        =   71
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label BottomScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   1
      Left            =   390
      TabIndex        =   32
      Top             =   1860
      Width           =   405
   End
   Begin VB.Label BottomScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   2
      Left            =   990
      TabIndex        =   30
      Top             =   1860
      Width           =   405
   End
   Begin VB.Label BottomScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   1590
      TabIndex        =   28
      Top             =   1860
      Width           =   405
   End
   Begin VB.Label BottomScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   4
      Left            =   2190
      TabIndex        =   26
      Top             =   1860
      Width           =   405
   End
   Begin VB.Label BottomScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   5
      Left            =   2790
      TabIndex        =   24
      Top             =   1860
      Width           =   405
   End
   Begin VB.Label BottomScore 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   6
      Left            =   3390
      TabIndex        =   22
      Top             =   1860
      Width           =   405
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "JOKER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   4200
      TabIndex        =   20
      Top             =   840
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label LabBONUS 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "35 Point  BONUS!"
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
      Left            =   4200
      TabIndex        =   19
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Top Score"
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
      Left            =   4200
      TabIndex        =   18
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label TopScores 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   1
      Left            =   360
      TabIndex        =   17
      Top             =   840
      Width           =   435
   End
   Begin VB.Label TopScores 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   15
      Top             =   840
      Width           =   435
   End
   Begin VB.Label TopScores 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   3
      Left            =   1560
      TabIndex        =   13
      Top             =   840
      Width           =   435
   End
   Begin VB.Label TopScores 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   4
      Left            =   2160
      TabIndex        =   11
      Top             =   840
      Width           =   435
   End
   Begin VB.Label TopScores 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   5
      Left            =   2760
      TabIndex        =   9
      Top             =   840
      Width           =   435
   End
   Begin VB.Label TopScores 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "88"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   285
      Index           =   6
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "On Roll #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   4
      Top             =   2420
      Width           =   1935
   End
   Begin VB.Label EndLABEL 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "GAME      OVER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   7320
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label finTOT 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   525
      Left            =   7080
      TabIndex        =   1
      Top             =   600
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Score:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Top Score"
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
      Left            =   4220
      TabIndex        =   78
      Top             =   270
      Width           =   1695
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "On Roll #"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7100
      TabIndex        =   79
      Top             =   2450
      Width           =   1935
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "GAME      OVER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   7350
      TabIndex        =   80
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "35 Point  BONUS!"
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
      Left            =   4220
      TabIndex        =   81
      Top             =   623
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5920
      TabIndex        =   107
      Top             =   2080
      Width           =   615
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Total Score:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   7080
      TabIndex        =   121
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu Hide1 
      Caption         =   "HiddenMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuNewGAME 
         Caption         =   "New Game"
      End
      Begin VB.Menu mnuspace 
         Caption         =   "-"
      End
      Begin VB.Menu mnuspeed 
         Caption         =   "Speed Yahtzee"
      End
      Begin VB.Menu mnuSetTime 
         Caption         =   "Set Speed Challenge Time"
      End
      Begin VB.Menu mnuDiv3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSpotHelp 
         Caption         =   "Move Assistant"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuRand 
         Caption         =   "Random Roll Order"
      End
      Begin VB.Menu mnuSound 
         Caption         =   "Sounds"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSndTyps 
         Caption         =   "Sound Types"
         Begin VB.Menu mnuNormal 
            Caption         =   "Normal Sounds"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuFight 
            Caption         =   "Fight Sounds"
         End
      End
      Begin VB.Menu mnuDiv5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuScores 
         Caption         =   "Scores"
      End
      Begin VB.Menu mnuDivider2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Solo Yahtzee"
      End
      Begin VB.Menu mnuspacer5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Solo Yahtzee - (AKA Shaspeehtzee) -  Version 3.00.83  -  Source Code released on 011610


'                                       Programming By
'                              Ken Slater 2007/2008/2009/2010
'                                            0x34
'                                      (My Pet Project)

'                                 SPEED Yahtzee enhancement:
'               During game play, right click on the main FORM to access the menu.

'                                    Sound Effects Added
'                                      in February 2009

'                                Enjoy this very addicting game!
'                                Original game concept by HASBRO


'Max/Best possible Score (all yahtzees in six's) = 1575
'My personal best score = 624 (4 Yahtzees)

Dim LG As Boolean
Dim ASH As Integer

Private Sub AllScores_Click()   'Open Scores Screen
On Error GoTo Error
    iPlay ("ModClick2.wav")
    frmScores.Show vbModal
    Exit Sub
Error:
    MsgBox "Error while opening the HighScore Screen.        " & vbNewLine & _
        "ERROR #" & mError & " - " & Error$(mError), vbCritical, "Program Error"
End Sub

Private Sub FillLast()  'If this is the last one to be filled, do it automatically.
Dim T As Integer
    For T = 1 To 6
        If TopScores(T) = "" Then
            Call Picture6_Click(T)
            Exit Sub
        End If
    Next
    If BottomScore(1) = "" Then Call ThreeK_Click(0): Exit Sub
    If BottomScore(2) = "" Then Call FourK_Click(0): Exit Sub
    If BottomScore(3) = "" Then Call FullH_Click(0): Exit Sub
    If BottomScore(4) = "" Then Call SmStr_Click(0): Exit Sub
    If BottomScore(5) = "" Then Call LgStr_Click(0): Exit Sub
    If BottomScore(6) = "" Then Call Chance_Click(0): Exit Sub
    Call HandleYahtzee(0)
End Sub

Private Sub BonSNDTim_Timer()   'Timer to play sound when BONUS is achieved
    BonSNDTim.Enabled = False
    If GameOver = True Then
        iPlay ("EndGame.WAV")
    Else
        iPlay ("Bonus.WAV")
    End If
End Sub

Private Sub Chance_Click(Index As Integer) ' Chance Button
Dim aTOT As Integer
Dim T As Integer
Dim RESP As Long
    If Roll = 0 Then Exit Sub
    If Label2 = "Rolling..." Then Exit Sub
    If DiePlayed(12) = True Then Exit Sub
    DiePlayed(12) = True
    BottomScore(6).ForeColor = vbYellow
    YatzTIM.Enabled = False
    For T = 0 To 2
        Chance(T).ToolTipText = ""
    Next
    SelectSlp
    If Joker Then
        JokerHandler
        RESP = Index * 5
        KillYahtzee = False
    Else
        KillYahtzee = True
    End If
    If Yahtzee = True And KillYahtzee = False Then
        Call HandleYahtzee(1)
        Exit Sub
    Else
        KillYahtzee = False
    End If
    aTOT = DieSTAT(1) + DieSTAT(2) + DieSTAT(3) + DieSTAT(4) + DieSTAT(5)
    BottomScore(6) = aTOT
    Scores.Bottom = Scores.Bottom + aTOT
    Tally
End Sub

Private Sub Chance_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Right$((CStr(Label2)), 1) = "1" Then Exit Sub
    If RollDice.BackColor = &HFFFF& Then Exit Sub
    If FixEm = True And Chance(2).Visible = False Then
        CleanUp
    End If
    FixEm = True
    If BottomScore(6) = "" Then
        Chance(2).Visible = True
        BottomScore(6).ForeColor = vbGreen
        BottomScore(6).Caption = DieSTAT(1) + DieSTAT(2) + DieSTAT(3) + DieSTAT(4) + DieSTAT(5)
    End If
End Sub

Private Sub EndLABEL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ShowMenu
    End If
End Sub

Private Sub FlashTIM_Timer()
    If RollDice.BackColor = &H8000000F Then
        RollDice.BackColor = vbRed
    Else
        RollDice.BackColor = &H8000000F
    End If
    FlashCounter = FlashCounter + 1
    If FlashCounter > 2 Then
        RollDice.BackColor = &H8000000F
        FlashTIM.Enabled = False
        FlashCounter = 0
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ShowMenu
    End If
End Sub

Public Sub ShowMenu()
    iPlay ("ModClick5.wav")
    PopupMenu Hide1
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim T As Integer
    If SHTim.Enabled = True Then Exit Sub
    If FixEm Then
        FixEm = False
        CleanUp
    End If
End Sub

Private Sub CleanUp()
Dim T As Integer
    For T = 1 To 18
        If T < 7 Then
            If DiePlayed(T) = False Then
                TopScores(T).Caption = ""
            End If
            If DiePlayed(T + 6) = False Then
                BottomScore(T).Caption = ""
            End If
        End If
        If T < 7 Then
            Picture6(T).Visible = True
            TopScores(T).ForeColor = vbYellow
            BottomScore(T).ForeColor = vbYellow
        Else
            Picture6(T).Visible = False
        End If
    Next
    For T = 1 To 2
        ThreeK(T).Visible = False
        FourK(T).Visible = False
        FullH(T).Visible = False
        SmStr(T).Visible = False
        LgStr(T).Visible = False
        Chance(T).Visible = False
    Next
    If mnuSpotHelp = True Then
        For T = 1 To 6
            If SptHlp(T) Then Picture6(T + 12).Visible = True
        Next
        If SptHlp(7) Then ThreeK(2).Visible = True
        If SptHlp(8) Then FourK(2).Visible = True
        If SptHlp(9) Then FullH(2).Visible = True
        If SptHlp(10) Then SmStr(2).Visible = True
        If SptHlp(11) Then LgStr(2).Visible = True
        If SptHlp(12) Then Chance(2).Visible = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    SaveHighScore
    Unload frmScores
    Unload NHS
    Unload AboutScrn
    Unload SpeedSet
End Sub

Private Sub FourK_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Right$((CStr(Label2)), 1) = "1" Then Exit Sub
    If RollDice.BackColor = &HFFFF& Then Exit Sub
    If FixEm = True And (FourK(2).Visible = False And FourK(1).Visible = False) Then
        CleanUp
    End If
    If BottomScore(2) = "" Then
        FixEm = True
        If FourKind Or Joker = True Then
            FourK(2).Visible = True
            BottomScore(2).ForeColor = vbGreen
            BottomScore(2).Caption = DieSTAT(1) + DieSTAT(2) + DieSTAT(3) + DieSTAT(4) + DieSTAT(5)
        Else
            FourK(1).Visible = True
            BottomScore(2).ForeColor = vbRed
            BottomScore(2).Caption = "0"
        End If
    End If
End Sub

Private Sub FullH_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Right$((CStr(Label2)), 1) = "1" Then Exit Sub
    If RollDice.BackColor = &HFFFF& Then Exit Sub
    If FixEm = True And (FullH(2).Visible = False And FullH(1).Visible = False) Then
        CleanUp
    End If
    If BottomScore(3) = "" Then
        FixEm = True
        If FullHouse Or Joker = True Then
            FullH(2).Visible = True
            BottomScore(3).ForeColor = vbGreen
            BottomScore(3).Caption = "25"
        Else
            FullH(1).Visible = True
            BottomScore(3).ForeColor = vbRed
            BottomScore(3).Caption = "0"
        End If
    End If
End Sub

Private Sub JokerTIM_Timer() 'Flash the JOKER label
    If Label5.Visible = True Then
        Label5.Visible = False
    Else
        Label5.Visible = True
    End If
End Sub

Private Sub LabBONUS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ShowMenu
    End If
End Sub

Private Sub Label14_Click()
    ShowMenu
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ShowMenu
    End If
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ShowMenu
    End If
End Sub

Private Sub Label5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ShowMenu
    End If
End Sub

Private Sub mnuRand_Click()
    iPlay ("ModClick7.wav")
    If mnuRand.Checked = True Then
        mnuRand.Checked = False
        Pref.RndStop = 0
    Else
        mnuRand.Checked = True
        Pref.RndStop = 1
    End If
End Sub

Private Sub TESTBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then  'Turns on/off TEST MODE
        iPlay ("ModClick2.wav")
        If Check1.Visible = True Then
            Check1.Visible = False
        Else
            Check1.Visible = True
            iTEST = True
        End If
    End If
End Sub

Private Sub LgStr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Right$((CStr(Label2)), 1) = "1" Then Exit Sub
    If RollDice.BackColor = &HFFFF& Then Exit Sub
    If FixEm = True And (LgStr(2).Visible = False And LgStr(1).Visible = False) Then
        CleanUp
    End If
    If BottomScore(5) = "" Then
        FixEm = True
        If LgStreight Or Joker = True Then
            LgStr(2).Visible = True
            BottomScore(5).ForeColor = vbGreen
            BottomScore(5).Caption = "40"
        Else
            LgStr(1).Visible = True
            BottomScore(5).ForeColor = vbRed
            BottomScore(5).Caption = "0"
        End If
    End If
End Sub

Private Sub MenuNotifyTIM_Timer()
    MenuNotifyTIM.Enabled = False
    RollDice.Caption = "Roll Dice"
    RollDice.FontSize = 20
End Sub

Private Sub mnuAbout_Click()
    iPlay ("About.wav")
    AboutScrn.Show vbModal
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFight_Click()
    If mnuNormal.Checked = True Then
        mnuNormal.Checked = False
        mnuFight.Checked = True
        FightSND = True
        Pref.SndType = 1
        iPlay ("PUNCH5.WAV")
    End If
End Sub

Private Sub mnuNewGAME_Click()
    Call NewGameSet
End Sub

Private Sub mnuNormal_Click()
    If mnuNormal.Checked = False Then
        mnuNormal.Checked = True
        mnuFight.Checked = False
        FightSND = False
        Pref.SndType = 0
        iPlay ("ModClick3.wav")
    End If
End Sub

Private Sub mnuScores_Click()
    AllScores_Click
End Sub

Private Sub mnuSetTime_Click()
    iPlay ("ModClick6.wav")
    SpeedSet.Show vbModal
End Sub

Private Sub mnuSound_Click()
    If mnuSound.Checked = True Then
        mnuSound.Checked = False
        mnuSndTyps.Enabled = False
        Pref.Sounds = 0
    Else
        mnuSound.Checked = True
        mnuSndTyps.Enabled = True
        iPlay ("ModClick7.wav")
        Pref.Sounds = 1
    End If
End Sub

Private Sub mnuspeed_Click()
    If mnuspeed.Checked = True Then
        mnuspeed.Checked = False
        Pref.SpeedYaht = 0
    Else
        mnuspeed.Checked = True
        Pref.SpeedYaht = 1
    End If
    iPlay ("ModClick7.wav")
End Sub

Private Sub NewGame_Click()
    Call NewGameSet
End Sub

Private Sub NewGameSet()    'RESET / New Game
Dim T As Integer
    If RollTIM.Enabled Then Exit Sub
    If iInit = True Then
        iPlay ("ModClick4.wav")
    Else
        iPlay ("StartUp.wav")
    End If
    iInit = True
    AmtYaht = 0
    If SpTimer.Enabled = True Then
        SpTimer.Enabled = False
        Label11 = ""
        Label12 = ""
        SpDial = 0
    End If
    RollDice.Caption = "Roll Dice"
    Yahtzeed = False
    Yatz(2).Visible = False
    Yatz(1).Visible = False
    YatzTIM.Enabled = False
    Yatz(0).Visible = True
    GameOver = False
    OverTIM.Enabled = False
    BonSNDTim.Enabled = False
    EndLABEL.Visible = False
    Label9.Visible = False
    For T = 1 To 2
        ThreeK(T).Visible = False
        FourK(T).Visible = False
        FullH(T).Visible = False
        SmStr(T).Visible = False
        LgStr(T).Visible = False
        Chance(T).Visible = False
    Next
    For T = 1 To 3
        Shape2(T).BackColor = vbWhite
    Next
    If Check1.Visible = False And Check1.Value = 0 Then
        iTEST = False
    Else
        iTEST = True
    End If
    For T = 7 To 18
        Picture6(T).Visible = False
    Next
    AutomateEnd = False
    Roll = 0
    YatSpent = False
    Joker = False
    JokerTIM.Enabled = False
    Label5.Visible = False
    Label2 = "On Roll #" & (Roll + 1)
    Label8 = Label2
    finTOT = ""
    For T = 0 To 12
        Picture7(T).Visible = False
        DiePlayed(T + 1) = False
    Next
    BonusScore = 0
    LabBONUS.Visible = False
    Label10.Visible = False
    For T = 1 To 5
        Label1(T).BackColor = CTL_OFF
        DieSELECT(T) = False
    Next
    For T = 1 To 6
        Picture1(T).Visible = False
    Next
    Picture1(7).Visible = True
    For T = 1 To 6
        Picture2(T).Visible = False
    Next
    Picture2(7).Visible = True
    For T = 1 To 6
        Picture3(T).Visible = False
    Next
    Picture3(7).Visible = True
    For T = 1 To 6
        Picture4(T).Visible = False
    Next
    Picture4(7).Visible = True
    For T = 1 To 6
        Picture5(T).Visible = False
    Next
    Picture5(7).Visible = True
    Scores.Ones = 0
    Scores.Twos = 0
    Scores.Threes = 0
    Scores.Fours = 0
    Scores.Fives = 0
    Scores.Sixes = 0
    Scores.Top = 0
    Scores.Bottom = 0
    Label3 = ""
    Label7 = ""
    For T = 1 To 6
        TopScores(T) = ""
        BottomScore(T) = ""
    Next
End Sub

Private Sub Form_Load()
Dim T As Integer
    PUB = False  'Make this FALSE to hide the "MORE" button on ScoreBoard
    RollDice.Caption = "Roll Dice"
    RollDice.FontSize = 20
    FightSND = False
    Label13.Top = Label4.Top + 30
    Label13.Left = Label4.Left + 20
    SPSetting = 4
    Label9.Top = EndLABEL.Top + 40
    Yatz(2).Visible = False
    Yatz(1).Visible = False
    Yatz(0).Visible = True
    LabBONUS.Visible = False
    Label10.Visible = False
    GameScores(500, 0) = 0
    Label11 = ""
    Label12 = ""
    For T = 1 To 3
        Shape2(T).BackColor = vbWhite
    Next
    For T = 0 To 2
        ThreeK(T).Top = 1320
        FourK(T).Top = 1320
        FullH(T).Top = 1320
        SmStr(T).Top = 1320
        LgStr(T).Top = 1320
        Chance(T).Top = 1320
    Next
    For T = 1 To 2
        ThreeK(T).Visible = False
        FourK(T).Visible = False
        FullH(T).Visible = False
        SmStr(T).Visible = False
        LgStr(T).Visible = False
        Chance(T).Visible = False
    Next
    For T = 1 To 18
        Picture6(T).Top = 360
    Next
    For T = 1 To 6
        Picture1(T).Top = TopPosition
        Picture2(T).Top = TopPosition
        Picture3(T).Top = TopPosition
        Picture4(T).Top = TopPosition
        Picture5(T).Top = TopPosition
        TopScores(T) = ""
        BottomScore(T) = ""
        If T < 6 Then
            Label1(T).BackColor = CTL_OFF
        End If
    Next
    Roll = 0
    Label2 = "On Roll #" & (Roll + 1)
    Label8 = Label2
    Scores.HighScore = GetHighScore
    Me.Caption = " Solo Yahtzee                         High Score = " & Scores.HighScore & _
    " by " & Scores.HSHolder
    Call NewGameSet
End Sub

Public Sub CheckForEnd()
Dim T As Integer
Dim B As Integer
Dim RunTot As Integer
    RunTot = 0
    If Picture7(0).Visible = True Or Yatz(2).Visible = True Then
        RunTot = RunTot + 1
    End If
    For T = 1 To 6
        If TopScores(T) <> "" Then RunTot = RunTot + 1
        If BottomScore(T) <> "" Then RunTot = RunTot + 1
    Next
    If RunTot = 12 Then AutomateEnd = True
    Debug.Print "CheckForEnd:  Automate = " & AutomateEnd & " : Run Total = " & RunTot
    If RunTot > 12 Then
        EndLABEL.Visible = True
        Label9.Visible = True
        OverTIM.Enabled = True
        GameOver = True
        Label2 = "Done"
        Label8 = Label2
        B = CInt(GameScores(500, 0))
        B = B + 1: GameScores(500, 0) = CStr(B)
        If iTEST Then
            GameScores(B, 0) = "Game #" & B & ": " & finTOT & " (Test Mode)"
        Else
            GameScores(B, 0) = "Game #" & B & ": " & finTOT
        End If
        GameScores(B, 1) = AmtYaht
        Totals(B) = finTOT
        If finTOT >= ATS(5).Score Then
            If iTEST = False And Check1.Value = 0 Then
                PreFOT = finTOT
                BonSNDTim.Enabled = False
                iPlay ("NewHighScore.wav")
                NHS.Show vbModal
            Else
                frmMain.Caption = " Solo Yahtzee                Test Mode High Score = " & finTOT & "  (Not Saved)"
                BonSNDTim.Enabled = True
            End If
        Else
            BonSNDTim.Enabled = True
        End If
    End If
End Sub

Private Sub FourK_Click(Index As Integer)
Dim aTOT As Integer
Dim T As Integer
    If Roll = 0 Then Exit Sub
    If Label2 = "Rolling..." Then Exit Sub
    If DiePlayed(8) = True Then Exit Sub
    DiePlayed(8) = True
    BottomScore(2).ForeColor = vbYellow
    YatzTIM.Enabled = False
    For T = 0 To 2
        FourK(T).ToolTipText = ""
    Next
    SelectSlp
    If Joker Then
        JokerHandler
        KillYahtzee = False
    Else
        KillYahtzee = True
    End If
    If Yahtzee = True And KillYahtzee = False Then
        Call HandleYahtzee(1)
        Exit Sub
    Else
        KillYahtzee = False
    End If
    If FourKind = True Then
        aTOT = DieSTAT(1) + DieSTAT(2) + DieSTAT(3) + DieSTAT(4) + DieSTAT(5)
        BottomScore(2) = aTOT
        Scores.Bottom = Scores.Bottom + aTOT
    Else
        BottomScore(2) = "0"
    End If
    Tally
End Sub

Private Sub FullH_Click(Index As Integer)
Dim T As Integer
    If Roll = 0 Then Exit Sub
    If Label2 = "Rolling..." Then Exit Sub
    If DiePlayed(9) = True Then Exit Sub
    DiePlayed(9) = True
    BottomScore(3).ForeColor = vbYellow
    YatzTIM.Enabled = False
    For T = 0 To 2
        FullH(T).ToolTipText = ""
    Next
    SelectSlp
    If Joker Then
        YatzTIM.Enabled = False
        JokerHandler
        BottomScore(3) = "25"
        Scores.Bottom = Scores.Bottom + 25
        Tally
        KillYahtzee = False
        Exit Sub
    Else
        KillYahtzee = True
    End If
    If Yahtzee = True And KillYahtzee = False Then
        Call HandleYahtzee(1)
        Exit Sub
    Else
        KillYahtzee = False
    End If
    If FullHouse = True Then
        BottomScore(3) = "25"
        Scores.Bottom = Scores.Bottom + 25
    Else
        BottomScore(3) = "0"
    End If
    Tally
End Sub

Private Sub LgStr_Click(Index As Integer)
Dim T As Integer
    If Roll = 0 Then Exit Sub
    If Label2 = "Rolling..." Then Exit Sub
    If DiePlayed(11) = True Then Exit Sub
    DiePlayed(11) = True
    BottomScore(5).ForeColor = vbYellow
    YatzTIM.Enabled = False
    For T = 0 To 2
        LgStr(T).ToolTipText = ""
    Next
    SelectSlp
    If Joker Then
        JokerHandler
        BottomScore(5) = "40"
        Scores.Bottom = Scores.Bottom + 40
        Tally
        KillYahtzee = False
        Exit Sub
    Else
        KillYahtzee = True
    End If
    If Yahtzee = True And KillYahtzee = False Then
        Call HandleYahtzee(1)
        Exit Sub
    Else
        KillYahtzee = False
    End If
    If LgStreight = True Then
        BottomScore(5) = "40"
        Scores.Bottom = Scores.Bottom + 40
    Else
        BottomScore(5) = "0"
    End If
    Tally
End Sub

Private Sub mnuSpotHelp_Click()
    iPlay ("ModClick7.wav")
    If mnuSpotHelp.Checked = False Then
        mnuSpotHelp.Checked = True
        Pref.MovAss = 1
    Else
        mnuSpotHelp.Checked = False
        Pref.MovAss = 0
    End If
End Sub

Private Sub OverTIM_Timer()
    If EndLABEL.Visible = True Then
        EndLABEL.Visible = False
        Label9.Visible = False
    Else
        EndLABEL.Visible = True
        Label9.Visible = True
    End If
End Sub

Private Sub Picture1_Click(Index As Integer)
    SelectDie (1)
End Sub
Private Sub Picture2_Click(Index As Integer)
    SelectDie (2)
End Sub
Private Sub Picture3_Click(Index As Integer)
    SelectDie (3)
End Sub
Private Sub Picture4_Click(Index As Integer)
    SelectDie (4)
End Sub
Private Sub Picture5_Click(Index As Integer)
    SelectDie (5)
End Sub
Public Sub SelectDie(Index As Integer)
    If Right$((CStr(Label2)), 1) = "1" Then
        RollDice.BackColor = vbRed
        FlashTIM.Enabled = True
        FlashCounter = 0
        iPlay ("Error.wav")
        Exit Sub
    End If
    If RollDice.BackColor = &HFFFF& Then Exit Sub
    If Roll < 1 Or Roll > 2 Then Exit Sub
    If Label1(Index).BackColor = CTL_ON Then
        Label1(Index).BackColor = CTL_OFF
        DieSELECT(Index) = False
        iPlay ("UnSelect.wav")
    Else
        Label1(Index).BackColor = CTL_ON
        DieSELECT(Index) = True
        SelectSnd
    End If
End Sub

Private Sub Picture6_Click(Index As Integer)
Dim T As Integer
Dim RESP As Integer
    If Index > 6 Then
        If Index > 12 Then
            Index = Index - 12
        Else
            Index = Index - 6
        End If
    End If
    If Roll = 0 Then Exit Sub
    If Label2 = "Rolling..." Then Exit Sub
    If DiePlayed(Index) = True Then Exit Sub
    RESP = 0
    YatzTIM.Enabled = False
    For T = 1 To 5
        If DieSTAT(T) = Index Then
            RESP = RESP + Index
        End If
    Next
    DiePlayed(Index) = True
    SelectSlp
    If Joker Then
        JokerHandler
        RESP = Index * 5
        KillYahtzee = False
    Else
        KillYahtzee = True
    End If
    If Yahtzee = True And KillYahtzee = False Then
        Call HandleYahtzee(1)
        Exit Sub
    Else
        KillYahtzee = False
    End If
    TopScores(Index) = RESP
    For T = 1 To 6 '
        TopScores(T).ForeColor = vbYellow
    Next
    Select Case Index
        Case 1
            Scores.Ones = RESP
        Case 2
            Scores.Twos = RESP
        Case 3
            Scores.Threes = RESP
        Case 4
            Scores.Fours = RESP
        Case 5
            Scores.Fives = RESP
        Case 6
            Scores.Sixes = RESP
    End Select
    Scores.Top = Scores.Top + RESP
    Label3 = "Top: " & Scores.Top
    Label7 = Label3
    Tally
End Sub

Public Sub Tally()
Dim T As Integer
    Roll = 0
    Label2 = "On Roll # " & (Roll + 1)
    Label8 = Label2
    If Roll = 0 Then
        For T = 1 To 3
            Shape2(T).BackColor = vbWhite
        Next
    End If
    If SpTimer.Enabled = True Then
        SpTimer.Enabled = False
        Label11 = ""
        Label12 = ""
        SpDial = 0
    End If
    If BonusScore < 1 Then
        If Scores.Top >= 63 Then
            BonusScore = 35
            LabBONUS.Visible = True
            Label10.Visible = True
            BonSNDTim.Enabled = True
        End If
    End If
    For T = 1 To 5
        Label1(T).BackColor = CTL_OFF
        DieSELECT(T) = False
    Next
    finTOT = Scores.Top + Scores.Bottom + BonusScore
    CheckForEnd
    If Yahtzee = True Then
        Yahtzee = False
        If Yahtzeed = False Then
            Yatz(0).Visible = True
            Yatz(1).Visible = False
            Yatz(2).Visible = False
        Else
            Yatz(0).Visible = False
            Yatz(1).Visible = True
            Yatz(2).Visible = False
        End If
    End If
    For T = 1 To 18
        If T < 7 Then
            Picture6(T).Visible = True
        Else
            Picture6(T).Visible = False
        End If
    Next
    For T = 1 To 2
        ThreeK(T).Visible = False
        FourK(T).Visible = False
        FullH(T).Visible = False
        SmStr(T).Visible = False
        LgStr(T).Visible = False
        Chance(T).Visible = False
    Next
    For T = 1 To 14
        SptHlp(T) = False
    Next
End Sub

Private Sub Picture6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Right$((CStr(Label2)), 1) = "1" Then Exit Sub
    If RollDice.BackColor = &HFFFF& Then Exit Sub
    If Index > 6 Then
        If Index > 12 Then
            Index = Index - 12
        Else
            Index = Index - 6
        End If
    End If
    If TopScores(Index) = "" Then
        FixEm = True
        If GoodTrip(Index) Or Joker = True Then
            Picture6(Index).Visible = False
            Picture6(Index + 12).Visible = True
            TopScores(Index).ForeColor = vbGreen
            TopScores(Index).Caption = Qtally(Index)
        Else
            Picture6(Index).Visible = False
            Picture6(Index + 6).Visible = True
            TopScores(Index).ForeColor = vbRed
            TopScores(Index).Caption = Qtally(Index)
        End If
    End If
End Sub

Private Sub RollDice_Click()
Dim T As Integer
    If RollDice.BackColor = &HFFFF& Then Exit Sub    ' Already rolling
    If Roll > 2 Then Exit Sub
    If GameOver Then
        RollDice.FontSize = 14
        RollDice.Caption = "Right Click for Menu"
        MenuNotifyTIM.Enabled = True
        Exit Sub
    End If
    Yahtzee = False
    Joker = False
    JokerTIM.Enabled = False
    Label5.Visible = False
    If SpTimer.Enabled = True Then
        SpTimer.Enabled = False
        Label11 = ""
        Label12 = ""
        SpDial = 0
    End If
    YatzTIM.Enabled = False
    If Yahtzeed = False Then
        If Yatz(2).Visible = False Then
            Yatz(0).Visible = True
            Yatz(1).Visible = False
        End If
    End If
    Label2 = "Rolling..."
    Label8 = Label2 'Shadow
    RollDice.BackColor = &HFFFF&
    Roll = Roll + 1
    Cntr1 = 0
    For T = 1 To 5
        Die(T) = False
    Next
    For T = 1 To 14
        SptHlp(T) = False
    Next
    CleanUp
    RollSoundProcessor
    RollTIM.Enabled = True
End Sub

Private Sub SetYatz(PicNumb As Integer) ' Part of TEST routine
Dim T As Integer
    iTEST = True
    Select Case PicNumb 'This sub places a 6 in every spot (Not a Cheat, a Test)
        Case 1
            For T = 1 To 7
                Picture1(T).Visible = False
            Next
            Picture1(6).Visible = True
        Case 2
            For T = 1 To 7
                Picture2(T).Visible = False
            Next
            Picture2(6).Visible = True
        Case 3
            For T = 1 To 7
                Picture3(T).Visible = False
            Next
            Picture3(6).Visible = True
        Case 4
            For T = 1 To 7
                Picture4(T).Visible = False
            Next
            Picture4(6).Visible = True
        Case 5
            For T = 1 To 7
                Picture5(T).Visible = False
            Next
            Picture5(6).Visible = True
    End Select
End Sub

Private Sub RollDice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        ShowMenu
    End If
End Sub

Private Sub RollTIM_Timer() ' Standard Order Die Roll and select
    If mnuRand.Checked = True Then
        Call RandRoller
        Exit Sub
    End If
Dim Die1 As Integer
Dim Die2 As Integer
Dim Die3 As Integer
Dim Die4 As Integer
Dim Die5 As Integer
Dim T As Integer
Dim KJ As Integer
Dim CntStepped As Boolean
    CntStepped = False
    Randomize
    Die1 = GetRandomValue
    Die2 = GetRandomValue
    Die3 = GetRandomValue
    Die4 = GetRandomValue
    Die5 = GetRandomValue
    EndCheck
    If Die(1) = False And DieSELECT(1) = False Then
        For T = 1 To 7
            Picture1(T).Visible = False
        Next
        Picture1(Die1).Visible = True
        Picture1(Die1).Refresh
        If CntStepped = False Then
            CntStepped = True
            Cntr1 = Cntr1 + 1
        End If
        If Cntr1 > 10 Then
            If Check1 Then
                Call SetYatz(1)
                Die1 = 6
            End If
            Cntr1 = 0
            Die(1) = True
            DieSTAT(1) = Die1
            EndCheck
            If GameOver = True Then
                RollTIM.Enabled = False
                Exit Sub
            End If
        End If
    End If
    If Die(2) = False And DieSELECT(2) = False Then
        For T = 1 To 7
            Picture2(T).Visible = False
        Next
        Picture2(Die2).Visible = True
        Picture2(Die2).Refresh
        If CntStepped = False Then
            CntStepped = True
            Cntr1 = Cntr1 + 1
        End If
        If Cntr1 > 10 Then
            If Check1 Then
                Call SetYatz(2)
                Die2 = 6
            End If
            Cntr1 = 0
            Die(2) = True
            DieSTAT(2) = Die2
            EndCheck
            If GameOver = True Then
                RollTIM.Enabled = False
                Exit Sub
            End If
        End If
    End If
    If Die(3) = False And DieSELECT(3) = False Then
        For T = 1 To 7
            Picture3(T).Visible = False
        Next
        Picture3(Die3).Visible = True
        Picture3(Die3).Refresh
        If CntStepped = False Then
            CntStepped = True
            Cntr1 = Cntr1 + 1
        End If
        If Cntr1 > 10 Then
            If Check1 Then
                Call SetYatz(3)
                Die3 = 6
            End If
            Cntr1 = 0
            Die(3) = True
            DieSTAT(3) = Die3
            EndCheck
            If GameOver = True Then
                RollTIM.Enabled = False
                Exit Sub
            End If
        End If
    End If
    If Die(4) = False And DieSELECT(4) = False Then
        For T = 1 To 7
            Picture4(T).Visible = False
        Next
        Picture4(Die4).Visible = True
        Picture4(Die4).Refresh
        If CntStepped = False Then
            CntStepped = True
            Cntr1 = Cntr1 + 1
        End If
        If Cntr1 > 10 Then
            If Check1 Then
                Call SetYatz(4)
                Die4 = 6
            End If
            Cntr1 = 0
            Die(4) = True
            DieSTAT(4) = Die4
            EndCheck
            If GameOver = True Then
                RollTIM.Enabled = False
                Exit Sub
            End If
        End If
    End If
    If Die(5) = False And DieSELECT(5) = False Then
        For T = 1 To 7
            Picture5(T).Visible = False
        Next
        Picture5(Die5).Visible = True
        Picture5(Die5).Refresh
        If CntStepped = False Then
            CntStepped = True
            Cntr1 = Cntr1 + 1
        End If
        If Cntr1 > 10 Then
            If Check1 Then
                Call SetYatz(5)
                Die5 = 6
            End If
            Cntr1 = 0
            Die(5) = True
            DieSTAT(5) = Die5
            EndCheck
            If GameOver = True Then
                RollTIM.Enabled = False
                Exit Sub
            End If
        End If
    End If
End Sub

Private Function GetRandomValue() As Integer    'Generate a random value for a DIE
SptA:
    Randomize '                     Five numbers from inside a wider group (12). This ensures a good random
    GetRandomValue = Int(Rnd * 12) 'dice roll every time.
    If GetRandomValue < 3 Or GetRandomValue > 8 Then GoTo SptA
    GetRandomValue = GetRandomValue - 2
End Function

Private Sub RandRoller() 'This optional SUB randomizes the roll stop positions
Dim RandDieValue(5) As Integer
Dim RandomPick As Integer
Dim T As Integer
Dim CntStepped As Boolean
    CntStepped = False
    Randomize
    For T = 1 To 5
        RandDieValue(T) = GetRandomValue
    Next
    EndCheck
    If Die(1) = False And DieSELECT(1) = False Then
        For T = 1 To 7
            Picture1(T).Visible = False
        Next
        Picture1(RandDieValue(1)).Visible = True
        Picture1(RandDieValue(1)).Refresh
        If CntStepped = False Then
            CntStepped = True
            Cntr1 = Cntr1 + 1
        End If
    End If
    If Die(2) = False And DieSELECT(2) = False Then
        For T = 1 To 7
            Picture2(T).Visible = False
        Next
        Picture2(RandDieValue(2)).Visible = True
        Picture2(RandDieValue(2)).Refresh
        If CntStepped = False Then
            CntStepped = True
            Cntr1 = Cntr1 + 1
        End If
    End If
    If Die(3) = False And DieSELECT(3) = False Then
        For T = 1 To 7
            Picture3(T).Visible = False
        Next
        Picture3(RandDieValue(3)).Visible = True
        Picture3(RandDieValue(3)).Refresh
        If CntStepped = False Then
            CntStepped = True
            Cntr1 = Cntr1 + 1
        End If
    End If
    If Die(4) = False And DieSELECT(4) = False Then
        For T = 1 To 7
            Picture4(T).Visible = False
        Next
        Picture4(RandDieValue(4)).Visible = True
        Picture4(RandDieValue(4)).Refresh
        If CntStepped = False Then
            CntStepped = True
            Cntr1 = Cntr1 + 1
        End If
    End If
    If Die(5) = False And DieSELECT(5) = False Then
        For T = 1 To 7
            Picture5(T).Visible = False
        Next
        Picture5(RandDieValue(5)).Visible = True
        Picture5(RandDieValue(5)).Refresh
        If CntStepped = False Then
            CntStepped = True
            Cntr1 = Cntr1 + 1
        End If
    End If
    If Cntr1 > 10 Then
SptQ:
        RandomPick = (Int(Rnd * 5) + 1)
        If Die(RandomPick) = True Or DieSELECT(RandomPick) = True Then GoTo SptQ
        If Check1 Then  'Always a Yahtzee TEST mode
            Call SetYatz(RandomPick)
            RandDieValue(RandomPick) = 6
        End If
        Cntr1 = 0
        Die(RandomPick) = True
        DieSTAT(RandomPick) = RandDieValue(RandomPick)
        EndCheck
        If GameOver = True Then
            RollTIM.Enabled = False
            Exit Sub
        End If
    End If
End Sub

Private Sub EndCheck()
Dim T As Integer
Dim fDie(5) As Boolean
Dim DoneD As Boolean
    If Die(1) = True Or DieSELECT(1) = True Then
        fDie(1) = True
    Else
        fDie(1) = False
    End If
    If Die(2) = True Or DieSELECT(2) = True Then
        fDie(2) = True
    Else
        fDie(2) = False
    End If
    If Die(3) = True Or DieSELECT(3) = True Then
        fDie(3) = True
    Else
        fDie(3) = False
    End If
    If Die(4) = True Or DieSELECT(4) = True Then
        fDie(4) = True
    Else
        fDie(4) = False
    End If
    If Die(5) = True Or DieSELECT(5) = True Then
        fDie(5) = True
    Else
        fDie(5) = False
    End If
    DoneD = True
    For T = 1 To 5
        If fDie(T) = False Then DoneD = False
    Next
    If DoneD Then
        RollTIM.Enabled = False
        For T = 1 To 5
            Debug.Print "Dice #" & T & " = " & DieSTAT(T)
        Next
        RollDice.BackColor = &H8000000F
        If DieSTAT(1) = DieSTAT(2) And DieSTAT(2) = DieSTAT(3) And DieSTAT(3) = DieSTAT(4) _
        And DieSTAT(4) = DieSTAT(5) And YatSpent = False Then
            Yahtzee = True
            'Yahtzeed = True    'When active, causes "YAHZEE" text to remain in box after first yahtzee. Aesthetic Preferences
            Yatz(0).Visible = False 'Black
            Yatz(1).Visible = True  'Yahtzee
            Yatz(2).Visible = False 'Scratched
            If Picture7(0).Visible = True And (Label2 <> "Done" Or AutomateEnd = False) Then
                Joker = True
                JokerTIM.Enabled = True
                Label5.Visible = True
            ElseIf Picture7(0).Visible = True And AutomateEnd = True Then
                Joker = True
                Call FillLast
            End If
            If Joker = False And AutomateEnd = False Then
                YatzTIM.Enabled = True
            ElseIf Joker = False And AutomateEnd = True Then
                Call FillLast
            End If
            iPlay ("Yahtzee.wav")
            AmtYaht = AmtYaht + 1
        End If
        If Roll < 3 Then
            Label2 = "On Roll #" & (Roll + 1)
            Label8 = Label2
            If mnuspeed.Checked = True Then
                SpTimer = True
                Label11 = "1"
                Label12 = "1"
            End If
        Else
            Label2 = "DONE"
            Label8 = Label2
            If AutomateEnd = True Then  'Auto End Game
                If Roll > 0 Then
                    If Roll = 1 Then
                        Shape2(1).BackColor = vbRed
                        Shape2(2).BackColor = vbWhite
                        Shape2(3).BackColor = vbWhite
                    ElseIf Roll = 2 Then
                        Shape2(2).BackColor = vbRed
                    ElseIf Roll = 3 Then
                        Shape2(3).BackColor = vbRed
                    End If
                End If
                Call FillLast
                Exit Sub
            End If
        End If
        If Roll > 0 Then
            If Roll = 1 Then
                Shape2(1).BackColor = vbRed
                Shape2(2).BackColor = vbWhite
                Shape2(3).BackColor = vbWhite
            ElseIf Roll = 2 Then
                Shape2(2).BackColor = vbRed
            ElseIf Roll = 3 Then
                Shape2(3).BackColor = vbRed
            End If
        End If
        If mnuSpotHelp.Checked = True Then
            If Not (Yahtzee = True And Joker = False) Then
                CheckPlays
            End If
        End If
    End If
End Sub

Private Sub CheckPlays()    ' Automated Placement Assistance
Dim T As Integer
    For T = 1 To 6
        If TopScores(T) = "" Then
            If GoodTrip(T) Or Joker = True Then
                SptHlp(T) = True
            Else
                SptHlp(T) = False
            End If
        End If
    Next
    If BottomScore(1) = "" Then
        If ThreeKind = True Or Joker = True Then
            SptHlp(7) = True
        Else
            SptHlp(7) = False
        End If
    End If
    If BottomScore(2) = "" Then
        If FourKind = True Or Joker = True Then
            SptHlp(8) = True
        Else
            SptHlp(8) = False
        End If
    End If
    If BottomScore(3) = "" Then
        If FullHouse = True Or Joker = True Then
            SptHlp(9) = True
        Else
            SptHlp(9) = False
        End If
    End If
    If BottomScore(4) = "" Then
        If SmStreight = True Or Joker = True Then
            SptHlp(10) = True
        Else
            SptHlp(10) = False
        End If
    End If
    If BottomScore(5) = "" Then
        If LgStreight = True Or Joker = True Then
            SptHlp(11) = True
        Else
            SptHlp(11) = False
        End If
    End If
    If BottomScore(6) = "" Then
        SptHlp(12) = True
    Else
        SptHlp(12) = False
    End If
    ASH = 0
    SHTim.Enabled = True
End Sub

Private Sub SHTim_Timer()
Dim T As Integer
    If LG = False Or ASH = 0 Then
        LG = True
    Else
        LG = False
    End If
    For T = 1 To 6
        If SptHlp(T) Then Picture6(T + 12).Visible = LG
    Next
    If SptHlp(7) Then ThreeK(2).Visible = LG
    If SptHlp(8) Then FourK(2).Visible = LG
    If SptHlp(9) Then FullH(2).Visible = LG
    If SptHlp(10) Then SmStr(2).Visible = LG
    If SptHlp(11) Then LgStr(2).Visible = LG
    If SptHlp(12) Then Chance(2).Visible = LG
    ASH = ASH + 1
    If ASH >= 9 Then
        SHTim.Enabled = False
        ASH = 0
    End If
End Sub

Private Sub SmStr_Click(Index As Integer)
Dim T As Integer
    If Roll = 0 Then Exit Sub
    If Label2 = "Rolling..." Then Exit Sub
    If DiePlayed(10) = True Then Exit Sub
    DiePlayed(10) = True
    BottomScore(4).ForeColor = vbYellow
    YatzTIM.Enabled = False
    For T = 0 To 2
        SmStr(T).ToolTipText = ""
    Next
    SelectSlp
    If Joker Then
        JokerHandler
        BottomScore(4) = "30"
        Scores.Bottom = Scores.Bottom + 30
        Tally
        KillYahtzee = False
        Exit Sub
    Else
        KillYahtzee = True
    End If
    If Yahtzee = True And KillYahtzee = False Then
        Call HandleYahtzee(1)
        Exit Sub
    Else
        KillYahtzee = False
    End If
    If SmStreight = True Then
        BottomScore(4) = "30"
        Scores.Bottom = Scores.Bottom + 30
    Else
        BottomScore(4) = "0"
    End If
    Tally
End Sub

Private Sub SmStr_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ChkSmSt
End Sub

Private Sub ChkSmSt()
    If Right$((CStr(Label2)), 1) = "1" Then Exit Sub
    If RollDice.BackColor = &HFFFF& Then Exit Sub
    If FixEm = True And (SmStr(2).Visible = False And SmStr(1).Visible = False) Then
        CleanUp
    End If
    If BottomScore(4) = "" Then
        FixEm = True
        If SmStreight Or Joker = True Then
            SmStr(2).Visible = True
            BottomScore(4).ForeColor = vbGreen
            BottomScore(4).Caption = "30"
        Else
            SmStr(1).Visible = True
            BottomScore(4).ForeColor = vbRed
            BottomScore(4).Caption = "0"
        End If
    End If
End Sub

Private Sub SpTimer_Timer() 'Speed Yahtzee Timer
    If Label2 = "Done" Then
        SpTimer.Enabled = False
        Label11 = ""
        Label12 = ""
        Exit Sub
    End If
    SpDial = SpDial + 1
    Label11 = (SpDial + 1)
    Label12 = Label11
    If SpDial > (SPSetting - 1) Then
        Call RollDice_Click
        SpDial = 0
        Label11 = ""
        Label12 = ""
        SpTimer.Enabled = False
    End If
End Sub

Private Sub ThreeK_Click(Index As Integer)
Dim T As Integer
Dim aTOT As Integer
    If Roll = 0 Then Exit Sub
    If Label2 = "Rolling..." Then Exit Sub
    If DiePlayed(7) = True Then Exit Sub
    DiePlayed(7) = True
    BottomScore(1).ForeColor = vbYellow
    YatzTIM.Enabled = False
    For T = 0 To 2
        ThreeK(T).ToolTipText = ""
    Next
    SelectSlp
    If Joker Then
        JokerHandler
        KillYahtzee = False
    Else
        KillYahtzee = True
    End If
    If Yahtzee = True And KillYahtzee = False Then
        Call HandleYahtzee(1)
        Exit Sub
    Else
        KillYahtzee = False
    End If
    If ThreeKind = True Then
        aTOT = DieSTAT(1) + DieSTAT(2) + DieSTAT(3) + DieSTAT(4) + DieSTAT(5)
        BottomScore(1) = aTOT
        Scores.Bottom = Scores.Bottom + aTOT
    Else
        BottomScore(1) = "0"
    End If
    Tally
End Sub

Private Sub ThreeK_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Right$((CStr(Label2)), 1) = "1" Then Exit Sub
    If RollDice.BackColor = &HFFFF& Then Exit Sub
    If FixEm = True And (ThreeK(2).Visible = False And ThreeK(1).Visible = False) Then
        CleanUp
    End If
    If BottomScore(1) = "" Then
        FixEm = True
        If ThreeKind Or Joker = True Then
            ThreeK(2).Visible = True
            BottomScore(1).ForeColor = vbGreen
            BottomScore(1).Caption = DieSTAT(1) + DieSTAT(2) + DieSTAT(3) + DieSTAT(4) + DieSTAT(5)
        Else
            ThreeK(1).Visible = True
            BottomScore(1).ForeColor = vbRed
            BottomScore(1).Caption = "0"
        End If
    End If
End Sub

Private Sub Yatz_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    HandleYahtzee (Index)
End Sub

Public Sub HandleYahtzee(Index As Integer)
Dim T As Integer
    If Roll = 0 Then Exit Sub
    If YatSpent Then Exit Sub
    If Label2 = "Rolling..." Then Exit Sub
    If AutomateEnd = False Then
        If Joker = True Then Exit Sub
        If Picture7(0).Visible = True Then Exit Sub
    End If
    If Yahtzee = False Then
        If Picture7(0).Visible = True Then Exit Sub
        YatzTIM.Enabled = False
        YatSpent = True
        Yatz(0).Visible = False
        Yatz(1).Visible = False
        Yatz(2).Visible = True  'Show "SCRATCH"
        Tally
        If GameOver = False Then
            iPlay ("SCRATCH.wav")  'Play loser sound here (Scratch Yahtzee)
        End If
        Exit Sub
    End If
    YatzTIM.Enabled = False
    Yatz(0).Visible = False
    Yatz(1).Visible = True  'Show "YAHTZEE"
    Yatz(2).Visible = False
    YatzTIM.Enabled = False
    Yahtzee = False
    If DieSTAT(1) = DieSTAT(2) And DieSTAT(2) = DieSTAT(3) And DieSTAT(3) = DieSTAT(4) _
    And DieSTAT(4) = DieSTAT(5) Then
        If Picture7(0).Visible = False Then ' First Yahtzee = +50 Points
            Picture7(0).Visible = True
            Scores.Bottom = Scores.Bottom + 50
            Tally
            Exit Sub
        End If
    End If
End Sub

Public Sub JokerHandler()
Dim T As Integer
    Yahtzee = False
    YatzTIM.Enabled = False
    If Yahtzeed = False Then
        Yatz(0).Visible = True
        Yatz(1).Visible = False
        Yatz(2).Visible = False
    Else
        Yatz(0).Visible = False
        Yatz(1).Visible = True
        Yatz(2).Visible = False
    End If
    Joker = False
    JokerTIM.Enabled = False
    Label5.Visible = False
    Scores.Bottom = Scores.Bottom + 100 ' Each additional Yahtzee is worth 100 Points
    For T = 1 To 12
        If Picture7(T).Visible = False Then
            Picture7(T).Visible = True
            Exit For
        End If
    Next
End Sub

Private Sub YatzTIM_Timer() 'Flash YAHTZEE text in box
    If Yatz(0).Visible = True Then
        Yatz(0).Visible = False
        Yatz(1).Visible = True
    Else
        Yatz(0).Visible = True
        Yatz(1).Visible = False
    End If
End Sub
