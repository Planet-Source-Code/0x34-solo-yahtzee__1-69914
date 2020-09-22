VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form SpeedSet 
   BackColor       =   &H00C00000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   " Speed Challenge Set"
   ClientHeight    =   720
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3570
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   3570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   661
      _Version        =   393216
      Value           =   2
      Min             =   1
      Enabled         =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Seconds before Auto Roll:"
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
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   180
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "5"
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
      Left            =   3120
      TabIndex        =   1
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "SpeedSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Solo Yahtzee

'                                       Programming By
'                                    Ken Slater 2007/2008
'                                            0x34

Private Sub Form_Load()
    UpDown1 = SPSetting
    Label1 = UpDown1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SPSetting = UpDown1
    Pref.SpeedValue = SPSetting
End Sub

Private Sub UpDown1_Change()
    Label1 = UpDown1
End Sub
