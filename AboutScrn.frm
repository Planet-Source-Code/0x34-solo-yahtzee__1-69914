VERSION 5.00
Begin VB.Form AboutScrn 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   Picture         =   "AboutScrn.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5640
      Top             =   240
   End
End
Attribute VB_Name = "AboutScrn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X(1000), Y(1000), Z(1000) As Integer
Dim tmpX(1000), tmpY(1000), tmpZ(1000) As Integer
Dim K As Integer
Dim Zoom As Integer
Dim Speed As Integer
Dim DrSty As Integer
Dim ExitView As Boolean
Dim Density As Integer
Dim JustStarted As Boolean

Private Sub Form_Click()
    If JustStarted = True Then
        JustStarted = False
        Exit Sub
    End If
    If Timer1.Enabled = False Then
        Unload Me
        Exit Sub
    End If
    ExitView = True
    For i = 0 To 100
        AboutScrn.DrawMode = DrSty
        Me.Circle (tmpX(i), tmpY(i)), DrSty, BackColor
    Next
End Sub

Private Sub Form_Load()
Dim i As Integer
    Density = 200
    DrSty = 6
    AboutScrn.DrawMode = DrSty
    ExitView = False
    Me.Refresh
    Me.AutoRedraw = True
    Speed = -3  '   <-----------------  (Range = -1 to -50)
    K = 2038
    Zoom = 256
    Timer1.Interval = 1
    For i = 0 To 1000
        X(i) = Int(Rnd * 1024) - 512
        Y(i) = Int(Rnd * 1024) - 512
        Z(i) = Int(Rnd * 512) - 256
        tmpX(i) = 0
        tmpY(i) = 0
        tmpZ(i) = 0
    Next i
    JustStarted = False
    If frmMain.mnuSound.Checked = True Then
        JustStarted = False
        Timer1.Enabled = True
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Timer1.Enabled = False Then
            Timer1.Enabled = True
            JustStarted = True
        End If
    End If
End Sub

Private Sub Timer1_Timer()
Dim i As Integer
Dim Radius As Integer
Dim StarColor As Integer
    If ExitView Then GoTo Rexx
    For i = 0 To Density
        AboutScrn.DrawMode = DrSty
        Me.Circle (tmpX(i), tmpY(i)), DrSty - 1, BackColor
        Z(i) = Z(i) + Speed
        If Z(i) > 255 Then Z(i) = -255
        If Z(i) < -255 Then Z(i) = 255
        tmpZ(i) = Z(i) + Zoom
        tmpX(i) = (X(i) * K / tmpZ(i)) + (AboutScrn.Width / 2)
        tmpY(i) = (Y(i) * K / tmpZ(i)) + (AboutScrn.Height / 2)
        Radius = 1
        StarColor = 256 - Z(i)
        Me.Circle (tmpX(i), tmpY(i)), DrSty - 1, RGB(StarColor, StarColor, StarColor)
        If ExitView = True Then
            GoTo Rexx
        End If
    Next
    Exit Sub
Rexx:
    Timer1.Enabled = False
    Unload Me
End Sub
