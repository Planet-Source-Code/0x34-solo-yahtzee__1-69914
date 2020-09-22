VERSION 5.00
Begin VB.Form Graph 
   ClientHeight    =   1155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5025
   Icon            =   "Graph.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1155
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton SaveButt 
      Caption         =   "Save History"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox P1 
      BackColor       =   &H00000000&
      Height          =   975
      Left            =   120
      ScaleHeight     =   915
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Graph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MaxiScore As Integer
Dim PointCount As Integer
Dim iAverage As Integer

            '   Scoring Graph added 122809
            '   Code by Ken Slater - 0x34
            '
            '   Ain't coding fun??  Good times!

Private Sub Form_Load()
Dim Y As Integer
Dim T As Integer
    Me.Width = 11160
    Me.Height = 5270
    P1.Left = 0
    P1.Top = 0
    PointCount = CInt(GameScores(500, 0))
    MaxiScore = GetMax(PointCount)
    If PointCount = 1 Then
        Me.Caption = " Solo Yahtzee - Session Scoring Graph  (1 game played)"
    Else
        Me.Caption = " Solo Yahtzee - Session Scoring Graph  (" & PointCount & " games played)"
    End If
    Y = 0
    If PointCount > 1 Then
        For T = 1 To PointCount
            Y = Y + Totals(T)
        Next
        iAverage = (Y / PointCount)
    Else
        iAverage = 0
    End If
    Debug.Print "MaxiScore = " & MaxiScore & " - PointCount = " & PointCount
    P1.AutoRedraw = True
    If GraphX > 1 Then  ' Set to previously adjusted size
        Graph.Width = GraphX
        Graph.Height = GraphY
    End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
    P1.Width = Graph.Width - 105
    P1.Height = Graph.Height - 495 '375 if using a Tool Window
    P1.ScaleWidth = 100 + (PointCount * 100)
    P1.ScaleHeight = 50 + MaxiScore
    P1.Cls
    GridIt
    GraphIt
End Sub

Private Function GetMax(A As Integer) As Integer
Dim T As Integer
    GetMax = 0
    For T = 1 To A
        If Totals(T) > GetMax Then
            GetMax = Totals(T)
        End If
    Next
End Function

Private Sub GridIt()
Dim T As Integer
Dim E As Long
Dim K As Integer
Dim Divs As Integer
    Divs = Round(P1.ScaleHeight / 10, 0)
    E = Divs
    For T = 0 To 10
        P1.Line (0, E)-(P1.ScaleWidth, E), RGB(0, 0, 100)
        E = E + Divs
    Next
    P1.FontSize = 5
    P1.ForeColor = RGB(0, 255, 255)
    K = 1
    For T = 100 To (PointCount * 100) + 100 Step 100
        P1.Line (T, 0)-(T, P1.ScaleHeight), RGB(0, 0, 100)
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

Public Sub GraphIt()
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
        P1.CurrentX = K
        P1.CurrentY = (P1.ScaleHeight - 40) - Totals(T)
        P1.FontSize = 7
        P1.ForeColor = RGB(255, 255, 0)
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
        P1.ForeColor = RGB(255, 0, 0)
        P1.CurrentY = (((P1.ScaleHeight - 5) - iAverage) - ((P1.ScaleHeight / 100) * 5))
        P1.CurrentX = 10
        P1.FontSize = 6
        P1.Print "AVG"
        P1.CurrentY = P1.ScaleHeight - (iAverage - ((P1.ScaleHeight / 100) / 6))
        P1.CurrentX = 10
        P1.FontSize = 6
        P1.Print iAverage
        P1.DrawStyle = vbDot   '2
        P1.Line (65, ((P1.ScaleHeight - 5) - iAverage))-((P1.ScaleWidth - 65), ((P1.ScaleHeight - 5) - iAverage)), RGB(200, 0, 0)
        P1.DrawStyle = vbSolid  '0
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    GraphX = Graph.Width
    GraphY = Graph.Height
    iPlay ("ModClick7.wav")
End Sub

Private Sub P1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Unload Graph
End Sub
