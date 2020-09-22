Attribute VB_Name = "modMain"
Option Explicit

'Solo Yahtzee

'                                       Programming By
'                              Ken Slater 2007/2008/2009/2010
'                                            0x34

'                                 SPEED Yahtzee enhancement:
'               During game play, right click on the main FORM to access the menu.

'                                    Sound Effects Added
'                                      in February 2009

Global Yahtzee As Boolean
Global Yahtzeed As Boolean
Global YahtzeeCNT As Integer
Global YatSpent As Boolean
Global BonusScore As Integer
Global Joker As Boolean
Global KillYahtzee As Boolean
Global PreFOT As Integer
Global FixEm As Boolean
Global SpDial As Integer
Global SPSetting As Integer
Global DiePlayed(13) As Boolean
Global PUB As Boolean
Global mError As Long
Global SavePath As String
Global SptHlp(14) As Boolean
Global iInit As Boolean
Global FightSND As Boolean
Global CustFound As Boolean
Global GraphX As Long
Global GraphY As Long
Global GraphOpen As Boolean
Dim LE As Boolean


Global AutomateEnd As Boolean
Global Roll As Integer
Global FlashCounter As Integer
Global ScorePlaced As Boolean
Global NHSplace As Integer
Global AmtYaht As Integer
Global ColRTB(4) As ColorConstants
Global GameOver As Boolean

Global Const CTL_ON = &HFFFF&       'Color of selected dice
Global Const CTL_OFF = &H500000     'Color of non-selected dice
Global Const TopPosition = 2640     'Dice Picture Position (Top) in frmMain

Global iTEST As Boolean             'True if in TEST MODE

'// Dice Roll Variables
Global Die(5) As Boolean
Global DieSTAT(5) As Integer
Global Cntr1 As Long
Global DieSELECT(5) As Boolean

Global Scores As Scrs
Global DICE(5) As DiceStat
Global ATS(5) As HScores    'All Time Scores
Global mATS(5) As HScores   'For Merge Function
Global CL As ColorResults
Global CustomColor As CustCol
Global Pref As iPrefs

Global GameScores(500, 2) As String   'Track up to 499 games per session (the Shaw rule)
Global Totals(500) As Integer

Const csndsync = &H0, csndasync = &H1
Const csndnodefault = &H2
Const csndloop = &H8, csndnostop = &H10

Global Clik As Long

Type CustCol
    Position As ColorConstants
    Score As ColorConstants
    Name As ColorConstants
    Date As ColorConstants
    YatCnt As ColorConstants
End Type

Type iPrefs
    SpeedYaht As Integer
    SpeedValue As Integer
    MovAss  As Integer
    Sounds As Integer
    SndType As Integer
    RndStop As Integer
End Type
    
Type Scrs
    Ones As Integer
    Twos As Integer
    Threes As Integer
    Fours As Integer
    Fives As Integer
    Sixes As Integer
    Chance As Integer
    Top As Integer
    Bottom As Integer
    Total As Integer
    HighScore As Integer
    HSHolder As String
End Type

Type ColorResults
    Red As Long
    Green As Long
    Blue As Long
End Type

Type HScores
    Score As Integer
    Name As String
    Qnt As Integer
    Date As String
End Type

Type DiceStat
    one As Boolean
    Two As Boolean
    Three As Boolean
    Four As Boolean
    Five As Boolean
    Six As Boolean
End Type

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
     As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory _
     As String, ByVal nShowCmd As Long) As Long
     
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long

Public Function DiceUp(A As Integer, B As Integer)
    DICE(B).one = False
    DICE(B).Two = False
    DICE(B).Three = False
    DICE(B).Four = False
    DICE(B).Five = False
    DICE(B).Six = False
    If A = 1 Then
        DICE(B).one = True
        Exit Function
    End If
    If A = 2 Then
        DICE(B).Two = True
        Exit Function
    End If
    If A = 3 Then
        DICE(B).Three = True
        Exit Function
    End If
    If A = 4 Then
        DICE(B).Four = True
        Exit Function
    End If
    If A = 5 Then
        DICE(B).Five = True
        Exit Function
    End If
    If A = 6 Then
        DICE(B).Six = True
        Exit Function
    End If
End Function

Public Function SmStreight() As Boolean 'Is there a Small Streight?
Dim T As Integer
Dim Q As Integer
Dim IsThere As Boolean
    SmStreight = False
    For T = 1 To 3
        Q = T
        If DieSTAT(1) = Q Or DieSTAT(2) = Q Or DieSTAT(3) = Q Or DieSTAT(4) = Q Or DieSTAT(5) = Q Then
            Q = Q + 1
            If DieSTAT(1) = Q Or DieSTAT(2) = Q Or DieSTAT(3) = Q Or DieSTAT(4) = Q Or DieSTAT(5) = Q Then
                Q = Q + 1
                If DieSTAT(1) = Q Or DieSTAT(2) = Q Or DieSTAT(3) = Q Or DieSTAT(4) = Q Or DieSTAT(5) = Q Then
                    Q = Q + 1
                    If DieSTAT(1) = Q Or DieSTAT(2) = Q Or DieSTAT(3) = Q Or DieSTAT(4) = Q Or DieSTAT(5) = Q Then
                        SmStreight = True
                        Exit For
                    End If
                End If
            End If
        End If
    Next
End Function

Public Function LgStreight() As Boolean 'Is there a Large Streight?
Dim T As Integer
Dim Q As Integer
Dim IsThere As Boolean
    LgStreight = False
    For T = 1 To 2
        Q = T
        If DieSTAT(1) = Q Or DieSTAT(2) = Q Or DieSTAT(3) = Q Or DieSTAT(4) = Q Or DieSTAT(5) = Q Then
            Q = Q + 1
            If DieSTAT(1) = Q Or DieSTAT(2) = Q Or DieSTAT(3) = Q Or DieSTAT(4) = Q Or DieSTAT(5) = Q Then
                Q = Q + 1
                If DieSTAT(1) = Q Or DieSTAT(2) = Q Or DieSTAT(3) = Q Or DieSTAT(4) = Q Or DieSTAT(5) = Q Then
                    Q = Q + 1
                    If DieSTAT(1) = Q Or DieSTAT(2) = Q Or DieSTAT(3) = Q Or DieSTAT(4) = Q Or DieSTAT(5) = Q Then
                        Q = Q + 1
                        If DieSTAT(1) = Q Or DieSTAT(2) = Q Or DieSTAT(3) = Q Or DieSTAT(4) = Q Or DieSTAT(5) = Q Then
                            LgStreight = True
                            Exit For
                        End If
                    End If
                End If
            End If
        End If
    Next
End Function

Public Function FullHouse() As Boolean ' Is there a full house?
Dim hThree As Boolean
Dim hTwo As Boolean
Dim Tots(6) As Integer
Dim T As Integer
Dim Q As Integer
    hThree = False
    hTwo = False
    For T = 1 To 5
        For Q = 1 To 6
            If DieSTAT(T) = Q Then Tots(Q) = Tots(Q) + 1
        Next
    Next
    For T = 1 To 6
        If Tots(T) = 3 Then hThree = True
    Next
    For T = 1 To 6
        If Tots(T) = 2 Then hTwo = True
    Next
    If hThree = True And hTwo = True Then
        FullHouse = True
    Else
        FullHouse = False
    End If
End Function

Public Function FourKind() As Boolean
Dim Tots(6) As Integer
Dim T As Integer
Dim Q As Integer
    FourKind = False
    For T = 1 To 5
        For Q = 1 To 6
            If DieSTAT(T) = Q Then Tots(Q) = Tots(Q) + 1
        Next
    Next
    For T = 1 To 6
        If Tots(T) >= 4 Then FourKind = True
    Next
End Function

Public Function ThreeKind() As Boolean
Dim Tots(6) As Integer
Dim T As Integer
Dim Q As Integer
    ThreeKind = False
    For T = 1 To 5
        For Q = 1 To 6
            If DieSTAT(T) = Q Then Tots(Q) = Tots(Q) + 1
        Next
    Next
    For T = 1 To 6
        If Tots(T) >= 3 Then ThreeKind = True
    Next
End Function

Public Function GetHighScore() As Integer ' Retrieve High Score from file on HDD
Dim Y As String
Dim X As String
Dim Z As String
Dim T As Integer
On Error GoTo Error
    LE = False
    Open App.Path & "\YHS.bin" For Input As 1
        For T = 1 To 5
            Input #1, Y: ATS(T).Score = CInt(Y) 'Score
            Input #1, X: ATS(T).Name = X        'Name
            Input #1, X: ATS(T).Qnt = X         'Yahtzee Quantity
            Input #1, Z: ATS(T).Date = Z        'date
        Next
        Input #1, ColRTB(0)
        Input #1, ColRTB(1)
        Input #1, ColRTB(2)
        Input #1, ColRTB(3)
        Input #1, ColRTB(4)
        Input #1, Pref.SpeedYaht
        Input #1, Pref.SpeedValue
        Input #1, Pref.MovAss
        Input #1, Pref.Sounds
        Input #1, Pref.SndType
        Input #1, Pref.RndStop
    Close #1
    GetHighScore = ATS(1).Score
    If ATS(1).Name = "" Then
        Scores.HSHolder = "(Unknown)"
    Else
        Scores.HSHolder = ATS(1).Name
    End If
    AdjustPrefs
    Exit Function
Error:
    LE = True   'Load Error
    AdjustPrefs
    On Error Resume Next
    Close #1
    GetHighScore = 0
    Scores.HSHolder = "Nobody"
End Function

Public Function OpenYHS(mPath As String) As Boolean ' Import High Score from file on HDD
Dim Y As String
Dim X As String
Dim Z As String
Dim T As Integer
On Error GoTo Error
    Open mPath For Input As 1
        For T = 1 To 5
            Input #1, Y: ATS(T).Score = CInt(Y) 'Score
            Input #1, X: ATS(T).Name = X        'Name
            Input #1, X: ATS(T).Qnt = X         'Yahtzee Quantity
            Input #1, Z: ATS(T).Date = Z        'date
        Next
    Close #1
    OpenYHS = True
    AdjustPrefs
    Exit Function
Error:
    mError = Error.Number
    On Error Resume Next
    Close #1
    OpenYHS = False
End Function

Public Sub AdjustPrefs()
    If LE = False Then  'Load Error - True if unable to open HighScore File
        If Pref.SpeedYaht = 1 Then
            frmMain.mnuspeed.Checked = True
        Else
            frmMain.mnuspeed.Checked = False
        End If
        SPSetting = Pref.SpeedValue
        If SPSetting < 1 Or SPSetting > 10 Then
            SPSetting = 4
            Pref.SpeedValue = 4
        End If
        If Pref.MovAss = 1 Then
            frmMain.mnuSpotHelp.Checked = True
        Else
            frmMain.mnuSpotHelp.Checked = False
        End If
        If Pref.Sounds = 1 Then
            frmMain.mnuSound.Checked = True
        Else
            frmMain.mnuSound.Checked = False
        End If
        If Pref.SndType = 1 Then
            frmMain.mnuFight.Checked = True
            frmMain.mnuNormal.Checked = False
            FightSND = True
        Else
            frmMain.mnuFight.Checked = False
            frmMain.mnuNormal.Checked = True
            FightSND = False
        End If
        If Pref.RndStop = 1 Then
            frmMain.mnuRand.Checked = True
        Else
            frmMain.mnuRand.Checked = False
        End If
        CustomColor.Position = ColRTB(0)
        CustomColor.Score = ColRTB(1)
        CustomColor.Name = ColRTB(2)
        CustomColor.Date = ColRTB(3)
        CustomColor.YatCnt = ColRTB(4)
        If CustomColor.Position <> 0 Then
            CustFound = True
        Else
            CustFound = False
        End If
    Else
        LE = False
        Pref.SpeedYaht = 0
        frmMain.mnuspeed.Checked = False
        SPSetting = 4
        Pref.SpeedValue = 4
        Pref.MovAss = 1
        Pref.RndStop = 0
        frmMain.mnuSpotHelp.Checked = True
        Pref.Sounds = 1
        frmMain.mnuSound.Checked = True
        Pref.SndType = 0
        frmMain.mnuFight.Checked = False
        frmMain.mnuNormal.Checked = True
        FightSND = False
        CustFound = False
    End If
End Sub

Public Function OpenYHSmrg(mPath As String) As Boolean ' Import High Score to MERGE into our scores
Dim Y As String
Dim X As String
Dim Z As String
Dim T As Integer
On Error GoTo Error
    Open mPath For Input As 1
        For T = 1 To 5
            Input #1, Y: mATS(T).Score = CInt(Y) 'Score
            Input #1, X: mATS(T).Name = X        'Name
            Input #1, X: mATS(T).Qnt = X         'Yahtzee Quantity
            Input #1, Z: mATS(T).Date = Z        'date
        Next
    Close #1
    OpenYHSmrg = True
    Exit Function
Error:
    mError = Error.Number
    On Error Resume Next
    Close #1
    OpenYHSmrg = False
End Function

Public Sub SaveHighScore() ' Save High Score to file on HDD
Dim T As Integer
On Error GoTo Error
    If Scores.HSHolder = "" Then
        Scores.HSHolder = "(Unknown)"
    End If
    Open App.Path & "\YHS.bin" For Output As 1
        For T = 1 To 5
            Print #1, CStr(ATS(T).Score)
            Print #1, ATS(T).Name
            Print #1, ATS(T).Qnt
            Print #1, ATS(T).Date
        Next
        Print #1, CustomColor.Position
        Print #1, CustomColor.Score
        Print #1, CustomColor.Name
        Print #1, CustomColor.Date
        Print #1, CustomColor.YatCnt
        Print #1, Pref.SpeedYaht
        Print #1, Pref.SpeedValue
        Print #1, Pref.MovAss
        Print #1, Pref.Sounds
        Print #1, Pref.SndType
        Print #1, Pref.RndStop
    Close #1
    Exit Sub
Error:
    On Error Resume Next
    Close #1
    MsgBox "Save High Score function FAILED!    ", vbCritical, "Ken must have Goofed Up!"
End Sub

Public Function ExportYHS(mPath As String) As Boolean ' Export High Score
Dim T As Integer
On Error GoTo Error
    Open mPath For Output As 1
        For T = 1 To 5
            Print #1, CStr(ATS(T).Score)
            Print #1, ATS(T).Name
            Print #1, ATS(T).Qnt
            Print #1, ATS(T).Date
        Next
        Print #1, CustomColor.Position
        Print #1, CustomColor.Score
        Print #1, CustomColor.Name
        Print #1, CustomColor.Date
        Print #1, CustomColor.YatCnt
        Print #1, Pref.SpeedYaht
        Print #1, Pref.SpeedValue
        Print #1, Pref.MovAss
        Print #1, Pref.Sounds
        Print #1, Pref.SndType
        Print #1, Pref.RndStop
    Close #1
    ExportYHS = True
    Exit Function
Error:
    On Error Resume Next
    Close #1
    ExportYHS = False
End Function

Public Function GetCustomColors() As Boolean
On Error GoTo Error
Dim Y As String
Dim X As String
Dim Z As String
Dim T As Integer
    Open App.Path & "\YHS.bin" For Input As 1
        For T = 1 To 5
            Input #1, Y: ATS(T).Score = CInt(Y) 'Score
            Input #1, X: ATS(T).Name = X        'Name
            Input #1, X: ATS(T).Qnt = X
            Input #1, Z: ATS(T).Date = Z        'date
        Next
        Input #1, ColRTB(0)
        Input #1, ColRTB(1)
        Input #1, ColRTB(2)
        Input #1, ColRTB(3)
        Input #1, ColRTB(4)
        
        Input #1, Pref.SpeedYaht
        Input #1, Pref.SpeedValue
        Input #1, Pref.MovAss
        Input #1, Pref.Sounds
        Input #1, Pref.SndType
        Input #1, Pref.RndStop
    Close #1
    GetCustomColors = True
    Exit Function
Error:
    On Error Resume Next
    Close #1
    GetCustomColors = False
End Function

Public Function WriteCustomColors(A As ColorConstants, B As ColorConstants, C As ColorConstants, D As ColorConstants, E As ColorConstants) As Boolean
On Error GoTo Error
Dim Y As String
Dim X As String
Dim Z As String
Dim T As Integer
    Open App.Path & "\YHS.bin" For Output As 1
        For T = 1 To 5
            Print #1, CStr(ATS(T).Score)
            Print #1, ATS(T).Name
            Print #1, ATS(T).Qnt
            Print #1, ATS(T).Date
        Next
        Print #1, A
        Print #1, B
        Print #1, C
        Print #1, D
        Print #1, E
        Print #1, Pref.SpeedYaht
        Print #1, Pref.SpeedValue
        Print #1, Pref.MovAss
        Print #1, Pref.Sounds
        Print #1, Pref.SndType
        Print #1, Pref.RndStop
    Close #1
    WriteCustomColors = True
    Exit Function
Error:
    On Error Resume Next
    Close #1
    WriteCustomColors = False
End Function

Public Function GoodTrip(Index As Integer) As Boolean 'Show pink or green icon
Dim T As Integer
Dim RESP As Integer
    For T = 0 To 5
        If DieSTAT(T) = Index Then
            RESP = RESP + Index
        End If
    Next
    If RESP >= (Index * 3) Then
        GoodTrip = True
    Else
        GoodTrip = False
    End If
End Function

Public Function Qtally(Index As Integer) As Integer 'Advise player of possible score if played
Dim T As Integer
Dim RESP As Integer
    For T = 0 To 5
        If DieSTAT(T) = Index Then
            RESP = RESP + Index
        End If
    Next
    If Joker Then
        Qtally = (Index * 5)
    Else
        Qtally = RESP
    End If
End Function

Public Sub GetColors(R As ColorConstants)
Dim L As String
    L = CStr(Hex(R))
    If Len(L) > 4 Then
        CL.Blue = HexToDec(Left$(L, 2))
        CL.Green = HexToDec(Mid$(L, 3, 2))
        CL.Red = HexToDec(Right$(L, 2))
    ElseIf Len(L) > 2 Then
        CL.Blue = 0
        CL.Green = HexToDec(Left$(L, 2))
        CL.Red = HexToDec(Right$(L, 2))
    Else
        CL.Blue = 0
        CL.Green = 0
        CL.Red = HexToDec(L)
    End If
End Sub

Public Function HexToDec(HexValue As String) As Integer
    HexToDec = Val("&H" & HexValue)
End Function

Public Sub MergeEm()    ' This sub Merges a HighScore file into your existing HighScore file (Mixer).
Dim mTMP(5) As HScores  ' And of course, it sorts them as well.
Dim T As Integer
Dim H As Integer
Dim K As Integer
Dim Q As Integer
Dim mPOS(5) As Boolean
Dim yPOS(5) As Boolean
    For T = 1 To 5
        mPOS(T) = False
        yPOS(T) = False
    Next
    For T = 1 To 5
        For H = 1 To 5
            If mATS(H).Score > ATS(T).Score Then
                If mATS(H).Score > mTMP(T).Score Then
                    If yPOS(H) = False Then
                        mTMP(T).Score = mATS(H).Score
                        mTMP(T).Name = mATS(H).Name
                        mTMP(T).Qnt = mATS(H).Qnt
                        mTMP(T).Date = mATS(H).Date
                        mPOS(T) = True
                        K = H
                    End If
                End If
            End If
        Next
        If mPOS(T) = False Then
            mTMP(T).Score = ATS(T).Score
            mTMP(T).Name = ATS(T).Name
            mTMP(T).Qnt = ATS(T).Qnt
            mTMP(T).Date = ATS(T).Date
        Else
            For Q = 5 To T Step -1
                ATS(Q).Date = ATS(Q - 1).Date: ATS(Q).Score = ATS(Q - 1).Score: ATS(Q).Name = ATS(Q - 1).Name: ATS(Q).Qnt = ATS(Q - 1).Qnt
            Next
        End If
        yPOS(K) = True
    Next
    For T = 1 To 5
        ATS(T).Score = mTMP(T).Score
        ATS(T).Name = mTMP(T).Name
        ATS(T).Qnt = mTMP(T).Qnt
        ATS(T).Date = mTMP(T).Date
    Next
    For T = 1 To 5  'Clean-Up
        If ATS(T).Score < 7 Then
            ATS(T).Score = 0
            ATS(T).Name = ""
            ATS(T).Qnt = 0
            ATS(T).Date = ""
        End If
    Next
End Sub

Public Sub SortEm() ' This Sub sorts HighScore Data into the Proper Positions.
Dim mTMP(5) As HScores
Dim HS As Integer
Dim T As Integer
Dim H As Integer
Dim K As Integer
Dim mPOS(5) As Boolean
    For T = 1 To 5
        mPOS(T) = False
    Next
    H = 1
    For T = 1 To 5
        If ATS(T).Score = 0 Then ATS(T).Score = H: H = H + 1
    Next
    For T = 1 To 5
        For H = 1 To 5
            If mPOS(H) = False Then
                If mTMP(T).Score < ATS(H).Score Then
                    K = H
                    mTMP(T).Score = ATS(H).Score
                    mTMP(T).Name = ATS(H).Name
                    mTMP(T).Qnt = ATS(H).Qnt
                    mTMP(T).Date = ATS(H).Date
                End If
            End If
        Next
        mPOS(K) = True
    Next
    For T = 1 To 5
        ATS(T).Score = mTMP(T).Score
        ATS(T).Name = mTMP(T).Name
        ATS(T).Qnt = mTMP(T).Qnt
        ATS(T).Date = mTMP(T).Date
    Next
    For T = 1 To 5  'Clean-Up
        If ATS(T).Score < 7 Then
            ATS(T).Score = 0
            ATS(T).Name = ""
            ATS(T).Qnt = 0
            ATS(T).Date = ""
        End If
    Next
End Sub

Public Function iPlay(Tune As String) As Boolean
    If frmMain.mnuSound.Checked = True Then
        If Dir(App.Path & "\Snds\" & Tune) = Tune Then  'NOTE**   File names are CASE SENSITIVE!!!!
            Clik = sndPlaySound(App.Path & "\Snds\" & Tune, 1)
            iPlay = True
        Else
            iPlay = False
        End If
    End If
End Function

Public Sub SelectSnd() ' Randomly select a Punch Sound Effect
Dim K As Integer
    If FightSND Then
        Randomize
        K = Int(Rnd * 5)
        Select Case K
            Case Is = 0
                iPlay ("PUNCH1.WAV")
            Case Is = 1
                iPlay ("PUNCH2.WAV")
            Case Is = 2
                iPlay ("PUNCH3.WAV")
            Case Is = 3
                iPlay ("PUNCH4.WAV")
            Case Is = 4
                iPlay ("PUNCH5.WAV")
        End Select
    Else
        iPlay ("ModClick3.wav")
    End If
End Sub

Public Sub SelectSlp() ' Randomly select a Slap Sound Effect
Dim K As Integer
    If FightSND Then
        Randomize
        K = Int(Rnd * 4)
        Select Case K
            Case Is = 0
                iPlay ("SLAP1.WAV")
            Case Is = 1
                iPlay ("SLAP2.WAV")
            Case Is = 2
                iPlay ("SLAP3.WAV")
            Case Is = 3
                iPlay ("SLAP4.WAV")
        End Select
    Else
        iPlay ("ModClick6.wav")
    End If
End Sub

Public Sub RollSoundProcessor() ' Play Dice Rolling Sound Effects
Dim T As Integer
    T = 0
    If DieSELECT(1) = False Then T = 1      'To best handle this routine, sound effects
    If DieSELECT(2) = False Then T = T + 1  'have been custom sized to match the time it
    If DieSELECT(3) = False Then T = T + 1  'takes for each roll to finish, based on how
    If DieSELECT(4) = False Then T = T + 1  'many dice are currently rolling.
    If DieSELECT(5) = False Then T = T + 1  'If you change these, you'll need to custom
    If T = 0 Then Exit Sub                  'size your new sound effects to sync them with
    Select Case T                           'the rolling action of the die.
        Case Is = 1
            iPlay ("OneDieRoll.WAV")
        Case Is = 2
            iPlay ("TwoDieRoll.WAV")
        Case Is = 3
            iPlay ("ThreeDieRoll.WAV")
        Case Is = 4
            iPlay ("FourDieRoll.WAV")
        Case Is = 5
            iPlay ("FiveDieRoll.WAV")
    End Select
End Sub
