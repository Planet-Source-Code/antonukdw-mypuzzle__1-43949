VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "MyPuzzle"
   ClientHeight    =   5385
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   5850
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   359
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   600
      ScaleHeight     =   129
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   6
      Top             =   3120
      Width           =   1935
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   0
         X2              =   128
         Y1              =   128
         Y2              =   128
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   128
         X2              =   129
         Y1              =   0
         Y2              =   257
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   128
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   128
      End
      Begin VB.Label lblElapsed 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1530
         Width           =   1455
      End
      Begin VB.Image imgMaster 
         Height          =   975
         Left            =   240
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   1215
         Left            =   120
         Top             =   120
         Width           =   1695
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00C0C0C0&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   120
         Top             =   1440
         Width           =   1695
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3240
      Top             =   3120
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   3720
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.PictureBox picMaster 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   1200
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   0
      Top             =   1800
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox picMasks 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   600
      Left            =   360
      ScaleHeight     =   40
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   44
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   660
   End
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   2640
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picPat 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   3960
      ScaleHeight     =   80
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   80
      TabIndex        =   2
      Top             =   1800
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   -45
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   3
      Top             =   -30
      Width           =   5175
   End
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4935
      Left            =   0
      ScaleHeight     =   329
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   345
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   5175
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNew 
         Caption         =   "&New Game"
         Begin VB.Menu mnuNewGame 
            Caption         =   "&Easy"
            Index           =   0
         End
         Begin VB.Menu mnuNewGame 
            Caption         =   "&Normal"
            Index           =   1
         End
         Begin VB.Menu mnuNewGame 
            Caption         =   "&Hard"
            Index           =   2
         End
      End
      Begin VB.Menu mnuLoadGame 
         Caption         =   "&Load Game"
      End
      Begin VB.Menu mnuSaveGame 
         Caption         =   "&Save Game"
      End
      Begin VB.Menu mnu0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuHints 
         Caption         =   "&Hints"
      End
      Begin VB.Menu mnuShuffle 
         Caption         =   "Sh&uffle"
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSounds 
         Caption         =   "&Sounds"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuStatus 
         Caption         =   "S&tatus Window"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuBackground 
         Caption         =   "&Background Color"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMoveBoard 
         Caption         =   "&Move Board"
         Begin VB.Menu mnuMove 
            Caption         =   "Move &Left"
            Index           =   0
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuMove 
            Caption         =   "Move &Right"
            Index           =   1
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuMove 
            Caption         =   "Move &Up"
            Index           =   2
            Shortcut        =   ^W
         End
         Begin VB.Menu mnuMove 
            Caption         =   "Move &Down"
            Index           =   3
            Shortcut        =   ^Z
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHowToPlay 
         Caption         =   "&How To Play"
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim XLast           As Integer  'last x mouse pointer down
Dim YLast           As Integer  'last y mouse pointer down
Dim GameStart       As Boolean  'game status
Dim FileName        As String   'current picture file name
Dim ElapsedSeconds  As Long     'time elapsed

Private Sub Form_Load()
    Randomize Timer
    
    'load sound from resource - MSDN
    Sound1 = LoadResData("CLICK", "WAVE")
    Sound2 = LoadResData("ROTATE", "WAVE")
    Sound3 = LoadResData("COMPLETE", "WAVE")
    
    'load masks picture
    picMasks.Picture = LoadResPicture("MASK", vbResBitmap)
    SavePicture picMasks.Picture, "c:\mask.bmp"
    
    'set game status
    GameStart = False
    
    AppPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
    
    'load config
    Dim C1 As Integer, C2 As Integer, C3 As Long
    On Error GoTo errHandle
    Open AppPath & "MyPuzzle.cfg" For Input As #1
    Input #1, C1, C2, C3
    mnuSounds.Checked = (C1 = 1)
    mnuStatus.Checked = (C2 = 1)
    picStatus.Visible = (C2 = 1)
    picScreen.BackColor = C3
    picBG.BackColor = C3
errHandle:
    Close #1
End Sub

Private Sub Form_Resize()
    'adjust size and position
    picBG.Move 0, 0, ScaleWidth, ScaleHeight
    picScreen.Move 0, 0, ScaleWidth, ScaleHeight
    picStatus.Move ScaleWidth - picStatus.Width - 30, ScaleHeight - picStatus.Height - 30
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'save config
    On Error GoTo errHandle
    Open AppPath & "MyPuzzle.cfg" For Output As #1
    Print #1, IIf(mnuSounds.Checked, 1, 0), IIf(mnuStatus.Checked, 1, 0), picScreen.BackColor
errHandle:
    Close #1
End Sub

Private Sub mnuAbout_Click()
    MsgBox LoadResString(201) & vbNewLine & vbNewLine _
        & LoadResString(202) & vbNewLine _
        & LoadResString(203) & vbNewLine & vbNewLine _
        & LoadResString(204) & vbNewLine _
        & LoadResString(205) & vbNewLine, vbInformation
End Sub

Private Sub mnuBackground_Click()
    On Error GoTo errHandle
    Dialog.ShowColor
    picBG.BackColor = Dialog.Color
    picScreen.BackColor = Dialog.Color
    DrawBoard picScreen, picBG
errHandle:
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuHints_Click()
    Dim i As Integer
    Dim x As Integer
    Dim y As Integer
    
    For i = 0 To UBound(Pattern)
        If Pattern(i).Col = 0 And Pattern(i).Row = 0 And Pattern(i).Rotation = 0 Then
            'find left and top of row,col (0,0)
            x = Pattern(i).x
            y = Pattern(i).y
            Exit For
        End If
    Next
    
    For i = 0 To UBound(Pattern)
        If Pattern(i).Rotation <> 0 _
            Or x + Pattern(i).Col * RectSize <> Pattern(i).x _
            Or y + Pattern(i).Row * RectSize <> Pattern(i).y Then
            picScreen.Line (x + Pattern(i).Col * RectSize + (PatternSize - RectSize) \ 2, y + Pattern(i).Row * RectSize + (PatternSize - RectSize) \ 2)-Step(RectSize, RectSize), vbYellow, B
            picScreen.Line (Pattern(i).x + (PatternSize - RectSize) \ 2, Pattern(i).y + (PatternSize - RectSize) \ 2)-Step(RectSize, RectSize), vbYellow, B
            Exit For
        End If
    Next
End Sub

Private Sub mnuLoadGame_Click()
    Dim i As Integer
    Dim r As Integer
    Dim M As Integer
    On Error GoTo errHandle
    
    Timer1.Enabled = False
    Dialog.FileName = ""
    Dialog.DialogTitle = "Load Game"
    Dialog.Filter = "My Puzzles|*.MPZ|All Files|*.*"
    Dialog.ShowOpen

    Me.MousePointer = vbHourglass
    Open Dialog.FileName For Input As #1
    Input #1, ElapsedSeconds
    Input #1, FileName
    Input #1, Level
    
    picMaster.Picture = LoadPicture(FileName)
    
    'calculate max picture size
    WindowState = vbMaximized
    If picMaster.Width > ScaleWidth - PatternSize * 2 Then picMaster.Width = ScaleWidth - PatternSize * 2
    If picMaster.Height > ScaleHeight - PatternSize * 2 Then picMaster.Height = ScaleHeight - PatternSize * 2
    Set imgMaster.Picture = picMaster.Picture
    
    'create patterns table
    InitBoard picMaster, picTmp, picMasks, picPat
    
    'read patterns attributes
    Do While Not EOF(1)
        Input #1, Pattern(i).x, Pattern(i).y, Pattern(i).Col, Pattern(i).Row, r, Pattern(i).Pic, M
        'adjust rotation
        Do While Pattern(i).Rotation <> r
            RotatePattern i, picTmp
        Loop
        i = i + 1
    Loop
    Close #1
    
    DrawBoard picScreen, picBG
    GameStart = True
    
errHandle:
    Me.MousePointer = vbCustom
    Timer1.Enabled = GameStart
End Sub

Private Sub mnuMove_Click(Index As Integer)
    Dim i As Integer
    Me.MousePointer = vbHourglass
    For i = 0 To UBound(Pattern)
        If Index = 0 Then Pattern(i).x = Pattern(i).x - RectSize
        If Index = 1 Then Pattern(i).x = Pattern(i).x + RectSize
        If Index = 2 Then Pattern(i).y = Pattern(i).y - RectSize
        If Index = 3 Then Pattern(i).y = Pattern(i).y + RectSize
    Next
    DrawBoard picScreen, picBG
    Me.MousePointer = vbCustom
End Sub

Private Sub mnuNewGame_Click(Index As Integer)
    On Error GoTo errHandle
    
    Timer1.Enabled = False
    Dialog.FileName = FileName
    Dialog.DialogTitle = "Select Picture"
    Dialog.Filter = "Picture Files|*.gif;*.jpg;*.bmp|All Files|*.*"
    Dialog.ShowOpen
    
    FileName = Dialog.FileName
    Me.MousePointer = vbHourglass
    picMaster.Picture = LoadPicture(FileName)
    
    'cancel if picture too small
    If picMaster.Width < RectSize * 2 Or picMaster.Height < RectSize * 2 Then
        MsgBox "Picture too small!", vbCritical
    Else
        'reset elapsed time
        ElapsedSeconds = 0
        
        'calculate max picture size
        WindowState = vbMaximized
        If picMaster.Width > ScaleWidth - PatternSize * 2 Then picMaster.Width = ScaleWidth - PatternSize * 2
        If picMaster.Height > ScaleHeight - PatternSize * 2 Then picMaster.Height = ScaleHeight - PatternSize * 2
        Set imgMaster.Picture = picMaster.Picture
        
        Level = Index
        InitBoard picMaster, picTmp, picMasks, picPat
        ShuffleBoard picScreen, picTmp
        DrawBoard picScreen, picBG
        SelPattern = -1
        GameStart = True
    End If
    
errHandle:
    Me.MousePointer = vbCustom
    Timer1.Enabled = GameStart
End Sub

Private Sub mnuGame_Click()
    mnuSaveGame.Enabled = GameStart
End Sub

Private Sub mnuOptions_Click()
    mnuHints.Enabled = GameStart
    mnuShuffle.Enabled = picPat.UBound > 0
    mnuMoveBoard.Enabled = picPat.UBound > 0
End Sub

Private Sub mnuHowToPlay_Click()
    MsgBox LoadResString(101) & vbNewLine _
        & LoadResString(102) & vbNewLine _
        & LoadResString(103) & vbNewLine & vbNewLine _
        & LoadResString(104) & vbNewLine _
        & LoadResString(105) & vbNewLine _
        & LoadResString(106) & vbNewLine _
        & LoadResString(107), vbInformation, "How to Play"
End Sub

Private Sub mnuSaveGame_Click()
    Dim P As PatternData
    Dim i As Integer
    Dim j As Integer
    On Error GoTo errHandle
    
    Timer1.Enabled = False
    Dialog.FileName = ""
    Dialog.DialogTitle = "Save Game"
    Dialog.Filter = "My Puzzles|*.MPZ|All Files|*.*"
    Dialog.ShowSave

    Me.MousePointer = vbHourglass
    
    'sort pattern by col and row
    'in order to get correct picture index when load saved data
    For i = 0 To UBound(Pattern)
        For j = i + 1 To UBound(Pattern)
            If Pattern(i).Row * 100 + Pattern(i).Col > Pattern(j).Row * 100 + Pattern(j).Col Then
                P = Pattern(i)
                Pattern(i) = Pattern(j)
                Pattern(j) = P
            End If
        Next
    Next
    
    'write data
    Open Dialog.FileName For Output As #1
    Print #1, ElapsedSeconds
    Print #1, FileName
    Print #1, Level
    For i = 0 To UBound(Pattern)
        Print #1, Pattern(i).x, Pattern(i).y, Pattern(i).Col, Pattern(i).Row, Pattern(i).Rotation, Pattern(i).Pic, Pattern(i).Mask
    Next
    Close #1
    
errHandle:
    Me.MousePointer = vbCustom
    Timer1.Enabled = GameStart
End Sub

Private Sub mnuShuffle_Click()
    If MsgBox("Shuffle patterns?", vbYesNo + vbQuestion) = vbYes Then
        Me.MousePointer = vbHourglass
        ShuffleBoard picScreen, picTmp
        DrawBoard picScreen, picBG
        ElapsedSeconds = 0
        GameStart = True
        Me.MousePointer = vbCustom
    End If
End Sub

Private Sub mnuSounds_Click()
    mnuSounds.Checked = Not mnuSounds.Checked
End Sub

Private Sub mnuStatus_Click()
    mnuStatus.Checked = Not mnuStatus.Checked
    picStatus.Visible = mnuStatus.Checked
End Sub

Private Sub picScreen_DragDrop(Source As Control, x As Single, y As Single)
    'move status window
    If Source Is picStatus Then picStatus.Move x - XLast, y - YLast
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim j As Integer
    Dim P As PatternData
    
    If Not GameStart Then Exit Sub
    
    'left click
    If Button = 1 Then
        For i = 0 To UBound(Pattern)
            If x >= Pattern(i).x And x <= Pattern(i).x + PatternSize And y >= Pattern(i).y And y <= Pattern(i).y + PatternSize Then
                If CheckInside(i, x, y) Then
                    'save current pattern
                    P = Pattern(i)
                    
                    'move pattern center at mouse pointer
                    P.x = x - PatternSize \ 2
                    P.y = y - PatternSize \ 2
                    
                    'rearrange z-order (0 = topmost)
                    For j = i - 1 To 0 Step -1
                        Pattern(j + 1) = Pattern(j)
                    Next
                    
                    'put current pattern as topmost
                    Pattern(0) = P
                    SelPattern = 0
                    
                    'save last mouse pointer position
                    XLast = x
                    YLast = y
                    
                    DrawBoard picScreen, picBG
                    If mnuSounds.Checked Then sndPlaySound Sound1(0), SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY

                    Exit Sub
                End If
            End If
        Next
    ElseIf SelPattern >= 0 Then
        'rotate when drag and right click
        RotatePattern SelPattern, picTmp
        DrawBoard picScreen, picBG
        If mnuSounds.Checked Then sndPlaySound Sound2(0), SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY
    End If
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not GameStart Then Exit Sub
    
    If SelPattern >= 0 Then
        If XLast <> x Or YLast <> y Then
            'move pattern center at mouse pointer
            Pattern(SelPattern).x = x - PatternSize \ 2
            Pattern(SelPattern).y = y - PatternSize \ 2
            
            'save position
            XLast = x
            YLast = y
            DrawBoard picScreen, picBG
        End If
    End If
End Sub

Private Sub picScreen_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    Dim Complete As Boolean
    
    If Not GameStart Then Exit Sub
    
    If Button = 1 And SelPattern >= 0 Then
        'snap to grid
        x = Round((x - PatternSize \ 2) / (RectSize \ 4)) * (RectSize \ 4)
        y = Round((y - PatternSize \ 2) / (RectSize \ 4)) * (RectSize \ 4)
        Pattern(SelPattern).x = x
        Pattern(SelPattern).y = y
        
        SelPattern = -1
        DrawBoard picScreen, picBG
        If mnuSounds.Checked Then sndPlaySound Sound1(0), SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY
        
        'check complete
        For i = 0 To UBound(Pattern)
            If Pattern(i).Col = 0 And Pattern(i).Row = 0 And Pattern(i).Rotation = 0 Then
                'find left and top of row,col (0,0)
                x = Pattern(i).x
                y = Pattern(i).y
                Exit For
            End If
        Next
        
        Complete = True
        For i = 0 To UBound(Pattern)
            If Pattern(i).Rotation <> 0 _
                Or x + Pattern(i).Col * RectSize <> Pattern(i).x _
                Or y + Pattern(i).Row * RectSize <> Pattern(i).y Then
                Complete = False
                Exit For
            End If
        Next
        
        If Complete Then
            'end of game
            GameStart = False
            Timer1.Enabled = False
            If mnuSounds.Checked Then sndPlaySound Sound3(0), SND_NODEFAULT Or SND_ASYNC Or SND_MEMORY
            MsgBox "Puzzle Completed! You're awesome!", vbExclamation
        End If
    End If
End Sub

Private Sub picStatus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    XLast = x
    YLast = y
    picStatus.Drag
End Sub

Private Sub Timer1_Timer()
    Dim T As Date
    Dim M As Integer
    Dim S As Integer
    
    'increase elapsed time
    ElapsedSeconds = ElapsedSeconds + 1
    
    'show elapsed time in status window
    T = TimeSerial(0, 0, ElapsedSeconds)
    lblElapsed.Caption = IIf(Level = 0, "Easy", IIf(Level = 1, "Normal", "Hard")) & " - " & Format(T, "hh:nn:ss")
End Sub
