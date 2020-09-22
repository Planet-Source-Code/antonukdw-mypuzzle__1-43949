Attribute VB_Name = "PuzzleModule"
Option Explicit

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long

'PlaySound const
Public Const SND_SYNC = &H0        ' Play synchronously (default)
Public Const SND_NODEFAULT = &H2   ' Don't use default sound
Public Const SND_MEMORY = &H4      ' lpszSoundName points to a
Public Const SND_LOOP = &H8        ' Loop the sound until next
Public Const SND_NOSTOP = &H10     ' Don't stop any currently
Public Const SND_ASYNC = &H1         '  play asynchronously

'Copy Mode
Public Const BLACKNESS = &H42           ' (DWORD) dest = BLACK
Public Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)
Public Const MERGECOPY = &HC000CA       ' (DWORD) dest = (source AND pattern)
Public Const MERGEPAINT = &HBB0226      ' (DWORD) dest = (NOT source) OR dest
Public Const NOTSRCCOPY = &H330008      ' (DWORD) dest = (NOT source)
Public Const NOTSRCERASE = &H1100A6     ' (DWORD) dest = (NOT src) AND (NOT dest)
Public Const PATCOPY = &HF00021         ' (DWORD) dest = pattern
Public Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Public Const PATPAINT = &HFB0A09        ' (DWORD) dest = DPSnoo
Public Const SRCAND = &H8800C6          ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020         ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const WHITENESS = &HFF0062       ' (DWORD) dest = WHITE

'Puzzle Const
Public Const PatternSize = 80
Public Const RectSize = 44

Public Type PatternData
    X As Integer
    Y As Integer
    Rotation As Integer
    Mask As Integer
    Row As Integer
    Col As Integer
    Pic As Integer
End Type

Public Pattern() As PatternData
Public SelPattern As Integer
Public Level As Integer
Public Sound1() As Byte
Public Sound2() As Byte
Public Sound3() As Byte
Public AppPath As String

Public Sub ShuffleBoard(PicSrc As PictureBox, picTmp As PictureBox)
    Dim i As Integer
    Dim j As Integer
    Dim n As Integer
    
    For i = 0 To UBound(Pattern)
        Pattern(i).X = Int(Rnd() * (PicSrc.Width - PatternSize * 3)) + PatternSize
        Pattern(i).Y = Int(Rnd() * (PicSrc.Height - PatternSize * 3)) + PatternSize
        
        'snap to grid
        Pattern(i).X = Round((Pattern(i).X - PatternSize \ 4) / (RectSize \ 4)) * (RectSize \ 4)
        Pattern(i).Y = Round((Pattern(i).Y - PatternSize \ 4) / (RectSize \ 4)) * (RectSize \ 4)
    Next

    'add some rotation
    n = Int(Rnd() * UBound(Pattern) * IIf(Level = 0, 0.2, IIf(Level = 1, 0.5, 0.7)))
    For i = 0 To n
        j = Int(Rnd() * UBound(Pattern))
        RotatePattern j, picTmp
    Next
End Sub

Public Sub InitBoard(PicSrc As PictureBox, picTmp As PictureBox, picMsk As PictureBox, picPat As Object)
    Dim Pic As Integer
    Dim Dest As Long
    Dim msk As Integer
    Dim mskCtr As Integer
    Dim Row As Integer
    Dim Col As Integer
    Dim MaxRow As Integer
    Dim MaxCol As Integer
    Dim w As Long
    Dim h As Long
    Dim ret As Long
    Dim X As Integer
    Dim Y As Integer
    
    'unload existing pattern
    For X = 1 To frmMain.picPat.UBound
        Unload frmMain.picPat(X)
    Next
    
    'prepare vars
    w = PicSrc.Width
    h = PicSrc.Height
    
    Row = 0     'pattern row number
    Col = 0     'pattern column number
    msk = 0     'index reference of pattern mask picture
    mskCtr = 0  'toggle counter 1,0
    Pic = 0     'index reference of pattern picture
    
    'makesure the master size is correct
    w = w - (w Mod RectSize)
    h = h - (h Mod RectSize)
    MaxCol = ((w) \ RectSize)
    MaxRow = ((h) \ RectSize)
    
    Do
        'get mask picture
        If Level = 0 Then
            msk = 18
        Else
            If Row = 0 And Col = 0 Then 'top left
                msk = IIf(mskCtr = 0, 2, 9)
            ElseIf Row = 0 And Col = MaxCol Then 'top right
                msk = IIf(mskCtr = 0, 6, 3)
            ElseIf Row = MaxRow And Col = 0 Then 'bottom left
                msk = IIf(mskCtr = 0, 8, 5)
            ElseIf Row = MaxRow And Col = MaxCol Then 'bottom right
                msk = IIf(mskCtr = 0, 4, 7)
            ElseIf Row = 0 Then 'top middle
                msk = IIf(mskCtr = 0, 10, 14)
            ElseIf Row = MaxRow Then 'bottom middle
                msk = IIf(mskCtr = 0, 12, 16)
            ElseIf Col = 0 Then 'left middle
                msk = IIf(mskCtr = 0, 17, 13)
            ElseIf Col = MaxCol Then 'right middle
                msk = IIf(mskCtr = 0, 15, 11)
            Else
                msk = IIf(mskCtr = 0, 0, 1)
            End If
        End If
        
        'add new pattern
        ReDim Preserve Pattern(Pic)
        Pattern(Pic).Col = Col
        Pattern(Pic).Row = Row
        Pattern(Pic).Mask = msk
        Pattern(Pic).Rotation = 0
        Pattern(Pic).Pic = Pic
        
        If Pic > 0 Then Load frmMain.picPat(Pic)
        
        'get position of pattern from master picture
        X = Col * RectSize - (PatternSize - RectSize) \ 2
        Y = Row * RectSize - (PatternSize - RectSize) \ 2
        
        Dest = picPat(Pic).hdc
        ret = BitBlt(Dest, 0, 0, PatternSize, PatternSize, PicSrc.hdc, X, Y, SRCCOPY)
        ret = BitBlt(Dest, 0, 0, PatternSize, PatternSize, picMsk.hdc, msk * PatternSize, 0, MERGEPAINT)
        
        'next col, pic
        mskCtr = (mskCtr + 1) Mod 2
        Col = Col + 1
        Pic = Pic + 1
        If Col > MaxCol Then
            Col = 0
            Row = Row + 1
            mskCtr = Row Mod 2
        End If
    Loop Until Row > MaxRow
End Sub

Public Sub DrawBoard(picScreen As PictureBox, picBG As PictureBox)
    Dim i As Integer
    Dim ret As Long
    
    'LockWindowUpdate picScreen.hdc
    'picScreen.Visible = False
    'put background
    ret = BitBlt(picScreen.hdc, 0, 0, picScreen.Width, picScreen.Height, picBG.hdc, 0, 0, SRCCOPY)
    
    'put some shadow
    For i = UBound(Pattern) To 0 Step -1
        If i <> SelPattern Then ret = BitBlt(picScreen.hdc, Pattern(i).X + 1, Pattern(i).Y + 1, PatternSize, PatternSize, frmMain.picMasks.hdc, Pattern(i).Mask * PatternSize, 0, SRCINVERT)
    Next
    
    'put pattern must be order by z-order
    For i = UBound(Pattern) To 0 Step -1
        If i = SelPattern Then ret = BitBlt(picScreen.hdc, Pattern(i).X + 1, Pattern(i).Y + 1, PatternSize, PatternSize, frmMain.picMasks.hdc, Pattern(i).Mask * PatternSize, 0, SRCINVERT)
        ret = BitBlt(picScreen.hdc, Pattern(i).X, Pattern(i).Y, PatternSize, PatternSize, frmMain.picMasks.hdc, Pattern(i).Mask * PatternSize, 0, SRCPAINT)
        ret = BitBlt(picScreen.hdc, Pattern(i).X, Pattern(i).Y, PatternSize, PatternSize, frmMain.picPat(Pattern(i).Pic).hdc, 0, 0, SRCAND)
    Next
    picScreen.Refresh
    'LockWindowUpdate 0
End Sub

Public Sub RotatePattern(Index As Integer, picTmp As PictureBox)
    Dim X As Long
    Dim Y As Long
    Dim c As Long
    Dim S As Long
    Dim Src As Long
    Dim Tmp As Long
    Dim ret As Long
    
    Src = frmMain.picPat(Pattern(Index).Pic).hdc
    Tmp = picTmp.hdc
    
    ret = BitBlt(Tmp, 0, 0, PatternSize, PatternSize, Src, 0, 0, SRCCOPY)
    S = PatternSize - 1
    
    For X = 0 To S
        For Y = 0 To S
            c = GetPixel(Tmp, X, Y)
            ret = SetPixel(Src, S - Y, X, c)
        Next
    Next

    Pattern(Index).Rotation = (Pattern(Index).Rotation + 1) Mod 4
    
    Select Case Pattern(Index).Mask
        Case 0, 2 To 4, 6 To 8, 10 To 12, 14 To 16
            Pattern(Index).Mask = Pattern(Index).Mask + 1
        Case 1: Pattern(Index).Mask = 0
        Case 5, 9, 13, 17: Pattern(Index).Mask = Pattern(Index).Mask - 3
    End Select
End Sub

Public Function CheckInside(ByVal Index As Integer, ByVal X As Integer, ByVal Y As Integer) As Boolean
    X = X - Pattern(Index).X
    Y = Y - Pattern(Index).Y
    
    CheckInside = frmMain.picMasks.Point(Pattern(Index).Mask * PatternSize + X, Y) <> vbBlack
End Function
