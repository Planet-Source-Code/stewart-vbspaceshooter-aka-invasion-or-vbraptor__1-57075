Attribute VB_Name = "modGame"
Option Explicit


Private ScrollY As Single

Dim TempTime As Long

Private tGrdX As Integer, tGrdY As Integer
Private Tiles2() As Tile_Data



'Some sound buffers
Public DDS_Shot As DirectSoundBuffer
Public DDS_ExplodeW As DirectSoundBuffer
Public DDS_Thrust As DirectSoundBuffer
Public DDS_Menu As DirectSoundBuffer

Public clsDDraw As New cDDraw
Public clsDSound As New cDSound

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Type POINTAPI
        x As Long
        y As Long
End Type


Public mMoney(20) As Money


'Runtime framerate stuff
Dim framesDone As Integer, LastTimeChecked As Long, FrameText As String

Private rYou As RECT, rBack1 As RECT, rCursor As RECT, rBullet As RECT, rBottom As RECT
Private rE1 As RECT, rHit As RECT, rHealth As RECT, rExplode As RECT, rHull As RECT
Private rW1 As RECT, rW2 As RECT, rB1 As RECT, rB2 As RECT, rB3 As RECT, rB4 As RECT
Private rMB As RECT, rTiles As RECT, rMissile As RECT, rNewBack As RECT, rHanger As RECT
Private rHBar As RECT, rEnergy As RECT, rHeader As RECT, rLetters As RECT, rMoney As RECT

Private bLeft As Boolean, bRight As Boolean, bForward As Boolean

Private LastMenuMoveTick As Long
Private Const MenuMoveTick = 120

Dim mlngTimer As Long
Dim mlngFrameTimer As Long
Dim mintFPSCounter As Long
Dim mintFPS As Long

Private Sub Timer()

    mlngElapsed = clsDDraw.TickCount() - mlngTimer
    mlngTimer = clsDDraw.TickCount()
    If clsDDraw.TickCount() - mlngFrameTimer >= 1000 Then
        mlngFrameTimer = clsDDraw.TickCount()
        mintFPS = mintFPSCounter
        mintFPSCounter = 0
    Else
        mintFPSCounter = mintFPSCounter + 1
    End If

End Sub


Public Sub CheckInput()
  If bShowingHanger = True Then
    If GetAsyncKeyState(DoEscape) < 0 And InGame = True And clsDDraw.TickCount - LastMenuMoveTick > MenuMoveTick Then
      bShowingHanger = False
      LastMenuMoveTick = clsDDraw.TickCount
      InMenu = True
      bShowMenu = True
    End If
  End If
  If GetAsyncKeyState(DoEscape) < 0 And clsDDraw.TickCount - LastMenuMoveTick > MenuMoveTick Then
    LastMenuMoveTick = clsDDraw.TickCount
    InMenu = True
    IsPaused = True
    Menu
  End If
  bLeft = False
  bRight = False
  bForward = False
  If GetAsyncKeyState(GoLeft) < 0 Then
    You.CurX = You.CurX - 3
    bLeft = True
  End If
  If GetAsyncKeyState(GoRight) < 0 Then
    You.CurX = You.CurX + 3
    bRight = True
  End If
  If GetAsyncKeyState(GoForward) < 0 Then
    bForward = True
    You.CurY = You.CurY - 3
    DDS_Thrust.Play DSBPLAY_DEFAULT
  Else
  End If
  If GetAsyncKeyState(GoBack) < 0 Then
    You.CurY = You.CurY + 3
  End If
  If (GetAsyncKeyState(DoShoot) < 0) Then
'    CreateShot FindShot
'    CreateMissile FindMissile
    Firing = True
    'DDS_Shot.Play DSBPLAY_LOOPING
  Else
    DDS_Shot.Stop
    Firing = False
  End If
End Sub

Public Function FindShot() As Integer
  Dim x As Integer
  For x = 0 To MaxShots
    If PShotL(x).active = False And PShotR(x).active = False Then
      FindShot = x
      Exit Function
    End If
  Next x
End Function

Public Function FindMissile() As Integer
  Dim x As Integer
  For x = 0 To MaxShots
    If MissileL(x).active = False And MissileR(x).active = False Then
      FindMissile = x
      Exit Function
    End If
  Next x
End Function

Public Sub Load()
  FrameText = "Still Checking Frames"
  Difficulty = ReadINI("Settings", "Difficulty", App.Path & "\data\game.ini")
  PixelPerfect = ReadINI("Settings", "Collision", App.Path & "\data\game.ini")
  MusVol = ReadINI("Settings", "MusicVolume", App.Path & "\data\game.ini")
  SndVol = ReadINI("Settings", "SoundVolume", App.Path & "\data\game.ini")
  clsDDraw.Init frmMain.hWnd, ScreenWidth, ScreenHeight, 16, DDS_Primary, DDS_Buffer, DDSD_Buffer
  clsDSound.Init frmMain.hWnd
  Dim iTmp As Double
  iTmp = (MusVol / 100)
  iTmp = cVol.VolumeMax * iTmp
  cVol.VolumeLevel = iTmp
  iTmp = (SndVol / 100)
  iTmp = cVol.WaveMax * iTmp
  cVol.WaveLevel = iTmp
    'This will allow us to play a lot of shot sounds at once.
    ' An array of
    clsDSound.LoadSoundFile DDS_Shot, App.Path & "\sounds\fire.wav"
  clsDSound.LoadSoundFile DDS_ExplodeW, App.Path & "\sounds\explo2.wav"
  clsDSound.LoadSoundFile DDS_Menu, App.Path & "\sounds\menu.wav"
  clsDSound.LoadSoundFile DDS_Thrust, App.Path & "\sounds\thrust.wav"
End Sub
Public Sub CloseGame()
  'frmMsg.lblCaption.Caption = "Chickening out already?!"
  'frmMsg.Left = (Screen.Width / 2) - (frmMsg.Width / 2)
  'frmMsg.Top = (Screen.Height / 2) - (frmMsg.Height / 2)
  'frmMsg.Show vbModal
  'If frmMsg.ReturnValue = True Then
  ShowCursor 1
    clsDDraw.EndIt
  'Else
  '
  ' Exit Sub
  'End If
End Sub


Public Sub Start()
  ShowCursor 0
  Load
  InitEnemyShips
  LoadData
End Sub

Public Sub Options()
   Dim mBackX As Integer, mBackY As Integer
   Dim rBack As RECT
   With rBack
     .Left = 0
     .Right = ScreenWidth
     .Top = 0
     .Bottom = ScreenHeight
   End With
   OptSel = 0
   mBackX = ((ScreenWidth - 430) \ 2)
   mBackY = ((ScreenHeight - 400) \ 2)
   
   Do While InOptions = True
    'CheckInput
     DDS_Buffer.BltColorFill rBack, 0
    BackYPos1 = BackYPos1 + 1
    BackYPos2 = BackYPos2 + 1
    Draw DDS_MenuBack, rBack1, 0, BackYPos1, False, True
    Draw DDS_MenuBack, rBack1, 0, BackYPos2, False, True
    If BackYPos1 >= ScreenHeight Then BackYPos1 = BackYPos2 - ScreenHeight
    If BackYPos2 >= ScreenHeight Then BackYPos2 = BackYPos1 - ScreenHeight
     
     Draw DDS_MenuBack2, rMB, mBackX, mBackY, True, True
     clsDDraw.DrawText DDS_Buffer, mBackX + 101, mBackY + 156, "Difficulty", &H404040
     If OptSel = 0 Then
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 155, "Difficulty", vbGreen
     Else
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 155, "Difficulty", vbBlue
     End If
     clsDDraw.DrawText DDS_Buffer, mBackX + 295, mBackY + 156, ReturnDifficulty, &H404040
     clsDDraw.DrawText DDS_Buffer, mBackX + 294, mBackY + 155, ReturnDifficulty, vbRed
     
     clsDDraw.DrawText DDS_Buffer, mBackX + 101, mBackY + 176, "Music Volume", &H404040
     If OptSel = 1 Then
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 175, "Music Volume", vbGreen
     Else
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 175, "Music Volume", vbBlue
     End If
     clsDDraw.DrawText DDS_Buffer, mBackX + 291, mBackY + 176, Str(MusVol) & "%", &H404040
     clsDDraw.DrawText DDS_Buffer, mBackX + 290, mBackY + 175, Str(MusVol) & "%", vbRed
     
     clsDDraw.DrawText DDS_Buffer, mBackX + 101, mBackY + 196, "Sound volume", &H404040
     If OptSel = 2 Then
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 195, "Sound Volume", vbGreen
     Else
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 195, "Sound Volume", vbBlue
     End If
     clsDDraw.DrawText DDS_Buffer, mBackX + 291, mBackY + 196, Str(SndVol) & "%", &H404040
     clsDDraw.DrawText DDS_Buffer, mBackX + 290, mBackY + 195, Str(SndVol) & "%", vbRed
     
     clsDDraw.DrawText DDS_Buffer, mBackX + 101, mBackY + 216, "Collision Detection", &H404040
     If OptSel = 3 Then
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 215, "Collision Detection", vbGreen
     Else
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 215, "Collision Detection", vbBlue
     End If
     If PixelPerfect = True Then
       clsDDraw.DrawText DDS_Buffer, mBackX + 295, mBackY + 216, "Pixel Perfect", &H404040
       clsDDraw.DrawText DDS_Buffer, mBackX + 294, mBackY + 215, "Pixel Perfect", vbRed
     Else
       clsDDraw.DrawText DDS_Buffer, mBackX + 295, mBackY + 216, "Intersect Rect", &H404040
       clsDDraw.DrawText DDS_Buffer, mBackX + 294, mBackY + 215, "Intersect Rect", vbRed
     End If
       
     
     DDS_Primary.Flip Nothing, DDFLIP_WAIT
     DoEvents
   Loop
   Menu
End Sub

Private Function ReturnDifficulty() As String
  If Difficulty = 0 Then
    ReturnDifficulty = "Rookie"
  ElseIf Difficulty = 1 Then
    ReturnDifficulty = "Novice"
  ElseIf Difficulty = 2 Then
    ReturnDifficulty = "Skilled"
  ElseIf Difficulty = 3 Then
    ReturnDifficulty = "Ace"
  ElseIf Difficulty = 4 Then
    ReturnDifficulty = "Death Incarnate"
  End If
End Function

Public Sub Menu()
  Dim ShowWall As Boolean
  InMenu = True
  BackYPos1 = 0
  BackYPos2 = BackYPos1 - ScreenHeight
  Dim rBack As RECT
  ShowWall = True
  SelButton = 1
  Dim mBackX As Integer, mBackY As Integer
  StopBGSound
  DDS_Menu.Play DSBPLAY_LOOPING
  With rB1
    .Left = 0
    .Right = ButWidth
    .Top = 0
    .Bottom = ButHeight
  End With
  With rB2
    .Left = 0
    .Right = ButWidth
    .Top = (ButHeight * 3)
    .Bottom = .Top + ButHeight
  End With
  With rB3
    .Left = 0
    .Right = ButWidth
    .Top = ButHeight * 5
    .Bottom = .Top + ButHeight
  End With
  With rB4
    .Left = 0
    .Right = ButWidth
    .Top = ButHeight * 7
    .Bottom = .Top + ButHeight
  End With
  Dim W1X As Integer, W2X As Integer, i As Integer
  W1X = 0
  W2X = ScreenWidth / 2
  With rW1
    .Left = 0
    .Top = 0
    .Right = WallWidth
    .Bottom = WallHeight
  End With
  With rW2
    .Left = 0
    .Top = 0
    .Right = WallWidth
    .Bottom = WallHeight
  End With
  With rBack
    .Left = 0
    .Right = ScreenWidth
    .Top = 0
    .Bottom = ScreenHeight
  End With
  
  Do Until InMenu = False
    If ShowWall = True Then
      W1X = W1X - 3
      W2X = W2X + 3
      If (W1X + WallWidth) <= 0 Then ShowWall = False
    End If
    Dim A As POINTAPI
    GetCursorPos A
    
    'CheckInput
    'Draw DDS_BACK, rBack1, 0, 0, False, False
    
    BackYPos1 = BackYPos1 + 1
    BackYPos2 = BackYPos2 + 1
    Draw DDS_MenuBack, rBack1, 0, BackYPos1, False, True
    Draw DDS_MenuBack, rBack1, 0, BackYPos2, False, True
    If BackYPos1 >= ScreenHeight Then BackYPos1 = BackYPos2 - ScreenHeight
    If BackYPos2 >= ScreenHeight Then BackYPos2 = BackYPos1 - ScreenHeight
    
    If (SelButton = 1) Then
      SetBut1
    ElseIf (SelButton = 2) Then
      SetBut2
    ElseIf (SelButton = 3) Then
      SetBut3
    Else
      SetBut4
    End If
    mBackX = ((ScreenWidth - 430) / 2)
    mBackY = ((ScreenHeight - 400) \ 2)
    Draw DDS_HEADER, rHeader, mBackX + 178, mBackY + 14, True, True
    Draw DDS_MenuBack2, rMB, mBackX, mBackY, True, True
    
    Draw DDS_Button, rB1, mBackX + 100, mBackY + 155, True, False
    Draw DDS_Button, rB2, mBackX + 100, mBackY + 155 + (ButHeight), True, False
    Draw DDS_Button, rB3, mBackX + 100, mBackY + 155 + (ButHeight * 2), True, False
    Draw DDS_Button, rB4, mBackX + 100, mBackY + 155 + (ButHeight * 3), True, False
    Draw DDS_WALL1, rW1, W1X, 0, False, True
    Draw DDS_WALL2, rW2, W2X, 0, False, True
    Draw DDS_Cursor, rCursor, A.x - 10, A.y - 10, True, True
    DDS_Primary.Flip Nothing, DDFLIP_WAIT
    DoEvents
    
  Loop
  DDS_Menu.Stop
  If WhatToDo = 0 Then
    IsPaused = False
    NewGame
  ElseIf WhatToDo = 1 Then
    InOptions = True
    Options
  ElseIf WhatToDo = 2 Then
    CloseGame
  End If
End Sub

Public Sub ExecuteClick(x As Integer, y As Integer)
  Dim mBackX As Integer, mBackY As Integer
  mBackX = ((ScreenWidth - 430) / 2)
  mBackY = ((ScreenHeight - 400) \ 2)
  
  If x > (mBackX + 100) And x < (mBackX + 100 + ButWidth) Then
    If y > (mBackY + 155) And y < (mBackY + 155 + ButHeight) And SelButton = 1 Then
      InMenu = False
      WhatToDo = 0
    End If
    If y > (mBackY + 155 + ButHeight) And y < (mBackY + 155 + (ButHeight * 2)) And SelButton = 2 Then
          
    End If
    If y > (mBackY + 155 + (ButHeight * 2)) And y < (mBackY + 155 + (ButHeight * 3)) And SelButton = 3 Then
      InMenu = False
      WhatToDo = 1
    End If
    If y > (mBackY + 155 + (ButHeight * 3)) And y < (mBackY + 155 + (ButHeight * 4)) And SelButton = 4 Then
      InMenu = False
      WhatToDo = 2
    End If
  End If
End Sub

Public Sub HighlightBut(x As Integer, y As Integer)
  Dim mBackX As Integer, mBackY As Integer
  mBackX = ((ScreenWidth - 430) / 2)
  mBackY = ((ScreenHeight - 400) \ 2)
  If x > (mBackX + 100) And x < (mBackX + 100 + ButWidth) Then
    'The potential to be in a button exists. Their all aligned on x coords so
    'after this check were good
    
    If y > (mBackY + 155) And y < (mBackY + 155 + ButHeight) Then
      SelButton = 1
    End If
    If y > (mBackY + 155 + ButHeight) And y < (mBackY + 155 + (ButHeight * 2)) Then
      SelButton = 2
    End If
    If y > (mBackY + 155 + (ButHeight * 2)) And y < (mBackY + 155 + (ButHeight * 3)) Then
      SelButton = 3
    End If
    If y > (mBackY + 155 + (ButHeight * 3)) And y < (mBackY + 155 + (ButHeight * 4)) Then
      SelButton = 4
    End If
  End If
End Sub

Private Sub SetBut1()
  With rB1
    .Top = 0
    .Bottom = .Top + ButHeight
  End With
  With rB2
    .Top = ButHeight * 3
    .Bottom = .Top + ButHeight
  End With
  With rB3
    .Top = ButHeight * 5
    .Bottom = .Top + ButHeight
  End With
  With rB4
    .Top = ButHeight * 7
    .Bottom = .Top + ButHeight
  End With
End Sub

Private Sub SetBut2()
  With rB1
    .Top = ButHeight
    .Bottom = .Top + ButHeight
  End With
  With rB2
    .Top = ButHeight * 2
    .Bottom = .Top + ButHeight
  End With
  With rB3
    .Top = ButHeight * 5
    .Bottom = .Top + ButHeight
  End With
  With rB4
    .Top = ButHeight * 7
    .Bottom = .Top + ButHeight
  End With
End Sub

Private Sub SetBut3()
  With rB1
    .Top = ButHeight
    .Bottom = .Top + ButHeight
  End With
  With rB2
    .Top = ButHeight * 3
    .Bottom = .Top + ButHeight
  End With
  With rB3
    .Top = ButHeight * 4
    .Bottom = .Top + ButHeight
  End With
  With rB4
    .Top = ButHeight * 7
    .Bottom = .Top + ButHeight
  End With
End Sub

Private Sub SetBut4()
  With rB1
    .Top = ButHeight
    .Bottom = .Top + ButHeight
  End With
  With rB2
    .Top = ButHeight * 3
    .Bottom = .Top + ButHeight
  End With
  With rB3
    .Top = ButHeight * 5
    .Bottom = .Top + ButHeight
  End With
  With rB4
    .Top = ButHeight * 6
    .Bottom = .Top + ButHeight
  End With
End Sub



Public Sub NewGame()
  StartNewChar
End Sub

Public Sub StartNewGame()
  InGame = True
  LoadWorld "lvl1.lvl"
  'bAtBoss = True
  BBMoveDir = 0
  BackYPos1 = 0
  BackYPos2 = -ScreenHeight
  NewGuy
  Main
End Sub

Public Sub DrawScrollBar(PosX1 As Integer, PosY1 As Integer, PosX2 As Integer, PosY2 As Integer, Max As Integer, Value As Integer)
  Dim tRect As RECT
  Dim tPnt As POINTAPI
  DDS_Buffer.SetForeColor vbGreen
  DDS_Buffer.DrawBox PosX1, PosY1, PosX2, PosY2
  tRect.Left = PosX1 + 1
  tRect.Right = PosX2 - 1
  tRect.Top = PosY1 + 1
  tRect.Bottom = PosY2 - 1
  DDS_Buffer.SetForeColor vbBlue
  DDS_Buffer.BltColorFill tRect, &H404000
  
  'the buttons
  DDS_Buffer.DrawBox PosX1 + 2, PosY1 + 2, PosX2 - 2, PosY1 + 18
  DDS_Buffer.DrawBox PosX1 + 2, PosY2 - 18, PosX2 - 2, PosY2 - 2
  DDS_Buffer.SetForeColor vbRed
  
  'the scroller
  DDS_Buffer.DrawBox PosX1 + 2, PosY1 + 1 + 18 + Value, PosX2 - 2, PosY1 + 18 + Value + 10
  
  
  'Determine if the mouse is over either of the buttons and if so highlight them
  GetCursorPos tPnt
    tRect.Left = PosX1 + 3
    tRect.Right = PosX2 - 3
    tRect.Top = PosY1 + 3
    tRect.Bottom = PosY1 + 17
  
  If tPnt.x > PosX1 + 2 And tPnt.x < PosX2 - 2 And tPnt.y > PosY1 + 2 And tPnt.y < PosY1 + 18 Then
    DDS_Buffer.SetForeColor vbBlue
    DDS_Buffer.BltColorFill tRect, vbGreen
  Else
    DDS_Buffer.BltColorFill tRect, 0
  End If
    tRect.Left = PosX1 + 3
    tRect.Right = PosX2 - 3
    tRect.Top = PosY2 - 17
    tRect.Bottom = PosY2 - 3
  
  If tPnt.x > PosX1 + 2 And tPnt.x < PosX2 - 2 And tPnt.y > PosY2 - 18 And tPnt.y < PosY2 - 2 Then
    DDS_Buffer.SetForeColor vbBlue
    DDS_Buffer.BltColorFill tRect, vbGreen
  Else
    DDS_Buffer.BltColorFill tRect, 0
  End If
  
End Sub

Public Sub StartNewChar()
  Dim tmpRect As RECT
  Dim lJX As Long, lJY As Long
  Dim BufFont As IFont
  Dim tmpBufFont As StdFont
  bInNewChar = True
  bAtCallsign = False
  strCallSign = ""
  strUserName = ""
  lJX = (ScreenWidth - JWidth) \ 2
  lJY = (ScreenHeight - JHeight) \ 2
  Set tmpBufFont = New StdFont
  tmpBufFont.Name = "Courier New"
  tmpBufFont.Size = 10
  tmpBufFont.Bold = False
  Set BufFont = tmpBufFont
  DDS_Buffer.SetFont BufFont
  Do
    DDS_Buffer.BltColorFill rBack1, 0
    DDS_Buffer.SetForeColor RGB(192, 192, 192)
    'DDS_Buffer.SetFont
    Draw DDS_NEWBACK, rNewBack, lJX, lJY, False, True
    If clsDDraw.TickCount - lBlinkTime > BlinkTime Then
      lBlinkTime = clsDDraw.TickCount
      bBlinkOn = Not bBlinkOn
    End If
    If bBlinkOn And Not bAtCallsign Then
      DDS_Buffer.DrawLine lJX + 34 + frmMain.TextWidth(strUserName) + 1, lJY + 103, lJX + 34 + frmMain.TextWidth(strUserName) + 1, lJY + 115
      DDS_Buffer.DrawLine lJX + 34 + frmMain.TextWidth(strUserName) + 2, lJY + 103, lJX + 34 + frmMain.TextWidth(strUserName) + 2, lJY + 115
      
    End If
    DDS_Buffer.DrawText lJX + 34, lJY + 101, strUserName, False
    If bAtCallsign Then
      If bBlinkOn Then
        DDS_Buffer.DrawLine lJX + 34 + frmMain.TextWidth(strCallSign) + 1, lJY + 153, lJX + 34 + frmMain.TextWidth(strCallSign) + 1, lJY + 165
        DDS_Buffer.DrawLine lJX + 34 + frmMain.TextWidth(strCallSign) + 2, lJY + 153, lJX + 34 + frmMain.TextWidth(strCallSign) + 2, lJY + 165
      End If
      DDS_Buffer.DrawText lJX + 34, lJY + 151, strCallSign, False
    End If
    DrawScrollBar 183, 224, 200, 397, 0, 0
    DDS_Primary.Flip Nothing, DDFLIP_WAIT
    DoEvents
  Loop
End Sub
Public Sub ShowHanger()
  bInHanger = True
  bShowingHanger = True
  Do Until bShowingHanger = False
    If InMenu Then Exit Do
    
    Draw DDS_HANGER, rHanger, 0, 0, False, True
    If strHangerText = "" Then strHangerText = " "
    DDS_Buffer.SetForeColor RGB(192, 192, 192)
    DDS_Buffer.DrawText 11, 453, strHangerText, False
    DDS_Buffer.SetForeColor RGB(256, 0, 0)
    DDS_Buffer.DrawText 10, 452, strHangerText, False
    Dim A As POINTAPI
    GetCursorPos A
    CheckInput
    Draw DDS_Cursor, rCursor, A.x - 10, A.y - 10, True, True
    DDS_Primary.Flip Nothing, DDFLIP_WAIT
    DoEvents
  Loop
  If bShowMenu = True Then
    InMenu = True
    bShowMenu = False
    Menu
    
  End If
End Sub
Public Sub Main()
  Do
    'draw current game screen
    Timer
    If IsPaused = False Then CheckInput
    If InMenu = True Then
      Exit Do
    End If
    If IsPaused = False Then DrawFrame
    If IsPaused = False Then BackYPos1 = BackYPos1 + 1
    If IsPaused = False Then BackYPos2 = BackYPos2 + 1
    If clsDDraw.TickCount - LastEnergyUpdateTick > ShieldEnergyUpdate Then
      LastEnergyUpdateTick = clsDDraw.TickCount
      You.CurEnergy = You.CurEnergy + 1
      If You.Shield < You.MaxShield Then
        If You.CurEnergy > 250 Then
          You.Shield = You.Shield + 1
          You.CurEnergy = You.CurEnergy - 2
        End If
      End If
      If You.CurEnergy > 500 Then You.CurEnergy = 500
      If You.Shield > You.MaxShield Then You.Shield = You.MaxShield
    End If
  DoEvents
  Loop
  If InMenu = True Then
    Menu
  End If
  
End Sub

Public Sub DrawWord(x As Long, y As Long, strText As String)
  Dim rTmp As RECT, i As Integer, strTmp As String, iASC As Integer
  For i = 1 To Len(strText)
    strTmp = Mid(strText, i, 1)
    iASC = Asc(strTmp)
    If IsNumeric(strTmp) Then
      iASC = iASC - 48
      With rTmp
        .Left = ((iASC + 26) * 9)
        .Right = ((iASC + 26) * 9) + 9
        .Bottom = rLetters.Bottom
        .Top = 0
      End With
    Else
      iASC = iASC - 97
      With rTmp
        .Left = (iASC * 9)
        .Right = (iASC * 9) + 9
        .Bottom = rLetters.Bottom
        .Top = 0
      End With
    End If
    Draw DDS_Letters, rTmp, x + ((i - 1) * 9), y, True, False
  Next
End Sub

Private Sub LoadData()
  'Load data
  'Clear all the variables if they exist
  Dim i As Integer
  Set DDS_YOU = Nothing
  Set DDS_PSHOT = Nothing
  Set DDS_Missile = Nothing
  Set DDS_ESHOT = Nothing
  Set DDS_Bottom = Nothing
  Set DDS_Bottom = Nothing
  Set DDS_HEALTH = Nothing
  Set DDS_Explode = Nothing
  Set DDS_WALL1 = Nothing
  Set DDS_WALL2 = Nothing
  Set DDS_Button = Nothing
  Set DDS_MenuBack = Nothing
  InitSound
  Set DDS_Money = Nothing
  clsDDraw.DDCreateSurface DDS_Money, App.Path & "\images\money.bmp", rMoney, 16, 16, 0
  clsDDraw.DDCreateSurface DDS_YOU, App.Path & "\images\newmain.bmp", rYou, , , 0
  clsDDraw.DDCreateSurface DDS_Letters, App.Path & "\images\letters.bmp", rLetters, , , 0
  clsDDraw.DDCreateSurface DDS_HEADER, App.Path & "\images\newhead.bmp", rHeader, , , 0
  clsDDraw.DDCreateSurface DDS_NEWBACK, App.Path & "\images\jreg.bmp", rNewBack, , , 0
  clsDDraw.DDCreateSurface DDS_HIT, App.Path & "\images\hit.bmp", rHit, , , 0
  clsDDraw.DDCreateSurface DDS_PSHOT, App.Path & "\images\bullet.bmp", rBullet, , , 0
  clsDDraw.DDCreateSurface DDS_Missile, App.Path & "\images\missile.bmp", rMissile, , , 0
  'clsDDraw.DDCreateSurface DDS_Back, App.Path & "\images\land1.bmp", rBack1, , , 0
  clsDDraw.DDCreateSurface DDS_Cursor, App.Path & "\images\cursor.bmp", rCursor, , , 0
  clsDDraw.DDCreateSurface DDS_MenuBack, App.Path & "\images\space03.bmp", rBack1, , , 0
  clsDDraw.DDCreateSurface DDS_ESHOT, App.Path & "\images\bullete.bmp", rBullet, , , 0
  clsDDraw.DDCreateSurface DDS_Bottom, App.Path & "\images\newbot.bmp", rBottom, DDSD_Buffer.lWidth, 64, 0
  clsDDraw.DDCreateSurface DDS_HEALTH, App.Path & "\images\filler.bmp", rHealth
  clsDDraw.DDCreateSurface DDS_HBAR, App.Path & "\images\hbar.bmp", rHBar
  clsDDraw.DDCreateSurface DDS_HANGER, App.Path & "\images\hanger.bmp", rHanger
  clsDDraw.DDCreateSurface DDS_Explode, App.Path & "\images\explode.bmp", rExplode
  clsDDraw.DDCreateSurface DDS_WALL1, App.Path & "\images\wall1.bmp", rW1
  clsDDraw.DDCreateSurface DDS_WALL2, App.Path & "\images\wall2.bmp", rW2
  clsDDraw.DDCreateSurface DDS_Button, App.Path & "\images\buttons.bmp", rB1
  clsDDraw.DDCreateSurface DDS_MenuBack2, App.Path & "\images\menuback.bmp", rMB
End Sub

Public Sub UpdateTiles()
  DDS_Buffer.BltColorFill rBack1, 0
  
  Call Draw(DDS_Back, rBack1, 0, BackYPos1, False, True)
  Call Draw(DDS_Back, rBack1, 0, BackYPos2, False, True)
'  Dim x As Integer, y As Integer, intY As Integer
'  ScrollY = ScrollY - 0.5
'  For x = 0 To (ScreenWidth \ TileWidth) - 1
'    For y = 0 To (ScreenHeight \ TileHeight)
'      intY = (y * TileHeight) - ScrollY Mod TileHeight
'      With rTiles
'        .Left = (Tiles2(x, (intY + TileHeight \ 2 + ScrollY - ScreenHeight \ 2) \ TileHeight).TileNumX * TileWidth)
'        .Right = .Left + TileWidth
'        .Top = (Tiles2(x, (intY + TileHeight \ 2 + ScrollY - ScreenHeight \ 2) \ TileHeight).TileNumY * TileHeight)
'        .Bottom = .Top + TileHeight
'      End With
'
'
'        Draw DDS_Tiles, rTiles, x * TileWidth, intY, False, True
'    Next y
'  Next x
End Sub

Private Sub DrawHealth()
  With rHealth
    .Right = ((You.Shield) / (You.MaxShield)) * 166
    .Bottom = 10
  End With
  With rHull
    .Right = ((You.Hull) / (You.MaxHull)) * 166
    .Bottom = 10
  End With
  With rEnergy
    .Right = ((You.CurEnergy) / (500)) * 166
    .Bottom = 10
  End With
  'DDS_Buffer.DrawText 450, 429, "Shield:", False
  'DDS_Buffer.DrawText 450, 446, "    Hull:", False
  DDS_Buffer.DrawText 200, 420, "Shield", False
  DDS_Buffer.DrawText 200, 440, "Energy", False
  DDS_Buffer.DrawText 200, 460, "Hull", False
  Draw DDS_HBAR, rHBar, 18, 460, False, False
  Draw DDS_HBAR, rHBar, 18, 420, False, False
  Draw DDS_HBAR, rHBar, 18, 440, False, False
  Draw DDS_HEALTH, rHealth, 20, 422, False, False
  Draw DDS_HEALTH, rEnergy, 20, 442, False, False
  Draw DDS_HEALTH, rHull, 20, 462, False, False
  'Draw DDS_HEALTH, rHull, 512, 450, False, False
End Sub


Public Sub UpdateShots()
  Dim x As Integer, y As Integer
  Dim rMoney As Integer
  Dim ShotTickCount As Long
  ShotTickCount = clsDDraw.TickCount
  For x = 0 To MaxShots
    If PShotL(x).CurY <= 0 Then PShotL(x).active = False
    If PShotR(x).CurY < 0 Then PShotR(x).active = False
    If MissileL(x).CurY < 0 Then MissileR(x).active = False
    If MissileR(x).CurY < 0 Then MissileR(x).active = False
  Next
  For x = 0 To MaxShots
    If (PShotL(x).active = True Or PShotR(x).active = True) Or (MissileL(x).active = True Or MissileR(x).active = True) Then
      If PShotL(x).active = True Then PShotL(x).CurY = PShotL(x).CurY - ShotVelocity
      If PShotR(x).active = True Then PShotR(x).CurY = PShotR(x).CurY - ShotVelocity
      If PShotL(x).CurY <= 0 Then PShotL(x).active = False
      If PShotR(x).CurY <= 0 Then PShotR(x).active = False
      If PShotL(x).active = True Then
        Draw DDS_PSHOT, rBullet, PShotL(x).CurX, PShotL(x).CurY, True, False
      End If
      If PShotR(x).active = True Then
        Draw DDS_PSHOT, rBullet, PShotR(x).CurX, PShotR(x).CurY, True, False
      End If
      If MissileL(x).active = True Then MissileL(x).CurY = MissileL(x).CurY - ShotVelocity
      If MissileR(x).active = True Then MissileR(x).CurY = MissileR(x).CurY - ShotVelocity
      If MissileL(x).CurY <= 0 Then MissileL(x).active = False
      If MissileR(x).CurY <= 0 Then MissileR(x).active = False
      
      If MissileL(x).active = True Then
        Draw DDS_Missile, rMissile, MissileL(x).CurX, MissileL(x).CurY, True, False
      End If
      If MissileR(x).active = True Then
        Draw DDS_Missile, rMissile, MissileR(x).CurX, MissileR(x).CurY, True, False
      End If
      
      For y = 0 To MaxEnemies
        If PShotL(x).active = True Or PShotR(x).active = True Then
          If BadGuys(y).active = True And PShotL(x).active = True Then
            If CheckCollide(DDS_PSHOT, BadGuys(y).Surface, PShotL(x).CurX, PShotL(x).CurY, ShotWidth, ShotHeight, BadGuys(y).x, BadGuys(y).y, BadGuys(y).Width, BadGuys(y).Height, 0) Then
              CreateHit PShotL(x).CurX, PShotL(x).CurY, BadGuys(y).Velocity
'              PlayWav "hitp.wav"
                If BadGuys(y).Shield > 0 Then
                  BadGuys(y).Shield = BadGuys(y).Shield - 30
                Else
                  BadGuys(y).Hull = BadGuys(y).Hull - 50
                End If
                If BadGuys(y).Shield <= 0 And BadGuys(y).Hull <= 0 Then
                  BadGuys(y).active = False
                  If bAtBoss Then
                    bAtBoss = False
                    bEndLevel = True
                  End If
                  AddExplode BadGuys(y).x, BadGuys(y).y
                  You.CurMoney = You.CurMoney + BadGuys(x).Value
                  rMoney = Int(Rnd * 100)
                  If rMoney > 1 Then
                    AddMoney (BadGuys(y).x + (BadGuys(y).Width \ 2)), (BadGuys(y).y + (BadGuys(y).Height \ 2))
                  End If
                End If
                PShotL(x).active = False
              End If
            End If
          If BadGuys(y).active = True And PShotR(x).active = True Then
            If CheckCollide(DDS_PSHOT, BadGuys(y).Surface, PShotR(x).CurX, PShotR(x).CurY, ShotWidth, ShotHeight, BadGuys(y).x, BadGuys(y).y, BadGuys(y).Width, BadGuys(y).Height, 0) Then
              CreateHit PShotR(x).CurX, PShotR(x).CurY, BadGuys(y).Velocity
'             PlayWav "hitp.wav"
              If BadGuys(y).Shield > 0 Then
                BadGuys(y).Shield = BadGuys(y).Shield - 30
              Else
                BadGuys(y).Hull = BadGuys(y).Hull - 50
              End If
              If BadGuys(y).Shield <= 0 And BadGuys(y).Hull <= 0 Then
                BadGuys(y).active = False
                  If bAtBoss Then
                    bAtBoss = False
                    bEndLevel = True
                  End If
                AddExplode BadGuys(y).x, BadGuys(y).y
                rMoney = Int(Rnd * 100)
                If rMoney > 1 Then
                  AddMoney (BadGuys(y).x + (BadGuys(y).Width \ 2)), (BadGuys(y).y + (BadGuys(y).Height \ 2))
                End If
              
                You.CurMoney = You.CurMoney + BadGuys(x).Value
              End If
              PShotR(x).active = False
            End If
          End If
        End If
        If MissileL(x).active Or MissileR(x).active = True Then
          If BadGuys(y).active = True And MissileL(x).active = True Then
            If CheckCollide(DDS_Missile, BadGuys(y).Surface, MissileL(x).CurX, MissileL(x).CurY, ShotWidth, ShotHeight, BadGuys(y).x, BadGuys(y).y, BadGuys(y).Width, BadGuys(y).Height, 0) Then
              CreateHit MissileL(x).CurX, MissileL(x).CurY, BadGuys(y).Velocity
'             PlayWav "hitp.wav"
                If BadGuys(y).Shield > 0 Then
                  BadGuys(y).Shield = BadGuys(y).Shield - 30
                Else
                  BadGuys(y).Hull = BadGuys(y).Hull - 50
                End If
                If BadGuys(y).Shield <= 0 And BadGuys(y).Hull <= 0 Then
                  BadGuys(y).active = False
                  If bAtBoss Then
                    bAtBoss = False
                    bEndLevel = True
                  End If
                  AddExplode BadGuys(y).x, BadGuys(y).y
                  rMoney = Int(Rnd * 100)
                  If rMoney > 1 Then
                    AddMoney (BadGuys(y).x + (BadGuys(y).Width \ 2)), (BadGuys(y).y + (BadGuys(y).Height \ 2))
                  End If
                  
                  You.CurMoney = You.CurMoney + BadGuys(x).Value
                End If
                MissileL(x).active = False
              End If
            End If
          If BadGuys(y).active = True And MissileR(x).active = True Then
            If CheckCollide(DDS_Missile, BadGuys(y).Surface, MissileR(x).CurX, MissileR(x).CurY, ShotWidth, ShotHeight, BadGuys(y).x, BadGuys(y).y, BadGuys(y).Width, BadGuys(y).Height, 0) Then
              CreateHit MissileR(x).CurX, MissileR(x).CurY, BadGuys(y).Velocity
'             PlayWav "hitp.wav"
              If BadGuys(y).Shield > 0 Then
                BadGuys(y).Shield = BadGuys(y).Shield - 30
              Else
                BadGuys(y).Hull = BadGuys(y).Hull - 50
              End If
              If BadGuys(y).Shield <= 0 And BadGuys(y).Hull <= 0 Then
                BadGuys(y).active = False
                If bAtBoss Then
                  bAtBoss = False
                  bEndLevel = True
                End If
                AddExplode BadGuys(y).x, BadGuys(y).y
                rMoney = Int(Rnd * 100)
                If rMoney > 1 Then
                  AddMoney (BadGuys(y).x + (BadGuys(y).Width \ 2)), (BadGuys(y).y + (BadGuys(y).Height \ 2))
                End If
                
                You.CurMoney = You.CurMoney + BadGuys(x).Value
              End If
              MissileR(x).active = False
            End If
          End If
        End If
      Next y
    Else
      If Firing = True And ((ShotTickCount - LastShotTick) > ShotTick) Then
        CreateShot x
        CreateMissile x
      End If
        
    End If
      
  Next x
    
End Sub

Public Sub CreateShot(ShotI As Integer)
  Dim i As Integer
  If You.CurEnergy <= 0 Then Exit Sub
  PShotL(ShotI).active = True
  PShotR(ShotI).active = True
  You.CurEnergy = You.CurEnergy - 2
  PShotL(ShotI).CurX = (You.CurX + 14)
  PShotR(ShotI).CurX = (You.CurX + 34)
  PShotL(ShotI).CurY = (You.CurY + 2)
  PShotR(ShotI).CurY = (You.CurY + 2)
  LastShotTick = clsDDraw.TickCount
End Sub

Public Sub CreateMissile(ShotI As Integer)
  Dim i As Integer
  If You.CurEnergy <= 0 Then Exit Sub
  You.CurEnergy = You.CurEnergy - 2
  MissileL(ShotI).active = True
  MissileR(ShotI).active = True
  MissileL(ShotI).CurX = (You.CurX + 10)
  MissileR(ShotI).CurX = (You.CurX + 38)
  MissileL(ShotI).CurY = (You.CurY + 10)
  MissileR(ShotI).CurY = (You.CurY + 10)
  LastShotTick = clsDDraw.TickCount
End Sub


Public Sub DrawFrame()
    
    Dim ddrVal As Long 'Every drawing procedure returns a value, so we must have a
                       'var able to hold it. From this value we can check for errors.
    Dim bRestore As Boolean
    bRestore = False
    Do Until clsDDraw.ExModeActive
      DoEvents
      bRestore = True
    Loop
    'DoEvents
    If bRestore Then
      bRestore = False
      clsDDraw.RestoreAllSurfaces 'this just re-allocates memory back to us. we must
                               'still reload all the surfaces.
      LoadData ' must init the surfaces again if they we're lost
    End If
    If BackYPos1 >= ScreenHeight Then BackYPos1 = BackYPos2 - ScreenHeight
    If BackYPos2 >= ScreenHeight Then BackYPos2 = BackYPos1 - ScreenHeight
    UpdateTiles
    If bEndLevel Then
      If You.CurX > ((ScreenWidth - YouWidth) \ 2) Then
        You.CurX = You.CurX - 1
      ElseIf You.CurX < ((ScreenWidth - YouWidth) \ 2) Then
        You.CurX = You.CurX + 1
      End If
      If You.CurX = ((ScreenWidth - YouWidth) \ 2) Then
        You.CurY = You.CurY + 1
      End If
      If You.CurY > ScreenHeight Then
        InGame = False
        StopBGSound
        bEndLevel = False
        ShowHanger
      End If
    Else
      If bForward = True Then
        If bLeft = True Then
          rYou.Left = YouWidth * 3
          rYou.Right = YouWidth * 4
        ElseIf bRight = True Then
          rYou.Left = YouWidth * 5
          rYou.Right = YouWidth * 6
        Else
          rYou.Left = YouWidth * 4
          rYou.Right = YouWidth * 5
        End If
      Else
        If bLeft = True Then
          rYou.Left = 0
          rYou.Right = YouWidth
        ElseIf bRight = True Then
          rYou.Left = YouWidth * 2
          rYou.Right = YouWidth * 3
        Else
          rYou.Left = YouWidth
          rYou.Right = YouWidth * 2
        End If
      End If
      rYou.Top = 0
      rYou.Bottom = YouHeight
      If You.CurX > (DDSD_Buffer.lWidth - YouWidth) Then
        You.CurX = (DDSD_Buffer.lWidth - YouWidth)
      End If
      If You.CurY > (ScreenHeight - 67 - YouHeight) Then
        You.CurY = (414 - YouHeight)
      End If
     
      If You.CurX < 0 Then You.CurX = 0
      If You.CurY < 0 Then You.CurY = 0
    
      'ddrval = Draw(DDS_YOU, rYou, You.CurX, You.CurY, True, True)
    End If
    
    DDS_Buffer.BltFast You.CurX, You.CurY, DDS_YOU, rYou, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    UpdateEnemyShips
    UpdateMoney
    UpdateHits
    UpdateShots
    UpdateEFiring
    UpdateExplode
    ParseLevel
    DDS_Buffer.SetForeColor RGB(150, 150, 150)
    Call DDS_Buffer.DrawLine(0, ScreenHeight - 66, ScreenWidth, ScreenHeight - 66)
    Call DDS_Buffer.DrawLine(0, ScreenHeight - 65, ScreenWidth, ScreenHeight - 65)
    DDS_Buffer.SetForeColor RGB(180, 180, 180)
    Call DDS_Buffer.DrawLine(0, ScreenHeight - 67, ScreenWidth, ScreenHeight - 67)
    ddrVal = Draw(DDS_Bottom, rBottom, 0, ScreenHeight - 64, False, True)
    
    DrawHealth
    If clsDDraw.TickCount - LastTimeChecked >= 1000 Then
        LastTimeChecked = clsDDraw.TickCount
        FrameText = "FPS: " & CStr(framesDone) & " fps"
        framesDone = 0
    End If
    Dim rOver As RECT
    DDS_Buffer.SetForeColor vbRed
    DDS_Buffer.DrawText 300, 420, FrameText, False
    'DrawWord 300, 420, "hello there 0123456"
    'DDS_Buffer.DrawText 250, 460, "$" & Str(You.CurMoney), False
    'DDS_Buffer.DrawText 18, 460, clrBack, False
    DDS_Primary.Flip Nothing, DDFLIP_WAIT   'Flip the secondary surface to the primary
    framesDone = framesDone + 1
End Sub

Private Sub NewGuy()
  Dim x As Integer, ddrVal As Long
  
  ShipInvincible = True
  You.Shield = 100
  You.MaxShield = 100
  You.Hull = 300
  You.MaxHull = 300
  You.CurX = (ScreenWidth - YouWidth) \ 2
  rYou.Left = YouWidth
  rYou.Right = YouWidth * 2
  You.CurEnergy = 500
  For x = (ScreenHeight + YouHeight) To (414 - YouHeight) Step -1
    You.CurY = x
    
    UpdateTiles
    DDS_Buffer.BltFast You.CurX, You.CurY, DDS_YOU, rYou, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
    UpdateEnemyShips
    UpdateEFiring
    UpdateShots
    UpdateExplode
    DDS_Buffer.SetForeColor RGB(150, 150, 150)
    Call DDS_Buffer.DrawLine(0, ScreenHeight - 66, ScreenWidth, ScreenHeight - 66)
    Call DDS_Buffer.DrawLine(0, ScreenHeight - 65, ScreenWidth, ScreenHeight - 65)
    DDS_Buffer.SetForeColor RGB(180, 180, 180)
    Call DDS_Buffer.DrawLine(0, ScreenHeight - 67, ScreenWidth, ScreenHeight - 67)

    ddrVal = Draw(DDS_Bottom, rBottom, 0, ScreenHeight - 64, False, True)
    DrawHealth
    BackYPos1 = BackYPos1 + 1
    BackYPos2 = BackYPos2 + 1
   
    DDS_Primary.Flip Nothing, DDFLIP_WAIT  'Flip the secondary surface to the primary
  Next
  ShipInvincible = False
End Sub



Public Sub UpdateEnemyShips()
  Dim x As Integer
  Dim bShipOnScreen As Boolean
  bShipOnScreen = False
  For x = 0 To MaxEnemies 'MaxEnemyShips
    If BadGuys(x).active = True Then
      bShipOnScreen = True
      If bAtBoss = False Then BadGuys(x).y = BadGuys(x).y + BadGuys(x).Velocity
      If BadGuys(x).AI = 1 Then
        AIBounce (x)
      ElseIf BadGuys(x).AI = 2 Then
        DownOff (x)
      ElseIf BadGuys(x).AI = 3 Then
        BigBoss (x)
      End If
      If (BadGuys(x).FrameX >= BadGuys(x).FramesX) Then
        If (BadGuys(x).FrameY < BadGuys(x).FramesY) Then
          BadGuys(x).FrameY = BadGuys(x).FrameY + 1
          BadGuys(x).FrameX = 0
          BadGuys(x).Tick = clsDDraw.TickCount
        Else
          BadGuys(x).FrameY = 0
          BadGuys(x).FrameX = 0
          BadGuys(x).Tick = clsDDraw.TickCount
        End If
      End If
      If (BadGuys(x).FrameX < BadGuys(x).FramesX) And (clsDDraw.TickCount - BadGuys(x).Tick) > UpdateEnemyAnimTick Then
        BadGuys(x).FrameX = BadGuys(x).FrameX + 1
        BadGuys(x).Tick = clsDDraw.TickCount
      End If
      With BadGuys(x).RECT
        .Left = BadGuys(x).FrameX * BadGuys(x).Width
        .Right = BadGuys(x).Width + .Left
        .Top = BadGuys(x).FrameY * BadGuys(x).Height
        .Bottom = BadGuys(x).Height + .Top
      End With
      Draw BadGuys(x).Surface, BadGuys(x).RECT, BadGuys(x).x, BadGuys(x).y, True, True
      If CheckCollide(DDS_YOU, BadGuys(x).Surface, You.CurX, You.CurY, YouWidth, YouHeight, BadGuys(x).x, BadGuys(x).y, BadGuys(x).Width, BadGuys(x).Height, 0) And BadGuys(x).active = True Then
        If You.Shield > 0 Then
          If ShipInvincible = False Then You.Shield = You.Shield - 30
          'Beep
          BadGuys(x).active = False
                If bAtBoss Then
                  bAtBoss = False
                  bEndLevel = True
                End If
          AddExplode BadGuys(x).x, BadGuys(x).y
          You.CurMoney = You.CurMoney + BadGuys(x).Value
        Else
          If You.Hull > 0 Then
            If ShipInvincible = False Then You.Hull = You.Hull - 50
            BadGuys(x).active = False
                If bAtBoss Then
                  bAtBoss = False
                  bEndLevel = True
                End If
            AddExplode BadGuys(x).x, BadGuys(x).y
          End If
        End If
        If You.Hull <= 0 And You.Shield <= 0 Then
          AddExplode You.CurX, You.CurY
          NewGuy
        End If
      End If
      If BadGuys(x).y >= ScreenHeight Then
        BadGuys(x).active = False
        
      End If
      
      
      'If ((clsDDraw.TickCount - LastGuyTick) > NewBadGuyTick) Then CreateEnemy x
    End If
  Next
  If bShipOnScreen = False And bAtBoss = True Then
    'Create Boss :)
    CreateBoss
  End If
End Sub


Private Function Draw(Surface As DirectDrawSurface7, RECTvar As RECT, ByVal x As Integer, ByVal y As Integer, Optional transparent As Boolean = True, Optional Clip As Boolean = True) As Long
    'This subroutine will BltFast a surface to the
    'backbuffer.
    
    'CLIPPING
    'Temporary rect
    Dim RectTEMP As RECT
    RectTEMP = RECTvar
    
    If Clip = True Then
        'Set up screen rect for clipping
        Dim ScreenRECT As RECT
        With ScreenRECT
            .Top = y
            .Left = x
            .Bottom = y + RECTvar.Bottom - RECTvar.Top
            .Right = x + RECTvar.Right - RECTvar.Left
        End With
        'Clip surface
        With ScreenRECT
            If .Bottom > ScreenHeight Then
                RectTEMP.Bottom = RectTEMP.Bottom - (.Bottom - ScreenHeight)
                .Bottom = ScreenHeight
            End If
            If .Left < 0 Then
                RectTEMP.Left = RectTEMP.Left - .Left
                .Left = 0
                x = 0
            End If
            If .Right > ScreenWidth Then
                RectTEMP.Right = RectTEMP.Right - (.Right - ScreenWidth)
                .Right = ScreenWidth
                
            End If
            If .Top < 0 Then
                RectTEMP.Top = RectTEMP.Top - .Top
                .Top = 0
                y = 0
            End If
        End With
    
    End If
    If transparent = False Then
        Draw = DDS_Buffer.BltFast(x, y, Surface, RectTEMP, DDBLTFAST_WAIT)
    Else
        Draw = DDS_Buffer.BltFast(x, y, Surface, RectTEMP, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Function


Private Function FindHit() As Integer
  Dim x As Integer
  For x = 0 To MaxHit
    If (Hits(x).active = False) Then
      FindHit = x
      Exit Function
    End If
  Next
  FindHit = MaxHit
End Function


Private Sub CreateHit(x As Integer, y As Integer, speed As Integer)
  Dim NewHit As Integer
  NewHit = FindHit()
  Hits(NewHit).active = True
  Hits(NewHit).frame = 0
  Hits(NewHit).LastTick = clsDDraw.TickCount
  Hits(NewHit).x = x
  Hits(NewHit).y = y
  Hits(NewHit).speed = speed
End Sub

Public Sub UpdateHits()
  Dim x As Integer, NewTick As Long
  NewTick = clsDDraw.TickCount
  For x = 0 To MaxHit
    If (Hits(x).active = True) Then
      Hits(x).y = Hits(x).y + Hits(x).speed
      With rHit
        .Top = 0
        .Bottom = HitHeight
        .Left = Hits(x).frame * HitWidth
        .Right = .Left + HitWidth
      End With
      'BitBlt frmMain.picBak.hDC, Hits(X).X, Hits(X).y, HitWidth, HitHeight, HitDC, Hits(X).Frame * HitWidth, 12, vbSrcAnd
      'BitBlt frmMain.picBak.hDC, Hits(X).X, Hits(X).y, HitWidth, HitHeight, HitDC, Hits(X).Frame * HitWidth, 0, vbSrcPaint
      Draw DDS_HIT, rHit, Hits(x).x, Hits(x).y, True, False
    End If
    If ((NewTick - Hits(x).LastTick) > HitTick) Then
      Hits(x).LastTick = clsDDraw.TickCount
      Hits(x).frame = Hits(x).frame + 1
      If Hits(x).frame = 6 Then
        Hits(x).active = False
      End If
    End If
    
  Next
End Sub

Private Function FindExplode() As Integer
  Dim x As Integer
  For x = 0 To MaxExplode
    If (Explodes(x).active = False) Then
      FindExplode = x
      Exit Function
    End If
  Next
End Function

Public Sub AddExplode(x, y)
  Dim NewExplode As Integer
  DDS_ExplodeW.Play DSBPLAY_DEFAULT
  NewExplode = FindExplode
  Explodes(NewExplode).active = True
  Explodes(NewExplode).x = x
  Explodes(NewExplode).y = y
  Explodes(NewExplode).LastTick = clsDDraw.TickCount
  Explodes(NewExplode).frame = 0
End Sub

Public Sub UpdateExplode()
  Dim x As Integer, NewTick As Long
  Dim rTemp As RECT
  NewTick = clsDDraw.TickCount

  For x = 0 To MaxExplode
    With rTemp
      .Left = Explodes(x).frame * ExplodeWidth
      .Right = .Left + ExplodeWidth
      .Bottom = ExplodeHeight
      .Top = 0
    End With
    If (Explodes(x).active = True) Then

      Draw DDS_Explode, rTemp, Explodes(x).x, Explodes(x).y, True, True
      If (NewTick - Explodes(x).LastTick > ExplodeTick) Then
        Explodes(x).LastTick = clsDDraw.TickCount
        Explodes(x).frame = Explodes(x).frame + 1
        If Explodes(x).frame = 13 Then Explodes(x).active = False
      End If
    End If
  Next
End Sub
 


Public Sub UpdateEFiring()
  Dim x As Integer
  For x = 0 To MaxEShots
    If (EShotL(x).active = True Or EShotR(x).active = True) Then
      If EShotL(x).active = True Then EShotL(x).CurY = EShotL(x).CurY + ShotVelocity
      If EShotR(x).active = True Then EShotR(x).CurY = EShotR(x).CurY + ShotVelocity
      If EShotL(x).CurY >= ScreenHeight Then EShotL(x).active = False
      If EShotR(x).CurY >= ScreenHeight Then EShotR(x).active = False
      If EShotR(x).active = True And CheckCollide(DDS_ESHOT, DDS_YOU, EShotR(x).CurX, EShotR(x).CurY, ShotWidth, ShotHeight, You.CurX, You.CurY, YouWidth, YouHeight, 0) Then
        CreateHit EShotL(x).CurX, (EShotL(x).CurY + ShotHeight), 0
        EShotR(x).active = False
        If You.Shield > 0 Then
          If ShipInvincible = False Then You.Shield = You.Shield - 10
        Else
          If ShipInvincible = False Then You.Hull = You.Hull - 15
        End If
        If You.Shield <= 0 And You.Hull <= 0 Then
          AddExplode You.CurX, You.CurY
          Firing = False
          NewGuy
        End If
      End If
      If EShotL(x).active = True And CheckCollide(DDS_ESHOT, DDS_YOU, EShotL(x).CurX, EShotL(x).CurY, ShotWidth, ShotHeight, You.CurX, You.CurY, YouWidth, YouHeight, 0) Then
        CreateHit EShotL(x).CurX, (EShotL(x).CurY + ShotHeight), 0
        EShotL(x).active = False
        If You.Shield > 0 Then
          If ShipInvincible = False Then You.Shield = You.Shield - 10
        Else
          If ShipInvincible = False Then You.Hull = You.Hull - 15
        End If
        If You.Shield <= 0 And You.Hull <= 0 Then
          AddExplode You.CurX, You.CurY
          Firing = False
          NewGuy
        End If
      End If
        
      If EShotL(x).active = True Then
        Draw DDS_ESHOT, rBullet, EShotL(x).CurX, EShotL(x).CurY, True, False
      End If
      If EShotR(x).active = True Then
        Draw DDS_ESHOT, rBullet, EShotR(x).CurX, EShotR(x).CurY, True, False
      End If
    End If
  Next
End Sub

Public Sub LoadWorld(lvl As String)
  Dim iFile As Integer
  iFile = FreeFile()
  OpenMap App.Path & "\levels\test.map"
  Open App.Path & "\levels\" & lvl For Binary Access Read As #iFile
    Get #iFile, , Level
  Close #iFile
  LoadLevel "lvl1.enm" 'Level.EnemyFile
  clsDDraw.DDCreateSurface DDS_Back, App.Path & "\images\" & Level.TileBitmap, rBack1, , , 0
  PlayBGSound App.Path & "\sounds\" & Level.MusicFile
  
End Sub

Public Sub OpenMap(strFilePath As String)

  Dim Counter As Byte
  Dim lngXSize As Long
  Dim lngYSize As Long
  Dim byteData(8) As Byte
  Dim byteInputData1 As Byte
  Dim byteInputData2 As Byte
  Dim byteInputData3 As Byte
  Dim byteInputData4 As Byte
  Dim iFile As Integer
  iFile = FreeFile
  Open strFilePath For Binary As #iFile
        For Counter = 1 To 8
            byteData(Counter) = CByte(Asc(Input(1, #iFile)))
            DoEvents
        Next Counter
        tGrdX = byteData(3) * 256 + byteData(4) - 1
        tGrdY = byteData(7) * 256 + byteData(8) - 1
        ScrollY = ((TileHeight * tGrdY) - (TileHeight * 2) - 100)
        mlngElapsed = 0
        ReDim Tiles2(tGrdX, tGrdY)
        For lngXSize = 0 To tGrdX
            For lngYSize = 0 To tGrdY
                DoEvents
                byteInputData1 = CByte(Asc(Input(1, #iFile)))
                byteInputData2 = CByte(Asc(Input(1, #iFile)))
                byteInputData3 = CByte(Asc(Input(1, #iFile)))
                byteInputData4 = CByte(Asc(Input(1, #iFile)))
                Tiles2(lngXSize, lngYSize).TileNumX = byteInputData1 * 256 + byteInputData2
                Tiles2(lngXSize, lngYSize).TileNumY = byteInputData3 * 256 + byteInputData4
                Tiles2(lngXSize, lngYSize).y = lngYSize * -(TileHeight)
            Next
        Next
  Close #iFile
End Sub

Public Sub AddMoney(x As Integer, y As Integer)
  Dim i As Integer
  ' First find an innactive money
  For i = 0 To 20
    If mMoney(i).active = False Then Exit For ' We've found it
  Next
  If i > 20 Then i = 20
  mMoney(i).frame = 0
  mMoney(i).LastTick = clsDDraw.TickCount()
  mMoney(i).x = x
  mMoney(i).y = y
  mMoney(i).active = True
  mMoney(i).RECT = rMoney
End Sub

Public Sub UpdateMoney()
  Dim i As Integer
  For i = 0 To 20
    If mMoney(i).active = True Then
      If clsDDraw.TickCount() - mMoney(i).LastTick > 300 Then
        If mMoney(i).frame = 0 Then
          mMoney(i).frame = 1
          mMoney(i).LastTick = clsDDraw.TickCount()
        Else
          mMoney(i).frame = 0
          mMoney(i).LastTick = clsDDraw.TickCount()
        End If
        
      End If
      mMoney(i).y = mMoney(i).y + 1
      With mMoney(i).RECT
        .Left = mMoney(i).frame * 16
        .Right = .Left + 16
        .Top = 0
        .Bottom = 16
      End With
      Draw DDS_Money, mMoney(i).RECT, mMoney(i).x, mMoney(i).y, True, True
      If mMoney(i).y > ScreenHeight Then mMoney(i).active = False
    End If
  Next i
End Sub
