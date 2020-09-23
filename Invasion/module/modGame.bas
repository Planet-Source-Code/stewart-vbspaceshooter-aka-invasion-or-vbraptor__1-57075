Attribute VB_Name = "modGame"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modGame.bas                                           |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+


Private ScrollY As Single
Public ShowWall As Boolean
Dim TempTime As Long
Public W1X As Integer, W2X As Integer, i As Integer
Private tGrdX As Integer, tGrdY As Integer
Private Tiles2() As Tile_Data

Private cFrameLimit As New clsFrameLimiter

'Some sound buffers
Public DDS_Shot As DirectSoundBuffer
Public DDS_ExplodeW As DirectSoundBuffer
Public DDS_Thrust As DirectSoundBuffer
Public DDS_Menu As DirectSoundBuffer
Public DDS_Click As DirectSoundBuffer
Public DDS_GetMoney As DirectSoundBuffer

Public clsDDraw As New cDDraw
Public clsDSound As New cDSound

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Type POINTAPI
        x As Long
        Y As Long
End Type


Public mMoney(20) As Money


'Runtime framerate stuff
Dim framesDone As Integer, LastTimeChecked As Long, FrameText As String

Private rYou As RECT, rBack1 As RECT, rBackEarth As RECT, rCursor As RECT, rBullet As RECT, rBottom As RECT
Private rE1 As RECT, rHit As RECT, rHealth As RECT, rExplode As RECT, rHull As RECT
Private rW1 As RECT, rW2 As RECT, rB1 As RECT, rB2 As RECT, rB3 As RECT, rB4 As RECT
Private rMB As RECT, rTiles As RECT, rMissile As RECT, rMissileE As RECT, rNewBack As RECT, rHanger As RECT
Private rHBar As RECT, rEnergy As RECT, rHeader As RECT, rLetters As RECT, rMoney As RECT
Private rSave As RECT, rYouSmall As RECT, rDisplay As RECT, rPulse As RECT, rShop As RECT, rReactor As RECT
Private rPulseC As RECT, rPlasma As RECT, rMicro As RECT, rPlasmaS As RECT, rMessage As RECT
Private rShield As RECT, rShopSell As RECT

Private bLeft As Boolean, bRight As Boolean, bForward As Boolean

Public LastMenuMoveTick As Long
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
  bLeft = False
  bRight = False
  bForward = False
'  If bShowMsg = True Then
'    If GetAsyncKeyState(DoReturn) < 0 And clsDDraw.TickCount - LastMenuMoveTick > MenuMoveTick Then
'      bShowMsg = False
'    End If
'    Exit Sub
'  End If
  If Scene = fGame Then
    If bEndLevel = False Then
      If GetAsyncKeyState(GoLeft) < 0 Then
        You.CurX = You.CurX - 5
        bLeft = True
      End If
      If GetAsyncKeyState(GoRight) < 0 Then
        You.CurX = You.CurX + 5
        bRight = True
      End If
      If GetAsyncKeyState(GoForward) < 0 Then
        bForward = True
        You.CurY = You.CurY - 5
        DDS_Thrust.Play DSBPLAY_DEFAULT
      Else
      End If
      If GetAsyncKeyState(GoBack) < 0 Then
        You.CurY = You.CurY + 5
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
    End If
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
  lColorDepth = ReadINI("Settings", "ColorDepth", App.Path & "\data\game.ini")
  bShowFPS = ReadINI("Settings", "ShowFPS", App.Path & "\data\game.ini")
  clsDDraw.Init frmMain.hWnd, ScreenWidth, ScreenHeight, lColorDepth, DDS_Primary, DDS_Buffer, DDSD_Buffer
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
  clsDSound.LoadSoundFile DDS_GetMoney, App.Path & "\sounds\click.wav"
  clsDSound.LoadSoundFile DDS_ExplodeW, App.Path & "\sounds\explo2.wav"
  clsDSound.LoadSoundFile DDS_Menu, App.Path & "\sounds\menu.wav"
  clsDSound.LoadSoundFile DDS_Click, App.Path & "\sounds\btnx.wav"
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
   'OptSel = 0
   mBackX = ((ScreenWidth - 430) \ 2)
   mBackY = ((ScreenHeight - 400) \ 2)
   
'   Do While InOptions = True
    CheckInput
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
     
     If OptSel = 4 Then
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 236, "Show FPS", vbGreen
     Else
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 236, "Show FPS", vbBlue
     End If
     
     
     If bShowFPS = True Then
       clsDDraw.DrawText DDS_Buffer, mBackX + 295, mBackY + 236, "True", &H404040
       clsDDraw.DrawText DDS_Buffer, mBackX + 294, mBackY + 235, "True", vbRed
     Else
       clsDDraw.DrawText DDS_Buffer, mBackX + 295, mBackY + 236, "False", &H404040
       clsDDraw.DrawText DDS_Buffer, mBackX + 294, mBackY + 235, "False", vbRed
     End If
       
     If OptSel = 5 Then
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 256, "Color Depth", vbGreen
     Else
       clsDDraw.DrawText DDS_Buffer, mBackX + 100, mBackY + 256, "ColorDepth", vbBlue
     End If

     clsDDraw.DrawText DDS_Buffer, mBackX + 291, mBackY + 256, Str(lColorDepth), &H404040
     clsDDraw.DrawText DDS_Buffer, mBackX + 290, mBackY + 255, Str(lColorDepth), vbRed
     
   'Loop
   'Menu
End Sub

Public Sub RunTheGame()
    Dim bRestore As Boolean

  Do
    DoEvents
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
    
    If Scene <> fMenu Then
      DDS_Menu.Stop
    End If
    Select Case Scene
      Case fGame
        Main
      Case fMenu
        Menu
      Case fPlayer
        StartNewChar
      Case fEntrance
        ShowHanger
      Case fOptions
        Options
      Case fLoad
        LoadCharProfile
      Case fShop
        ShowShop
    End Select
    If bShowMsg = True Then
      ShowMsg strMsgText
    End If
    If BShipDestroyed = True And bGamePause = False Then
      bInHanger = False
      InGame = False
      BShipDestroyed = False
      bShowMenu = True
      StartMenu
    End If
    DDS_Primary.Flip Nothing, DDFLIP_WAIT
    
    'cFrameLimit.LimitFrames 60
  Loop
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

Public Sub StartMenu()
  StopBGSound
  WhatToDo = 0
  bInNewChar = False
  SelButton = 1
  ShowWall = True
  BackYPos1 = 0
  BackYPos2 = BackYPos1 - ScreenHeight
  
  W1X = 0
  W2X = ScreenWidth / 2
  Scene = fMenu
  DDS_Menu.Play DSBPLAY_LOOPING
End Sub

Public Sub Menu()
  InMenu = True
  Dim rBack As RECT
  
  Dim mBackX As Integer, mBackY As Integer
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
    If ShowWall = True Then
      W1X = W1X - 3
      W2X = W2X + 3
      If (W1X + WallWidth) <= 0 Then ShowWall = False
    End If
    Dim a As POINTAPI
    GetCursorPos a
    
    'CheckInput
    'Draw DDS_Back, rBack1, 0, 0, False, False
    
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
    Draw DDS_HEADER, rHeader, mBackX + 167, mBackY + 40, False, True
    Draw DDS_MenuBack2, rMB, mBackX, mBackY, True, True
    
    Draw DDS_Button, rB1, mBackX + 100, mBackY + 155, True, False
    Draw DDS_Button, rB2, mBackX + 100, mBackY + 155 + (ButHeight), True, False
    Draw DDS_Button, rB3, mBackX + 100, mBackY + 155 + (ButHeight * 2), True, False
    Draw DDS_Button, rB4, mBackX + 100, mBackY + 155 + (ButHeight * 3), True, False
    Draw DDS_WALL1, rW1, W1X, 0, False, True
    Draw DDS_WALL2, rW2, W2X, 0, False, True
    Draw DDS_Cursor, rCursor, a.x - 10, a.Y - 10, True, True
    
  
  If WhatToDo = 1 Then
    bGamePause = False
    strUserName = ""
    strCallSign = ""
    bInNewChar = True
    bAtCallsign = False
    Scene = fPlayer
  ElseIf WhatToDo = 2 Then
    OptSel = 0
    WhatToDo = 0
    Scene = fOptions
  ElseIf WhatToDo = 3 Then
    CloseGame
  End If
End Sub

Public Sub ExecuteClick(x As Integer, Y As Integer)
  Dim mBackX As Integer, mBackY As Integer
  mBackX = ((ScreenWidth - 430) / 2)
  mBackY = ((ScreenHeight - 400) \ 2)
  
  If x > (mBackX + 100) And x < (mBackX + 100 + ButWidth) Then
    If Y > (mBackY + 155) And Y < (mBackY + 155 + ButHeight) And SelButton = 1 Then
      InMenu = False
      WhatToDo = 1
    End If
    If Y > (mBackY + 155 + ButHeight) And Y < (mBackY + 155 + (ButHeight * 2)) And SelButton = 2 Then
      LoadCursorPos = 0
      Scene = fLoad
      WhatToDo = 4
          
    End If
    If Y > (mBackY + 155 + (ButHeight * 2)) And Y < (mBackY + 155 + (ButHeight * 3)) And SelButton = 3 Then
      InMenu = False
      WhatToDo = 2
    End If
    If Y > (mBackY + 155 + (ButHeight * 3)) And Y < (mBackY + 155 + (ButHeight * 4)) And SelButton = 4 Then
      InMenu = False
      WhatToDo = 3
    End If
  End If
End Sub


Public Sub HighlightBut(x As Integer, Y As Integer)
  Dim mBackX As Integer, mBackY As Integer
  mBackX = ((ScreenWidth - 430) / 2)
  mBackY = ((ScreenHeight - 400) \ 2)
  If x > (mBackX + 100) And x < (mBackX + 100 + ButWidth) Then
    'The potential to be in a button exists. Their all aligned on x coords so
    'after this check were good
    
    If Y > (mBackY + 155) And Y < (mBackY + 155 + ButHeight) Then
      SelButton = 1
    End If
    If Y > (mBackY + 155 + ButHeight) And Y < (mBackY + 155 + (ButHeight * 2)) Then
      SelButton = 2
    End If
    If Y > (mBackY + 155 + (ButHeight * 2)) And Y < (mBackY + 155 + (ButHeight * 3)) Then
      SelButton = 3
    End If
    If Y > (mBackY + 155 + (ButHeight * 3)) And Y < (mBackY + 155 + (ButHeight * 4)) Then
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
  ClearAllEnemies
  LoadWorld "lvl1.lvl"

  'bAtBoss = True
  BBMoveDir = 0
  BackYPos1 = 0
  BackYPos2 = -ScreenHeight
  NewGuy
  bEndLevel = False
  bAtBoss = False
  
  Scene = fGame
End Sub


Public Sub StartNewChar()
  Dim tmpRect As RECT
  Dim lJX As Long, lJY As Long
  Dim BufFont As IFont
  Dim iFile As Integer
  Dim tmpBufFont As StdFont
  lJX = (ScreenWidth - JWidth) \ 2
  lJY = (ScreenHeight - JHeight) \ 2
  Set tmpBufFont = New StdFont
  tmpBufFont.Name = "Courier New"
  tmpBufFont.Size = 10
  tmpBufFont.Bold = False
  You.ReactorPower = 5
  You.CurMoney = 10000
  You.ShieldR = 1
  You.lPlasma = 0
  You.lPulse = 0
  You.lMicro = 0
  iFile = FreeFile
  Open App.Path & "\levels\lvl1.lgf" For Input As #iFile
    Input #iFile, You.Level
  Close #iFile
  You.iLevel = 1
  Set BufFont = tmpBufFont
  DDS_Buffer.SetFont BufFont
  
    DDS_Buffer.BltColorFill rBack1, 0
    DDS_Buffer.SetForeColor RGB(192, 192, 192)
    'DDS_Buffer.SetFont
    Draw DDS_NEWEARTH, rBackEarth, 0, 0, False, True
    Draw DDS_NEWBACK, rNewBack, lJX, lJY, True, True
    
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
  
  
End Sub
Public Sub ShowHanger()
  Dim rctBottom As RECT
  bInHanger = True
  bShowingHanger = True
  bEndLevel = False
    
    Draw DDS_HANGER, rHanger, 0, 0, False, True
    With rctBottom
      .Left = 1
      .Top = ScreenHeight - 49
      .Right = ScreenWidth - 1
      .Bottom = ScreenHeight - 1
    End With
    DDS_Buffer.BltColorFill rctBottom, 0
    If strHangerText = "" Then strHangerText = " "
    DDS_Buffer.SetForeColor RGB(192, 192, 192)
    DDS_Buffer.DrawText 11, 453, strHangerText, False
    DDS_Buffer.SetForeColor RGB(256, 0, 0)
    DDS_Buffer.DrawText 10, 452, strHangerText, False
    DDS_Buffer.DrawBox 0, ScreenHeight - 50, ScreenWidth, ScreenHeight
    
    Dim a As POINTAPI
    GetCursorPos a
    CheckInput
    Draw DDS_Cursor, rCursor, a.x - 10, a.Y - 10, True, True
    DoEvents
End Sub
Public Sub Main()
    'draw current game screen
    Timer
    If bGamePause = False Then CheckInput
    DrawFrame
    If bGamePause = False Then BackYPos1 = BackYPos1 + 1
    If bGamePause = False Then BackYPos2 = BackYPos2 + 1
    If bGamePause = False Then
    If clsDDraw.TickCount - LastEnergyUpdateTick > ShieldEnergyUpdate Then
      LastEnergyUpdateTick = clsDDraw.TickCount
      You.CurEnergy = You.CurEnergy + You.ReactorPower
      If You.Shield < You.MaxShield Then
        If You.CurEnergy > 250 Then
          You.Shield = You.Shield + You.ShieldR
          You.CurEnergy = You.CurEnergy - (You.ShieldR * 6)
        End If
      End If
      If You.CurEnergy > 500 Then You.CurEnergy = 500
      If You.Shield > You.MaxShield Then You.Shield = You.MaxShield
    End If
    End If
End Sub

Public Sub DrawWord(x As Long, Y As Long, strText As String)
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
    Draw DDS_Letters, rTmp, x + ((i - 1) * 9), Y, True, False
  Next
End Sub

Private Sub LoadData()
  'Load data
  'Clear all the variables if they exist
  Dim i As Integer
  Set DDS_YOU = Nothing
  Set DDS_PULSEC = Nothing
  Set DDS_PLASMAB = Nothing
  Set DDS_PLASMA = Nothing
  Set DDS_MICRO = Nothing
  Set DDS_REACTOR = Nothing
  Set DDS_SHOP = Nothing
  Set DDS_PULSE = Nothing
  Set DDS_YOUSMALL = Nothing
  Set DDS_DISPLAY = Nothing
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
  Set DDS_Save = Nothing
  Set DDS_MESSAGE = Nothing
  clsDDraw.DDCreateSurface DDS_Money, App.Path & "\images\money.bmp", rMoney, 32, 16, 0
  clsDDraw.DDCreateSurface DDS_SHIELD, App.Path & "\images\shield.bmp", rShield, , , , 0
  clsDDraw.DDCreateSurface DDS_MESSAGE, App.Path & "\images\msg.bmp", rMessage, , , 0
  clsDDraw.DDCreateSurface DDS_NEWEARTH, App.Path & "\images\earthmoon.bmp", rBackEarth, , , 0
  clsDDraw.DDCreateSurface DDS_YOU, App.Path & "\images\newmain.bmp", rYou, , , 0
  clsDDraw.DDCreateSurface DDS_PULSEC, App.Path & "\images\pulsec.bmp", rPulseC, , , 0
  clsDDraw.DDCreateSurface DDS_PLASMAB, App.Path & "\images\plasmab.bmp", rPlasma, , , 0
  clsDDraw.DDCreateSurface DDS_PLASMA, App.Path & "\images\plasma.bmp", rPlasmaS, , , 0
  clsDDraw.DDCreateSurface DDS_MICRO, App.Path & "\images\micro.bmp", rMicro, , , 0
  clsDDraw.DDCreateSurface DDS_REACTOR, App.Path & "\images\reactor.bmp", rReactor, , , 0
  clsDDraw.DDCreateSurface DDS_SHOP, App.Path & "\images\shope.bmp", rShop, , , 0
  clsDDraw.DDCreateSurface DDS_SHOPSELL, App.Path & "\images\shopsell.bmp", rShopSell, , , 0
  clsDDraw.DDCreateSurface DDS_PULSE, App.Path & "\images\pulse.bmp", rPulse, , , 0
  clsDDraw.DDCreateSurface DDS_DISPLAY, App.Path & "\images\display.bmp", rDisplay, , , 0
  clsDDraw.DDCreateSurface DDS_YOUSMALL, App.Path & "\images\mainsmall.bmp", rYouSmall, , , 0
  clsDDraw.DDCreateSurface DDS_HEADER, App.Path & "\images\head.bmp", rHeader, , , 0
  clsDDraw.DDCreateSurface DDS_NEWBACK, App.Path & "\images\jreg.bmp", rNewBack, , , 0
  clsDDraw.DDCreateSurface DDS_HIT, App.Path & "\images\hit.bmp", rHit, , , 0
  clsDDraw.DDCreateSurface DDS_PSHOT, App.Path & "\images\bullet.bmp", rBullet, , , 0
  clsDDraw.DDCreateSurface DDS_Missile, App.Path & "\images\missile.bmp", rMissile, , , 0
  clsDDraw.DDCreateSurface DDS_MissileE, App.Path & "\images\missilee.bmp", rMissileE, , , 0
  clsDDraw.DDCreateSurface DDS_Cursor, App.Path & "\images\cursor.bmp", rCursor, , , 0
  clsDDraw.DDCreateSurface DDS_MenuBack, App.Path & "\images\space03.bmp", rBack1, , , 0
  clsDDraw.DDCreateSurface DDS_ESHOT, App.Path & "\images\bullete.bmp", rBullet, , , 0
  clsDDraw.DDCreateSurface DDS_Bottom, App.Path & "\images\newbot.bmp", rBottom, DDSD_Buffer.lWidth, 64, 0
  clsDDraw.DDCreateSurface DDS_HEALTH, App.Path & "\images\health.bmp", rHealth
  clsDDraw.DDCreateSurface DDS_HBAR, App.Path & "\images\hbar.bmp", rHBar
  clsDDraw.DDCreateSurface DDS_HANGER, App.Path & "\images\hanger.bmp", rHanger
  clsDDraw.DDCreateSurface DDS_Explode, App.Path & "\images\explode.bmp", rExplode
  clsDDraw.DDCreateSurface DDS_WALL1, App.Path & "\images\wall1.bmp", rW1
  clsDDraw.DDCreateSurface DDS_WALL2, App.Path & "\images\wall2.bmp", rW2
  clsDDraw.DDCreateSurface DDS_Button, App.Path & "\images\buttons.bmp", rB1
  clsDDraw.DDCreateSurface DDS_MenuBack2, App.Path & "\images\menuback.bmp", rMB
  clsDDraw.DDCreateSurface DDS_Save, App.Path & "\images\save.bmp", rSave
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
  Dim x As Integer, Y As Integer
  Dim rMoney As Integer
  
  ShotTickCount = clsDDraw.TickCount
  For x = 0 To MaxShots
    If (PShotL(x).active = True Or PShotR(x).active = True) Then
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
           
      For Y = 0 To MaxEnemies
        If PShotR(x).active = True And BadGuys(Y).active Then
          If BadGuys(Y).active = True And PShotL(x).active = True Then
            If CheckCollide(DDS_PSHOT, BadGuys(Y).Surface, PShotL(x).CurX, PShotL(x).CurY, ShotWidth, ShotHeight, BadGuys(Y).x, BadGuys(Y).Y, BadGuys(Y).Width, BadGuys(Y).Height, 0) Then
              CreateHit PShotL(x).CurX, PShotL(x).CurY, BadGuys(Y).Velocity
'              PlayWav "hitp.wav"
                If BadGuys(Y).Shield > 0 Then
                  BadGuys(Y).Shield = BadGuys(Y).Shield - 30
                Else
                  BadGuys(Y).Hull = BadGuys(Y).Hull - 50
                End If
                If BadGuys(Y).Shield <= 0 And BadGuys(Y).Hull <= 0 Then
                  BadGuys(Y).active = False
                  
                  If BadGuys(Y).bBoss Then
                    bAtBoss = False
                    bEndLevel = True
                    You.iLevel = You.iLevel + 1
                  End If
                  AddExplode BadGuys(Y).x, BadGuys(Y).Y
                  You.CurMoney = You.CurMoney + BadGuys(Y).Value
                  rMoney = Int(Rnd * 100)
                  If rMoney > 1 Then
                    AddMoney (BadGuys(Y).x + (BadGuys(Y).Width \ 2)), (BadGuys(Y).Y + (BadGuys(Y).Height \ 2))
                  End If
                End If
                PShotL(x).active = False
              End If
            End If
          If PShotR(x).active = True And BadGuys(Y).active Then
            If CheckCollide(DDS_PSHOT, BadGuys(Y).Surface, PShotR(x).CurX, PShotR(x).CurY, ShotWidth, ShotHeight, BadGuys(Y).x, BadGuys(Y).Y, BadGuys(Y).Width, BadGuys(Y).Height, 0) Then
              CreateHit PShotR(x).CurX, PShotR(x).CurY, BadGuys(Y).Velocity
'             PlayWav "hitp.wav"
              If BadGuys(Y).Shield > 0 Then
                BadGuys(Y).Shield = BadGuys(Y).Shield - 30
              Else
                BadGuys(Y).Hull = BadGuys(Y).Hull - 50
              End If
              If BadGuys(Y).Shield <= 0 And BadGuys(Y).Hull <= 0 Then
                'BadShips.Remove y
                BadGuys(Y).active = False
                  If BadGuys(Y).bBoss Then
                    bAtBoss = False
                    bEndLevel = True
                  End If
                AddExplode BadGuys(Y).x, BadGuys(Y).Y
                rMoney = Int(Rnd * 100)
                If rMoney > 1 Then
                  AddMoney (BadGuys(Y).x + (BadGuys(Y).Width \ 2)), (BadGuys(Y).Y + (BadGuys(Y).Height \ 2))
                End If
              
                You.CurMoney = You.CurMoney + BadGuys(Y).Value
              End If
              PShotR(x).active = False
            End If
          End If
        End If
      Next
    Else
      If Firing = True And ((ShotTickCount - LastShotTick) > ShotTick) Then
        CreateShot x
      End If
        
    End If
      
  Next x
    
End Sub

Public Sub UpdateMissiles()
  Dim x As Integer, Y As Integer
  Dim rMoney As Integer
  ShotTickCount = clsDDraw.TickCount
  For x = 0 To MaxShots
    If (MissileL(x).active = True Or MissileR(x).active = True) Then
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
      
      For Y = 0 To MaxEnemies
        If MissileL(x).active Or MissileR(x).active = True Then
          If BadGuys(Y).active = True And MissileL(x).active = True Then
            If CheckCollide(DDS_Missile, BadGuys(Y).Surface, MissileL(x).CurX, MissileL(x).CurY, ShotWidth, ShotHeight, BadGuys(Y).x, BadGuys(Y).Y, BadGuys(Y).Width, BadGuys(Y).Height, 0) Then
              CreateHit MissileL(x).CurX, MissileL(x).CurY, BadGuys(Y).Velocity
'             PlayWav "hitp.wav"
                If BadGuys(Y).Shield > 0 Then
                  BadGuys(Y).Shield = BadGuys(Y).Shield - 30
                Else
                  BadGuys(Y).Hull = BadGuys(Y).Hull - 50
                End If
                If BadGuys(Y).Shield <= 0 And BadGuys(Y).Hull <= 0 Then
                  BadGuys(Y).active = False
                  If BadGuys(Y).bBoss Then
                    bAtBoss = False
                    bEndLevel = True
                  End If
                  AddExplode BadGuys(Y).x, BadGuys(Y).Y
                  rMoney = Int(Rnd * 100)
                  If rMoney > 1 Then
                    AddMoney (BadGuys(Y).x + (BadGuys(Y).Width \ 2)), (BadGuys(Y).Y + (BadGuys(Y).Height \ 2))
                  End If
                  
                  You.CurMoney = You.CurMoney + BadGuys(x).Value
                End If
                MissileL(x).active = False
              End If
            End If
          If BadGuys(Y).active = True And MissileR(x).active = True Then
            If CheckCollide(DDS_Missile, BadGuys(Y).Surface, MissileR(x).CurX, MissileR(x).CurY, ShotWidth, ShotHeight, BadGuys(Y).x, BadGuys(Y).Y, BadGuys(Y).Width, BadGuys(Y).Height, 0) Then
              CreateHit MissileR(x).CurX, MissileR(x).CurY, BadGuys(Y).Velocity
'             PlayWav "hitp.wav"
              If BadGuys(Y).Shield > 0 Then
                BadGuys(Y).Shield = BadGuys(Y).Shield - 30
              Else
                BadGuys(Y).Hull = BadGuys(Y).Hull - 50
              End If
              If BadGuys(Y).Shield <= 0 And BadGuys(Y).Hull <= 0 Then
                BadGuys(Y).active = False
                If BadGuys(Y).bBoss Then
                  bAtBoss = False
                  bEndLevel = True
                End If
                AddExplode BadGuys(Y).x, BadGuys(Y).Y
                rMoney = Int(Rnd * 100)
                If rMoney > 1 Then
                  AddMoney (BadGuys(Y).x + (BadGuys(Y).Width \ 2)), (BadGuys(Y).Y + (BadGuys(Y).Height \ 2))
                End If
                
                You.CurMoney = You.CurMoney + BadGuys(x).Value
              End If
              MissileR(x).active = False
            End If
          End If
        End If
      Next Y
    Else
      If Firing = True And ((ShotTickCount - MissileShotTick) > MissileShotCount) Then
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
  You.CurEnergy = You.CurEnergy - 5
  PShotL(ShotI).CurX = (You.CurX + 14)
  PShotR(ShotI).CurX = (You.CurX + 34)
  PShotL(ShotI).CurY = (You.CurY + 2)
  PShotR(ShotI).CurY = (You.CurY + 2)
  LastShotTick = clsDDraw.TickCount
End Sub

Public Sub CreatePulse(ShotI As Integer)
  Dim i As Integer
  If You.CurEnergy <= 0 Then Exit Sub
  Pulse(ShotI).active = True
  If You.lPulse < 1 Then Exit Sub
  You.CurEnergy = You.CurEnergy - 20
  Pulse(ShotI).CurX = (You.CurX + ((YouWidth - 26) \ 2)) + 2
  Pulse(ShotI).CurY = (You.CurY + 10)
  PulseShotTick = clsDDraw.TickCount
End Sub

Public Sub CreatePlasma(ShotI As Integer)
  Dim i As Integer
  If You.lPlasma < 1 Then Exit Sub
  If You.CurEnergy <= 0 Then Exit Sub
  Plasma(ShotI).active = True
  You.CurEnergy = You.CurEnergy - 2
  Plasma(ShotI).CurX = (You.CurX + ((YouWidth - (rPlasmaS.Right)) \ 2)) + 2
  Plasma(ShotI).CurY = (You.CurY + 10)
  PlasmaShotTick = clsDDraw.TickCount
End Sub


Public Sub CreateMissile(ShotI As Integer)
  Dim i As Integer
  If You.lMicro < 1 Then Exit Sub
  If You.CurEnergy <= 0 Then Exit Sub
  You.CurEnergy = You.CurEnergy - 2
  MissileL(ShotI).active = True
  MissileR(ShotI).active = True
  MissileL(ShotI).CurX = (You.CurX + 10)
  MissileR(ShotI).CurX = (You.CurX + 38)
  MissileL(ShotI).CurY = (You.CurY + 10)
  MissileR(ShotI).CurY = (You.CurY + 10)
  MissileShotTick = clsDDraw.TickCount
End Sub


Public Sub DrawFrame()
    
    Dim ddrVal As Long 'Every drawing procedure returns a value, so we must have a
                       'var able to hold it. From this value we can check for errors.
    If BackYPos1 >= ScreenHeight Then BackYPos1 = BackYPos2 - ScreenHeight
    If BackYPos2 >= ScreenHeight Then BackYPos2 = BackYPos1 - ScreenHeight
    UpdateTiles
    If bEndLevel Then
      Firing = False
      If You.CurX > ((ScreenWidth - YouWidth) \ 2) Then
        You.CurX = You.CurX - 1
      ElseIf You.CurX < ((ScreenWidth - YouWidth) \ 2) Then
        You.CurX = You.CurX + 2
      End If
      If You.CurX = ((ScreenWidth - YouWidth) \ 2) Then
        You.CurY = You.CurY + 2
      End If
      If You.CurY > ScreenHeight Then
        InGame = False
        bInHanger = True
        Scene = fEntrance
        You.iLevel = You.iLevel + 1
        'bEndLevel = False
        StopBGSound
        bEndLevel = False
        RunTheGame
        'ShowHanger
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
    
    If bGamePause = False Then
      UpdateEnemyShips
      UpdateMoney
      UpdateHits
    
      UpdateShots
      UpdateMissiles
      UpdatePulseShots
      UpdatePlasmaShots
      UpdateEFiring
      UpdateEMissile
      UpdateExplode
      DDS_Buffer.BltFast You.CurX, You.CurY, DDS_YOU, rYou, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT
      
    Else
      DrawInGameMenu
    End If
    ParseLevel
    DDS_Buffer.SetForeColor RGB(150, 150, 150)
    Call DDS_Buffer.DrawLine(0, ScreenHeight - 66, ScreenWidth, ScreenHeight - 66)
    Call DDS_Buffer.DrawLine(0, ScreenHeight - 65, ScreenWidth, ScreenHeight - 65)
    DDS_Buffer.SetForeColor RGB(180, 180, 180)
    Call DDS_Buffer.DrawLine(0, ScreenHeight - 67, ScreenWidth, ScreenHeight - 67)
    ddrVal = Draw(DDS_Bottom, rBottom, 0, ScreenHeight - 64, False, True)
    
    DrawHealth
    
    DrawYouDisplay
    If clsDDraw.TickCount - LastTimeChecked >= 1000 Then
        LastTimeChecked = clsDDraw.TickCount
        FrameText = "FPS: " & CStr(framesDone) & " fps"
        framesDone = 0
    End If
    Dim rOver As RECT
    DDS_Buffer.SetForeColor vbRed
    If bShowFPS = True Then DDS_Buffer.DrawText 5, 5, FrameText, False
    'DrawWord 300, 420, "hello there 0123456"
    'DDS_Buffer.DrawText 250, 460, "$" & Str(You.CurMoney), False
    'DDS_Buffer.DrawText 18, 460, clrBack, False
    framesDone = framesDone + 1
End Sub

Public Sub DrawYouDisplay()
  Dim x As Double
  Draw DDS_YOUSMALL, rYouSmall, ScreenWidth - rYouSmall.Right - 10, ScreenHeight - rYouSmall.Bottom - 5, False, False
  x = (You.Shield / You.MaxShield)
  x = x * 100
  If (x > 90) Then
    DDS_Buffer.SetForeColor vbGreen
  ElseIf (x > 50) And (x < 90) Then
    DDS_Buffer.SetForeColor vbYellow
  ElseIf (x > 10) And (x < 50) Then
    DDS_Buffer.SetForeColor vbRed
  Else
    DDS_Buffer.SetForeColor &H404040
  End If
  DDS_Buffer.DrawCircle (ScreenWidth - 41), (ScreenHeight - 31), 22
  DDS_Buffer.SetForeColor vbYellow
  
  Draw DDS_DISPLAY, rDisplay, 300, ScreenHeight - 60, False, False
  DDS_Buffer.DrawText 305, ScreenHeight - 57, "Cash", False
  DDS_Buffer.DrawText 350, ScreenHeight - 57, You.CurMoney, False
  Draw DDS_DISPLAY, rDisplay, 300, ScreenHeight - 30, False, False
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
    DrawYouDisplay
    BackYPos1 = BackYPos1 + 1
    BackYPos2 = BackYPos2 + 1
    DDS_Primary.Flip Nothing, DDFLIP_WAIT
    DoEvents
  Next
  DoEvents
  ShipInvincible = False
End Sub



Public Sub UpdateEnemyShips()
  Dim x As Integer, i As Integer
  Dim bShipOnScreen As Boolean
  bShipOnScreen = False
  For x = 0 To MaxEnemies 'MaxEnemyShips
    If BadGuys(x).active = True Then
      bShipOnScreen = True
      If bAtBoss = False Then BadGuys(x).Y = BadGuys(x).Y + BadGuys(x).Velocity
      If BadGuys(x).AI = 1 Then
        AIBounce (x)
      ElseIf BadGuys(x).AI = 2 Then
        DownOff (x)
      ElseIf BadGuys(x).AI = 3 Then
        BigBoss (x)
      Else
        If BadGuys(x).CanShoot Then
          If (35 + (Difficulty * 50) > Int(Rnd * 1000)) Then
            CreateEShot FindEShot, x
          End If
        End If
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
      Draw BadGuys(x).Surface, BadGuys(x).RECT, BadGuys(x).x, BadGuys(x).Y, True, True
      If CheckCollide(DDS_YOU, BadGuys(x).Surface, You.CurX, You.CurY, YouWidth, YouHeight, BadGuys(x).x, BadGuys(x).Y, BadGuys(x).Width, BadGuys(x).Height, 0) And BadGuys(x).active = True And (BadGuys(x).bBoss = False) Then
        If You.Shield > 0 Then
          If ShipInvincible = False Then You.Shield = You.Shield - (30 + (Difficulty * 5))
          'Beep
          BadGuys(x).active = False
                If bAtBoss Then
                  bAtBoss = False
                  bEndLevel = True
                End If
          AddExplode BadGuys(x).x, BadGuys(x).Y
          You.CurMoney = You.CurMoney + BadGuys(x).Value
        Else
          If You.Hull > 0 Then
            If ShipInvincible = False Then You.Hull = You.Hull - (50 + (Difficulty * 5))
            BadGuys(x).active = False
                If bAtBoss Then
                  bAtBoss = False
                  bEndLevel = True
                End If
            AddExplode BadGuys(x).x, BadGuys(x).Y
          End If
        End If
        If You.Hull <= 0 And You.Shield <= 0 Then
          AddExplode You.CurX, You.CurY
          strMsgText = "Your ship has been destroyed.  You have lost all posessions."
          bShowMsg = True
          BShipDestroyed = True
        End If
      End If
      If BadGuys(x).Y >= ScreenHeight Then

        BadGuys(x).active = False
        If LineNum = MaxRows Then
          For i = 0 To MaxEnemies
            If BadGuys(i).active = False Then
              bCanDoBoss = True
            End If
          Next
        End If
      End If
      
      
      'If ((clsDDraw.TickCount - LastGuyTick) > NewBadGuyTick) Then CreateEnemy x
    End If
  Next
  If bShipOnScreen = False And bAtBoss = True Then
    'Create Boss :)
    CreateBoss
  End If
End Sub



Private Function Draw(Surface As DirectDrawSurface7, RECTvar As RECT, ByVal x As Integer, ByVal Y As Integer, Optional transparent As Boolean = True, Optional Clip As Boolean = True) As Long
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
            .Top = Y
            .Left = x
            .Bottom = Y + RECTvar.Bottom - RECTvar.Top
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
                Y = 0
            End If
        End With
    
    End If
    If transparent = False Then
        Draw = DDS_Buffer.BltFast(x, Y, Surface, RectTEMP, DDBLTFAST_WAIT)
    Else
        Draw = DDS_Buffer.BltFast(x, Y, Surface, RectTEMP, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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


Private Sub CreateHit(x As Integer, Y As Integer, speed As Integer)
  Dim NewHit As Integer
  NewHit = FindHit()
  Hits(NewHit).active = True
  Hits(NewHit).frame = 0
  Hits(NewHit).LastTick = clsDDraw.TickCount
  Hits(NewHit).x = x
  Hits(NewHit).Y = Y
  Hits(NewHit).speed = speed
End Sub

Public Sub UpdateHits()
  Dim x As Integer, NewTick As Long
  NewTick = clsDDraw.TickCount
  For x = 0 To MaxHit
    If (Hits(x).active = True) Then
      Hits(x).Y = Hits(x).Y + Hits(x).speed
      With rHit
        .Top = 0
        .Bottom = HitHeight
        .Left = Hits(x).frame * HitWidth
        .Right = .Left + HitWidth
      End With
      'BitBlt frmMain.picBak.hDC, Hits(X).X, Hits(X).y, HitWidth, HitHeight, HitDC, Hits(X).Frame * HitWidth, 12, vbSrcAnd
      'BitBlt frmMain.picBak.hDC, Hits(X).X, Hits(X).y, HitWidth, HitHeight, HitDC, Hits(X).Frame * HitWidth, 0, vbSrcPaint
      Draw DDS_HIT, rHit, Hits(x).x, Hits(x).Y, True, False
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

Public Sub AddExplode(x, Y)
  Dim NewExplode As Integer
  DDS_ExplodeW.Play DSBPLAY_DEFAULT
  NewExplode = FindExplode
  Explodes(NewExplode).active = True
  Explodes(NewExplode).x = x
  Explodes(NewExplode).Y = Y
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

      Draw DDS_Explode, rTemp, Explodes(x).x, Explodes(x).Y, True, True
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
          If ShipInvincible = False Then You.Shield = You.Shield - (10 + (Difficulty * 5))
        Else
          If ShipInvincible = False Then You.Hull = You.Hull - (15 + (Difficulty * 5))
        End If
        If You.Shield <= 0 And You.Hull <= 0 Then
          AddExplode You.CurX, You.CurY
          Firing = False
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

Public Sub UpdateEMissile()
  Dim x As Integer
  For x = 0 To MaxEShots
    If (EMisL(x).active = True Or EMisR(x).active = True) Then
      If EMisL(x).active = True Then EMisL(x).CurY = EMisL(x).CurY + ShotVelocity
      If EMisR(x).active = True Then EMisR(x).CurY = EMisR(x).CurY + ShotVelocity
      If EMisL(x).CurY >= ScreenHeight Then EMisL(x).active = False
      If EMisR(x).CurY >= ScreenHeight Then EMisR(x).active = False
      If EMisR(x).active = True And CheckCollide(DDS_MissileE, DDS_YOU, EMisR(x).CurX, EMisR(x).CurY, 4, 16, You.CurX, You.CurY, YouWidth, YouHeight, 0) Then
        CreateHit EMisL(x).CurX, (EMisL(x).CurY + 16), 0
        EMisR(x).active = False
        If You.Shield > 0 Then
          If ShipInvincible = False Then You.Shield = You.Shield - (10 + (Difficulty * 5))
        Else
          If ShipInvincible = False Then You.Hull = You.Hull - (35 + (Difficulty * 5))
        End If
        If You.Shield <= 0 And You.Hull <= 0 Then
          AddExplode You.CurX, You.CurY
          Firing = False
        End If
      End If
      If EMisL(x).active = True And CheckCollide(DDS_MissileE, DDS_YOU, EMisL(x).CurX, EMisL(x).CurY, 4, 16, You.CurX, You.CurY, YouWidth, YouHeight, 0) Then
        CreateHit EMisL(x).CurX, (EMisL(x).CurY + 16), 0
        EMisL(x).active = False
        If You.Shield > 0 Then
          If ShipInvincible = False Then You.Shield = You.Shield - 10
        Else
          If ShipInvincible = False Then You.Hull = You.Hull - (35 + (Difficulty * 5))
        End If
        If You.Shield <= 0 And You.Hull <= 0 Then
          AddExplode You.CurX, You.CurY
          Firing = False
        End If
      End If
        
      If EMisL(x).active = True Then
        Draw DDS_MissileE, rMissileE, EMisL(x).CurX, EMisL(x).CurY, True, False
      End If
      If EMisR(x).active = True Then
        Draw DDS_MissileE, rMissileE, EMisR(x).CurX, EMisR(x).CurY, True, False
      End If
    End If
  Next
End Sub


Public Sub LoadWorld(lvl As String)
  Dim iFile As Integer
  Dim i As Integer
  iFile = FreeFile()
  'Open App.Path & "\levels\" & lvl For Binary Access Read As #iFile
  '  Get #iFile, , Level
  'Close #iFile
  Open App.Path & "\levels\lvl1.lgf" For Input As #iFile
    For i = 1 To You.iLevel
      Input #iFile, You.Level
    Next i
  Close #1
  Debug.Print You.Level
  LoadLevel You.Level 'Level.EnemyFile
  clsDDraw.DDCreateSurface DDS_Back, App.Path & "\temp\tmpBack.bmp", rBack1, , , 0
  PlayBGSound App.Path & "\temp\tmpBGMusic.mid"
  
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
                Tiles2(lngXSize, lngYSize).Y = lngYSize * -(TileHeight)
            Next
        Next
  Close #iFile
End Sub

Public Sub AddMoney(x As Integer, Y As Integer)
  Dim i As Integer
  Dim v As Integer
  ' lets first determine if we should place money
  Randomize Time
  v = Int(Rnd * 150) + 1
  If v < 100 Then Exit Sub
  ' First find an innactive money
  For i = 0 To 20
    If mMoney(i).active = False Then Exit For ' We've found it
  Next
  If i > 20 Then i = 20
  mMoney(i).frame = 0
  mMoney(i).LastTick = clsDDraw.TickCount()
  mMoney(i).x = x
  mMoney(i).Y = Y
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
      mMoney(i).Y = mMoney(i).Y + 1
      
      With mMoney(i).RECT
        
        .Left = mMoney(i).frame * 16
        .Right = .Left + 16
        .Top = 0
        .Bottom = 16
      End With
      Draw DDS_Money, mMoney(i).RECT, mMoney(i).x, mMoney(i).Y, True, True
      If CheckCollide(DDS_YOU, DDS_Money, You.CurX, You.CurY, YouWidth, YouHeight, mMoney(i).x, mMoney(i).Y, 16, 16, 0) Then
        mMoney(i).active = False
        DDS_GetMoney.Play DSBPLAY_DEFAULT
        You.CurMoney = You.CurMoney + 1000
      End If
      
      If mMoney(i).Y > ScreenHeight Then mMoney(i).active = False
    End If
  Next i
End Sub

Public Sub RemoveBadGuy(iIndex As Integer)
  BadShips.Remove iIndex
End Sub

Public Sub DrawInGameMenu()
  Dim rMenuFill As RECT
  Dim rMenuI1 As RECT, rMenuI2 As RECT
  DDS_Buffer.SetForeColor vbBlue
  DDS_Buffer.DrawBox (ScreenWidth - 140) \ 2, (ScreenHeight - 80) \ 2, ((ScreenWidth - 140) \ 2) + 140, ((ScreenHeight - 80) \ 2) + 80
  With rMenuFill
    .Left = ((ScreenWidth - 140) \ 2) + 1
    .Top = ((ScreenHeight - 80) \ 2) + 1
    .Right = (((ScreenWidth - 140) \ 2) + 140) - 1
    .Bottom = (((ScreenHeight - 80) \ 2) + 80) - 1
  End With
  'DDS_Buffer.BltColorFill rMenuFill, 0
  With rMenuI1
    .Left = ((ScreenWidth - 140) \ 2) + 2
    .Top = ((ScreenHeight - 80) \ 2) + 2
    .Right = (((ScreenWidth - 140) \ 2) + 140) - 2
    .Bottom = (((ScreenHeight - 80) \ 2) + 39)
  End With
  
  With rMenuI2
    .Left = ((ScreenWidth - 140) \ 2) + 2
    .Top = ((ScreenHeight - 80) \ 2) + 41
    .Right = (((ScreenWidth - 140) \ 2) + 140) - 2
    .Bottom = (((ScreenHeight - 80) \ 2) + 78)
  End With
  If iInMenuSel = 1 Then
    DDS_Buffer.SetFillStyle 0
    DDS_Buffer.SetFillColor RGB(11, 19, 91)
    DDS_Buffer.SetForeColor vbRed
    DDS_Buffer.SetFillStyle 0
    DDS_Buffer.DrawBox rMenuI1.Left, rMenuI1.Top, rMenuI1.Right, rMenuI1.Bottom
    DDS_Buffer.SetFillColor RGB(192, 192, 192)
    DDS_Buffer.SetForeColor vbRed
    DDS_Buffer.DrawBox rMenuI2.Left, rMenuI2.Top, rMenuI2.Right, rMenuI2.Bottom
    DDS_Buffer.DrawText (rMenuI1.Left + 3), (rMenuI1.Top + 10), "Return to Hanger", False
    DDS_Buffer.SetForeColor vbBlue
    DDS_Buffer.DrawText rMenuI2.Left + 45, rMenuI2.Top + 10, "Cancel", False
    
  Else
    DDS_Buffer.SetFillStyle 0
    DDS_Buffer.SetFillColor RGB(192, 192, 192)
    DDS_Buffer.SetForeColor vbBlue
    DDS_Buffer.SetFillStyle 0
    DDS_Buffer.DrawBox rMenuI1.Left, rMenuI1.Top, rMenuI1.Right, rMenuI1.Bottom
    DDS_Buffer.SetFillColor RGB(11, 19, 91)
    DDS_Buffer.SetForeColor vbBlue
    DDS_Buffer.DrawBox rMenuI2.Left, rMenuI2.Top, rMenuI2.Right, rMenuI2.Bottom
    DDS_Buffer.DrawText (rMenuI1.Left + 3), (rMenuI1.Top + 10), "Return to Hanger", False
    DDS_Buffer.SetForeColor vbRed
    DDS_Buffer.DrawText rMenuI2.Left + 45, rMenuI2.Top + 10, "Cancel", False
  End If
  DDS_Buffer.SetFillStyle 1
  
  
End Sub

Public Sub SaveCharProfile()
  ' Just an ini file to store players at this time.  May upgrade with time.
  Dim iFile As Integer
  Dim strStore As String
  iFile = FreeFile
  Open App.Path & "\save\" & You.strUserName & ".sav" For Binary Access Write As #iFile
    Put #iFile, , You
  Close #iFile
End Sub

Public Sub LoadProfile()
  ' Just an ini file to store players at this time.  May upgrade with time.
  Dim strStore As String
  Dim strFiles(1000) As String
  Dim strFile As String
  Dim x As Integer
  strStore = Dir(App.Path & "\save\")
  Dim b As Integer
  x = 0
  
  Do Until LenB(strStore) = 0
    If Len(strStore) < 35 Then
      strFiles(x) = strStore
    Else
      strFiles(x) = Left(strStore, 32) & "..."
    End If
    x = x + 1
    DoEvents
    strStore = Dir()
  Loop
  
  strFile = strFiles(LoadCursorPos)
  Dim iFile As Integer
  iFile = FreeFile
  Open App.Path & "\save\" & strFile For Binary Access Read As #iFile
    Get #iFile, , You
  Close #iFile
  bInHanger = True
  bShowMenu = False
  Scene = fEntrance
  RunTheGame
End Sub


Public Sub UpdatePulseShots()
  Dim x As Integer, Y As Integer
  Dim rMoney As Integer
  ShotTickCount = clsDDraw.TickCount
  For x = 0 To MaxShots
    If Pulse(x).CurY < 0 Then Pulse(x).active = False
    If (Pulse(x).active = True) Then
      If Pulse(x).active = True Then Pulse(x).CurY = Pulse(x).CurY - ShotVelocity
      If Pulse(x).CurY <= 0 Then Pulse(x).active = False
      
      If Pulse(x).active = True Then
        Draw DDS_PULSE, rPulse, Pulse(x).CurX, Pulse(x).CurY, True, False
      End If
      
      For Y = 0 To MaxEnemies
        If Pulse(x).active Then
          If BadGuys(Y).active = True And Pulse(x).active = True Then
            If CheckCollide(DDS_PULSE, BadGuys(Y).Surface, Pulse(x).CurX, Pulse(x).CurY, PulseWidth, PulseHeight, BadGuys(Y).x, BadGuys(Y).Y, BadGuys(Y).Width, BadGuys(Y).Height, 0) Then
              CreateHit Pulse(x).CurX, Pulse(x).CurY, BadGuys(Y).Velocity
'             PlayWav "hitp.wav"
                If BadGuys(Y).Shield > 0 Then
                  BadGuys(Y).Shield = BadGuys(Y).Shield - 800
                End If
                BadGuys(Y).Hull = BadGuys(Y).Hull - 600
                If BadGuys(Y).Shield <= 0 And BadGuys(Y).Hull <= 0 Then
                  BadGuys(Y).active = False
                  If BadGuys(Y).bBoss Then
                    bAtBoss = False
                    bEndLevel = True
                  End If
                  AddExplode BadGuys(Y).x, BadGuys(Y).Y
                  rMoney = Int(Rnd * 100)
                  If rMoney > 1 Then
                    AddMoney (BadGuys(Y).x + (BadGuys(Y).Width \ 2)), (BadGuys(Y).Y + (BadGuys(Y).Height \ 2))
                  End If
                  
                  You.CurMoney = You.CurMoney + BadGuys(x).Value
                End If
                Pulse(x).active = False
              End If
            End If
          End If
      Next Y
    Else
      If Firing = True And ((ShotTickCount - PulseShotTick) > PulseShotCount) Then
        CreatePulse x
      End If
        
        
    End If
      
  Next x
    
End Sub

Public Sub UpdatePlasmaShots()
  Dim x As Integer, Y As Integer
  Dim rMoney As Integer
  ShotTickCount = clsDDraw.TickCount
  For x = 0 To MaxShots
    If (Plasma(x).active = True) Then
      If Plasma(x).active = True Then Plasma(x).CurY = Plasma(x).CurY - ShotVelocity
      If Plasma(x).CurY <= 0 Then Plasma(x).active = False
      
      If Plasma(x).active = True Then
        Draw DDS_PLASMA, rPlasmaS, Plasma(x).CurX, Plasma(x).CurY, True, False
      End If
      
      For Y = 0 To MaxEnemies
        If Plasma(x).active Then
          If BadGuys(Y).active = True And Plasma(x).active = True Then
            If CheckCollide(DDS_PLASMA, BadGuys(Y).Surface, Plasma(x).CurX, Plasma(x).CurY, PlasmaWidth, PlasmaHeight, BadGuys(Y).x, BadGuys(Y).Y, BadGuys(Y).Width, BadGuys(Y).Height, 0) Then
              CreateHit Plasma(x).CurX, Plasma(x).CurY, BadGuys(Y).Velocity
'             PlayWav "hitp.wav"
                If BadGuys(Y).Shield > 0 Then
                  BadGuys(Y).Shield = BadGuys(Y).Shield - 5
                End If
                If BadGuys(Y).Shield < 0 Then
                  BadGuys(Y).Hull = BadGuys(Y).Hull - 10
                End If
                If BadGuys(Y).Shield <= 0 And BadGuys(Y).Hull <= 0 Then
                  BadGuys(Y).active = False
                  If BadGuys(Y).bBoss Then
                    bAtBoss = False
                    bEndLevel = True
                  End If
                  AddExplode BadGuys(Y).x, BadGuys(Y).Y
                  rMoney = Int(Rnd * 100)
                  If rMoney > 1 Then
                    AddMoney (BadGuys(Y).x + (BadGuys(Y).Width \ 2)), (BadGuys(Y).Y + (BadGuys(Y).Height \ 2))
                  End If
                  
                  You.CurMoney = You.CurMoney + BadGuys(x).Value
                End If
                Plasma(x).active = False
              End If
            End If
          End If
      Next Y
    Else
      If Firing = True And ((ShotTickCount - PlasmaShotTick) > PlasmaShotCount) Then
        CreatePlasma x
      End If
        
        
    End If
      
  Next x
    
End Sub

Public Sub ShowShop()
  
  Dim rFill As RECT
  Dim a As POINTAPI
  Dim strPrice As String
  With rFill
    .Left = 0
    .Right = ScreenWidth
    .Bottom = ScreenHeight
    .Top = 0
  End With
  DDS_Buffer.BltColorFill rFill, 0
  If BShopSell = False Then
    Draw DDS_SHOP, rShop, 0, (ScreenHeight - rShop.Bottom) \ 2, False, False
  Else
    Draw DDS_SHOPSELL, rShopSell, 0, (ScreenHeight - rShop.Bottom) \ 2, False, False
  End If
  GetCursorPos a
  Draw DDS_Cursor, rCursor, a.x - 10, a.Y - 10, True, True
  DDS_Buffer.SetForeColor vbYellow
  If bStartShop = True Then
    bStartShop = False
    Set DDS_ITEMDISPLAY = DDS_MICRO
    rShopItem = rMicro
    strShopStr1 = "This micro missile launcher is a "
    strShopStr2 = "great equalizer. It delivers micro"
    strShopStr3 = "missiles at an impressive rate. "
    lShopPrice = 18000
    iShopItem = 1
  End If
    
  strPrice = "$" & lShopPrice
  DDS_Buffer.DrawText (238 + ((348 - frmMain.TextWidth(strPrice)) \ 2)), 240, strPrice, False
  DDS_Buffer.SetForeColor vbWhite
  DDS_Buffer.DrawText 33, ((ScreenHeight - rShop.Bottom) \ 2) + 283, "$" & You.CurMoney, False
  Draw DDS_ITEMDISPLAY, rShopItem, (238 + ((348 - rShopItem.Right) \ 2)), ((ScreenHeight - rShop.Bottom) \ 2) + 80, True, False
  Draw DDS_ITEMDISPLAY, rShopItem, (30 + ((130 - rShopItem.Right) \ 2)), ((ScreenHeight - rShop.Bottom) \ 2) + 150, True, False
  DDS_Buffer.SetForeColor vbYellow
  DDS_Buffer.DrawText (25 + ((130 - frmMain.TextWidth(ShowNum)) \ 2)), 290, ShowNum, False
  DDS_Buffer.SetForeColor vbWhite
  DDS_Buffer.DrawText 245, ((ScreenHeight - rShop.Bottom) \ 2) + 250, strShopStr1, False
  DDS_Buffer.DrawText 245, ((ScreenHeight - rShop.Bottom) \ 2) + 260, strShopStr2, False
  DDS_Buffer.DrawText 245, ((ScreenHeight - rShop.Bottom) \ 2) + 270, strShopStr3, False
  DDS_Buffer.DrawText (ScreenWidth - (frmMain.TextWidth(strShopText))) \ 2, ScreenHeight - 15, strShopText, False
  
  
End Sub

Public Function ShowMsg(txt As String)
  Dim iCnt As Integer, tmp As String
  Dim strs As String, a As String, i As Long
  Dim d As POINTAPI
  iCnt = 0
  Dim store(50) As String
  bGamePause = True
  For i = 1 To Len(txt)
    a = Mid(txt, i, 1)
    
    If a <> " " Then
      tmp = tmp & a
    End If
    'tmp = tmp + a
    
    If a = " " Then
      If frmMain.TextWidth(strs & tmp) < 180 And i < Len(txt) Then
        strs = strs & tmp & " "
        tmp = ""
      Else
        iCnt = iCnt + 1
        store(iCnt) = strs
        strs = tmp & " "
        tmp = ""
      End If
    End If
    If i = Len(txt) And (frmMain.TextWidth(strs & tmp) > 180) Then
      iCnt = iCnt + 1
      store(iCnt) = strs
      strs = tmp
      store(iCnt) = strs
    End If
    
    If i = Len(txt) And (frmMain.TextWidth(strs & tmp) < 180) Then
        iCnt = iCnt + 1
        strs = strs & tmp
        store(iCnt) = strs
        
    End If
    
    
  Next i

  Draw DDS_MESSAGE, rMessage, (ScreenWidth - (rMessage.Right)) \ 2, (ScreenHeight - (rMessage.Bottom)) \ 2, False, True
  For i = 1 To iCnt
    DDS_Buffer.DrawText ((ScreenWidth - (rMessage.Right)) \ 2) + 30, ((ScreenHeight - (rMessage.Bottom)) \ 2) + 40 + ((frmMain.TextHeight(store(i)) + 4) * (i - 1)), store(i), False
  Next
  GetCursorPos d
  Draw DDS_Cursor, rCursor, d.x - 10, d.Y - 10, True, True
  CheckInput
  DoEvents
End Function

Public Function ShowNum() As String
  Select Case iShopItem
    Case 1
      ShowNum = Str(You.lMicro)
    Case 2
      ShowNum = Str(You.lPlasma)
    Case 3
      ShowNum = Str(You.ReactorPower)
    Case 4
      ShowNum = Str(You.lPulse)
    Case 5
      ShowNum = Str(You.ShieldR)
  End Select
End Function

Public Sub SetItemDisplayUp()
  iShopItem = iShopItem + 1
  If iShopItem > 5 Then iShopItem = 1
  Select Case iShopItem
    Case 2
      Set DDS_ITEMDISPLAY = DDS_PLASMAB
      rShopItem = rPlasma
      If BShopSell = False Then
        lShopPrice = 30000
      Else
        lShopPrice = 17000
      End If
      strShopStr1 = "The plasma gun is an excellent additional"
      strShopStr2 = "Weapon Capable of shooting out rapid"
      strShopStr3 = "low yield plasma charges."
      
    Case 3
      Set DDS_ITEMDISPLAY = DDS_REACTOR
      rShopItem = rReactor
      strShopStr1 = "This is a reactor upgrade.  Each reactor"
      strShopStr2 = "upgrade you buy will make your reactor"
      strShopStr3 = "generate 2 more units of power per tick."
      If BShopSell = False Then
        lShopPrice = 32000
      Else
        lShopPrice = 18500
      End If
      
    Case 1
      Set DDS_ITEMDISPLAY = DDS_MICRO
      rShopItem = rMicro
      strShopStr1 = "This micro missile launcher is a "
      strShopStr2 = "great equalizer. It delivers micro"
      strShopStr3 = "missiles at an impressive rate. "
      If BShopSell = False Then
        lShopPrice = 18000
      Else
        lShopPrice = 3500
      End If
      
    Case 4
      Set DDS_ITEMDISPLAY = DDS_PULSEC
      rShopItem = rPulseC
      strShopStr1 = "The pulse cannon is the most feared weapon"
      strShopStr2 = "of all.It delivers semi rapid high energy"
      strShopStr3 = "pulse beams inflicting massive damage."
      If BShopSell = False Then
        lShopPrice = 95000
      Else
        lShopPrice = 63000
      End If
      
    Case 5
      Set DDS_ITEMDISPLAY = DDS_SHIELD
      rShopItem = rShield
      strShopStr1 = "The shield upgrade will upgrade how much"
      strShopStr2 = "your shields regenerate.  The shields draw"
      strShopStr3 = "whatever they use off the reactor."
      If BShopSell = False Then
        lShopPrice = 22000
      Else
        lShopPrice = 2500
      End If
  End Select
End Sub

Public Sub SetItemDisplay()
  iShopItem = iShopItem - 1
  If iShopItem < 1 Then iShopItem = 5
  Select Case iShopItem
    Case 2
      Set DDS_ITEMDISPLAY = DDS_PLASMAB
      rShopItem = rPlasma
      If BShopSell = False Then
        lShopPrice = 30000
      Else
        lShopPrice = 17000
      End If
      strShopStr1 = "The plasma gun is an excellent additional"
      strShopStr2 = "Weapon Capable of shooting out rapid"
      strShopStr3 = "low yield plasma charges."
      
    Case 3
      Set DDS_ITEMDISPLAY = DDS_REACTOR
      rShopItem = rReactor
      strShopStr1 = "This is a reactor upgrade.  Each reactor"
      strShopStr2 = "upgrade you buy will make your reactor"
      strShopStr3 = "generate 2 more units of power per tick."
      If BShopSell = False Then
        lShopPrice = 32000
      Else
        lShopPrice = 18500
      End If
      
    Case 1
      Set DDS_ITEMDISPLAY = DDS_MICRO
      rShopItem = rMicro
      strShopStr1 = "This micro missile launcher is a "
      strShopStr2 = "great equalizer. It delivers micro"
      strShopStr3 = "missiles at an impressive rate. "
      If BShopSell = False Then
        lShopPrice = 18000
      Else
        lShopPrice = 3500
      End If
      
    Case 4
      Set DDS_ITEMDISPLAY = DDS_PULSEC
      rShopItem = rPulseC
      strShopStr1 = "The pulse cannon is the most feared weapon"
      strShopStr2 = "of all.It delivers semi rapid high energy"
      strShopStr3 = "pulse beams inflicting massive damage."
      If BShopSell = False Then
        lShopPrice = 95000
      Else
        lShopPrice = 63000
      End If
      
    Case 5
      Set DDS_ITEMDISPLAY = DDS_SHIELD
      rShopItem = rShield
      strShopStr1 = "The shield upgrade will upgrade how much"
      strShopStr2 = "your shields regenerate.  The shields draw"
      strShopStr3 = "whatever they use off the reactor."
      If BShopSell = False Then
        lShopPrice = 22000
      Else
        lShopPrice = 2500
      End If
  End Select
End Sub


Public Sub LoadCharProfile()
  Dim strStore As String
  Dim rBack As RECT
  Dim i As Integer, x As Long
  Dim strFiles(1000) As String
  With rBack
    .Left = 0
    .Right = ScreenWidth
    .Top = 0
    .Bottom = ScreenHeight
  End With
  Draw DDS_MenuBack, rBack1, 0, 0, False, False
  Draw DDS_Save, rSave, (ScreenWidth - rSave.Right) \ 2, (ScreenHeight - rSave.Bottom) \ 2, False, False
  strStore = Dir(App.Path & "\save\")
  Dim b As Integer
  x = 0
  
  Do Until LenB(strStore) = 0
    If Len(strStore) < 35 Then
      strFiles(x) = strStore
    Else
      strFiles(x) = Left(strStore, 32) & "..."
    End If
    x = x + 1
    DoEvents
    strStore = Dir()
  Loop
  lLoadFileCount = x
  If x < 20 Then
    For i = 0 To x
      If strFiles(i) <> "" Then
        If i = LoadCursorPos Then
          DDS_Buffer.SetForeColor vbYellow
        Else
          DDS_Buffer.SetForeColor vbWhite
        End If
        DDS_Buffer.DrawText 210, (80 + (frmMain.TextHeight(strFiles(i)) * i) + 4), strFiles(i), False
      End If
      DoEvents
    Next
  Else
    If LoadCursorPos > 9 And ((x - 9) > LoadCursorPos) Then
      b = -1
      For i = LoadCursorPos - 9 To LoadCursorPos + 9
        b = b + 1
        If strFiles(i) <> "" Then
          If i = LoadCursorPos Then
            DDS_Buffer.SetForeColor vbYellow
          Else
            DDS_Buffer.SetForeColor vbWhite
          End If
          DDS_Buffer.DrawText 210, (80 + (frmMain.TextHeight(strFiles(i)) * (b)) + 4), strFiles(i), False
        End If
        DoEvents
      Next
    ElseIf LoadCursorPos < 10 Then
      b = -1
      For i = 0 To 18
        b = b + 1
        If strFiles(i) <> "" Then
          If i = LoadCursorPos Then
            DDS_Buffer.SetForeColor vbYellow
          Else
            DDS_Buffer.SetForeColor vbWhite
          End If
          DDS_Buffer.DrawText 210, (80 + (frmMain.TextHeight(strFiles(i)) * (b)) + 4), strFiles(i), False
        End If
        DoEvents
      Next
    ElseIf LoadCursorPos > 9 And ((x - 10) < LoadCursorPos) Then
      b = -1
      For i = LoadCursorPos - 9 To x
        b = b + 1
        If strFiles(i) <> "" Then
          If i = LoadCursorPos Then
            DDS_Buffer.SetForeColor vbYellow
          Else
            DDS_Buffer.SetForeColor vbWhite
          End If
          DDS_Buffer.DrawText 210, (80 + (frmMain.TextHeight(strFiles(i)) * (b)) + 4), strFiles(i), False
        End If
        DoEvents
      Next
    End If
  End If
    
  DoEvents
End Sub

Public Sub ClearAllEnemies()
  Dim x As Long
  For x = 0 To MaxEnemies
    BadGuys(x).active = False
  Next
  For x = 0 To MaxShots
    PShotL(x).active = False
    PShotR(x).active = False
    MissileL(x).active = False
    MissileR(x).active = False
    Pulse(x).active = False
    Plasma(x).active = False
    EShotL(x).active = False
    EShotR(x).active = False
  Next
  For x = 0 To MaxExplode
    Explodes(x).active = False
  Next
End Sub

Public Sub SetPixelsUp(bUp As Boolean)
  If bUp = True Then
    Select Case lColorDepth
      Case 8
        lColorDepth = 16
      Case 16
        lColorDepth = 8
    End Select
  Else
    Select Case lColorDepth
      Case 8
        lColorDepth = 16
      Case 16
        lColorDepth = 8
    End Select
  End If
End Sub
