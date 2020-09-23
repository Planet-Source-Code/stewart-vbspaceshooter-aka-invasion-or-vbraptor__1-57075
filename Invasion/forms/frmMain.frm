VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Alien Invasion"
   ClientHeight    =   8055
   ClientLeft      =   675
   ClientTop       =   1020
   ClientWidth     =   9075
   BeginProperty Font 
      Name            =   "Fixedsys"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmMain.frx":0CCA
   MousePointer    =   99  'Custom
   ScaleHeight     =   537
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - frmMain.frm                                           |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  If Scene = fEntrance Then
'    If KeyCode = vbKeyEscape Then
'      WhatToDo = 0
'      InGame = False
'      StartMenu
'    End If
  End If
  If Scene = fLoad Then
    If KeyCode = vbKeyEscape Then
      StartMenu
    End If
  End If
  If Scene = fGame And bGamePause = False Then
    If KeyCode = vbKeyEscape Then
      iInMenuSel = 1
      bGamePause = True
      Exit Sub
    End If
  End If
  If Scene = fPlayer Then
    If KeyCode = vbKeyEscape Then
      StartMenu
    End If
  End If
  If Scene = fGame And bGamePause = True Then
    
    If KeyCode = vbKeyUp Then
      iInMenuSel = iInMenuSel - 1
      If iInMenuSel < 1 Then iInMenuSel = 2
    End If
    If KeyCode = vbKeyDown Then
      iInMenuSel = iInMenuSel + 1
      If iInMenuSel > 2 Then iInMenuSel = 1
    End If
    If KeyCode = vbKeyEscape Then
      bGamePause = False
    End If
    If KeyCode = vbKeyReturn Then
      If iInMenuSel = 1 Then
        InGame = False
        bInHanger = True
        Scene = fEntrance
        StopBGSound
        bEndLevel = False
      Else
        bGamePause = False
      End If
    End If
  End If
  
  If Scene = fOptions Then
    If KeyCode = vbKeyEscape Then
      writeini "Settings", "MusicVolume", Str(MusVol), App.Path & "\data\game.ini"
      writeini "Settings", "SoundVolume", Str(SndVol), App.Path & "\data\game.ini"
      writeini "Settings", "Difficulty", Str(Difficulty), App.Path & "\data\game.ini"
      writeini "Settings", "Collision", Str(PixelPerfect), App.Path & "\data\game.ini"
      writeini "Settings", "ShowFPS", Str(bShowFPS), App.Path & "\data\game.ini"
      writeini "Settings", "ColorDepth", Str(lColorDepth), App.Path & "\data\game.ini"
      clsDDraw.RestoreAllSurfaces
      'clsDDraw.Init frmMain.hWnd, ScreenWidth, ScreenHeight, lColorDepth, DDS_Primary, DDS_Buffer, DDSD_Buffer
      InOptions = False
      
      
      Dim iTmp As Double
      iTmp = (MusVol / 100)
      iTmp = cVol.VolumeMax * iTmp
      cVol.VolumeLevel = iTmp
      iTmp = (SndVol / 100)
      iTmp = cVol.WaveMax * iTmp
      cVol.WaveLevel = iTmp
      StartMenu
    End If
    If KeyCode = vbKeyUp Then
      OptSel = OptSel - 1
      DDS_Click.Play DSBPLAY_DEFAULT
      If OptSel < 0 Then OptSel = 5
    End If
    If KeyCode = vbKeyDown Then
      DDS_Click.Play DSBPLAY_DEFAULT
      OptSel = OptSel + 1
      If OptSel > 5 Then OptSel = 0
    End If
    If KeyCode = vbKeyRight Then
      If OptSel = 0 Then
        If Difficulty < 4 Then Difficulty = Difficulty + 1
      End If
      If OptSel = 1 Then
        If MusVol < 100 Then MusVol = MusVol + 1
      End If
      If OptSel = 2 Then
        If SndVol < 100 Then SndVol = SndVol + 1
      End If
      If OptSel = 3 Then PixelPerfect = Not PixelPerfect
      If OptSel = 4 Then bShowFPS = Not bShowFPS
      If OptSel = 5 Then SetPixelsUp True
    End If
    If KeyCode = vbKeyLeft Then
      If OptSel = 0 Then
        If Difficulty > 0 Then Difficulty = Difficulty - 1
      End If
      If OptSel = 1 Then
        If MusVol > 0 Then MusVol = MusVol - 1
      End If
      If OptSel = 2 Then
        If SndVol > 0 Then SndVol = SndVol - 1
      End If
      If OptSel = 3 Then PixelPerfect = Not PixelPerfect
      If OptSel = 4 Then bShowFPS = Not bShowFPS
      If OptSel = 5 Then SetPixelsUp False
      
    End If
    'Exit Sub
    
  End If
  If Scene = fLoad Then
    If KeyCode = vbKeyDown Then
      If LoadCursorPos < lLoadFileCount - 1 Then
        LoadCursorPos = LoadCursorPos + 1
      End If
    End If
    If KeyCode = vbKeyUp Then
      If LoadCursorPos > 0 Then
        LoadCursorPos = LoadCursorPos - 1
      End If
    End If
    
    If KeyCode = vbKeyReturn Then
      LoadProfile
    End If
  End If
  If Scene = fShop Then
    If KeyCode = vbKeyEscape Then
      Scene = fEntrance
      bInHanger = True
    End If
  End If
  If Scene = fCreate Then
    If KeyCode = vbKeyEscape Then
      StartMenu
    End If
  End If
  If Scene = fMenu Then
    If KeyCode = vbKeyDown Then
      SelButton = SelButton + 1
      If SelButton > 4 Then SelButton = 1
      DDS_Click.Play DSBPLAY_DEFAULT
    End If
    If KeyCode = vbKeyUp Then
      SelButton = SelButton - 1
      If SelButton < 1 Then SelButton = 4
      DDS_Click.Play DSBPLAY_DEFAULT
    End If
    If KeyCode = vbKeyReturn Then
      If SelButton = 4 Then
        WhatToDo = 3
        InMenu = False
        CloseGame
        
      ElseIf SelButton = 3 Then
        'InOptions = True
        'InMenu = False
        'WhatToDo = 2
        XLoadPos = 0
        Scene = fOptions
        'Options
      ElseIf SelButton = 2 Then
        OptSel = 0
        LoadCursorPos = 0
        Scene = fLoad
      ElseIf SelButton = 1 Then
        InMenu = False
        WhatToDo = 1
        'NewGame
      End If
    End If
    Exit Sub
  End If
  
End Sub

Private Sub Form_KeyPress(keyascii As Integer)
  If Scene = fPlayer Then
  If bShowMsg = True Then
    If keyascii = 13 Then
      bShowMsg = False
    End If
    Exit Sub
  End If
  
    If bAtCallsign = False Then
      If keyascii = 8 Then
        If Len(strUserName) < 1 Then Exit Sub
        strUserName = Left(strUserName, Len(strUserName) - 1)
        If strUserName = "" Then strUserName = ""
      End If
      If keyascii = 13 Then
        If strUserName <> "" Then
          bAtCallsign = True
        Else
          strMsgText = "You must enter a username to continue."
          bShowMsg = True
        End If
      End If
      If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(keyascii))) <> 0 Then
        strUserName = strUserName + Chr(keyascii)
      End If
    Else
      If keyascii = 8 Then
        If Len(strCallSign) < 1 Then Exit Sub
        strCallSign = Left(strCallSign, Len(strCallSign) - 1)
        If strCallSign = "" Then strCallSign = ""
      End If
      If InStr(1, "ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(keyascii))) <> 0 Then
        strCallSign = strCallSign + Chr(keyascii)
      End If
      If keyascii = 13 Then
        If strCallSign <> "" Then
          You.strCallSign = strCallSign
          You.strUserName = strUserName
          Scene = fEntrance
        Else
          strMsgText = "You must enter a callsign to continue."
          bShowMsg = True
        End If

        'StartNewGame
      End If
    End If
    
  Else
    If bShowMsg = True Then
      If keyascii = 13 Then
        bShowMsg = False
      End If
      Exit Sub
    End If

  End If
  
  
End Sub

Private Sub Form_Load()
  Start
  StartMenu
  RunTheGame
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If ((clsDDraw.TickCount - LastClick) < ClickTime) Then
    LastClick = clsDDraw.TickCount
    Exit Sub
  End If
  If bShowMsg = True Then Exit Sub
  If Scene = fEntrance Then
    Select Case iHangerDo
      Case 0
        SaveCharProfile
      Case 1
        bShowingHanger = False
        bShowMenu = True
        
        StartMenu
        'CloseGame
      Case 2
        iShopItem = 0
        bStartShop = True
        bInHanger = False
        BShopSell = False
        Scene = fShop
        RunTheGame
      Case 3
        bShowingHanger = False
        bGamePause = False
        StartNewGame
    End Select
  End If
  If Scene = fShop And bInHanger = False Then
    Select Case iShopDo
      Case 0
        SetItemDisplay
      Case 1
        SetItemDisplayUp
      Case 2
        If iShopItem = 1 Then
          If BShopSell = False Then
            If You.CurMoney < lShopPrice Then
              strMsgText = "You need $" & lShopPrice & " to buy item (Micro Missiles)"
              bShowMsg = True
            Else
              You.lMicro = You.lMicro + 1
              You.CurMoney = You.CurMoney - lShopPrice
            End If
          Else
            If You.lMicro < 1 Then
              strMsgText = "You don't have any Micro missiles."
              bShowMsg = True
            Else
              You.lMicro = You.lMicro - 1
              You.CurMoney = You.CurMoney + lShopPrice
            End If
          End If
         
        ElseIf iShopItem = 2 Then
          If BShopSell = False Then
            If You.CurMoney < lShopPrice Then
              strMsgText = "You need $" & lShopPrice & " to buy item (Plasma Gun)"
              bShowMsg = True
            Else
              You.lPlasma = You.lPlasma + 1
              You.CurMoney = You.CurMoney - lShopPrice
            End If
          Else
            If You.lPlasma < 1 Then
              strMsgText = "You don't have any Plasma Guns."
              bShowMsg = True
            Else
              You.lPlasma = You.lPlasma - 1
              You.CurMoney = You.CurMoney + lShopPrice
            End If
          End If
        ElseIf iShopItem = 3 Then
          If BShopSell = False Then
            If You.CurMoney < lShopPrice Then
              strMsgText = "You need $" & lShopPrice & " to buy item (Reactor Upgrade)"
              bShowMsg = True
            Else
              You.ReactorPower = You.ReactorPower + 2
              You.CurMoney = You.CurMoney - lShopPrice
            End If
          Else
            If You.ReactorPower < 4 Then
              strMsgText = "You don't have any reactor upgrades."
              bShowMsg = True
            Else
              You.ReactorPower = You.ReactorPower - 2
              You.CurMoney = You.CurMoney + lShopPrice
            End If
          End If
              
        ElseIf iShopItem = 4 Then
          If BShopSell = False Then
            If You.CurMoney < lShopPrice Then
              strMsgText = "You need $" & lShopPrice & " to buy item (Pulse Cannon)"
              bShowMsg = True
            Else
              You.lPulse = You.lPulse + 1
              You.CurMoney = You.CurMoney - lShopPrice
            End If
          Else
            If You.lPulse < 1 Then
              strMsgText = "You don't have any Pulse Cannons."
              bShowMsg = True
            Else
              You.lPulse = You.lPulse - 1
              You.CurMoney = You.CurMoney + lShopPrice
            End If
          End If
          
        ElseIf iShopItem = 5 Then
          If BShopSell = False Then
            If You.CurMoney < lShopPrice Then
              strMsgText = "You need $" & lShopPrice & " to buy item (Shield Upgrade)"
              bShowMsg = True
            Else
              You.ShieldR = You.ShieldR + 2
              You.CurMoney = You.CurMoney - lShopPrice
            End If
          Else
            If You.ShieldR < 4 Then
              strMsgText = "You don't have any shield upgrades."
              bShowMsg = True
            Else
              You.ShieldR = You.ShieldR - 2
              You.CurMoney = You.CurMoney + lShopPrice
            End If
          End If
        End If
        
      Case 3
        bShowingHanger = True
        InGame = False
        bInHanger = True
        Scene = fEntrance
        StopBGSound
        bStartShop = False
        bEndLevel = False
      Case 4
        BShopSell = True
      Case 5
        BShopSell = False
    End Select
  End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If bShowMsg = True Then Exit Sub
  If InMenu Then HighlightBut Int(x), Int(Y)
  If bShowingHanger = True Then
    iHangerDo = 5
    If (x > 300 And x < 620) And (Y > 200 And Y < 400) Then
      strHangerText = "Save Game"
      iHangerDo = 0
    
    ElseIf (x > 104 And x < 166) And (Y > 336 And Y < 385) Then
      strHangerText = "Exit Character"
      iHangerDo = 1
    ElseIf (x > 334 And x < 636) And (Y > 50 And Y < 169) Then
      strHangerText = "Buy Parts"
      iHangerDo = 2
    ElseIf (x > 3 And x < 331) And (Y > 5 And Y < 132) Then
      strHangerText = "Launch"
      iHangerDo = 3
    Else
      strHangerText = " "
      iHangerDo = 4
    End If
  End If
  If Scene = fShop Then
    iShopDo = 10
    strShopText = ""
    If (x > 492) And (Y > 366) And (x < 548) And (Y < 390) Then
      strShopText = "Previous Item"
      iShopDo = 0
    ElseIf (x > 548) And (Y > 366) And (x < 606) And (Y < 390) Then
      strShopText = "Next Item"
      iShopDo = 1
    ElseIf (x > 352) And (Y > 347) And (x < 456) And (Y < 409) Then
      If BShopSell = True Then
        strShopText = "Sell Item"
      Else
        strShopText = "Buy Item"
      End If
      iShopDo = 2
    ElseIf (x > 0) And (Y > 145) And (x < 201) And (Y < 450) Then
      strShopText = "Exit Shop"
      iShopDo = 3
    ElseIf (x > 285) And (Y > 368) And (x < 320) And (Y < 386) Then
      strShopText = "Sell Items"
      iShopDo = 4
    ElseIf (x > 240) And (Y > 368) And (x < 275) And (Y < 386) Then
      strShopText = "Buy Items"
      iShopDo = 5
    End If
    
  End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
  If bShowMsg = True Then Exit Sub
  If InMenu = True Then ExecuteClick Int(x), Int(Y)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ShowCursor 1
End Sub

Private Sub Form_Terminate()
  ShowCursor 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ShowCursor 1
End Sub

