Attribute VB_Name = "modLvl"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modLvl.bas                                            |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+

Public Type Lines
  Ship(6) As Integer
  AI(6) As Integer
End Type

Public LineNum As Integer

Public lList() As Lines

Public Sub LoadLevel(lvl As String)
  Dim tmpStr As String, tmpStore As String, x As Integer, Y As Integer
  Dim iFile As Integer, l As Integer, lTotal As Long, z As Long
  Dim perLoad As Double
  Dim rScreen As RECT
  DeArchive App.Path & "\levels\" & lvl
  BShipDestroyed = False
  iFile = FreeFile
  Open App.Path & "\temp\tmpEnemy.enm" For Input As #iFile
  Y = -1
  lTotal = LOF(iFile)
  Input #iFile, l
  Erase lList
  ReDim lList(l)
  MaxRows = l
  With rScreen
    .Left = 0
    .Right = ScreenWidth
    .Top = 0
    .Bottom = ScreenHeight
  End With
       
  DDS_Buffer.BltColorFill rScreen, 0
  DDS_Buffer.DrawText 0, 0, "Loading...", False
  
  
  
  Input #iFile, Level.TileBitmap
  Input #iFile, Level.MusicFile
  Input #iFile, Level.EnemyFile
  Input #iFile, tmpStr
  Input #iFile, tmpStr
  Input #iFile, tmpStr
  Input #iFile, Level.BossShield
  Input #iFile, Level.BossHull
  Input #iFile, Level.BossLaserDamage
  Input #iFile, Level.BossMissileDamage
  Input #iFile, Level.BossLaser1X
  Input #iFile, Level.BossLaser2X
  Input #iFile, Level.BossMissile1X
  Input #iFile, Level.BossMissile2X
  For z = 0 To MaxRows
    Input #iFile, tmpStr
      Y = Y + 1
      For x = 1 To 12 Step 2
        tmpStore = Mid(tmpStr, x, 2)
        lList(Y).Ship(x \ 2) = Left(tmpStore, 1)
        lList(Y).AI(x \ 2) = Right(tmpStore, 1) - 1
      Next
  Next z
  Close #iFile
  LineNum = 190
  clsDDraw.DDCreateSurface DDS_BOSS, App.Path & "\temp\tmpBoss.bmp", rBoss
End Sub

Public Sub CreateBoss()
  Dim iship As Integer
  iship = 0
  BadGuys(iship).AI = 3
  BadGuys(iship).bBoss = True
  BadGuys(iship).Hull = Level.BossHull
  Set BadGuys(iship).Surface = DDS_BOSS
  BadGuys(iship).ShotL = Level.BossLaser1X
  'BadGuys(iShip).Velocity = Level.BossLaser2X
  BadGuys(iship).ShotR = Level.BossLaser2X
  'BadGuys(iShip).ShotY =
  BadGuys(iship).RECT = rBoss ' EnemyRect(Ship)
  BadGuys(iship).Shield = Level.BossShield 'EnemyShield(Ship)
  'BadGuys(iShip).FramesX = EnemyFramesX(Ship)
  'BadGuys(iShip).FramesY = EnemyFramesY(Ship)
  BadGuys(iship).Width = rBoss.Right 'EnemyWidth(Ship)
  BadGuys(iship).Value = 10000 'EnemyValue(Ship)
  BadGuys(iship).Height = rBoss.Bottom 'EnemyHeight(Ship)
  BadGuys(iship).Y = (0 - BadGuys(iship).Height)
  BadGuys(iship).x = (ScreenWidth - BadGuys(iship).Width) \ 2
  BadGuys(iship).active = True
  
End Sub

Private Function FindEnemy() As Integer
  Dim x As Integer
  For x = 0 To MaxEnemies
    If BadGuys(x).active = False Then
      FindEnemy = x
      Exit Function
    End If
  Next
End Function

Public Sub ParseLevel()
  On Error Resume Next
  If ((clsDDraw.TickCount - LastGuyTick) > NewBadGuyTick) Then
    If bAtBoss = True Then Exit Sub

    If bEndLevel Then Exit Sub
    
    LineNum = LineNum + 1
    EnemyDestroyed = clsDDraw.TickCount
    If lList(LineNum).Ship(0) > 0 Then CreateEnemy FindEnemy, 0, lList(LineNum).AI(0), lList(LineNum).Ship(0)
    If lList(LineNum).Ship(1) > 0 Then CreateEnemy FindEnemy, 0 + (EnemyWidth(lList(LineNum).Ship(1)) * 2), lList(LineNum).AI(1), lList(LineNum).Ship(1)
    If lList(LineNum).Ship(2) > 0 Then CreateEnemy FindEnemy, ((ScreenWidth - EnemyWidth(lList(LineNum).Ship(2))) \ 2) - EnemyWidth(lList(LineNum).Ship(2)), lList(LineNum).AI(2), lList(LineNum).Ship(2)
    If lList(LineNum).Ship(3) > 0 Then CreateEnemy FindEnemy, ((ScreenWidth - EnemyWidth(lList(LineNum).Ship(3))) \ 2) + EnemyWidth(lList(LineNum).Ship(3)), lList(LineNum).AI(3), lList(LineNum).Ship(3)
    If lList(LineNum).Ship(4) > 0 Then CreateEnemy FindEnemy, ScreenWidth - (EnemyWidth(lList(LineNum).Ship(4)) * 3), lList(LineNum).AI(4), lList(LineNum).Ship(4)
    If lList(LineNum).Ship(5) > 0 Then CreateEnemy FindEnemy, ScreenWidth - EnemyWidth(lList(LineNum).Ship(5)), lList(LineNum).AI(5), lList(LineNum).Ship(5)
    If LineNum = MaxRows Then
      If bCanDoBoss Then bAtBoss = True
      bAtBoss = True 'LineNum = -1
      BBMoveDir = 0
    End If
    LastGuyTick = clsDDraw.TickCount
  End If
End Sub
