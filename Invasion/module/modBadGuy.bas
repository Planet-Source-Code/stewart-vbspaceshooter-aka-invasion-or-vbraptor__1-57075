Attribute VB_Name = "modBadGuy"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modBadGuy.bas                                         |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+
Public EnemySurface(6) As DirectDrawSurface7
Public EnemyHeight(6) As Long
Public EnemyWidth(6) As Long
Public EnemyHull(6) As Long
Public EnemyShield(6) As Long
Public EnemyShotL(6) As Long
Public EnemyShotR(6) As Long
Public EnemyShotY(6) As Long
Public EnemyVelocity(6) As Integer
Public EnemyRect(6) As RECT
Public EnemyFramesX(6) As Integer
Public EnemyFramesY(6) As Integer
Public EnemyValue(6) As Integer
Public EnemyShoot(6) As Boolean

Private LastEShotTick
Private LastEMisTick

Public Const MaxEShots = 20

Public EShotL(0 To MaxEShots) As ShotDataP
Public EShotR(0 To MaxEShots) As ShotDataP
Public EMisL(0 To MaxEShots) As ShotDataP
Public EMisR(0 To MaxEShots) As ShotDataP

Public Sub CreateEnemy(iship As Integer, ShipX As Integer, ShipAI As Integer, Ship As Integer)
  Dim BadGuyNum As Integer

  BadGuys(iship).AI = ShipAI
  BadGuys(iship).Hull = EnemyHull(Ship)
  Set BadGuys(iship).Surface = EnemySurface(Ship)
  BadGuys(iship).ShotL = EnemyShotL(Ship)
  BadGuys(iship).Velocity = EnemyVelocity(Ship)
  BadGuys(iship).ShotR = EnemyShotR(Ship)
  BadGuys(iship).ShotY = EnemyShotY(Ship)
  BadGuys(iship).RECT = EnemyRect(Ship)
  BadGuys(iship).Shield = EnemyShield(Ship)
  BadGuys(iship).FramesX = EnemyFramesX(Ship)
  BadGuys(iship).FramesY = EnemyFramesY(Ship)
  BadGuys(iship).Width = EnemyWidth(Ship)
  BadGuys(iship).Value = EnemyValue(Ship)
  BadGuys(iship).bBoss = False
  BadGuys(iship).Height = EnemyHeight(Ship)
  BadGuys(iship).Y = (0 - BadGuys(Ship).Height)
  BadGuys(iship).CanShoot = EnemyShoot(Ship)
  BadGuys(iship).x = ShipX
  BadGuys(iship).active = True
End Sub

Public Sub InitEnemyShips()
  InitE1
  InitE2
  InitE3
  InitE4
  InitE5
  InitE6
End Sub

Public Sub CreateEShot(ShotI As Long, EnemyX As Integer)
  Dim i As Integer
  If BadGuys(EnemyX).CanShoot = False Then Exit Sub
  If (clsDDraw.TickCount - LastEShotTick < ShotTick) Then Exit Sub
  EShotL(ShotI).active = True
  EShotR(ShotI).active = True
  EShotL(ShotI).CurX = BadGuys(EnemyX).x + BadGuys(EnemyX).ShotL
  EShotR(ShotI).CurX = BadGuys(EnemyX).x + BadGuys(EnemyX).ShotR
  EShotL(ShotI).CurY = (BadGuys(EnemyX).Y + 1 + BadGuys(EnemyX).ShotY)
  EShotR(ShotI).CurY = (BadGuys(EnemyX).Y + 1 + BadGuys(EnemyX).ShotY)
  LastEShotTick = clsDDraw.TickCount
End Sub

Public Sub CreateEMissile(ShotI As Long, EnemyX As Integer)
  Dim i As Integer
  
  If BadGuys(EnemyX).CanShoot = False Then Exit Sub
  If (clsDDraw.TickCount - LastEMisTick < MissileShotCount) Then Exit Sub
  EMisL(ShotI).active = True
  EMisR(ShotI).active = True
  EMisL(ShotI).CurX = BadGuys(EnemyX).x + Level.BossMissile1X
  EMisR(ShotI).CurX = BadGuys(EnemyX).x + Level.BossMissile2X
  EMisL(ShotI).CurY = (BadGuys(EnemyX).Y + 1 + BadGuys(EnemyX).ShotY)
  EMisR(ShotI).CurY = (BadGuys(EnemyX).Y + 1 + BadGuys(EnemyX).ShotY)
  LastEMisTick = clsDDraw.TickCount
End Sub


Public Function FindEShot() As Integer
  Dim x As Integer
  For x = 0 To MaxEShots
    If EShotR(x).active = False And EShotL(x).active = False Then
      FindEShot = x
      Exit Function
    End If
  Next
End Function

Public Function FindEMissile() As Integer
  Dim x As Integer
  For x = 0 To MaxEShots
    If EMisR(x).active = False And EMisL(x).active = False Then
      FindEMissile = x
      Exit Function
    End If
  Next
End Function

