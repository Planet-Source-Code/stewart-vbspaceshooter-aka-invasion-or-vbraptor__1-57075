Attribute VB_Name = "modInitBadGuys"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modInitBadGuys.bas                                    |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+

Public Sub InitE1()
  EnemyHeight(1) = 33
  EnemyWidth(1) = 60
  EnemyShield(1) = 85
  EnemyShotL(1) = 27
  EnemyShotR(1) = 58
  EnemyFramesX(1) = 0
  EnemyFramesY(1) = 0
  EnemyShotY(1) = 37
  EnemyValue(1) = 300
  EnemyHull(1) = 85
  EnemyVelocity(1) = 2
  EnemyShoot(1) = True
  clsDDraw.DDCreateSurface EnemySurface(1), App.Path & "\images\e1.bmp", EnemyRect(1), EnemyWidth(1), EnemyHeight(1), 0
End Sub

Public Sub InitE2()
  EnemyHeight(2) = 51
  EnemyWidth(2) = 50
  EnemyShield(2) = 125
  EnemyHull(2) = 200
  EnemyShotL(2) = 12
  EnemyShotR(2) = 62
  EnemyFramesX(2) = 0
  EnemyFramesY(2) = 0
  EnemyShotY(2) = 62
  EnemyValue(2) = 600
  EnemyVelocity(2) = 2
  EnemyShoot(2) = True
  clsDDraw.DDCreateSurface EnemySurface(2), App.Path & "\images\e2.bmp", EnemyRect(2), EnemyWidth(2), EnemyHeight(2), 0
End Sub

Public Sub InitE3()
  EnemyHeight(3) = 40
  EnemyWidth(3) = 30
  EnemyShield(3) = 0
  EnemyHull(3) = 60
  EnemyVelocity(3) = 5
  EnemyFramesX(3) = 5
  EnemyValue(3) = 50
  EnemyFramesY(3) = 2
  EnemyShoot(3) = False
  clsDDraw.DDCreateSurface EnemySurface(3), App.Path & "\images\e3.bmp", EnemyRect(3), , , 0
End Sub

Public Sub InitE4()
  EnemyHeight(4) = 40
  EnemyWidth(4) = 50
  EnemyShield(4) = 0
  EnemyHull(4) = 130
  EnemyValue(4) = 100
  EnemyVelocity(4) = 5
  EnemyFramesX(4) = 5
  EnemyFramesY(4) = 2
  EnemyShoot(4) = False
  clsDDraw.DDCreateSurface EnemySurface(4), App.Path & "\images\e4.bmp", EnemyRect(4), , , 0
End Sub

Public Sub InitE5()
  EnemyHeight(5) = 50
  EnemyWidth(5) = 52
  EnemyShield(5) = 1200
  EnemyHull(5) = 1500
  EnemyShotL(5) = 12
  EnemyShotR(5) = 62
  EnemyFramesX(5) = 0
  EnemyFramesY(5) = 0
  EnemyShotY(5) = 62
  EnemyValue(5) = 2000
  EnemyVelocity(5) = 2
  EnemyShoot(5) = True
  clsDDraw.DDCreateSurface EnemySurface(5), App.Path & "\images\e5.bmp", EnemyRect(5), EnemyWidth(5), EnemyHeight(5), 0
End Sub

Public Sub InitE6()
  EnemyHeight(6) = 37
  EnemyWidth(6) = 50
  EnemyShield(6) = 200
  EnemyHull(6) = 400
  EnemyShotL(6) = 12
  EnemyShotR(6) = 62
  EnemyFramesX(6) = 0
  EnemyFramesY(6) = 0
  EnemyShotY(6) = 62
  EnemyValue(6) = 1000
  EnemyVelocity(6) = 2
  EnemyShoot(6) = True
  clsDDraw.DDCreateSurface EnemySurface(6), App.Path & "\images\e6.bmp", EnemyRect(6), EnemyWidth(6), EnemyHeight(6), 0
End Sub

