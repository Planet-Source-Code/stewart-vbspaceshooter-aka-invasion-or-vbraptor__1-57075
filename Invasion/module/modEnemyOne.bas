Attribute VB_Name = "modEnemyOne"
Public Sub InitE1()
  EnemyDC(1) = GenerateDC(App.Path & "\images\e1.bmp", EnemyDC(0))
  EnemyHeight(1) = 47
  EnemyWidth(1) = 85
  EnemyShield(1) = 85
  EnemyShotL(1) = 27
  EnemyShotR(1) = 58
  EnemyShotY(1) = 37
  EnemyHull(1) = 85
  EnemyVelocity(1) = 2
End Sub

Public Sub E1AI(EnemyX As Integer)
  If BadGuys(EnemyX).y > (frmMain.picBak.ScaleHeight \ 2) Then
    BadGuys(EnemyX).Velocity = ((0 - (Int(Rnd * 2))))
    BadGuys(EnemyX).movedir = (Int(Rnd * 2) - 1)
    If BadGuys(EnemyX).movedir < 0 Then BadGuys(EnemyX).movedir = 0
  End If
  If BadGuys(EnemyX).Velocity <= 0 Then
    If (BadGuys(EnemyX).y <= 0) Then
      BadGuys(EnemyX).Velocity = (Int(Rnd * 2)) + 1
      BadGuys(EnemyX).movedir = (Int(Rnd * 2) - 1)
      If BadGuys(EnemyX).movedir < 0 Then BadGuys(EnemyX).movedir = 0
    End If
  End If
  If (BadGuys(EnemyX).movedir = 0) Then
    BadGuys(EnemyX).x = BadGuys(EnemyX).x + 2
  ElseIf (BadGuys(EnemyX).movedir = 1) Then
    BadGuys(EnemyX).x = BadGuys(EnemyX).x - 1
  End If
  If (BadGuys(EnemyX).x <= 0) Then
    BadGuys(EnemyX).movedir = 0
  ElseIf BadGuys(EnemyX).x + BadGuys(EnemyX).xsize >= frmMain.picBak.ScaleWidth Then
    BadGuys(EnemyX).movedir = 1
  End If
  If (15 > Int(Rnd * 1000)) Then
    CreateEShot FindEShot, EnemyX
  End If
End Sub
