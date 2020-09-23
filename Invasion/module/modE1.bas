Attribute VB_Name = "modE1"
Public Sub InitE1()
  EnemyHeight(1) = 47
  EnemyWidth(1) = 85
  EnemyShield(1) = 85
  EnemyShotL(1) = 27
  EnemyShotR(1) = 58
  EnemyShotY(1) = 37
  EnemyHull(1) = 85
  EnemyVelocity(1) = 2
  clsDDraw.DDCreateSurface EnemySurface(1), App.Path & "\images\e1.bmp", EnemyRect(1), EnemyWidth(1), EnemyHeight(1), 0
End Sub
