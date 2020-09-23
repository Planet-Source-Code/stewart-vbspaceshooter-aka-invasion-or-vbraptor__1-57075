Attribute VB_Name = "modEnemyTwo"
Public Sub InitE2()
  EnemyDC(2) = GenerateDC(App.Path & "\images\e2.bmp", EnemyDC(2))
  EnemyHeight(2) = 78
  EnemyWidth(2) = 76
  EnemyShield(2) = 125
  EnemyHull(2) = 200
  EnemyVelocity(2) = 2
End Sub

