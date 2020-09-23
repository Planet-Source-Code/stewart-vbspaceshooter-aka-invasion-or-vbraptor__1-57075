Attribute VB_Name = "modEnemyAI"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modAI.bas                                             |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+

'First AI Format called ai bounce. Enemies move in a circular motion
Public Sub AIBounce(EnemyX As Integer)

  BadGuys(EnemyX).x = BadGuys(EnemyX).x - Cos(BadGuys(EnemyX).Y / 180 * Pi) * 1.3
  BadGuys(EnemyX).Y = BadGuys(EnemyX).Y + BadGuys(EnemyX).Velocity
  If (35 + (Difficulty * 20) > Int(Rnd * 1000)) Then
    If BadGuys(EnemyX).CanShoot Then CreateEShot FindEShot, EnemyX
  End If

End Sub
Public Sub BigBoss(EnemyX As Integer)
  If BBMoveDir = 0 Then
    If BadGuys(EnemyX).Y < ((ScreenHeight - BadGuys(EnemyX).Height - 67) \ 2) Then
      BadGuys(EnemyX).Y = BadGuys(EnemyX).Y + 2
    Else
      BBMoveDir = 1
    End If
    'Exit Sub
  End If
  If BBMoveDir = 1 Then
    If BadGuys(EnemyX).x > 0 Then
      BadGuys(EnemyX).x = BadGuys(EnemyX).x - 2
    Else
      BBMoveDir = 2
    End If
    'Exit Sub
  End If
  If BBMoveDir = 2 Then
    If BadGuys(EnemyX).Y > 5 Then
      BadGuys(EnemyX).Y = BadGuys(EnemyX).Y - 2
    Else
      BBMoveDir = 3
    End If
    'Exit Sub
  End If
  If BBMoveDir = 3 Then
    
    If BadGuys(EnemyX).x + BadGuys(EnemyX).Width < ScreenWidth Then
      BadGuys(EnemyX).x = BadGuys(EnemyX).x + 2
    Else
      BBMoveDir = 0
    End If
    'Exit Sub
  End If
  'If (BadGuys(EnemyX).y < 0) Then BBMoveDir = 0

  If (300 + (Difficulty * 50) > Int(Rnd * 1000)) Then
    If BadGuys(EnemyX).CanShoot Then CreateEShot FindEShot, EnemyX
    If BadGuys(EnemyX).CanShoot Then CreateEMissile FindEMissile, EnemyX
  End If
End Sub

Public Sub DownOff(EnemyX As Integer)
  If BadGuys(EnemyX).Y > (ScreenHeight \ 2) Then
    BadGuys(EnemyX).Velocity = 0
    If ((BadGuys(EnemyX).x <= (ScreenWidth \ 2))) Then
      BadGuys(EnemyX).MoveDir = 1
    Else
      BadGuys(EnemyX).MoveDir = 2
    End If
  End If
  If BadGuys(EnemyX).Velocity = 0 Then
    If BadGuys(EnemyX).MoveDir = 1 Then
      BadGuys(EnemyX).x = BadGuys(EnemyX).x - 1
    Else
      BadGuys(EnemyX).x = BadGuys(EnemyX).x + 1
    End If
  End If
  If (BadGuys(EnemyX).x + BadGuys(EnemyX).Width < 0) Or (BadGuys(EnemyX).x > ScreenWidth) Then
    BadGuys(EnemyX).active = False
    
  End If
  If (35 + (Difficulty * 20) > Int(Rnd * 1000)) Then
    If BadGuys(EnemyX).CanShoot Then CreateEShot FindEShot, EnemyX
  End If
  
End Sub

