Attribute VB_Name = "modBadGuys"
Public Sub CreateEnemy(iShip As Integer)
  EShip(iShip).AI = 2
  EShip(iShip).Hull = 30
  EShip(iShip).Shield = 15
  EShip(iShip).Y = (0 - EShip(iShip).Height)
  EShip(iShip).Width = 85
  EShip(iShip).Height = 47
  EShip(iShip).x = ((640 - EShip(iShip).Width) * Rnd) + 1
  EShip(iShip).Active = True
  LastGuyTick = clsDDraw.TickCount
End Sub

Public Sub UpdateEnemyShips()
  Dim x As Integer
  For x = 0 To MaxEnemyShips
    EShip(x).Y = EShip(x).Y + EShip(x).AI
    Draw DDS_E1, rE1, EShip(x).x, EShip(x).Y, True, True
    If EShip(x).Y >= 640 Or EShip(x).Active = False Then
      If ((clsDDraw.TickCount - LastGuyTick) > NewBadGuyTick) Then CreateEnemy x
    End If
  Next
End Sub
