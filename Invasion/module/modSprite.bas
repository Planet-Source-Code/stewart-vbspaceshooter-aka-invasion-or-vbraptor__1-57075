Attribute VB_Name = "modSprite"
Public Type Ship
  Lasers As Long
  Missiles As Long
  Shields As Long
  Hull As Long
  speed As Long
  Forward As Boolean
  CurX As Long
  Active As Boolean
  CurY As Long
End Type

Public You As Ship
Public YouDC As Long
Public BackDC As Long
Public Lives As Integer
Public ShipInvincible As Boolean

Public Sub StartGame()
  Lives = 6
  NewGuy
End Sub

Public Sub InitMainChar()
  YouDC = GenerateDC(App.Path & "\images\mainship.bmp", YouDC)
End Sub

Public Sub NewGuy()
  Dim X As Long
  GamePause = True
  You.Active = True
  You.Shields = 300
  You.Hull = 200
  If Lives = 0 Then
    MsgBox "Game Over"
    InPlay = False
    End
  End If
  ShipInvincible = True
  Lives = Lives - 1
  frmMain.PicLives.Picture = LoadPicture(App.Path & "\images\LIVE" & (Lives) & ".bmp")
  You.CurX = (frmMain.picBak.ScaleWidth - 100) \ 2
  You.Forward = True
  For X = frmMain.picBak.ScaleHeight To frmMain.picBak.ScaleHeight - 200 Step -1
    You.CurY = X
    UpdateBackTile
    UpdatePFiring
    UpdateEFiring
    UpdateEnemies
    UpdateExplode
    UpdateHits
    UpdateYou
    BackBufferToFront
    'Sleep 1
    'DoEvents
  Next X
  ShipInvincible = False
  You.Forward = False
End Sub
Public Sub UpdateYou()
  If You.Forward = False Then
    If You.Active = True Then
      BitBlt frmMain.picBak.hdc, You.CurX, You.CurY, 100, 96, YouDC, 0, 96, vbSrcAnd
      BitBlt frmMain.picBak.hdc, You.CurX, You.CurY, 100, 96, YouDC, 0, 0, vbSrcPaint
    End If
  Else
    If You.Active = True Then
      BitBlt frmMain.picBak.hdc, You.CurX, You.CurY, 100, 96, YouDC, 100, 96, vbSrcAnd
      BitBlt frmMain.picBak.hdc, You.CurX, You.CurY, 100, 96, YouDC, 100, 0, vbSrcPaint
    End If
  End If
End Sub
