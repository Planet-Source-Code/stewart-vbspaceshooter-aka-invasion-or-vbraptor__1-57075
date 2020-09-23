Attribute VB_Name = "modShooting"
Option Explicit
Private Const MaxPShots = 50
Private Const MaxEShots = 50
Public BulletDC As Long
Public BulletEDC As Long
Public ShotDC As Long
Public HitDC As Long
Public Firing As Boolean
Private Type Shot
  X As Long
  y As Long
  active As Boolean
End Type

Private Type Hit
  X As Integer
  y As Integer
  active As Boolean
  Frame As Integer
  LastTick As Long
  speed As Integer
End Type

Private Const HitTick = 25
Private Const MaxHit = 20
Private Const HitWidth = 12
Private Const HitHeight = 12

Private Hits(MaxHit) As Hit

Private Const ShotTick = 125

Private Const ShotWidth = 5
Private Const ShotHeight = 17

Private LastShotTick As Long
Public LastEShotTick

Private PShotL(0 To MaxPShots) As Shot
Private PShotR(0 To MaxPShots) As Shot
Private EShotL(0 To MaxEShots) As Shot
Private EShotR(0 To MaxEShots) As Shot

Public Sub InitBullet()
  BulletDC = GenerateDC(App.Path & "\images\bullet.bmp", BulletDC)
  BulletEDC = GenerateDC(App.Path & "\images\bullete.bmp", BulletEDC)
  ShotDC = GenerateDC(App.Path & "\images\shot.bmp", ShotDC)
  HitDC = GenerateDC(App.Path & "\images\hit.bmp", HitDC)
End Sub

Private Function FindHit() As Integer
  Dim X As Integer
  For X = 0 To MaxHit
    If (Hits(X).active = False) Then
      FindHit = X
      Exit Function
    End If
  Next
  FindHit = MaxHit
End Function

Public Sub UpdateHits()
  Dim X As Integer, NewTick As Long
  NewTick = GetTickCount
  For X = 0 To MaxHit
    If (Hits(X).active = True) Then
      Hits(X).y = Hits(X).y + Hits(X).speed
      BitBlt frmMain.picBak.hdc, Hits(X).X, Hits(X).y, HitWidth, HitHeight, HitDC, Hits(X).Frame * HitWidth, 12, vbSrcAnd
      BitBlt frmMain.picBak.hdc, Hits(X).X, Hits(X).y, HitWidth, HitHeight, HitDC, Hits(X).Frame * HitWidth, 0, vbSrcPaint
    End If
    If ((NewTick - Hits(X).LastTick) > HitTick) Then
      Hits(X).LastTick = GetTickCount
      Hits(X).Frame = Hits(X).Frame + 1
      If Hits(X).Frame = 6 Then
        Hits(X).active = False
      End If
    End If
    
  Next
End Sub

Private Sub CreateHit(X As Long, y As Long, speed As Integer)
  Dim NewHit As Integer
  NewHit = FindHit()
  Hits(NewHit).active = True
  Hits(NewHit).Frame = 0
  Hits(NewHit).LastTick = GetTickCount
  Hits(NewHit).X = X
  Hits(NewHit).y = y
  Hits(NewHit).speed = speed
End Sub

Public Sub UpdateEFiring()
  Dim X As Integer
  For X = 0 To MaxEShots
    If (EShotL(X).active = True Or EShotR(X).active = True) Then
      If EShotL(X).active = True Then EShotL(X).y = EShotL(X).y + 10
      If EShotR(X).active = True Then EShotR(X).y = EShotR(X).y + 10
      If EShotL(X).y >= frmMain.picBak.ScaleHeight Then EShotL(X).active = False
      If EShotR(X).y >= frmMain.picBak.ScaleHeight Then EShotR(X).active = False
      If CollisionDetect(EShotL(X).X, EShotL(X).y, ShotWidth, ShotHeight, ShotWidth, 0, BulletEDC, You.CurX, You.CurY, 100, 96, 0, 96, YouDC = True) And EShotL(X).active = True Then
        CreateHit EShotL(X).X, (EShotL(X).y + ShotHeight), 0
        EShotL(X).active = False
        If You.Shields > 0 Then
          If ShipInvincible = False Then You.Shields = You.Shields - 20
        Else
          If ShipInvincible = False Then You.Hull = You.Hull - 35
        End If
        If You.Shields <= 0 And You.Hull <= 0 Then
          AddExplode You.CurX, You.CurY
          Firing = False
          You.active = False
          NewGuy
        End If
      End If
      If CollisionDetect(EShotR(X).X, EShotR(X).y, ShotWidth, ShotHeight, ShotWidth, 0, BulletEDC, You.CurX, You.CurY, 100, 96, 0, 96, YouDC = True) And EShotR(X).active = True Then
        CreateHit EShotR(X).X, (EShotR(X).y + ShotHeight), 0
        EShotR(X).active = False
        If You.Shields > 0 Then
          You.Shields = You.Shields - 20
        Else
          You.Hull = You.Hull - 35
        End If
        If You.Shields <= 0 And You.Hull <= 0 Then
          AddExplode You.CurX, You.CurY
          Firing = False
          You.active = False
          NewGuy
        End If
      End If
        
      If EShotL(X).active = True Then
        BitBlt frmMain.picBak.hdc, EShotL(X).X, EShotL(X).y, ShotWidth, ShotHeight, BulletEDC, ShotWidth, 0, vbSrcAnd
        BitBlt frmMain.picBak.hdc, EShotL(X).X, EShotL(X).y, ShotWidth, ShotHeight, BulletEDC, 0, 0, vbSrcPaint
      End If
      If EShotR(X).active = True Then
        BitBlt frmMain.picBak.hdc, EShotR(X).X, EShotR(X).y, ShotWidth, ShotHeight, BulletEDC, ShotWidth, 0, vbSrcAnd
        BitBlt frmMain.picBak.hdc, EShotR(X).X, EShotR(X).y, ShotWidth, ShotHeight, BulletEDC, 0, 0, vbSrcPaint
      End If
    End If
  Next
End Sub

Public Function FindEShot() As Integer
  Dim X As Integer
  For X = 0 To MaxEShots
    If EShotR(X).active = False And EShotL(X).active = False Then
      FindEShot = X
      Exit Function
    End If
  Next
End Function

Public Sub UpdatePFiring()
  Dim X As Long, y As Long
  Dim ShotTickCount As Long
  ShotTickCount = GetTickCount()
  For X = 0 To MaxPShots
    If (PShotL(X).active = True Or PShotR(X).active = True) Then
      If PShotL(X).active = True Then PShotL(X).y = PShotL(X).y - 10
      If PShotR(X).active = True Then PShotR(X).y = PShotR(X).y - 10
      For y = 0 To MaxEnemies
        If CollisionDetect(PShotL(X).X, PShotL(X).y, ShotWidth, ShotHeight, ShotWidth, 0, ShotDC, BadGuys(y).X, BadGuys(y).y, BadGuys(y).xsize, BadGuys(y).ysize, 0, BadGuys(y).ysize, BadGuys(y).ImgDC = True) And PShotL(X).active = True And BadGuys(y).Activated = True Then
          CreateHit PShotL(X).X, PShotL(X).y, BadGuys(y).Velocity
          PlayWav "hitp.wav"
          If BadGuys(y).Shield > 0 Then
            BadGuys(y).Shield = BadGuys(y).Shield - 60
          Else
            BadGuys(y).Hull = BadGuys(y).Hull - 100
          End If
          If BadGuys(y).Shield <= 0 And BadGuys(y).Hull <= 0 Then
            BadGuys(y).Activated = False
            AddExplode BadGuys(y).X, BadGuys(y).y
          End If
          PShotL(X).active = False
        End If
        If CollisionDetect(PShotR(X).X, PShotR(X).y, ShotWidth, ShotHeight, ShotWidth, 0, ShotDC, BadGuys(y).X, BadGuys(y).y, BadGuys(y).xsize, BadGuys(y).ysize, 0, BadGuys(y).ysize, BadGuys(y).ImgDC = True) And PShotR(X).active = True And BadGuys(y).Activated = True Then
          CreateHit PShotR(X).X, PShotR(X).y, BadGuys(y).Velocity
          PlayWav "hitp.wav"
          If BadGuys(y).Shield > 0 Then
            BadGuys(y).Shield = BadGuys(y).Shield - 30
          Else
            BadGuys(y).Hull = BadGuys(y).Hull - 50
          End If
          If BadGuys(y).Shield <= 0 And BadGuys(y).Hull <= 0 Then
            BadGuys(y).Activated = False
            AddExplode BadGuys(y).X, BadGuys(y).y
          End If
          PShotR(X).active = False
        End If
      Next y
      If PShotL(X).y <= 0 Or PShotR(X).y <= 0 Then
        PShotL(X).active = False

        PShotR(X).active = False
      End If
      If PShotL(X).active = True Then
        BitBlt frmMain.picBak.hdc, PShotL(X).X, PShotL(X).y, ShotWidth, ShotHeight, BulletDC, ShotWidth, 0, vbSrcAnd
        BitBlt frmMain.picBak.hdc, PShotL(X).X, PShotL(X).y, ShotWidth, ShotHeight, BulletDC, 0, 0, vbSrcPaint
      End If
      If PShotR(X).active = True Then
        BitBlt frmMain.picBak.hdc, PShotR(X).X, PShotR(X).y, ShotWidth, ShotHeight, BulletDC, ShotWidth, 0, vbSrcAnd
        BitBlt frmMain.picBak.hdc, PShotR(X).X, PShotR(X).y, ShotWidth, ShotHeight, BulletDC, 0, 0, vbSrcPaint
      End If
    Else
      If Firing = True And ((ShotTickCount - LastShotTick) > ShotTick) Then CreateShot X
    End If
  Next
End Sub

Public Sub CreateEShot(ShotI As Long, EnemyX As Integer)
  Dim i As Integer
  PlayWav "shote.wav"
  If (GetTickCount - LastEShotTick < ShotTick) Then Exit Sub
    
  EShotL(ShotI).active = True
  EShotR(ShotI).active = True
  EShotL(ShotI).X = BadGuys(EnemyX).X + BadGuys(EnemyX).ShotL
  EShotR(ShotI).X = BadGuys(EnemyX).X + BadGuys(EnemyX).ShotR
  
  EShotL(ShotI).y = (BadGuys(EnemyX).y + 1 + BadGuys(EnemyX).ShotY)
  EShotR(ShotI).y = (BadGuys(EnemyX).y + 1 + BadGuys(EnemyX).ShotY)
  LastEShotTick = GetTickCount
End Sub

Public Sub CreateShot(ShotI As Long)
  Dim i As Integer
  PlayWav "shote.wav"
  PShotL(ShotI).active = True
  PShotR(ShotI).active = True
  PShotL(ShotI).X = (You.CurX + 3)
  PShotR(ShotI).X = (You.CurX + 96)
  PShotL(ShotI).y = (You.CurY + 1 + ShotHeight)
  PShotR(ShotI).y = (You.CurY + 1 + ShotHeight)
  BitBlt frmMain.picBak.hdc, You.CurX + 1, You.CurY + 32, 21, 12, ShotDC, 14, 12, vbSrcAnd
  BitBlt frmMain.picBak.hdc, You.CurX + 1, You.CurY + 32, 21, 12, ShotDC, 14, 0, vbSrcPaint
  BitBlt frmMain.picBak.hdc, You.CurX + 93, You.CurY + 32, 21, 12, ShotDC, 14, 12, vbSrcAnd
  BitBlt frmMain.picBak.hdc, You.CurX + 93, You.CurY + 32, 21, 12, ShotDC, 14, 0, vbSrcPaint
  LastShotTick = GetTickCount()
End Sub

Public Function PlayWav(wav As String)
  Dim ua As Long
        ua = mciSendString("open " & App.Path & "\sounds\" & wav & " Type sequencer Alias MFile", 0&, 0, 0)
        ua = mciSendString("play MFile", 0&, 0, 0)
  
End Function
