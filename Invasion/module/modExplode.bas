Attribute VB_Name = "modExplode"
Option Explicit
Private Const MaxExplode = 30
Private Const ExplodeTick = 30
Private Const ExplodeWidth = 77
Private Const ExplodeHeight = 64

Private ExplodeDC As Long

Private Type Explode
  X As Long
  y As Long
  active As Boolean
  Frame As Integer
  LastTick As Long
End Type

Private Explodes(MaxExplode) As Explode

Public Sub InitExplode()
  ExplodeDC = GenerateDC(App.Path & "\images\explode.bmp", ExplodeDC)
End Sub

Private Function FindExplode() As Integer
  Dim X As Integer
  For X = 0 To MaxExplode
    If (Explodes(X).active = False) Then
      FindExplode = X
      Exit Function
    End If
  Next
End Function

Public Sub AddExplode(X, y)
  Dim NewExplode As Integer
  PlayWav "explo2.wav"
  NewExplode = FindExplode
  Explodes(NewExplode).active = True
  Explodes(NewExplode).X = X
  Explodes(NewExplode).y = y
  Explodes(NewExplode).LastTick = GetTickCount
  Explodes(NewExplode).Frame = 0
End Sub

Public Sub UpdateExplode()
  Dim X As Integer, NewTick As Long
  NewTick = GetTickCount
  For X = 0 To MaxExplode
    If (Explodes(X).active = True) Then
      BitBlt frmMain.picBak.hdc, Explodes(X).X, Explodes(X).y, ExplodeWidth, ExplodeHeight, ExplodeDC, Explodes(X).Frame * ExplodeWidth, ExplodeHeight, vbSrcAnd
      BitBlt frmMain.picBak.hdc, Explodes(X).X, Explodes(X).y, ExplodeWidth, ExplodeHeight, ExplodeDC, Explodes(X).Frame * ExplodeWidth, 0, vbSrcPaint
      If (NewTick - Explodes(X).LastTick > ExplodeTick) Then
        Explodes(X).LastTick = GetTickCount
        Explodes(X).Frame = Explodes(X).Frame + 1
        If Explodes(X).Frame = 13 Then Explodes(X).active = False
      End If
    End If
  Next
End Sub
 
