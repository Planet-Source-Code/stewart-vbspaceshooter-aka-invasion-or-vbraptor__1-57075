Attribute VB_Name = "modBackTile"
Public Type Star
    x As Integer
    y As Integer
    bright As Byte
    speed As Byte
End Type

Const NumOfStars = 50
Public StarArray(0 To NumOfStars) As Star

Public BackTile1 As Long
Public BackTile2 As Long
Public Tile1Y As Long
Public Tile2Y As Long
Private Const TileHeight = 800
Private Const TileWidth = 600
Public Sub InitBackTile()
  BackTile1 = GenerateDC(App.Path & "\images\space1.bmp", BackTile1)
  BackTile2 = GenerateDC(App.Path & "\images\space1.bmp", BackTile2)
  Tile1Y = 0
  Tile2Y = (0 - TileHeight)
End Sub
Public Sub UpdateBackTile()
  Tile1Y = Tile1Y + 1
  Tile2Y = Tile2Y + 1
  If (Tile1Y >= frmMain.picBak.ScaleHeight) Then Tile1Y = (Tile2Y - TileHeight)
  If (Tile2Y >= frmMain.picBak.ScaleHeight) Then Tile2Y = (Tile1Y - TileHeight)
  BitBlt frmMain.picBak.hdc, Tile1X, Tile1Y, TileWidth, TileHeight, BackTile1, 0, 0, vbSrcCopy
  BitBlt frmMain.picBak.hdc, Tile2X, Tile2Y, TileWidth, TileHeight, BackTile2, 0, 0, vbSrcCopy
  DrawStars
End Sub

Public Sub DrawStars()

'Form1.PicScreenBuffer.Cls
'Draw the stars to their buffer
  For x = 0 To NumOfStars
    StarArray(x).y = StarArray(x).y + StarArray(x).speed
    If StarArray(x).y > frmMain.picBak.ScaleHeight Or StarArray(x).speed <= 0 Then BuildNewStar (x)
    SetPixelV frmMain.picBak.hdc, StarArray(x).x, StarArray(x).y, RGB(StarArray(x).bright, StarArray(x).bright, StarArray(x).bright)
  Next x

End Sub

Public Sub BuildNewStar1(ByVal ArrayVal As Integer)
    StarArray(ArrayVal).x = Rnd * frmMain.picBak.ScaleWidth
    StarArray(ArrayVal).y = Rnd * frmMain.picBak.ScaleHeight
    StarArray(ArrayVal).bright = Rnd * 255
    StarArray(ArrayVal).speed = Rnd * 5 + 2
End Sub

Public Sub BuildNewStar(ByVal ArrayVal As Integer)
    StarArray(ArrayVal).x = Rnd * frmMain.picBak.ScaleWidth
    StarArray(ArrayVal).y = 0
    StarArray(ArrayVal).bright = Rnd * 200 + 55
    StarArray(ArrayVal).speed = Rnd * 5 + 1
End Sub


