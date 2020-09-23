Attribute VB_Name = "modPubVar"
Public pDC As Long 'Primary DC
Public RadarDC As Long
Public diffWidth As Long
Public diffHeight As Long
Public InPlay As Boolean
Public GamePause As Boolean
Public Sub BackBufferToFront()
  BitBlt frmMain.picMain.hdc, 0, 0, frmMain.picMain.ScaleWidth, frmMain.picMain.ScaleHeight, frmMain.picBak.hdc, 0, 0, vbSrcCopy
End Sub

  
