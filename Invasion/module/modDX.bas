Attribute VB_Name = "modDX"
Option Explicit
Public objDx As New DirectX7
Public objDD As DirectDraw7
Public DDS_Primary As DirectDrawSurface7  'Front surface
Public DDSD_Primary As DDSURFACEDESC2
Public DDS_Buffer As DirectDrawSurface7   'Back buffer
Public DDSD_Buffer As DDSURFACEDESC2
Public DDS_You As DirectDrawSurface7      'This is the primary sprite
Public DDS_PShot As DirectDrawSurface7
Public DDS_Back As DirectDrawSurface7
Public DDSC_Back As DDSURFACEDESC2
Public bQuit As Boolean
Public BackYPos1 As Long
Public BackYPos2 As Long


Dim CurModeActiveStatus As Boolean 'This checks that we still have the correct display mode
Dim bRestore As Boolean 'If we don't have the correct display mode then this flag states that we need to restore the display

Public Sub NewGame()
  Set objDD = objDx.DirectDrawCreate("")
  frmMain.Show  'Show the form
  Call objDD.SetCooperativeLevel(frmMain.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    
  'This probably will be changed for more compatability in the future but at the
  'moment I'm making the assumption that most computers will support this rather
  'primitive state :)
  Call objDD.SetDisplayMode(800, 600, 16, 0, DDSDM_DEFAULT)
  
  'Setup the primary surface
  DDSD_Primary.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
  DDSD_Primary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
  DDSD_Primary.lBackBufferCount = 1
  Set DDS_Primary = objDD.CreateSurface(DDSD_Primary)
  
  'now grab the back surface (from the flipping chain)
  Dim caps As DDSCAPS2
  caps.lCaps = DDSCAPS_BACKBUFFER
  Set DDS_Buffer = DDS_Primary.GetAttachedSurface(caps)
  DDS_Buffer.GetSurfaceDesc DDSD_Buffer
  BackYPos1 = 0
  BackYPos2 = BackYPos1 - 600
  LoadData
  
  
  Main
End Sub

Private Sub LoadData()
  'Load the backdrop data in
  Dim CKey As DDCOLORKEY
  Set DDS_Back = Nothing   'Clear vars
  'Set DDS_You = Nothing   'Clear vars
  DDSC_Back.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
  DDSC_Back.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
  DDSC_Back.lWidth = DDSD_Buffer.lWidth
  DDSC_Back.lHeight = DDSD_Buffer.lHeight
  'Create the offscreen surface
  Set DDS_Back = objDD.CreateSurfaceFromFile(App.Path & "\images\space1.bmp", DDSC_Back)
  
  'Generate the sprites
  LoadSprite DDS_PShot, App.Path & "\images\bullet.bmp", 0, 0, 0
  LoadSprite DDS_You, App.Path & "\images\mainship.bmp", YouWidth, YouHeight, 0
  'Create the offscreen surface
  'Set DDS_You = objDD.CreateSurfaceFromFile(App.Path & "\images\mainship.bmp", DDSC_You)
  'CKey.low = vbBlack
  'CKey.high = vbBlack
  'DDS_You.SetColorKey DDCKEY_SRCBLT, CKey
End Sub

Public Sub Main()
  Do
    'draw current game screen
    CheckInput
    DrawFrame
    DoEvents
  Loop
  
End Sub

Public Sub EndIt()
    Call objDD.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
    'Stop the program:
    End
End Sub

Public Sub UpdateShots()
  Dim X As Integer, rShot As RECT
  rShot.Right = ShotWidth
  rShot.Bottom = ShotHeight
  For X = 0 To MaxShots
    If (PShotL(X).Active = True) Then
      PShotL(X).CurY = PShotL(X).CurY - 10
      If PShotL(X).CurY <= 0 Then PShotL(X).Active = False
      DDS_Buffer.BltFast PShotL(X).CurX, PShotL(X).CurY, DDS_PShot, rShot, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_DONOTWAIT
     End If
    If (PShotR(X).Active = True) Then
      PShotR(X).CurY = PShotR(X).CurY - 10
      If PShotR(X).CurY <= 0 Then PShotR(X).Active = False
      DDS_Buffer.BltFast PShotR(X).CurX, PShotR(X).CurY, DDS_PShot, rShot, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_DONOTWAIT
    End If
    If (PShotR(X).Active = False And PShotL(X).Active = False) Then
      If Firing = True And ((objDx.TickCount - LastShotTick) > ShotTick) Then CreateShot X
    End If
  Next
  
End Sub

Public Sub CreateShot(ShotI As Integer)
  Dim i As Integer
  PShotL(ShotI).Active = True
  PShotR(ShotI).Active = True
  PShotL(ShotI).CurX = (You.CurX + 3)
  PShotR(ShotI).CurX = (You.CurX + 95)
  PShotL(ShotI).CurY = (You.CurY + 1 + ShotHeight)
  PShotR(ShotI).CurY = (You.CurY + 1 + ShotHeight)
  LastShotTick = objDx.TickCount
End Sub

Public Sub DrawFrame()
    Dim rBack1 As RECT
    Dim rBack2 As RECT
    Dim rYou As RECT
    Dim ddrval As Long 'Every drawing procedure returns a value, so we must have a
                       'var able to hold it. From this value we can check for errors.
                       
    bRestore = False
    Do Until ExModeActive
      DoEvents
      bRestore = True
    Loop
    DoEvents
    If bRestore Then
      bRestore = False
      objDD.RestoreAllSurfaces 'this just re-allocates memory back to us. we must
                               'still reload all the surfaces.
      LoadData ' must init the surfaces again if they we're lost
    End If
    'rBack1.Left = 0
    'rBack1.Top = BackYPos1 + 600
    'rBack1.Bottom = (600 - rBack1.Top)
    'rBack1.Right = DDSD_Buffer.lWidth
    

    BackYPos1 = BackYPos1 - 1
    BackYPos2 = BackYPos2 - 1
    Dim YPos As Integer
    YPos = (rBack1.Bottom - rBack1.Top)
    rBack2.Top = 0
    rBack2.Left = 0
    rBack2.Bottom = (DDSD_Buffer.lHeight - YPos)
    rBack2.Right = DDSD_Buffer.lWidth
    If (BackYPos1 >= 600) Then BackYPos1 = BackYPos2 - 600
    If (BackYPos2 >= 600) Then BackYPos2 = BackYPos1 - 600
    ddrval = DDS_Buffer.BltFast(0, 0, DDS_Back, rBack1, DDBLTFAST_WAIT)
    'ddrval = DDS_Buffer.BltFast(0, 0, DDS_Back, rBack2, DDBLTFAST_WAIT)
    rYou.Left = 0
    rYou.Top = 0
    rYou.Right = YouWidth
    rYou.Bottom = YouHeight
    If You.CurX > (DDSD_Buffer.lWidth - rYou.Right) Then
      You.CurX = (DDSD_Buffer.lWidth - rYou.Right)
    End If
    If You.CurY > (DDSD_Buffer.lHeight - rYou.Bottom) Then
      You.CurY = (DDSD_Buffer.lHeight - rYou.Bottom)
    End If
    
    If You.CurX < 0 Then You.CurX = 0
    If You.CurY < 0 Then You.CurY = 0
    
    ddrval = DDS_Buffer.BltFast(You.CurX, You.CurY, DDS_You, rYou, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    UpdateShots
    DDS_Primary.Flip Nothing, DDFLIP_WAIT  'Flip the secondary surface to the primary
End Sub


Function ExModeActive() As Boolean
'This is used to test if we're in the correct resolution.
  Dim TestCoopRes As Long
  TestCoopRes = objDD.TestCooperativeLevel
  If (TestCoopRes = DD_OK) Then
    ExModeActive = True
  Else
    ExModeActive = False
  End If
End Function

Public Sub LoadSprite(ByRef Sprite As DirectDrawSurface7, File As String, bWidth As Integer, bHeight As Integer, ColourKey As Integer)

    Dim CKey As DDCOLORKEY
    Dim ddsdNewSprite As DDSURFACEDESC2
    
    'This routine loads sprites in the FObject file and sets their colour keys
    ddsdNewSprite.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT           'Set the surface description to include the Capabilities, Width and Height
    ddsdNewSprite.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN                    'Set the surface's capabilities to be an offscreen surface
    'ddsdNewSprite.lWidth = bWidth                                           'Set the width of the surface
    'ddsdNewSprite.lHeight = bHeight                                         'Set the height of the surface
    Set Sprite = objDD.CreateSurfaceFromFile(File, ddsdNewSprite)           'Load the bitmap from the resource file into the surface using the surface description
    CKey.low = ColourKey                                                    'Set the low value of the colour key
    CKey.high = ColourKey                                                   'and the high value (in this case they're the same because we're not using a range)
    Sprite.SetColorKey DDCKEY_SRCBLT, CKey                                  'Set the sprites colourkey using the key just created

End Sub


