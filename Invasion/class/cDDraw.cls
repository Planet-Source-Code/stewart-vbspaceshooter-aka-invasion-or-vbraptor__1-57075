VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - cDDraw.cls                                            |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+


Private objDx As New DirectX7
Private objDD As DirectDraw7


Dim CurModeActiveStatus As Boolean 'This checks that we still have the correct display mode
Dim bRestore As Boolean 'If we don't have the correct display mode then this flag states that we need to restore the display

Public Sub RestoreAllSurfaces()
  objDD.RestoreAllSurfaces
End Sub

Public Sub Init(lWnd As Long, lWidth As Long, lHeight As Long, lColors As Long, DDS_Primary As DirectDrawSurface7, DDS_Buffer As DirectDrawSurface7, DDSD_Buffer As DDSURFACEDESC2)
  Dim DDSD_Store As DDSURFACEDESC2
  Dim caps As DDSCAPS2
  Set objDD = objDx.DirectDrawCreate("")
  frmMain.Show  'Show the form
  
  Call objDD.SetCooperativeLevel(lWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
  'This probably will be changed for more compatability in the future but at the
  'moment I'm making the assumption that most computers will support this rather
  'primitive state :)
  
  
  
  
  Call objDD.SetDisplayMode(lWidth, lHeight, lColors, 0, DDSDM_DEFAULT)
  'Setup the primary surface
  DDSD_Store.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
  DDSD_Store.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
  DDSD_Store.lBackBufferCount = 1
  Set DDS_Primary = objDD.CreateSurface(DDSD_Store)
  
  'now grab the back surface (from the flipping chain)
  caps.lCaps = DDSCAPS_BACKBUFFER
  Set DDS_Buffer = DDS_Primary.GetAttachedSurface(caps)
  DDS_Buffer.GetSurfaceDesc DDSD_Buffer
End Sub

Public Sub EndIt()
    objDD.RestoreAllSurfaces
    Call objDD.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
    'Stop the program:
    End
End Sub

Public Function ExModeActive() As Boolean
'This is used to test if we're in the correct resolution.
  Dim TestCoopRes As Long
  TestCoopRes = objDD.TestCooperativeLevel
  If (TestCoopRes = DD_OK) Then
    ExModeActive = True
  Else
    ExModeActive = False
  End If
End Function

Public Sub DDCreateSurface(Surface As DirectDrawSurface7, BmpPath As String, RECTvar As RECT, Optional mWidth As Long, Optional mHeight As Long, Optional TransCol As Integer = 0, Optional UseSystemMemory As Boolean = True)
    'This sub will load a bitmap from a file
    'into a specified dd surface. Transparent
    'colour is black (0) by default.
    
    Dim tempddsd As DDSURFACEDESC2
    
    Set Surface = Nothing
    
    'Load sprite
    tempddsd.lFlags = DDSD_CAPS
    If UseSystemMemory = True Then
        tempddsd.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY Or DDSCAPS_OFFSCREENPLAIN
    Else
        tempddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    End If
    If mWidth > 0 Then
      tempddsd.lWidth = mWidth
    End If
    If mHeight > 0 Then
      tempddsd.lHeight = mHeight
    End If
    Set Surface = objDD.CreateSurfaceFromFile(BmpPath, tempddsd)
    
    'set the RECT dimensions
    RECTvar.Right = tempddsd.lWidth
    RECTvar.Bottom = tempddsd.lHeight
    'Colour key
    Dim ddckColourKey As DDCOLORKEY
    ddckColourKey.low = TransCol
    ddckColourKey.high = TransCol
    Surface.SetColorKey DDCKEY_SRCBLT, ddckColourKey
    
    
End Sub


Public Function TickCount() As Long
  TickCount = objDx.TickCount
End Function

Public Sub DrawText(Surface As DirectDrawSurface7, iX As Integer, iY As Integer, sText As String, lColor As Long)
  Surface.SetForeColor lColor
  Surface.DrawText iX, iY, sText, False
End Sub
