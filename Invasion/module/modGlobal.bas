Attribute VB_Name = "modGlobal"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modGlobal.bas                                         |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+

Global variables
Public EnemyDestroyed As Long
Public LastGuyTick As Long
Public LastShotTick As Long
Public LastMissileTick As Long
Public PixelPerfect As Boolean
Public strUserName As String
Public strCallSign As String
Public bAtCallsign As Boolean
Public bInNewChar As Boolean
Public bInHanger As Boolean
Public iHangerDo As Integer
Public strHangerText As String
Public bShowingHanger As Boolean
Public bShowMenu As Boolean
Public InMenu As Boolean
Public LastEnergyUpdateTick As Long
Public InOptions As Boolean
Public Difficulty As Integer

Public InGame As Boolean
Public OptSel As Integer
Public MusVol As Long, SndVol As Integer
Public SelButton As Integer
Public bGamePause As Boolean

Public WhatToDo As Integer
Public MaxRows As Integer


Public bAtBoss As Boolean
Public bEndLevel As Boolean

Public rBoss As RECT
Public DDS_BOSS As DirectDrawSurface7

Public BBMoveDir As Integer
Public ShipInvincible As Boolean


'Some surface variables
Public DDS_YOU As DirectDrawSurface7
Public DDS_PULSEC As DirectDrawSurface7
Public DDS_PLASMAB As DirectDrawSurface7
Public DDS_PLASMA As DirectDrawSurface7
Public DDS_MICRO As DirectDrawSurface7
Public DDS_SHOP As DirectDrawSurface7
Public DDS_REACTOR As DirectDrawSurface7
Public DDS_PULSE As DirectDrawSurface7
Public DDS_YOUSMALL As DirectDrawSurface7
Public DDS_DISPLAY As DirectDrawSurface7
Public DDS_PSHOT As DirectDrawSurface7
Public DDS_Letters As DirectDrawSurface7
Public DDS_Missile As DirectDrawSurface7
Public DDS_MissileE As DirectDrawSurface7
Public DDS_ESHOT As DirectDrawSurface7
Public DDS_HANGER As DirectDrawSurface7
Public DDS_HEADER As DirectDrawSurface7
Public DDS_Bottom As DirectDrawSurface7
Public DDS_HIT As DirectDrawSurface7
Public DDS_NEWBACK As DirectDrawSurface7
Public DDS_NEWEARTH As DirectDrawSurface7
Public DDS_Explode As DirectDrawSurface7
Public DDS_HEALTH As DirectDrawSurface7
Public DDS_HBAR As DirectDrawSurface7
Public DDS_MenuBack As DirectDrawSurface7
Public DDS_MenuBack2 As DirectDrawSurface7
Public DDS_Back As DirectDrawSurface7
Public DDS_Cursor As DirectDrawSurface7
Public DDS_Money As DirectDrawSurface7
Public DDS_Save As DirectDrawSurface7
Public DDS_ITEMDISPLAY As DirectDrawSurface7
Public DDS_MESSAGE As DirectDrawSurface7
Public DDS_SHIELD As DirectDrawSurface7
Public DDS_SHOPSELL As DirectDrawSurface7

'These are all menu surface variables.
Public DDS_WALL1 As DirectDrawSurface7
Public DDS_WALL2 As DirectDrawSurface7
Public DDS_Button As DirectDrawSurface7

Public BackYPos1 As Integer
Public BackYPos2 As Integer

Public mlngElapsed As Long


Public DDS_Primary As DirectDrawSurface7
Public DDS_Buffer As DirectDrawSurface7
Public DDSD_Buffer As DDSURFACEDESC2

Public cVol As New clsVolume


Public lBlinkTime As Long
Public bBlinkOn As Boolean
Public ShotTickCount As Long
Public MissileShotTick As Long
Public PulseShotTick As Long
Public PlasmaShotTick As Long

Public bCanDoBoss As Boolean

Public Scene As fScene

Public BadShips As New Collection

Public bShowFPS As Boolean

Public strShopText
Public iShopDo As Integer
Public iShopItem As Integer
Public rShopItem As RECT
Public strShopStr1 As String
Public strShopStr2 As String
Public strShopStr3 As String
Public lShopPrice As Long
Public bStartShop As Boolean


Public strMsgText As String
Public bShowMsg As Boolean

Public XLoadPos As Integer
Public LoadCursorPos As Integer

Public lLoadFileCount As Long

Public iInMenuSel As Integer

Public lColorDepth As Long

Public LastClick As Long

Public BShopSell As Boolean

Public BShipDestroyed As Boolean
