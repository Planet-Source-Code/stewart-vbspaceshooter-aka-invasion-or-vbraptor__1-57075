Attribute VB_Name = "modTypes"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modTypes.bas                                          |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+

Public Type Weapon
  WeaponSurface As DirectDrawSurface7
  PosX As Integer
  PosY As Integer
  Damage As Integer
End Type
Public Type MainShip
  iLevel As Integer
  ShieldR As Integer
  CurX As Integer
  CurY As Integer
  Missiles As Boolean
  Energy As Boolean
  speed As Long
  CurMoney As Long
  strUserName As String
  strCallSign As String
  Shield As Long
  MaxShield As Long
  CurEnergy As Long
  ReactorPower As Long
  Hull As Long
  MaxHull As Long
  lPulse As Long
  Level As String
  lPlasma As Long
  lMicro As Long
End Type

Public Type LevelData
  MapFile As String
  TileBitmap As String
  EnemyFile As String
  MusicFile As String
  Intro As String
  BossHull As Long
  BossShield As Long
  BossLaserDamage As Long
  BossMissileDamage As Long
  BossLaser1X As Long
  BossLaser2X As Long
  BossMissile1X As Long
  BossMissile2X As Long
End Type

Public Type Tile_Data
  'X As Integer
  Y As Integer
  TileNumX As Integer
  TileNumY As Integer
End Type

Public Type cRGB
  r As Byte
  g As Byte
  b As Byte
End Type

Public Type Explode
  x As Long
  Y As Long
  active As Boolean
  frame As Integer
  LastTick As Long
End Type

Public Type ShotDataP
  CurX As Integer
  CurY As Integer
  active As Boolean
End Type

Public Type MissileData
  CurX As Integer
  CurY As Integer
  active As Boolean
End Type

Public Type PlasmaData
  CurX As Integer
  CurY As Integer
  active As Boolean
End Type

Public Type PulseData
  CurX As Integer
  CurY As Integer
  active As Boolean
End Type


Public Type Star
  x As Long
  Y As Long
  Color As Long
  Velocity As Integer
End Type

Public Type EnemyShip
  x As Integer
  Y As Integer
  Shield As Long
  Hull As Long
  AI As Integer
  bBoss As Boolean
  MoveDir As Integer
  Velocity As Integer
  RECT As RECT
  Width As Integer
  Height As Integer
  active As Boolean
  FramesX As Integer
  FramesY As Integer
  FrameX As Integer
  FrameY As Integer
  CanShoot As Boolean
  Tick As Long
  Surface As DirectDrawSurface7
  Value As Long
  ShotL As Long
  ShotR As Long
  ShotY As Long
End Type

Public Type Hit
  x As Integer
  Y As Integer
  active As Boolean
  frame As Integer
  LastTick As Long
  speed As Integer
End Type

Public Const HitTick = 25
Public Const MaxHit = 20
Public Const HitWidth = 12
Public Const HitHeight = 12

Public Hits(MaxHit) As Hit

Public Stars(MaxStars) As Star

Public PShotL(MaxShots) As ShotDataP
Public PShotR(MaxShots) As ShotDataP

Public MissileL(MaxShots) As MissileData
Public MissileR(MaxShots) As MissileData

Public Pulse(MaxShots) As PulseData
Public Plasma(MaxShots) As PlasmaData

Public You As MainShip

Public BadGuys(MaxEnemies) As EnemyShip

Public Explodes(MaxExplode) As Explode

Public Const UpdateEnemyAnimTick = 45

Public Level As LevelData

Public Type Money
  x As Integer
  Y As Integer
  active As Boolean
  frame As Integer
  LastTick As Long
  RECT As RECT
End Type

