Attribute VB_Name = "modConst"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modConst.bas                                          |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+

' Key Codes
Public Const GoLeft = &H25
Public Const GoRight = &H27
Public Const GoForward = &H26
Public Const GoBack = &H28
Public Const DoShoot = &H20
Public Const DoReturn = 13
Public Const DoEscape = &H1B
Public Const DoShift = &H10
Public Const DoCtrl = &H11

Public Const MaxStars = 50

Public Const MaxShots = 20

Public Const MaxEnemies = 100

Public Const YouWidth = 51
Public Const YouHeight = 44

Public Const ShotWidth = 5
Public Const ShotHeight = 17

Public Const TileWidth = 64
Public Const TileHeight = 64
Public Const ShotTick = 125
Public Const MissileShotCount = 75

Public Const PulseShotCount = 200

Public Const NewBadGuyTick = 900
Public Const PlasmaShotCount = 25

Public Const MissileWidth = 4
Public Const MissileHeight = 16

Public Const ScreenWidth As Long = 640
Public Const ScreenHeight As Long = 480

Public Const MaxExplode = 30
Public Const ExplodeTick = 30
Public Const ExplodeWidth = 72
Public Const ExplodeHeight = 72

Public Const WallWidth = 320
Public Const WallHeight = 480

Public Const ButWidth = 300
Public Const ButHeight = 50


Public Const ShotVelocity = 10

Public Const ShieldEnergyUpdate = 45
Public Const Pi = 3.14159265358979 'define PI

Public Const JWidth = 499
Public Const JHeight = 302

Public Const BlinkTime = 90
Public Enum fScene
  fEntrance = 0
  fMenu = 1
  fOptions = 2
  fCreate = 3
  fPlayer = 4
  fGame = 5
  fBoss = 6
  fSave = 7
  fLoad = 8
  fShop = 9
End Enum


#If False Then 'Trick preserves Case of Enums when typing in IDE
Private fEntrance, fMenu, fOptions, fCreate, fPlayer, fGame, fBoss, fSave, fLoad
#End If

Public Const PulseWidth = 26
Public Const PulseHeight = 10

Public Const PlasmaWidth = 10
Public Const PlasmaHeight = 10

Public Const ClickTime = 1000
