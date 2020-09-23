VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBadGuy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public x As Integer
Public y As Integer
Public Shield As Integer
Public Hull As Integer
Public AI As Integer
Public bBoss As Boolean
Public MoveDir As Integer
Public Velocity As Integer
Private rRECT As RECT
Public Width As Integer
Public Height As Integer
Public active As Boolean
Public FramesX As Integer
Public FramesY As Integer
Public FrameX As Integer
Public FrameY As Integer
Public Tick As Long
Public Surface As DirectDrawSurface7
Public Value As Long
Public ShotL As Long
Public ShotR As Long
Public ShotY As Long
Public Index As Integer

Public Property Get lRECT() As RECT
  lRECT = rRECT
End Property

Public Property Let lRECT(NewRect As RECT)
  rRECT = NewRect
End Property


