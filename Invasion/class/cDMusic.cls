VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDMusic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - cDMusic.cls                                           |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+


Private objDx As New DirectX7
Private ObjDM As DirectMusicLoader

Public Sub Init(ddmperformance As DirectMusicPerformance)
  Set ObjDM = objDx.DirectMusicLoaderCreate
  Set ddmperformance = objDx.DirectMusicPerformanceCreate()
  ddmperformance.Init Nothing, 0
  ddmperformance.SetPort -1, 80
  ddmperformance.SetMasterAutoDownload True
  ddmperformance.SetMasterVolume 0
End Sub

Public Sub SetVolume(ddmperformance As DirectMusicPerformance, iVol As Long)
  ddmperformance.SetMasterVolume (iVol * 42 - 3000)
End Sub

Public Sub LoadMusic(DDMSegment As DirectMusicSegment, sFileName As String)
  Set DDMSegment = ObjDM.LoadSegment(sFileName)
  DDMSegment.SetStandardMidiFile
End Sub

Public Sub PlayMIDI(ddmperformance As DirectMusicPerformance, DDMSegment As DirectMusicSegment, DDMSegmentState As DirectMusicSegmentState)
  Set DDMSegmentState = ddmperformance.PlaySegment(DDMSegment, 0, 0)
End Sub

Public Sub StopMidi(ddmperformance As DirectMusicPerformance, DDMSegment As DirectMusicSegment, DDMSegmentState As DirectMusicSegmentState)
  ddmperformance.Stop DDMSegment, DDMSegmentState, 0, 0
End Sub
