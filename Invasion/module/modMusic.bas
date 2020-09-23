Attribute VB_Name = "modMusic"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - modMusic.bas                                          |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+

Private clsDMusic As New cDMusic
Private DDM_MainPerf As DirectMusicPerformance
Private DDM_MainSeg As DirectMusicSegment
Private DDM_MainSegState As DirectMusicSegmentState

Public Sub InitSound()
  clsDMusic.Init DDM_MainPerf
  SetVol MusVol
End Sub

Public Sub SetVol(iVol As Long)
  clsDMusic.SetVolume DDM_MainPerf, iVol
End Sub

Public Sub PlayBGSound(sFileName As String)
  clsDMusic.LoadMusic DDM_MainSeg, sFileName
  clsDMusic.PlayMIDI DDM_MainPerf, DDM_MainSeg, DDM_MainSegState
End Sub

Public Sub StopBGSound()
  clsDMusic.StopMidi DDM_MainPerf, DDM_MainSeg, DDM_MainSegState
End Sub

