VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'+------------------------------------------------------------------+
'| Invasion - cDSound.cls                                            |
'+------------------------------------------------------------------+
'| Design and code by Stewart (sobert81@devedit.com)                |
'+------------------------------------------------------------------+

Private objDx As New DirectX7
Private ObjDS As DirectSound
Private dsPrimaryBuffer As DirectSoundBuffer


Public Sub Init(lWnd As Long)
  Set ObjDS = objDx.DirectSoundCreate("")
  
  ObjDS.SetCooperativeLevel lWnd, DSSCL_NORMAL
  
End Sub

Public Sub LoadSoundFile(DDS_SoundBuffer As DirectSoundBuffer, File As String)
  Dim DsDesc As DSBUFFERDESC
  Dim bdesc As DSBUFFERDESC
  Dim DsWave As WAVEFORMATEX
  Dim s As WAVEFORMATEX
  If Dir(File) = "" Then Exit Sub
  With bdesc
    .lFlags = DSBCAPS_PRIMARYBUFFER Or DSBCAPS_CTRLPAN          'this will be the primary buffer, and have panning capabilities
  End With
  
  Set dsPrimaryBuffer = ObjDS.CreateSoundBuffer(bdesc, s)
  ObjDS.SetCooperativeLevel frmMain.hWnd, DSSCL_EXCLUSIVE
  
  DsDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
  DsWave.nFormatTag = WAVE_FORMAT_PCM 'Sound Must be PCM otherwise we get errors
  DsWave.nChannels = 2    '1= Mono, 2 = Stereo
  DsWave.lSamplesPerSec = 22050
  DsWave.nBitsPerSample = 16 '16 =16bit, 8=8bit
  DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
  DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign
  'DsWave.nFormatTag = WAVE_FORMAT_PCM 'Sound Must be PCM otherwise we get errors
  'DsWave.nChannels = 2 '1= Mono, 2 = Stereo
  'DsWave.lSamplesPerSec = 22050
  'DsWave.nBitsPerSample = 16 '16 =16bit, 8=8bit
  'DsWave.nBlockAlign = DsWave.nBitsPerSample / 8 * DsWave.nChannels
  'DsWave.lAvgBytesPerSec = DsWave.lSamplesPerSec * DsWave.nBlockAlign

  ' This line creates a sound buffer based on the information you have just
  'put in the two structures
  Set DDS_SoundBuffer = ObjDS.CreateSoundBufferFromFile(File, DsDesc, DsWave)
End Sub
