Attribute VB_Name = "modDXS"
Option Explicit

Public perf As DirectMusicPerformance                  'DirectMusic Performance object
Public seg As DirectMusicSegment                        'DirectMusic Segment
Public segstate As DirectMusicSegmentState              'DirectMusic Segment State
Public loader As DirectMusicLoader                      'DirectMusic Loader
Public Offset As Long

Public dx As New DirectX7
Public DS As DirectSound                            'Direct Sound object
Public dsPrimaryBuffer As DirectSoundBuffer         'Primary direct sound buffer
Public dsLaser2 As DirectSoundBuffer                'stage 2 laser fire buffer
Public dsLaser As DirectSoundBuffer                 'stage 1 laser fire
Public dsexplotion As DirectSoundBuffer             'explosion sound effect
Public dsPowerUp As DirectSoundBuffer               'power up sound effect buffer
Public dsMissile As DirectSoundBuffer               'missile sound effect buffer
Public dsEnergize As DirectSoundBuffer              'sound effect for when the player materializes
Public dsAlarm As DirectSoundBuffer                 'low shield alarm
Public dsEnemyFire As DirectSoundBuffer             'enemy fire direct sound buffer
Public dsnohit As DirectSoundBuffer                 'player hits an object that isn't destroyed
Public dsPulseCannon As DirectSoundBuffer           'sound for the pulse cannon
Public dsPlayerDies As DirectSoundBuffer            'sound for when the player dies
Public dsInvulnerability As DirectSoundBuffer       'sound for when the player is invulnerable
Public dsInvPowerDown As DirectSoundBuffer          'sound for when the invulnerability wears off
Public dsExtraLife As DirectSoundBuffer             'sound for when the player gets an extra life
Public DSEnemyFireIndex As Integer                  'Keeps track of how many sounds are playing at once
Public dsStart As DirectSoundBuffer
Public dsButX As DirectSoundBuffer
Public dsMusic As DirectSoundBuffer
'DX sound Module
Declare Sub ReleaseCapture Lib "user32" ()


Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long


Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub


Public Sub InitializeDS()
    
    'This sub initializes all of the sound effects used by SS2k
    
    Dim dsBDesc As DSBUFFERDESC                                     'variable to hold the direct sound buffer description
    Dim bdesc As DSBUFFERDESC                                       'variable to hold another direct sound buffer description
    Dim intCount As Integer                                         'standard count variable
    
    'On Local Error GoTo ErrorHandler                                'make sure to handle any errors
    
    Set DS = dx.DirectSoundCreate("")                               'create the direct sound object using the default sound device
    DS.SetCooperativeLevel frmMain.hwnd, DSSCL_NORMAL       'set the cooperative level to the space shooter form, and use normal mode
    With bdesc
        .lFlags = DSBCAPS_PRIMARYBUFFER Or DSBCAPS_CTRLPAN          'this will be the primary buffer, and have panning capabilities
    End With
    
    Dim s As WAVEFORMATEX, w As WAVEFORMATEX                        'Dim two waveformatex structures
    Set dsPrimaryBuffer = DS.CreateSoundBuffer(bdesc, s)            'create a primary sound buffer using the buffer desc and wave format structure
    
    With dsBDesc
        .lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLFREQUENCY          'allow for pan and frequency changes
    End With
    
    'The next lines load up all of the wave files using the default capabilites
    
    Set dsnohit = DS.CreateSoundBufferFromFile(App.Path & "\sounds\hitp.wav", dsBDesc, w)
    Set dsLaser = DS.CreateSoundBufferFromFile(App.Path & "\sounds\shote.wav", dsBDesc, w)
    Set dsexplotion = DS.CreateSoundBufferFromFile(App.Path & "\sounds\explo2.wav", dsBDesc, w)
    'The next lines initialize duplicate sound buffers from the existing ones
    


    
    Exit Sub

ErrorHandler:                                                       'Handle any errors
                                                             'release any objects
    MsgBox "Unable to create Direct Sound object." & vbCrLf & "Check to ensure that a sound card is installed and working properly."
                                                                    'display this message
    End                                                             'end the program
    
End Sub



