VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'                        The Fak-House
'
'             --------------------------------------------
'                   Microsoft Volume Sample Project
'                            January l998
'             -------------------------------------------
'                    (c) Microsoft Corporation, 1997
'
'
'
'                Realisé par M Fofana Abdel-Kader

Option Explicit

Private Const MMSYSERR_NOERROR = 0
Private Const MAXPNAMELEN = 32
Private Const MIXER_LONG_NAME_CHARS = 64
Private Const MIXER_SHORT_NAME_CHARS = 16
Private Const MIXER_GETLINEINFOF_COMPONENTTYPE = &H3&
Private Const MIXER_GETCONTROLDETAILSF_VALUE = &H0&
Private Const MIXER_GETLINECONTROLSF_ONEBYTYPE = &H2&
Private Const MIXERLINE_COMPONENTTYPE_DST_FIRST = &H0&
Private Const MIXERLINE_COMPONENTTYPE_SRC_FIRST = &H1000&
Private Const MIXER_SETCONTROLDETAILSF_VALUE = &H0&
      
Private Const MIXERCONTROL_CT_CLASS_FADER = &H50000000
Private Const MIXERCONTROL_CT_UNITS_UNSIGNED = &H30000
Private Const MIXERCONTROL_CT_CLASS_LIST = &H70000000
Private Const MIXERCONTROL_CT_SC_LIST_MULTIPLE = &H1000000
Private Const MIXERCONTROL_CT_UNITS_BOOLEAN = &H10000
Private Const MIXERCONTROL_CT_CLASS_SWITCH = &H20000000
Private Const MIXERCONTROL_CT_SC_SWITCH_BOOLEAN = &H0&
Private Const MIXERCONTROL_CT_CLASS_METER& = &H10000000
Private Const MIXERCONTROL_CT_SC_METER_POLLED& = &H0&
Private Const MIXERCONTROL_CT_UNITS_SIGNED& = &H20000

Private Const MIXERLINE_COMPONENTTYPE_DST_SPEAKERS = _
    (MIXERLINE_COMPONENTTYPE_DST_FIRST + 4)
                     
Private Const MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 3)
      
Private Const MIXERLINE_COMPONENTTYPE_SRC_LINE = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 2)
      
Private Const MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 5)
    
Private Const MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 4)

Private Const MIXERLINE_COMPONENTTYPE_DST_WAVEIN = _
    (MIXERLINE_COMPONENTTYPE_DST_FIRST + 7)
Private Const MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT = _
    (MIXERLINE_COMPONENTTYPE_SRC_FIRST + 8)

Private Const MIXERCONTROL_CONTROLTYPE_SIGNEDMETER = _
    (MIXERCONTROL_CT_CLASS_METER Or _
    MIXERCONTROL_CT_SC_METER_POLLED Or MIXERCONTROL_CT_UNITS_SIGNED)

Private Const MIXERCONTROL_CONTROLTYPE_BOOLEAN = _
    (MIXERCONTROL_CT_CLASS_SWITCH Or _
    MIXERCONTROL_CT_SC_SWITCH_BOOLEAN Or MIXERCONTROL_CT_UNITS_BOOLEAN)

Private Const MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT = _
    (MIXERCONTROL_CT_CLASS_LIST Or _
    MIXERCONTROL_CT_SC_LIST_MULTIPLE Or MIXERCONTROL_CT_UNITS_BOOLEAN)

Private Const MIXERCONTROL_CONTROLTYPE_FADER = _
    (MIXERCONTROL_CT_CLASS_FADER Or _
    MIXERCONTROL_CT_UNITS_UNSIGNED)
      
Private Const MIXERCONTROL_CONTROLTYPE_VOLUME = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 1)

Private Const MIXERCONTROL_CONTROLTYPE_BASS = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 2)
    
Private Const MIXERCONTROL_CONTROLTYPE_EQUALIZER = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 4)
    
Private Const MIXERCONTROL_CONTROLTYPE_MIXER = _
    (MIXERCONTROL_CONTROLTYPE_MULTIPLESELECT + 1)
    
Private Const MIXERCONTROL_CONTROLTYPE_MONO = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 3)
    
Private Const MIXERCONTROL_CONTROLTYPE_MUTE = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 2)
    
Private Const MIXERCONTROL_CONTROLTYPE_ONOFF = _
    (MIXERCONTROL_CONTROLTYPE_BOOLEAN + 1)
    
Private Const MIXERCONTROL_CONTROLTYPE_PEAKMETER = _
    (MIXERCONTROL_CONTROLTYPE_SIGNEDMETER + 1)
    
Private Const MIXERCONTROL_CONTROLTYPE_TREBLE = _
    (MIXERCONTROL_CONTROLTYPE_FADER + 3)

Private Declare Function mixerClose Lib "winmm.dll" _
    (ByVal hmx As Long) As Long
         
Private Declare Function mixerGetControlDetails Lib "winmm.dll" _
    Alias "mixerGetControlDetailsA" _
    (ByVal hmxobj As Long, _
    pMxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long) As Long
         
Private Declare Function mixerGetDevCaps Lib "winmm.dll" _
    Alias "mixerGetDevCapsA" _
    (ByVal uMxId As Long, _
    ByVal pmxcaps As MIXERCAPS, _
    ByVal cbmxcaps As Long) As Long
         
Private Declare Function mixerGetID Lib "winmm.dll" _
    (ByVal hmxobj As Long, _
    pumxID As Long, _
    ByVal fdwId As Long) As Long
                     
Private Declare Function mixerGetLineControls Lib "winmm.dll" _
    Alias "mixerGetLineControlsA" _
    (ByVal hmxobj As Long, _
    pmxlc As MIXERLINECONTROLS, _
    ByVal fdwControls As Long) As Long
                     
Private Declare Function mixerGetLineInfo Lib "winmm.dll" _
    Alias "mixerGetLineInfoA" _
    (ByVal hmxobj As Long, _
    pmxl As MIXERLINE, _
    ByVal fdwInfo As Long) As Long
                     
Private Declare Function mixerGetNumDevs Lib "winmm.dll" () As Long
      
Private Declare Function mixerMessage Lib "winmm.dll" _
    (ByVal hmx As Long, _
    ByVal uMsg As Long, _
    ByVal dwParam1 As Long, _
    ByVal dwParam2 As Long) As Long
                     
Private Declare Function mixerOpen Lib "winmm.dll" _
    (phmx As Long, _
    ByVal uMxId As Long, _
    ByVal dwCallback As Long, _
    ByVal dwInstance As Long, _
    ByVal fdwOpen As Long) As Long
                     
Private Declare Function mixerSetControlDetails Lib "winmm.dll" _
    (ByVal hmxobj As Long, _
    pMxcd As MIXERCONTROLDETAILS, _
    ByVal fdwDetails As Long) As Long
              
Private Declare Sub CopyStructFromPtr Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (struct As Any, _
    ByVal ptr As Long, ByVal cb As Long)
                     
Private Declare Sub CopyPtrFromStruct Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (ByVal ptr As Long, _
    struct As Any, _
    ByVal cb As Long)
                     
Private Declare Function GlobalAlloc Lib "kernel32" _
    (ByVal wFlags As Long, _
    ByVal dwBytes As Long) As Long
                     
Private Declare Function GlobalLock Lib "kernel32" _
    (ByVal hmem As Long) As Long
                     
Private Declare Function GlobalFree Lib "kernel32" _
    (ByVal hmem As Long) As Long
      
Private Type MIXERCAPS
    wMid As Integer                   '  manufacturer id
    wPid As Integer                   '  product id
    vDriverVersion As Long            '  version of the driver
    szPname As String * MAXPNAMELEN   '  product name
    fdwSupport As Long                '  misc. support bits
    cDestinations As Long             '  count of destinations
End Type
      
Private Type MIXERCONTROL
    cbStruct As Long                  '  size in Byte of MIXERCONTROL
    dwControlID As Long               '  unique control id for mixer device
    dwControlType As Long             '  MIXERCONTROL_CONTROLTYPE_xxx
    fdwControl As Long                '  MIXERCONTROL_CONTROLF_xxx
    cMultipleItems As Long            '  if MIXERCONTROL_CONTROLF_MULTIPLE set
    szShortName As String * MIXER_SHORT_NAME_CHARS  ' short name of control
    szName As String * MIXER_LONG_NAME_CHARS        ' long name of control
    lMinimum As Long                  '  Minimum value
    lMaximum As Long                  '  Maximum value
    Reserved(10) As Long              '  reserved structure space
End Type
      
Private Type MIXERCONTROLDETAILS
    cbStruct As Long       '  size in Byte of MIXERCONTROLDETAILS
    dwControlID As Long    '  control id to get/set details on
    cChannels As Long      '  number of channels in paDetails array
    item As Long           '  hwndOwner or cMultipleItems
    cbDetails As Long      '  size of _one_ details_XX struct
    paDetails As Long      '  pointer to array of details_XX structs
End Type
      
Private Type MIXERCONTROLDETAILS_UNSIGNED
    dwValue As Long        '  value of the control
End Type

Private Type MIXERCONTROLDETAILS_BOOLEAN
    fValue As Long
End Type

Private Type MIXERLINE
    cbStruct As Long               '  size of MIXERLINE structure
    dwDestination As Long          '  zero based destination index
    dwSource As Long               '  zero based source index (if source)
    dwLineID As Long               '  unique line id for mixer device
    fdwLine As Long                '  state/information about line
    dwUser As Long                 '  driver specific information
    dwComponentType As Long        '  component type line connects to
    cChannels As Long              '  number of channels line supports
    cConnections As Long           '  number of connections (possible)
    cControls As Long              '  number of controls at this line
    szShortName As String * MIXER_SHORT_NAME_CHARS
    szName As String * MIXER_LONG_NAME_CHARS
    dwType As Long
    dwDeviceID As Long
    wMid  As Integer
    wPid As Integer
    vDriverVersion As Long
    szPname As String * MAXPNAMELEN
End Type
      
Private Type MIXERLINECONTROLS
    cbStruct As Long       '  size in Byte of MIXERLINECONTROLS
    dwLineID As Long       '  line id (from MIXERLINE.dwLineID)
                                '  MIXER_GETLINECONTROLSF_ONEBYID or
    dwControl As Long      '  MIXER_GETLINECONTROLSF_ONEBYTYPE
    cControls As Long      '  count of controls pmxctrl points to
    cbmxctrl As Long       '  size in Byte of _one_ MIXERCONTROL
    pamxctrl As Long       '  pointer to first MIXERCONTROL array
End Type

Private hmixer As Long          ' mixer handle

Private volCtrl As MIXERCONTROL ' waveout volume control
Private micCtrl As MIXERCONTROL ' Microphon
Private linCtrl As MIXERCONTROL ' Line
Private cdCtrl As MIXERCONTROL  ' CD
Private synCtrl As MIXERCONTROL ' Synthisizer
Private wavCtrl As MIXERCONTROL ' Wave
'Mute
Private volMute As MIXERCONTROL ' waveout volume control
Private micMute As MIXERCONTROL ' Microphon
Private linMute As MIXERCONTROL ' Line
Private cdMute As MIXERCONTROL  ' CD
Private synMute As MIXERCONTROL ' Synthisizer
Private wavMute As MIXERCONTROL ' Wave

Private rc As Long              ' return code
Private ok As Boolean           ' boolean return code

' Audiolevel---------------------------------------------------------
Private Const CALLBACK_FUNCTION = &H30000
Private Const MM_WIM_DATA = &H3C0
Private Const WHDR_DONE = &H1         '  done bit
Private Const GMEM_FIXED = &H0         ' Global Memory Flag used by GlobalAlloc functin

Private Type WAVEHDR
' The WAVEHDR user-defined type defines the header used to identify a waveform-audio buffer.
   lpData As Long          ' Address of the waveform buffer.
   dwBufferLength As Long  ' Length, in bytes, of the buffer.
   dwBytesRecorded As Long ' When the header is used in input, this member specifies how much
                           ' data is in the buffer.

   dwUser As Long          ' User data.
   dwFlags As Long         ' Flags supplying information about the buffer. Set equal to zero.
   dwLoops As Long         ' Number of times to play the loop. Set equal to zero.
   lpNext As Long          ' Not used
   Reserved As Long        ' Not used
End Type

Private Type WAVEINCAPS
' The WAVEINCAPS user-defined variable describes the capabilities of a waveform-audio input
' device.
   wMid As Integer         ' Manufacturer identifier for the device driver for the
                           ' waveform-audio input device. Manufacturer identifiers
                           ' are defined in Manufacturer and Product Identifiers in
                           ' the Platform SDK product documentation.
   wPid As Integer         ' Product identifier for the waveform-audio input device.
                           ' Product identifiers are defined in Manufacturer and Product
                           ' Identifiers in the Platform SDK product documentation.
   vDriverVersion As Long  ' Version number of the device driver for the
                           ' waveform-audio input device. The high-order byte
                           ' is the major version number, and the low-order byte
                           ' is the minor version number.
   szPname As String * 32  ' Product name in a null-terminated string.
   dwFormats As Long       ' Standard formats that are supported. See the Platform
                           ' SDK product documentation for more information.
   wChannels As Integer    ' Number specifying whether the device supports
                           ' mono (1) or stereo (2) input.
End Type

Private Type WAVEFORMAT
' The WAVEFORMAT user-defined type describes the format of waveform-audio data. Only
' format information common to all waveform-audio data formats is included in this
' user-defined type.
   wFormatTag As Integer      ' Format type. Use the constant WAVE_FORMAT_PCM Waveform-audio data
                              ' to define the data as PCM.
   nChannels As Integer       ' Number of channels in the waveform-audio data. Mono data uses one
                              ' channel and stereo data uses two channels.
   nSamplesPerSec As Long     ' Sample rate, in samples per second.
   nAvgBytesPerSec As Long    ' Required average data transfer rate, in bytes per second. For
                              ' example, 16-bit stereo at 44.1 kHz has an average data rate of
                              ' 176,400 bytes per second (2 channels — 2 bytes per sample per
                              ' channel — 44,100 samples per second).
   nBlockAlign As Integer     ' Block alignment, in bytes. The block alignment is the minimum atomic unit of data. For PCM data, the block alignment is the number of bytes used by a single sample, including data for both channels if the data is stereo. For example, the block alignment for 16-bit stereo PCM is 4 bytes (2 channels — 2 bytes per sample).
   wBitsPerSample As Integer  ' For buffer estimation
   cbSize As Integer          ' Block size of the data.
End Type

Private Declare Function waveInOpen Lib "winmm.dll" (lphWaveIn As Long, _
                                             ByVal uDeviceID As Long, _
                                             lpFormat As WAVEFORMAT, _
                                             ByVal dwCallback As Long, _
                                             ByVal dwInstance As Long, _
                                             ByVal dwFlags As Long) As Long
' The waveInOpen function opens the given waveform-audio input device for recording. The function
' uses the following parameters
'     lphWaveIn-  a long value that is the handle identifying the open waveform-audio input
'                 device. Use this handle to identify the device when calling other
'                 waveform-audio input functions. This parameter can be NULL if WAVE_FORMAT_QUERY
'                 is specified for fdwOpen.
'     uDeviceID-  a long value that identifies the waveform-audio input device to open.
'                 This parameter can be either a device identifier or a handle of an open
'                 waveform-audio input device.
'     lpFormat-   the WAVEFORMAT user-defined typed that identifies the desired format for
'                 recording waveform-audio data.
'     dwCallback- a long value that is an event handle, a handle to a window, or the identifier
'                 of a thread to be called during waveform-audio recording to process messages
'                 related to the progress of recording. If no callback function is required,
'                 this value can be zero. For more information on the callback function,
'                 see waveInProc.
'     dwCallback- a long value that is the user-instance data passed to the callback mechanism.
'                 This parameter is not used with the window callback mechanism.
'     dwFlags-    Flags for opening the device. The following values are defined:
'                 CALLBACK_EVENT (&H50000)-event handle.
'                 CALLBACK_FUNCTION (&H30000)-callback procedure address.
'                 CALLBACK_NULL (&H00000)-No callback mechanism. This is the default setting.
'                 CALLBACK_THREAD (&H20000)-thread identifier.
'                 CALLBACK_WINDOW (&H10000)-window handle.
'                 WAVE_FORMAT_DIRECT (&H8)-ACM driver does not perform conversions on the
'                                            audio data.
'                 WAVE_FORMAT_QUERY (&H1)-queries the device to determine whether it supports
'                                         the given format, but it does not open the device.
'                 WAVE_MAPPED (&H4)-The uDeviceID parameter specifies a waveform-audio device
'                                   to be mapped to by the wave mapper.

Private Declare Function waveInPrepareHeader Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                      lpWaveInHdr As WAVEHDR, _
                                                      ByVal uSize As Long) As Long
' The waveInPrepareHeader function prepares a buffer for waveform-audio input. The function
' uses the following parameters:
'     hWaveIn-    a long value that is the handle of the waveform-audio input device.
'     lpWaveInHdr-the WAVEHDR user-defined type variable.
'     uSize-      the size in bytes of the WAVEHDR user-defined type variable. Use the
'                 results of the Len function for this parameter.


Private Declare Function waveInReset Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInReset function stops input on the given waveform-audio input device and resets
' the current position to zero. All pending buffers are marked as done and returned to
' the application. The function requires the handle to the waveform-audio input device.

Private Declare Function waveInStart Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInStart function starts input on the given waveform-audio input device. The function
' requires the handle of the waveform-audio input device.

Private Declare Function waveInStop Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInStop function stops waveform-audio input. The function requires the handle of
' the waveform-audio input device.

Private Declare Function waveInUnprepareHeader Lib "winmm.dll" _
                                          (ByVal hWaveIn As Long, _
                                          lpWaveInHdr As WAVEHDR, _
                                          ByVal uSize As Long) As Long
' The waveInUnprepareHeader function cleans up the preparation performed by the
' waveInPrepareHeader function. This function must be called after the device driver
' fills a buffer and returns it to the application. You must call this function before
' freeing the buffer. The function uses the following parameters:
'     hWaveIn-       a long value that is the handle of the waveform-audio input device.
'     lpWaveInHdr-   the variable typed as the WAVEHDR user-defined type identifying the
'                    buffer to be cleaned up.
'     uSize-         a long value that is the size in bytes, of the WAVEHDR varaible. Use
'                    the Len function with the WAVEHDR variable as the argument to get this
'                    value.

Private Declare Function waveInClose Lib "winmm.dll" (ByVal hWaveIn As Long) As Long
' The waveInClose function closes the given waveform-audio input device. The function
' requires the handle of the waveform-audio input device. If the function succeeds,
' the handle is no longer valid after this call.

Private Declare Function waveInGetDevCaps Lib "winmm.dll" Alias "waveInGetDevCapsA" _
                  (ByVal uDeviceID As Long, _
                  lpCaps As WAVEINCAPS, _
                  ByVal uSize As Long) As Long
' This function retrieves the capabilities of a given waveform-audio input device. You can
' use this function to determine the number of waveform-audio input devices present in the
' system. If the value specified by the uDeviceID parameter is a device identifier,
' it can vary from zero to one less than the number of devices present. The function uses
' the following parameters
'     uDeviceID-     long value that identifies waveform-audio output device. This value can be
'                    either a device identifier or a handle of an open waveform-audio input device.
'     lpCaps-user-   defined variable containing information about the capabilities of the device.
'     uSize-         the size in bytes of the user-defined variable used as the lpCaps parameter.
'                    Use the Len function to get this value.

Private Declare Function waveInGetNumDevs Lib "winmm.dll" () As Long
' The waveInGetNumDevs function returns the number of waveform-audio input devices present in the system.

Private Declare Function waveInGetErrorText Lib "winmm.dll" Alias "waveInGetErrorTextA" _
                     (ByVal err As Long, _
                     ByVal lpText As String, _
                     ByVal uSize As Long) As Long
'The waveInGetErrorText function retrieves a textual description of the error identified by
' the given error number. The function uses the following parameters:
'     Err-     a long value that is the error number.
'     lpText-  a string variable that contains the textual error description.
'     uSize-   the size in characters of the lpText string variable.

Private Declare Function waveInAddBuffer Lib "winmm.dll" (ByVal hWaveIn As Long, _
                                                   lpWaveInHdr As WAVEHDR, _
                                                   ByVal uSize As Long) As Long
' The waveInAddBuffer function sends an input buffer to the given waveform-audio input device.
' The function uses the following parameters:
'     hWaveIn-       a long value that is the handle of the waveform-audio input device.
'     lpWaveInHdr-   the variable typed as the WAVEHDR user-defined type.
'     uSize-         a long value that is the size in bytes of the variable typed as the
'                    WAVEHDR user-defined variable. Use the Len function with the WAVEHDR
'                    variable as the argument to get this value.




' The mixerClose function closes the specified mixer device. The function requires the
' handle of the mixer device. This handle must have been returned successfully by the
' mixerOpen function. If mixerClose is successful, the handle is no longer valid.

' The mixerGetControlDetails function retrieves details about a single control associated
' with an audio line. the function uses the following parameters:
'     hmxobj-     a long value that is the handle to the mixer device object being queried.
'     pMxcd-      the variable defined as the MIXERCONTROLDETAILS user-defined type.
'     fdwDetails- Flags for retrieving control details. The following values are defined:
'                    MIXER_GETCONTROLDETAILSF_LISTTEXT-The paDetails member of the
'                       MIXERCONTROLDETAILS user-defined variable points to one or more
'                       MIXERCONTROLDETAILS_LISTTEXT user-defined variables to receive text
'                       labels for multiple-item controls. An application must get all list
'                       text items for a multiple-item control at once. This flag cannot be
'                       used with MIXERCONTROL_CONTROLTYPE_CUSTOM controls.
'                    MIXER_GETCONTROLDETAILSF_VALUE-Current values for a control are
'                       retrieved. The paDetails member of the MIXERCONTROLDETAILS user-defined
'                       variable points to one or more details appropriate for the control class.
'                    MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-
'                       The hmxobj parameter is the handle of a MIDI (Musical Instrument Digital
'                       Interface) input device. This handle must have been returned by the
'                       midiInOpen function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The
'                       hmxobj parameter is a mixer device handle returned by the mixerOpen
'                       function. This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen function.
'                    MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000L)-The hmxobj parameter is the identifier
'                       of a MIDI input device. This identifier must be in the range of zero to
'                       one less than the number of devices returned by the midiInGetNumDevs
'                       function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier
'                       of a MIDI output device. This identifier must be in the range of zero
'                       to one less than the number of devices returned by the midiOutGetNumDevs
'                       function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is the identifier of a
'                       mixer device in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs function.


' The mixerGetDevCaps function queries a specified mixer device to determine its capabilities.
' The function uses the following parameters:
'     uMxId-      a long value that is the handle of an open mixer device.
'     pmxcaps-    a variable defined as the MIXERCAPS user-defined type to contain information
'                 about the capabilities of the device.
'     cbmxcaps-   a long value that is the size in bytes, of the variable defined as the
'                 MIXERCAPS user-defined type. Use the Len functions with the MIXERCAPS variable
'                 as the argument to get this value.


' The mixerGetID function retrieves the device identifier for a mixer device associated
' with a specified device handle.The function uses the following parameters:
'     hmxobj-  a long value that is the handle of the audio mixer object to map to a
'              mixer device identifier.
'     pumxID-  the long value to contain the mixer device identifier. If no mixer device
'              is available for the hmxobj object, the value – 1 is placed in this location
'              and the MMSYSERR_NODRIVER error value is returned.
'     fdwId-   Flags for mapping the mixer object hmxobj. The following values are defined:
'                 MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs function.
'                 MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-
'                       the hmxobj parameter is the handle of a MIDI input device. This handle
'                       must have been returned by the midiInOpen function.
'                 MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen function.
'                 MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The hmxobj
'                       parameter is a mixer device handle returned by the mixerOpen function.
'                       This flag is optional.
'                 MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen function.
'                 MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen function.
'                 MIXER_OBJECTF_MIDIIN (&H40000000L)-The hmxobj parameter is the identifier of
'                       a MIDI input device. This identifier must be in the range of zero to
'                       one less than the number of devices returned by the midiInGetNumDevs
'                       function.
'                 MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier of
'                       a MIDI output device. This identifier must be in the range of zero to
'                       one less than the number of devices returned by the midiOutGetNumDevs
'                       function.
'                 MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is the identifier of
'                       a mixer device in the range of zero to one less than the number of
'                       devices returned by the mixerGetNumDevs function. This flag is optional.
'                 MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs function.
'                 MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs function.


' The mixerGetLineControls function retrieves one or more controls associated with an audio
' line. The function uses the following parameters:
'     hmxobj-        a long value that is the handle of the mixer device object that is being
'                    queried.
'     pmxlc-         the variable defined as the MIXERLINECONTROLS user-defined type used to
'                    reference one or more variables defined as theMIXERCONTROL user-defined
'                    types to be filled with information about the controls associated with
'                    an audio line. The cbStruct member of the MIXERLINECONTROLS variable
'                    must always be initialized to be the size, in bytes, of the
'                    MIXERLINECONTROLS variable.
'     fdwControls-   Flags for retrieving information about one or more controls associated w
'                    with an audio line. The following values are defined:
'                    MIXER_GETLINECONTROLSF_ALL-The pmxlc parameter references a list of
'                       MIXERCONTROL variables that will receive information on all controls
'                       associated with the audio line identified by the dwLineID member of
'                       the MIXERLINECONTROLS structure. The cControls member must be initialized
'                       to the number of controls associated with the line. This number is
'                       retrieved from the cControls member of the MIXERLINE structure returned
'                       by the mixerGetLineInfo function. The cbmxctrl member must be
'                       initialized to the size, in bytes, of a single MIXERCONTROL variable.
'                       The pamxctrl member must point to the first MIXERCONTROL variable to be
'                       filled. The dwControlID and dwControlType members are ignored for this
'                       query.
'                    MIXER_GETLINECONTROLSF_ONEBYID-The pmxlc parameter references a single
'                       MIXERCONTROL variable that will receive information on the control
'                       identified by the dwControlID member of the MIXERLINECONTROLS variable.
'                       The cControls member must be initialized to 1. The cbmxctrl member must
'                       be initialized to the size, in bytes, of a single MIXERCONTROL variable.
'                       The pamxctrl member must point to a MIXERCONTROL structure to be filled.
'                       The dwLineID and dwControlType members are ignored for this query. This
'                       query is usually used to refresh a control after receiving a
'                       MM_MIXM_CONTROL_CHANGE control change notification message by the
'                       user-defined callback (see mixerOpen).
'                    MIXER_GETLINECONTROLSF_ONEBYTYPE-The mixerGetLineControls function
'                       retrieves information about the first control of a specific class for
'                       the audio line that is being queried. The pmxlc parameter references a
'                       single MIXERCONTROL structure that will receive information about the
'                       specific control. The audio line is identified by the dwLineID member.
'                       The control class is specified in the dwControlType member of the
'                       MIXERLINECONTROLS variable. The dwControlID member is ignored for this
'                       query. This query can be used by an application to get information on
'                       a single control associated with a line. For example, you might want
'                       your application to use a peak meter only from a waveform-audio output
'                       line.
'                    MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-The
'                       hmxobj parameter is the handle of a MIDI input device. This handle must
'                       have been returned by the midiInOpen function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The
'                       hmxobj parameter is a mixer device handle returned by the mixerOpen
'                       function. This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen function.
'                    MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000L)-The hmxobj parameter is the identifier of
'                       a MIDI input device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiInGetNumDevs function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier of
'                       a MIDI output device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiOutGetNumDevs function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is the identifier of a
'                       mixer device in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs function.


' The mixerGetLineInfo function retrieves information about a specific line of a mixer device.
' Uses the same parameters and constants as the mixerGetLineControls function.


' The mixerGetNumDevs function retrieves the number of mixer devices present in the system.


' The mixerMessage function sends a custom mixer driver message directly to a mixer driver.
' The function uses the following parameters:
'     hmx-     a long value that is the handle of an open instance of a mixer device. This
'              value is the result of the mixerOpen function.
'     uMsg-    Custom mixer driver message to send to the mixer driver. This message must
'              be above or equal to the MXDM_USER constant.
'     dwParam1 and dwParam2-Arguments associated with the message being sent.


' The mixerOpen function opens a specified mixer device and ensures that the device will
' not be removed until the application closes the handle. the function uses the following
' parameters:
'     phmx-       a long value that is the handle identifying the opened mixer device. Use
'                 this handle to identify the device when calling other audio mixer functions.
'                 This parameter cannot be NULL.
'     uMxId-      a long value that identifies the mixer device to open. Use a valid device
'                 identifier or any HMIXEROBJ (see the mixerGetID function for a description of
'                 mixer object handles). A "mapper" for audio mixer devices does not currently
'                 exist, so a mixer device identifier of – 1 is not valid.
'     dwCallback- Handle of a window called when the state of an audio line and/or control
'                 associated with the device being opened is changed. Specify zero for this
'                 parameter if no callback mechanism is to be used.
'     dwInstance- User instance data passed to the callback function. This parameter is not
'                 used with window callback functions.
'     fdwOpen-    Flags for opening the device. The following values are defined:
'                    CALLBACK_WINDOW-  The dwCallback parameter is assumed to be a window handle.
'                    MIXER_OBJECTF_AUX (&H50000000)-The uMxId parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-the
'                       uMxId parameter is the handle of a MIDI input device. This handle must
'                       have been returned by the midiInOpen function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       uMxId parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The uMxId
'                       parameter is a mixer device handle returned by the mixerOpen function.
'                       This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       uMxId parameter is a waveform-audio input handle returned by the
'                       waveInOpen function.
'                    MIXER_OBJECTF_HWAVEOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       uMxId parameter is a waveform-audio output handle returned by the
'                       waveOutOpen function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000)-The uMxId parameter is the identifier of
'                       a MIDI input device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiInGetNumDevs function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The uMxId parameter is the identifier of
'                       a MIDI output device. This identifier must be in the range of zero to one
'                       less than the number of devices returned by the midiOutGetNumDevs function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The uMxId parameter is a mixer device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The uMxId parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The uMxId parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs function.


' The mixerSetControlDetails function sets properties of a single control associated with an
' audio line. The function uses the following parameters
'     hmxobj-        a long value that is the handle of the mixer device object for which
'                    properties are being set.
'     pMxcd-         the variable declares as the MIXERCONTROLDETAILS user-defined type.
'                    This variable references the control detail structures that contain the
'                    desired state for the control.
'     fdwDetails-    Flags for setting properties for a control. The following values are
'                    defined:
'                    MIXER_OBJECTF_AUX (&H50000000)-The hmxobj parameter is an auxiliary device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the auxGetNumDevs function.
'                    MIXER_OBJECTF_HMIDIIN (MIXER_OBJECTF_HANDLE or MIXER_OBJECTF_MIDIIN)-
'                       The hmxobj parameter is the handle of a MIDI input device. This handle
'                       must have been returned by the midiInOpen function.
'                    MIXER_OBJECTF_HMIDIOUT (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIDIOUT)-The
'                       hmxobj parameter is the handle of a MIDI output device. This handle must
'                       have been returned by the midiOutOpen function.
'                    MIXER_OBJECTF_HMIXER (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_MIXER)-The hmxobj
'                       parameter is a mixer device handle returned by the mixerOpen function.
'                       This flag is optional.
'                    MIXER_OBJECTF_HWAVEIN (MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEIN)-The
'                       hmxobj parameter is a waveform-audio input handle returned by the
'                       waveInOpen function.
'                    MIXER_OBJECTF_HWAVEOUT ((MIXER_OBJECTF_HANDLE Or MIXER_OBJECTF_WAVEOUT)-The
'                       hmxobj parameter is a waveform-audio output handle returned by the
'                       waveOutOpen function.
'                    MIXER_OBJECTF_MIDIIN (&H40000000)-The hmxobj parameter is the identifier
'                       of a MIDI inputdevice. This identifier must be in the range of zero to
'                        one less than the number of devices returned by the midiInGetNumDevs
'                        function.
'                    MIXER_OBJECTF_MIDIOUT (&H30000000)-The hmxobj parameter is the identifier
'                       of a MIDI output device. This identifier must be in the range of zero
'                       to one less than the number of devices returned by the midiOutGetNumDevs
'                       function.
'                    MIXER_OBJECTF_MIXER (&H00000000)-The hmxobj parameter is a mixer device
'                       identifier in the range of zero to one less than the number of devices
'                       returned by the mixerGetNumDevs function. This flag is optional.
'                    MIXER_OBJECTF_WAVEIN (&H20000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio input device in the range of zero to one less than the
'                       number of devices returned by the waveInGetNumDevs function.
'                    MIXER_OBJECTF_WAVEOUT (&H10000000)-The hmxobj parameter is the identifier of a
'                       waveform-audio output device in the range of zero to one less than the
'                       number of devices returned by the waveOutGetNumDevs function.
'                    MIXER_SETCONTROLDETAILSF_CUSTOM-A custom dialog box for the specified
'                       custom mixer control is displayed. The mixer device gathers the required
'                       information from the user and returns the data in the specified buffer.
'                       The handle for the owning window is specified in the hwndOwner member
'                       of the MIXERCONTROLDETAILS structure. (This handle can be set to NULL.)
'                       The application can then save the data from the dialog box and use it
'                       later to reset the control to the same state by using the
'                       MIXER_SETCONTROLDETAILSF_VALUE flag.
'                    MIXER_SETCONTROLDETAILSF_VALUE (&H00000000)-The current value(s) for a control
'                       are set. The paDetails member of the MIXERCONTROLDETAILS structure points
'                       to one or more mixer-control details structures of the appropriate class for
'                       the control.



' The CopyStructFromPtr and CopyPtrFromStruct functions are user-defined versions of the
' RtlMoveMemory function. RtlMoveMemory moves memory either forward or backward, aligned or
' unaligned, in 4-byte blocks, followed by any remaining bytes. The function requires the
' following parameters:
'     Destination-   Pointer to the starting address of the copied block's destination.
'     Source-        Pointer to the starting address of the block of memory to copy.
'     Length-        Specifies the size, in bytes, of the block of memory to copy.


' The GlobalAlloc function allocates the specified number of bytes from the heap.
' Win32 memory management does not provide a separate local heap and global heap.
' This function is provided only for compatibility with 16-bit versions of Windows. The function
' uses the following parameters:
'     wFlags-     a long value that specifies how to allocate memory. If zero is specified,
'                 the default is allocate fixed memory. This value can be one or more of the
'                 following flags:
'                    GMEM_FIXED (&H0)- Allocates fixed memory. The return value is a pointer.
'                    GMEM_MOVEABLE (&H2)- Allocates movable memory. In Win32, memory blocks are
'                       never moved in physical memory, but they can be moved within the default .
'                       The return value is the handle of the memory object. To translate the
'                       heap handle into a pointer, use the GlobalLock function. This flag
'                       cannot be combined with the GMEM_FIXED flag.
'                    GPTR (GMEM_FIXED Or GMEM_ZEROINIT)-Combines the GMEM_FIXED and GMEM_ZEROINIT
'                       flags.
'                    GHND (GMEM_MOVEABLE Or GMEM_ZEROINIT)- Combines the GMEM_MOVEABLE and
'                       GMEM_ZEROINIT flags.
'                    GMEM_ZEROINIT (&H4)-Initializes memory contents to zero.
'     dwBytes-    Specifies the number of bytes to allocate. If this parameter is zero and
'                 the uFlags parameter specifies the GMEM_MOVEABLE flag, the function returns
'                 a handle to a memory object that is marked as discarded.


' The GlobalLock function locks a global memory object and returns a pointer to the first
' byte of the object's memory block. This function is provided only for compatibility with
' 16-bit versions of Windows. The function requires a handle to the global memory object. This
' handle is returned by either the GlobalAlloc or GlobalReAlloc function.


' The GlobalFree function frees the specified global memory object and invalidates its handle.
' This function is provided only for compatibility with 16-bit versions of Windows. The function
' requires a h andle to the global memory object. This handle is returned by either the
' GlobalAlloc or GlobalReAlloc function.

Private Type MIXERCONTROLDETAILS_SIGNED
' The MIXERCONTROLDETAILS_SIGNED user-defined type retrieves and
' sets signed type control
' properties for an audio mixer control.
   lValue As Long
End Type

Private i As Integer
Private j As Integer
Private msg As String * 200
Private hWaveIn As Long
Private format As WAVEFORMAT

Private Const NUM_BUFFERS = 2
Private Const BUFFER_SIZE = 8192
Private Const DEVICEID = 0
Private hmem(NUM_BUFFERS) As Long
Private inHdr(NUM_BUFFERS) As WAVEHDR

Private fRecording As Boolean

Private inputVolCtrl As MIXERCONTROL    ' waveout volume control
Private outputVolCtrl As MIXERCONTROL   ' microphone volume control

Private vol As MIXERCONTROLDETAILS_SIGNED   ' control's signed value
Private volume As Long                      ' volume value
Private volHmem As Long                     ' handle to volume memory
Private mxcd As MIXERCONTROLDETAILS         ' control info

Private Sub Class_Initialize()

    
    rc = mixerOpen(hmixer, 0, 0, 0, 0)
    If ((MMSYSERR_NOERROR <> rc)) Then
        Exit Sub
    End If
             
    ' Get the waveout volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, _
        volCtrl)
    
    ' Get the microphone volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, _
        micCtrl)
    
    ' Get tne Line volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_LINE, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, _
        linCtrl)
    
    ' Get tne CD volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, _
        cdCtrl)

    ' Get tne Synthesizer volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, _
        synCtrl)

    ' Get tne Wave volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
        MIXERCONTROL_CONTROLTYPE_VOLUME, _
        wavCtrl)
    
' Mute
    ' Get the waveout volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
        MIXERCONTROL_CONTROLTYPE_MUTE, _
        volMute)
    
    ' Get the microphone volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, _
        MIXERCONTROL_CONTROLTYPE_MUTE, _
        micMute)
    
    ' Get tne Line volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_LINE, _
        MIXERCONTROL_CONTROLTYPE_MUTE, _
        linMute)
    
    ' Get tne CD volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_COMPACTDISC, _
        MIXERCONTROL_CONTROLTYPE_MUTE, _
        cdMute)

    ' Get tne Synthesizer volume control
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_SYNTHESIZER, _
        MIXERCONTROL_CONTROLTYPE_MUTE, _
        synMute)

    ' Wave
    ok = GetVolumeControl(hmixer, _
        MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
        MIXERCONTROL_CONTROLTYPE_MUTE, _
        wavMute)

'Audiolevel--------------------------------------------------------
   ' Open the mixer specified by DEVICEID
   rc = mixerOpen(hmixer, DEVICEID, 0, 0, 0)
   
   If ((MMSYSERR_NOERROR <> rc)) Then
       Exit Sub
   End If
       
       
   ' Initialize mixercontrol structure
   mxcd.cbStruct = Len(mxcd)
   volHmem = GlobalAlloc(&H0, Len(volume))  ' Allocate a buffer for the volume value
   mxcd.paDetails = GlobalLock(volHmem)
   mxcd.cbDetails = Len(volume)
   mxcd.cChannels = 1

   ' Get the input volume meter
   ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_DST_WAVEIN, _
   MIXERCONTROL_CONTROLTYPE_PEAKMETER, inputVolCtrl)
   
   If (ok <> True) Then
       ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_MICROPHONE, MIXERCONTROL_CONTROLTYPE_PEAKMETER, inputVolCtrl)
   End If

   ' Get the output volume meter
   ok = GetVolumeControl(hmixer, MIXERLINE_COMPONENTTYPE_SRC_WAVEOUT, _
   MIXERCONTROL_CONTROLTYPE_PEAKMETER, outputVolCtrl)
   

End Sub

Private Sub Class_Terminate()
   If (fRecording = True) Then
       StopInput
   End If
   GlobalFree volHmem

End Sub


Private Function SetVolumeControl(ByVal hmixer As Long, _
                              mxc As MIXERCONTROL, _
                              ByVal volume As Long) As Boolean
    'This function sets the value for a volume control. Returns True if successful
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(vol)
         
    ' Allocate a buffer for the control value buffer
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    vol.dwValue = volume
         
    ' Copy the data into the control value buffer
    CopyPtrFromStruct mxcd.paDetails, vol, Len(vol)
         
    ' Set the control value
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
         
    GlobalFree (hmem)
    If (MMSYSERR_NOERROR = rc) Then
        SetVolumeControl = True
    Else
        SetVolumeControl = False
    End If
End Function

Private Function Wert(mxc As MIXERCONTROL) As Long
    'This function gets the value for a volume control.
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim vol As MIXERCONTROLDETAILS_UNSIGNED
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(vol)
    
    hmem = GlobalAlloc(&H40, Len(vol))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
         
    ' Get the control value
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    
    ' Copy the data into the control value buffer
    CopyStructFromPtr vol, mxcd.paDetails, Len(vol)
    
    Wert = vol.dwValue
    
    GlobalFree (hmem)
    
End Function

Private Function GetMute(mxc As MIXERCONTROL) As Boolean
    'This function gets the value for a volume control.
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Mut As MIXERCONTROLDETAILS_BOOLEAN
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Mut)
    
    hmem = GlobalAlloc(&H40, Len(Mut))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
         
    ' Get the control value
    rc = mixerGetControlDetails(hmixer, _
        mxcd, _
        MIXER_GETCONTROLDETAILSF_VALUE)
    
    ' Copy the data into the control value buffer
    CopyStructFromPtr Mut, mxcd.paDetails, Len(Mut)
    
    GetMute = IIf((Mut.fValue = 1), True, False)
    
    GlobalFree (hmem)
    
End Function
Private Function SetMute(mxc As MIXERCONTROL, _
                              ByVal Mute As Boolean)
    'This function sets the value for a volume control. Returns True if successful
                              
    Dim mxcd As MIXERCONTROLDETAILS
    Dim Mut As MIXERCONTROLDETAILS_BOOLEAN
    Dim hmem As Long
      
    mxcd.item = 0
    mxcd.dwControlID = mxc.dwControlID
    mxcd.cbStruct = Len(mxcd)
    mxcd.cbDetails = Len(Mut)
         
    ' Allocate a buffer for the control value buffer
    hmem = GlobalAlloc(&H40, Len(Mut))
    mxcd.paDetails = GlobalLock(hmem)
    mxcd.cChannels = 1
    Mut.fValue = IIf(Mute, 1, 0)
         
    ' Copy the data into the control value buffer
    CopyPtrFromStruct mxcd.paDetails, Mut, Len(Mut)
         
    ' Set the control value
    rc = mixerSetControlDetails(hmixer, _
        mxcd, _
        MIXER_SETCONTROLDETAILSF_VALUE)
         
    GlobalFree (hmem)
End Function

' Get Min-------------------------------------------------------
Public Property Get VolumeMin() As Long
'Minimaler Wert am Ausgang
    ' Sollte normalerweise 0 sein
    VolumeMin = volCtrl.lMinimum
End Property
Public Property Get MicMin() As Long
'Minimaler Wert des Mikrofons
    ' Sollte normalerweise 0 sein
    MicMin = micCtrl.lMinimum
End Property
Public Property Get LineMin() As Long
'Minimaler Wert des Line Eingangs
    ' Sollte normalerweise 0 sein
    LineMin = linCtrl.lMinimum
End Property
Public Property Get CDMin() As Long
'Minimaler Wert des CD Players
    ' Sollte normalerweise 0 sein
    CDMin = cdCtrl.lMinimum
End Property
Public Property Get WaveMin() As Long
'Minimaler Wert für Wave Dateien
    ' Sollte normalerweise 0 sein
    WaveMin = wavCtrl.lMinimum
End Property
Public Property Get SynMin() As Long
'Minimaler Wert des Synthesizers (MIDI)
    ' Sollte normalerweise 0 sein
    SynMin = synCtrl.lMinimum
End Property

' Get Max--------------------------------------------------------
Public Property Get VolumeMax() As Long
    VolumeMax = volCtrl.lMaximum
End Property
Public Property Get MicMax() As Long
'Maximaler Wert am Ausgang
    MicMax = micCtrl.lMaximum
End Property
Public Property Get LineMax() As Long
'Maximaler Wert am Line Eingang
    LineMax = linCtrl.lMaximum
End Property
Public Property Get CDMax() As Long
'Maximaler Wert des CD Players
    CDMax = cdCtrl.lMaximum
End Property
Public Property Get SynMax() As Long
'Maximaler Wert des Synthesizers (MIDI)
    SynMax = synCtrl.lMaximum
End Property
Public Property Get WaveMax() As Long
'Maximaler Wert Wave Dateien
    WaveMax = wavCtrl.lMaximum
End Property

' Get Level------------------------------------------------------
Public Property Get VolumeLevel() As Long
'Aktueller Wert am Ausgang
    VolumeLevel = Wert(volCtrl)
End Property
Public Property Get MicLevel() As Long
'Aktueller Wert des Mikrofoneingangs
    MicLevel = Wert(micCtrl)
End Property
Public Property Get LineLevel() As Long
'Aktueller Wert des Lineeingangs
    LineLevel = Wert(linCtrl)
End Property
Public Property Get SynLevel() As Long
'Aktueller Wert des Synthesizers (MIDI)
    SynLevel = Wert(synCtrl)
End Property
Public Property Get WaveLevel() As Long
'Aktueller Wert für Wave Dateien
    WaveLevel = Wert(wavCtrl)
End Property
Public Property Get CDLevel() As Long
'Aktueller Wert des CD Players
    CDLevel = Wert(cdCtrl)
End Property

' Let Level aktuellen Wert Setzen---------------------------------
Public Property Let MicLevel(ByVal New_Volume As Long)
    SetVolumeControl hmixer, micCtrl, New_Volume
End Property
Public Property Let LineLevel(ByVal New_Volume As Long)
    SetVolumeControl hmixer, linCtrl, New_Volume
End Property
Public Property Let CDLevel(ByVal New_Volume As Long)
    SetVolumeControl hmixer, cdCtrl, New_Volume
End Property
Public Property Let SynLevel(ByVal New_Volume As Long)
    SetVolumeControl hmixer, synCtrl, New_Volume
End Property
Public Property Let WaveLevel(ByVal New_Volume As Long)
    SetVolumeControl hmixer, wavCtrl, New_Volume
End Property
Public Property Let VolumeLevel(ByVal New_Volume As Long)
    SetVolumeControl hmixer, volCtrl, New_Volume
End Property

' Get Mute gibt True zurück wenn ein Aus-, Eingang stummgeschaltet ist-
Public Property Get VolumeMute() As Boolean
    VolumeMute = GetMute(volMute)
End Property
Public Property Get MikrofonMute() As Boolean
    MikrofonMute = GetMute(micMute)
End Property
Public Property Get LineMute() As Boolean
    LineMute = GetMute(linMute)
End Property
Public Property Get CompactDiscMute() As Boolean
    CompactDiscMute = GetMute(cdMute)
End Property
Public Property Get SynthesizerMute() As Boolean
    SynthesizerMute = GetMute(synMute)
End Property
Public Property Get WaveMute() As Boolean
    WaveMute = GetMute(wavMute)
End Property

' Let Mute True für Mute---------------------------------------------
Public Property Let VolumeMute(New_Mute As Boolean)
    SetMute volMute, New_Mute
End Property
Public Property Let MikrofonMute(New_Mute As Boolean)
    SetMute micMute, New_Mute
End Property
Public Property Let LineMute(New_Mute As Boolean)
    SetMute linMute, New_Mute
End Property
Public Property Let CompactDiscMute(New_Mute As Boolean)
    SetMute cdMute, New_Mute
End Property
Public Property Let SynthesizerMute(New_Mute As Boolean)
    SetMute synMute, New_Mute
End Property
Public Property Let WaveMute(New_Mute As Boolean)
    SetMute wavMute, New_Mute
End Property

' This function starts recording from the soundcard. The soundcard
' must be recording in order to
' monitor the input level. Without starting the recording from
' this application, input level
' can still be monitored if another application is recording audio
Private Function StartInput() As Boolean

    If fRecording Then
        StartInput = True
        Exit Function
    End If
    
    format.wFormatTag = 1
    format.nChannels = 1
    format.wBitsPerSample = 8
    format.nSamplesPerSec = 8000
    format.nBlockAlign = format.nChannels * format.wBitsPerSample / 8
    format.nAvgBytesPerSec = format.nSamplesPerSec * format.nBlockAlign
    format.cbSize = 0
    
    For i = 0 To NUM_BUFFERS - 1
        hmem(i) = GlobalAlloc(&H40, BUFFER_SIZE)
        inHdr(i).lpData = GlobalLock(hmem(i))
        inHdr(i).dwBufferLength = BUFFER_SIZE
        inHdr(i).dwFlags = 0
        inHdr(i).dwLoops = 0
    Next

    rc = waveInOpen(hWaveIn, DEVICEID, format, 0, 0, 0)
    If rc <> 0 Then
        waveInGetErrorText rc, msg, Len(msg)
        MsgBox msg, vbCritical + vbOKOnly, _
        App.Title & " - Mixercontrol Fehler"
        StartInput = False
        Exit Function
    End If

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInPrepareHeader(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical + vbOKOnly, _
            App.Title & " - Mixercontrol Fehler"
        End If
    Next

    For i = 0 To NUM_BUFFERS - 1
        rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
        If (rc <> 0) Then
            waveInGetErrorText rc, msg, Len(msg)
            MsgBox msg, vbCritical + vbOKOnly, _
            App.Title & " - Mixercontrol Fehler"
        End If
    Next

    fRecording = True
    rc = waveInStart(hWaveIn)
    StartInput = True
End Function

' Stop receiving audio input on the soundcard
Private Sub StopInput()

    fRecording = False
    waveInReset hWaveIn
    waveInStop hWaveIn
    For i = 0 To NUM_BUFFERS - 1
        waveInUnprepareHeader hWaveIn, inHdr(i), Len(inHdr(i))
        GlobalFree hmem(i)
    Next
    waveInClose hWaveIn
End Sub

' Function to process the wave recording notifications.
Private Sub waveInProc(ByVal hwi As Long, ByVal uMsg As Long, ByVal dwInstance As Long, ByRef hdr As WAVEHDR, ByVal dwParam2 As Long)
   If (uMsg = MM_WIM_DATA) Then
      If fRecording Then
         rc = waveInAddBuffer(hwi, hdr, Len(hdr))
      End If
   End If
End Sub

Private Function GetVolumeControl(ByVal hmixer As Long, _
                              ByVal componentType As Long, _
                              ByVal ctrlType As Long, _
                              ByRef mxc As MIXERCONTROL) As Boolean
                              
    ' This function attempts to obtain a mixer control. Returns True if successful.
    Dim mxlc As MIXERLINECONTROLS
    Dim mxl As MIXERLINE
    Dim hmem As Long
    Dim rc As Long
             
    mxl.cbStruct = Len(mxl)
    mxl.dwComponentType = componentType
      
    ' Obtain a line corresponding to the component type
    rc = mixerGetLineInfo(hmixer, mxl, _
        MIXER_GETLINEINFOF_COMPONENTTYPE)
         
    If (MMSYSERR_NOERROR = rc) Then
        mxlc.cbStruct = Len(mxlc)
        mxlc.dwLineID = mxl.dwLineID
        mxlc.dwControl = ctrlType
        mxlc.cControls = 1
        mxlc.cbmxctrl = Len(mxc)
             
        ' Allocate a buffer for the control
        hmem = GlobalAlloc(&H40, Len(mxc))
        mxlc.pamxctrl = GlobalLock(hmem)
        mxc.cbStruct = Len(mxc)
             
        ' Get the control
        rc = mixerGetLineControls(hmixer, _
            mxlc, _
            MIXER_GETLINECONTROLSF_ONEBYTYPE)
                  
        If (MMSYSERR_NOERROR = rc) Then
            GetVolumeControl = True
                 
            ' Copy the control into the destination structure
            CopyStructFromPtr mxc, mxlc.pamxctrl, Len(mxc)
        Else
            GetVolumeControl = False
        End If
        GlobalFree (hmem)
        Exit Function
    End If
      
    GetVolumeControl = False
End Function

Public Property Get MaxVolumeMeterInput() As Long
    MaxVolumeMeterInput = inputVolCtrl.lMaximum
End Property

Public Property Get MinVolumeMeterInput() As Long
    ' Sollte normalerweise 0 sein
    MinVolumeMeterInput = inputVolCtrl.lMinimum
End Property

Public Property Get MaxVolumeMeterOutput() As Long
    MaxVolumeMeterOutput = outputVolCtrl.lMaximum
End Property

Public Property Get MinVolumeMeterOutput() As Long
    ' Sollte normalerweise 0 sein
    MinVolumeMeterOutput = outputVolCtrl.lMinimum
End Property

Public Property Get CurrentVolumeMeterOutput() As Long
    ' Process sound buffer if recording
    If (fRecording) Then
       For i = 0 To (NUM_BUFFERS - 1)
          If inHdr(i).dwFlags And WHDR_DONE Then
             rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
          End If
       Next
    End If

    ' Get the current output level
    mxcd.dwControlID = outputVolCtrl.dwControlID
    mxcd.item = outputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, _
    MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
      
    If (volume < 0) Then volume = -volume
   
    CurrentVolumeMeterOutput = volume

End Property

Public Property Get CurrentVolumeMeterInput() As Long
    ' Process sound buffer if recording
    If Not StartInput Then StartInput


    If (fRecording) Then
        For i = 0 To (NUM_BUFFERS - 1)
            If inHdr(i).dwFlags And WHDR_DONE Then
                rc = waveInAddBuffer(hWaveIn, inHdr(i), Len(inHdr(i)))
            End If
        Next
    End If

   ' Get the current input level
    mxcd.dwControlID = inputVolCtrl.dwControlID
    mxcd.item = inputVolCtrl.cMultipleItems
    rc = mixerGetControlDetails(hmixer, mxcd, _
    MIXER_GETCONTROLDETAILSF_VALUE)
    CopyStructFromPtr volume, mxcd.paDetails, Len(volume)
    If (volume < 0) Then
        volume = -volume
    End If
    CurrentVolumeMeterInput = volume
End Property
