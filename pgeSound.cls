VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "pgeSound"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Pab Game Engine Sound Class
'---------------------------
'Created by Paul Berlin 2002-2003
'Handles sound effects & music using the fmod sound system.
'
'Filetypes supported for music:
'Type Stream - MP3, MP2, OGG, WMA, ASF, MID, MIDI, RMI, WAV, SGM
'Type Music - MOD, S3M, XM, IT
'Filetypes supported for soundeffects:
'MP3, MP2, OGG, WAV

Private Declare Function FSOUND_Init Lib "fmod.dll" Alias "_FSOUND_Init@12" (ByVal mixrate As Long, ByVal maxchannels As Long, ByVal Flags As Long) As Byte
Private Declare Function FSOUND_Close Lib "fmod.dll" Alias "_FSOUND_Close@0" () As Long
Private Declare Function FSOUND_SetBufferSize Lib "fmod.dll" Alias "_FSOUND_SetBufferSize@4" (ByVal lenms As Long) As Byte
Private Declare Function FSOUND_SetOutput Lib "fmod.dll" Alias "_FSOUND_SetOutput@4" (ByVal outputtype As FSOUND_OUTPUTTYPES) As Byte
Private Declare Function FMUSIC_FreeSong Lib "fmod.dll" Alias "_FMUSIC_FreeSong@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_PlaySong Lib "fmod.dll" Alias "_FMUSIC_PlaySong@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_StopSong Lib "fmod.dll" Alias "_FMUSIC_StopSong@4" (ByVal module As Long) As Byte
Private Declare Function FMUSIC_LoadSong Lib "fmod.dll" Alias "_FMUSIC_LoadSong@4" (ByVal name As String) As Long
Private Declare Function FMUSIC_SetLooping Lib "fmod.dll" Alias "_FMUSIC_SetLooping@8" (ByVal module As Long, ByVal looping As Byte) As Byte
Private Declare Function FMUSIC_SetMasterVolume Lib "fmod.dll" Alias "_FMUSIC_SetMasterVolume@8" (ByVal module As Long, ByVal Volume As Long) As Byte
Private Declare Function FSOUND_Stream_PlayEx Lib "fmod.dll" Alias "_FSOUND_Stream_PlayEx@16" (ByVal channel As Long, ByVal Stream As Long, ByVal DSP As Long, ByVal startpaused As Byte) As Long
Private Declare Function FSOUND_Stream_OpenFile Lib "fmod.dll" Alias "_FSOUND_Stream_OpenFile@12" (ByVal FileName As String, ByVal Mode As FSOUND_MODES, ByVal memlength As Long) As Long
Private Declare Function FSOUND_Stream_Stop Lib "fmod.dll" Alias "_FSOUND_Stream_Stop@4" (ByVal Stream As Long) As Byte
Private Declare Function FSOUND_SetFrequency Lib "fmod.dll" Alias "_FSOUND_SetFrequency@8" (ByVal channel As Long, ByVal Freq As Long) As Byte
Private Declare Function FSOUND_SetVolume Lib "fmod.dll" Alias "_FSOUND_SetVolume@8" (ByVal channel As Long, ByVal Vol As Long) As Byte
Private Declare Function FSOUND_SetPan Lib "fmod.dll" Alias "_FSOUND_SetPan@8" (ByVal channel As Long, ByVal pan As Long) As Byte
Private Declare Function FSOUND_SetLoopMode Lib "fmod.dll" Alias "_FSOUND_SetLoopMode@8" (ByVal channel As Long, ByVal loopmode As Byte) As Byte
Private Declare Function FSOUND_PlaySound Lib "fmod.dll" Alias "_FSOUND_PlaySound@8" (ByVal channel As Long, ByVal sptr As Long) As Long
Private Declare Function FSOUND_StopSound Lib "fmod.dll" Alias "_FSOUND_StopSound@4" (ByVal channel As Long) As Byte
Private Declare Function FSOUND_Sample_SetDefaults Lib "fmod.dll" Alias "_FSOUND_Sample_SetDefaults@20" (ByVal sptr As Long, ByVal deffreq As Long, ByVal defvol As Long, ByVal defpan As Long, ByVal defpri As Long) As Byte
Private Declare Function FSOUND_Sample_SetMode Lib "fmod.dll" Alias "_FSOUND_Sample_SetMode@8" (ByVal sptr As Long, ByVal Mode As FSOUND_MODES) As Byte
Private Declare Function FSOUND_Sample_Load Lib "fmod.dll" Alias "_FSOUND_Sample_Load@16" (ByVal Index As Long, ByVal name As String, ByVal Mode As FSOUND_MODES, ByVal memlength As Long) As Long
Private Declare Function FSOUND_SetPaused Lib "fmod.dll" Alias "_FSOUND_SetPaused@8" (ByVal channel As Long, ByVal Paused As Byte) As Byte
Private Declare Function FSOUND_IsPlaying Lib "fmod.dll" Alias "_FSOUND_IsPlaying@4" (ByVal channel As Long) As Byte

Private Enum FSOUND_OUTPUTTYPES
    FSOUND_OUTPUT_NOSOUND    'NoSound driver, all calls to this succeed but do nothing.
    FSOUND_OUTPUT_WINMM      'Windows Multimedia driver.
    FSOUND_OUTPUT_DSOUND     'DirectSound driver.  You need this to get EAX2 or EAX3 support, or FX api support.
    FSOUND_OUTPUT_A3D        'A3D driver.  You need this to get geometry and EAX reverb support.

    FSOUND_OUTPUT_OSS        'Linux/Unix OSS (Open Sound System) driver, i.e. the kernel sound drivers.
    FSOUND_OUTPUT_ESD        'Linux/Unix ESD (Enlightment Sound Daemon) driver.
    FSOUND_OUTPUT_ALSA       'Linux Alsa driver.
    
    FSOUND_OUTPUT_ASIO       ' Low latency ASIO driver
    FSOUND_OUTPUT_XBOX       ' Xbox driver
    FSOUND_OUTPUT_PS2        ' PlayStation 2 driver
    FSOUND_OUTPUT_MAC        ' Mac SoundMager driver
End Enum

Private Enum FSOUND_MODES
    FSOUND_LOOP_OFF = 1            ' For non looping samples.
    FSOUND_LOOP_NORMAL = 2         ' For forward looping samples.
    FSOUND_LOOP_BIDI = 4           ' For bidirectional looping samples.  (no effect if in hardware).
    FSOUND_8BITS = 8               ' For 8 bit samples.
    FSOUND_16BITS = 16             ' For 16 bit samples.
    FSOUND_MONO = 32               ' For mono samples.
    FSOUND_STEREO = 64             ' For stereo samples.
    FSOUND_UNSIGNED = 128          ' For source data containing unsigned samples.
    FSOUND_SIGNED = 256            ' For source data containing signed data.
    FSOUND_DELTA = 512             ' For source data stored as delta values.
    FSOUND_IT214 = 1024            ' For source data stored using IT214 compression.
    FSOUND_IT215 = 2048            ' For source data stored using IT215 compression.
    FSOUND_HW3D = 4096             ' Attempts to make samples use 3d hardware acceleration. (if the card supports it)
    FSOUND_2D = 8192               ' Ignores any 3d processing.  overrides FSOUND_HW3D.  Located in software.
    FSOUND_STREAMABLE = 16384      ' For realtime streamable samples.  If you dont supply this sound may come out corrupted.
    FSOUND_LOADMEMORY = 32768      ' For FSOUND_Sample_Load - name will be interpreted as a pointer to data
    FSOUND_LOADRAW = 65536         ' For FSOUND_Sample_Load/FSOUND_Stream_Open - will ignore file format and treat as raw pcm.
    FSOUND_MPEGACCURATE = 131072   ' For FSOUND_Stream_Open - scans MP2/MP3 (VBR also) for accurate FSOUND_Stream_GetLengthMs/FSOUND_Stream_SetTime.
    FSOUND_FORCEMONO = 262144      ' For forcing stereo streams and samples to be mono - needed with FSOUND_HW3D - incurs speed hit
    FSOUND_HW2D = 524288           ' 2d hardware sounds.  allows hardware specific effects
    FSOUND_ENABLEFX = 1048576      ' Allows DX8 FX to be played back on a sound.  Requires DirectX 8 - Note these sounds cant be played more than once, or have a changing frequency
    FSOUND_MPEGHALFRATE = 2097152  ' For FMODCE only - decodes mpeg streams using a lower quality decode, but faster execution
    FSOUND_XADPCM = 4194304        ' For XBOX only - Describes a user sample that its contents are compressed as XADPCM
    FSOUND_VAG = 8388608           ' For PS2 only - Describes a user sample that its contents are compressed as Sony VAG format.
    FSOUND_NONBLOCKING = 16777216  ' For FSOUND_Stream_OpenFile - Causes stream to open in the background and not block the foreground app - stream plays only when ready.
    
    FSOUND_NORMAL = FSOUND_16BITS Or FSOUND_SIGNED Or FSOUND_MONO
End Enum

Private Enum FSOUND_CHANNELSAMPLEMODE
    FSOUND_FREE = -1                 ' definition for dynamically allocated channel or sample
    FSOUND_UNMANAGED = -2            ' definition for allocating a sample that is NOT managed by fsound
    FSOUND_ALL = -3                  ' for a channel index or sample index, this flag affects ALL channels or samples available!  Not supported by all functions.
    FSOUND_STEREOPAN = -1            ' definition for full middle stereo volume on both channels
    FSOUND_SYSTEMCHANNEL = -1000     ' special channel ID for channel based functions that want to alter the global FSOUND software mixing output channel
    FSOUND_SYSTEMSAMPLE = -1000      ' special sample ID for all sample based functions that want to alter the global FSOUND software mixing output sample
End Enum

Public Enum enumPgePlayMode
  LOOP_OFF = 1
  LOOP_NORMAL = 2
  LOOP_BIDI = 4
End Enum

Private Type tSfxPool
  sID As String
  lPointer As Long
End Type

Private Type tMusicPool
  sID As String
  lPointer As Long
  bStream As Byte 'If true, the file is of type stream, else of type music
End Type

Private lMusicPlaying As Long 'Pointer to music pool of currently playing music
Private lMusicChannel As Long 'Pointer to channel stream music is playing on
Private MusicPool() As tMusicPool
Private SfxPool() As tSfxPool

Option Explicit

Public Function Init(ByVal lFreq As Long, lChannels As Long, ByVal lBufferMs As Long)
  'This function must be called before any other
  'It initializes the fmod system
  On Error Resume Next
  
  Call FSOUND_SetBufferSize(lBufferMs)
  Call FSOUND_SetOutput(FSOUND_OUTPUT_DSOUND)
  Init = FSOUND_Init(lFreq, lChannels, 0)

End Function

Private Sub Class_Initialize()
  ReDim MusicPool(1 To 1)
  ReDim SfxPool(1 To 1)
End Sub

Private Sub Class_Terminate()
  Dim x As Long
  FSOUND_Close
  For x = 1 To UBound(MusicPool)
    If Not MusicPool(x).bStream Then Call FMUSIC_FreeSong(MusicPool(x).lPointer)
  Next
End Sub

Public Sub MusicLoad(ByVal sFilename As String, ByVal sID As String, ByVal bStream As Boolean)
  'Loads an music file into music pool with ID
  
  ReDim Preserve MusicPool(1 To UBound(MusicPool) + 1)
  With MusicPool(UBound(MusicPool))
    
    .bStream = bStream
    .sID = sID
    If bStream Then
      .lPointer = FSOUND_Stream_OpenFile(sFilename, FSOUND_LOOP_NORMAL, 0)
    Else
      .lPointer = FMUSIC_LoadSong(sFilename)
      Call FMUSIC_SetLooping(.lPointer, True)
    End If
  
  End With
  
End Sub

Public Sub MusicPlayID(ByVal sID As String, Optional bVolume As Byte = 255)
  'This plays the specified music at specified volume
  Dim x As Long, y As Long
  
  For x = 1 To UBound(MusicPool)
    If MusicPool(x).sID = sID Then
      y = x
      Exit For
    End If
  Next
  
  If y > 0 And y <> lMusicPlaying Then
    MusicStop
    lMusicPlaying = y
    
    With MusicPool(y)
    
      If .bStream Then
        lMusicChannel = FSOUND_Stream_PlayEx(FSOUND_FREE, .lPointer, 0, True)
        Call FSOUND_SetVolume(lMusicChannel, bVolume)
        Call FSOUND_SetPaused(lMusicChannel, False)
      Else
        Call FMUSIC_PlaySong(.lPointer)
        Call FMUSIC_SetMasterVolume(.lPointer, bVolume)
      End If
    
    End With
  End If
  
End Sub

Public Sub MusicStop()
  'Stops playing music
  If lMusicPlaying > 0 Then
    With MusicPool(lMusicPlaying)
      If .bStream Then
        Call FSOUND_Stream_Stop(.lPointer)
      Else
        Call FMUSIC_StopSong(.lPointer)
      End If
      lMusicPlaying = 0
    End With
  End If
End Sub

Public Sub MusicVolume(ByVal bVolume As Byte)
  'Sets volume of playing music
  If lMusicPlaying > 0 Then
    With MusicPool(lMusicPlaying)
      If .bStream Then
        Call FSOUND_SetVolume(lMusicChannel, bVolume)
      Else
        Call FMUSIC_SetMasterVolume(.lPointer, bVolume)
      End If
    End With
  End If
End Sub

Public Sub SfxLoad(ByVal sFilename As String, ByVal sID As String)
  'Loads an sound effect file into sfx pool with ID
  
  ReDim Preserve SfxPool(1 To UBound(SfxPool) + 1)
  With SfxPool(UBound(SfxPool))
    .sID = sID
    .lPointer = FSOUND_Sample_Load(FSOUND_FREE, sFilename, 0, 0)
  End With
  
End Sub

Public Sub SfxLoadEx(ByVal sFilename As String, ByVal sID As String, ByVal lFreq As Long, ByVal lVolume As Long, ByVal lPan As Long, ByVal lPriority As Long)
  'Loads an sound effect file into sfx pool with ID
  'And also sets options for the sfx:
  'lFreq = Default sample frequency. The value here is specified in hz. -1 to ignore.
  'lVolume = Default sample volume. This is a value from 0 to 255. -1 to ignore.
  'lPan = Default sample pan position. This is a vale from 0 to 255. -1 to ignore.
  'lPriority = Default sample priority. This is a value from 0 to 255. -1 to ignore.
  
  ReDim Preserve SfxPool(1 To UBound(SfxPool) + 1)
  With SfxPool(UBound(SfxPool))
    .sID = sID
    .lPointer = FSOUND_Sample_Load(FSOUND_FREE, sFilename, 0, 0)
    Call FSOUND_Sample_SetDefaults(.lPointer, lFreq, lVolume, lPan, lPriority)
  End With
  
End Sub

Public Sub SfxPlayID(ByVal sID As String)
  'Plays sfx with ID
  'It will be played at default frequency, volume & panning
  Dim x As Long, y As Long
  
  For x = 1 To UBound(SfxPool)
    If SfxPool(x).sID = sID Then
      y = x
      Exit For
    End If
  Next
  
  If y > 0 Then
    With SfxPool(y)
      Call FSOUND_PlaySound(FSOUND_FREE, .lPointer)
    End With
  End If
End Sub

Public Sub SfxSetupID(ByVal sID As String, ByVal lFreq As Long, ByVal lVolume As Long, ByVal lPan As Long, ByVal lPriority As Long)
  'Sets up the sfx
  'lFreq = Default sample frequency. The value here is specified in hz. -1 to ignore.
  'lVolume = Default sample volume. This is a value from 0 to 255. -1 to ignore.
  'lPan = Default sample pan position. This is a vale from 0 to 255. -1 to ignore.
  'lPriority = Default sample priority. This is a value from 0 to 255. -1 to ignore.
  Dim x As Long, y As Long
  
  'NOTE: This does not seem to work right now, but using SfxChangePlaying works...
  
  For x = 1 To UBound(SfxPool)
    If SfxPool(x).sID = sID Then
      y = x
      Exit For
    End If
  Next
  
  If y > 0 Then
    With SfxPool(y)
      Call FSOUND_Sample_SetDefaults(.lPointer, lFreq, lVolume, lPan, lPriority)
    End With
  End If
  
End Sub

Public Function SfxPlayExID(ByVal sID As String, ByVal lMode As enumPgePlayMode, Optional ByVal lVolume As Long = -1) As Long
  'Plays sfx with ID, Looping or not
  'Returns ID of loop which is used to let you change
  'Frequency, Panning, Volume and stop it
  Dim x As Long, y As Long
  
  For x = 1 To UBound(SfxPool)
    If SfxPool(x).sID = sID Then
      y = x
      Exit For
    End If
  Next
  
  If y > 0 Then
    With SfxPool(y)
      Call FSOUND_Sample_SetMode(.lPointer, lMode)
      SfxPlayExID = FSOUND_PlaySound(FSOUND_FREE, .lPointer)
      If lVolume > -1 Then
        FSOUND_SetVolume SfxPlayExID, tob(lVolume)
      End If
    End With
  End If

End Function

Public Function SfxIsPlaying(ByVal lPlayExPointer As Long) As Boolean
  'This will return true if specified sound is playing
  SfxIsPlaying = FSOUND_IsPlaying(lPlayExPointer)
End Function

Public Sub SfxStop(ByVal lPlayExPointer As Long)
  'This will stop any sound started with SfxPlayExID
  Call FSOUND_StopSound(lPlayExPointer)
End Sub

Public Sub SfxEndLoop(ByVal lPlayExPointer As Long)
  'This will end the loop of the sound so that it wont loop any more
  'the sound will eventually stop, when it has played
  FSOUND_SetLoopMode lPlayExPointer, False
End Sub

Public Sub SfxChangePlaying(ByVal lPlayExPointer As Long, ByVal lFreq As Long, ByVal lVolume As Long, ByVal lPan As Long)
  'This will setup any sound started with SfxPlayExID while playing
  'These changes will not affect the sfx, just the playing sound
  If lFreq > 0 Then Call FSOUND_SetFrequency(lPlayExPointer, lFreq)
  If lVolume > 0 Then Call FSOUND_SetVolume(lPlayExPointer, lVolume)
  If lPan > 0 Then Call FSOUND_SetPan(lPlayExPointer, lPan)
End Sub

