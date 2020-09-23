Attribute VB_Name = "modDXSound"
Option Explicit
'SOUNDS? Well, in the vacuum is impossible to hear sounds
'but I think that the game is better with them, isn't it?

Public DX As New DirectX7
'DirectX Variables
Public ds As DirectSound

'User defined type to determine a buffer's capabilities
Private Type BufferCaps
    Volume As Boolean               'Can this buffer's volume be changed?
    Frequency As Boolean            'Can the frequency be altered?
    Pan As Boolean                  'Can we pan the sound from left to right?
    Loop As Boolean                 'Is this sound looping?
    Delete As Boolean               'Should this sound be deleted after playing?
End Type

'User defined type to contain sound data
Private Type SoundArray
    DSBuffer As DirectSoundBuffer   'The buffer that contains the sound
    DSState As String               'Describes the current state of the buffer (ie. "Playing", "Stopped")
    DSNotification As Long          'Contains the event reference returned by the DirectX7 object
    DSCaps As BufferCaps            'Describes the buffer's capabilities
    DSSourceName As String          'The name of the source file
    DSFile As Boolean               'Is the source in a seperate file?
    DSResource As Boolean           'Or is it in a resource?
    DSEmpty As Boolean              'Is this SoundArray index empty?
End Type

'This will contain the return values given by the LoadSound function
Dim Sound() As SoundArray          'Contains all the data needed for sound manipulation
Global SoundID() As Integer        'IDs of sounds

'Wave Format Setting Contants
Public Const NumChannels = 4              'How many channels will we be playing on?
Public Const SamplesPerSecond = 22050     'How many cycles per second (hertz)?
Public Const BitsPerSample = 16           'What bit-depth will we use?

'Constant that contains the path inside the app.path in which the sounds are stored
Public Const DataLocation = "\Audio\"

Public SoundVolume As Long

'This will contain the name of the two soundfiles we wish to use
Public SoundFileName(7) As String

'Constants for the index value of each sound
Public Const dsXpl1 = 0
Public Const dsXpl2 = 1
Public Const dsXpl3 = 2
Public Const dsXpl4 = 3
Public Const dsGun1 = 4
Public Const dsGun2 = 5
Public Const dsFire1 = 6
Public Const dsFire2 = 7

Public Sub Init_DSound()
    'Set up DirectSound for this window
    Initialize_DSound Form1.hWnd
    
    'Specify the sound files we'll be using
    SoundFileName(0) = "Xplode1.wav"
    SoundFileName(1) = "Xplode2.wav"
    SoundFileName(2) = "Xplode3.wav"
    SoundFileName(3) = "Xplode4.wav"
    SoundFileName(4) = "gun1.wav"
    SoundFileName(5) = "gun2.wav"
    SoundFileName(6) = "fire1.wav"
    SoundFileName(7) = "fire2.wav"
    
    'Load the specified sound into a buffer
    Dim Index As Integer
    ReDim SoundID(0 To UBound(SoundFileName))
    
    For Index = 0 To UBound(SoundFileName)
        SoundID(Index) = LoadSound(SoundFileName(Index), True, False, False, False, False, True, False, Form1)
        SetVolume Index, SoundVolume
    Next
End Sub


Public Sub Initialize_DSound(ByRef Handle As Long)

    'If we can't initialize properly, trap the error
    On Local Error GoTo ErrOut

    'Make the DirectSound object
    Set ds = DX.DirectSoundCreate("")
    
    'Set the DirectSound object's cooperative level (Priority gives us sole control)
    ds.SetCooperativeLevel Handle, DSSCL_PRIORITY
    
    'Initialize our Sound array to zero
    ReDim Sound(0)
    Sound(0).DSEmpty = True
    Sound(0).DSState = "empty"
    
    'Exit sub before the error code
    Exit Sub
    
ErrOut:
    'Display an error message and exit if initialization failed
    MsgBox "Unable to initialize DirectSound."
    End

End Sub

Public Function LoadSound(SourceName As String, IsFile As Boolean, IsResource As Boolean, IsDelete As Boolean, IsFrequency As Boolean, IsPan As Boolean, IsVolume As Boolean, IsLoop As Boolean, FormObject As Form) As Integer

Dim i As Integer
Dim Index As Integer
Dim DSBufferDescription As DSBUFFERDESC
Dim DSFormat As WAVEFORMATEX
Dim DSPosition(0) As DSBPOSITIONNOTIFY

    'Search the sound array for any empty spaces
    Index = -1
    For i = 0 To UBound(Sound)
        If Sound(i).DSEmpty = True Then 'If there is an empty space, us it
            Index = i
            Exit For
        End If
    Next
    If Index = -1 Then                  'If there's no empty space, make a new spot
        ReDim Preserve Sound(UBound(Sound) + 1)
        Index = UBound(Sound)
    End If
    LoadSound = Index                   'Set the return value of the function
    
    'Load the Sound array with the data given
    With Sound(Index)
        .DSEmpty = False                'This Sound(index) is now occupied with data
        .DSFile = IsFile                'Is this sound to be loaded from a file?
        .DSResource = IsResource        'Or is it to be loaded from a resource?
        .DSSourceName = SourceName      'What is the name of the source?
        .DSState = "Stopped"            'Set the current state to "Stopped"
        .DSCaps.Delete = IsDelete       'Is this sound to be deleted after it is played?
        .DSCaps.Frequency = IsFrequency 'Is this sound to have frequency altering capabilities?
        .DSCaps.Loop = IsLoop           'Is this sound to be looped?
        .DSCaps.Pan = IsPan             'Is this sound to have Left and Right panning capabilities?
        .DSCaps.Volume = IsVolume       'Is this sound capable of altered volume settings?
    End With
    
    'Set the buffer description according to the data provided
    With DSBufferDescription
        If Sound(Index).DSCaps.Delete = True Then .lFlags = .lFlags Or DSBCAPS_CTRLPOSITIONNOTIFY
        If Sound(Index).DSCaps.Frequency = True Then .lFlags = .lFlags Or DSBCAPS_CTRLFREQUENCY
        If Sound(Index).DSCaps.Pan = True Then .lFlags = .lFlags Or DSBCAPS_CTRLPAN
        If Sound(Index).DSCaps.Volume = True Then .lFlags = .lFlags Or DSBCAPS_CTRLVOLUME
    End With

    'Set the Wave Format
    With DSFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = NumChannels
        .lSamplesPerSec = SamplesPerSecond
        .nBitsPerSample = BitsPerSample
        .nBlockAlign = .nBitsPerSample / 8 * .nChannels
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With
    
    'Load the sound into the buffer
    If Sound(Index).DSFile = True Then          'If it's in a file...
        Set Sound(Index).DSBuffer = ds.CreateSoundBufferFromFile(App.path & DataLocation & Sound(Index).DSSourceName, DSBufferDescription, DSFormat)
    ElseIf Sound(Index).DSResource = True Then  'If it's in a resource...
        Set Sound(Index).DSBuffer = ds.CreateSoundBufferFromResource("", Sound(Index).DSSourceName, DSBufferDescription, DSFormat)
    End If
    
    'If the sound is to be deleted after it plays, we must create an event for it
    If Sound(Index).DSCaps.Delete = True Then
        Sound(Index).DSNotification = DX.CreateEvent(FormObject)        'Make the event (has to be created in a Form Object) and get its handle
        DSPosition(0).hEventNotify = Sound(Index).DSNotification        'Place this event handle in an DSBPOSITIONNOTIFY variable
        DSPosition(0).lOffset = DSBPN_OFFSETSTOP                        'Define the position within the wave file at which you would like the event to be triggered
        Sound(Index).DSBuffer.SetNotificationPositions 1, DSPosition()  'Set the "notification position" by passing the DSBPOSITIONNOTIFY variable
    End If
    
End Function

Public Sub RemoveSound(Index As Integer)

    'Destroy the event associated with the ending of this sound, if there was one
    If Sound(Index).DSCaps.Delete = True And Sound(Index).DSNotification <> 0 Then DX.DestroyEvent Sound(Index).DSNotification
    
    'Reset all the variables in the sound array
    With Sound(Index)
        Set .DSBuffer = Nothing
        .DSCaps.Delete = False
        .DSCaps.Frequency = False
        .DSCaps.Loop = False
        .DSCaps.Pan = False
        .DSCaps.Volume = False
        .DSEmpty = True
        .DSFile = False
        .DSNotification = 0
        .DSResource = False
        .DSSourceName = ""
        .DSState = "empty"
    End With
        
End Sub

Public Sub PlaySound(Index As Integer)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'If the sound is not "paused" then reset it's position to the beginning
    If Sound(Index).DSState <> "paused" Then Sound(Index).DSBuffer.SetCurrentPosition 0
    
    'Play looped or singly, as appropriate
    If Sound(Index).DSCaps.Loop = False Then Sound(Index).DSBuffer.Play DSBPLAY_DEFAULT
    If Sound(Index).DSCaps.Loop = True Then Sound(Index).DSBuffer.Play DSBPLAY_LOOPING
    
    'Set the state to "playing"
    Sound(Index).DSState = "playing"

End Sub

Public Sub StopSound(Index As Integer)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'Stop the buffer and reset to the beginning
    Sound(Index).DSBuffer.Stop
    Sound(Index).DSBuffer.SetCurrentPosition 0
    Sound(Index).DSState = "stopped"

End Sub

Public Sub PauseSound(Index As Integer)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'Stop the buffer
    Sound(Index).DSBuffer.Stop
    Sound(Index).DSState = "paused"

End Sub

Public Sub SetFrequency(Index As Integer, freq As Long)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'Check to make sure that the buffer has the capability of altering its frequency
    If Sound(Index).DSCaps.Frequency = False Then Exit Sub

    'Alter the frequency according to the Freq provided
    Sound(Index).DSBuffer.SetFrequency freq

End Sub

Public Sub SetVolume(Index As Integer, Vol As Long)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'Check to make sure that the buffer has the capability of altering its volume
    If Sound(Index).DSCaps.Volume = False Then Exit Sub

    'Alter the volume according to the Vol provided
    Sound(Index).DSBuffer.SetVolume Vol

End Sub

Public Sub SetPan(Index As Integer, Pan As Long)

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Sub
    
    'Check to make sure that the buffer has the capability of altering its pan
    If Sound(Index).DSCaps.Pan = False Then Exit Sub

    'Alter the pan according to the Pan provided
    Sound(Index).DSBuffer.SetPan Pan

End Sub

Public Function GetFrequency(Index As Integer) As Long

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Function
    
    'Check to make sure that the buffer has the capability of altering its frequency
    If Sound(Index).DSCaps.Frequency = False Then Exit Function
    
    'Return the frequency value
    GetFrequency = Sound(Index).DSBuffer.GetFrequency()

End Function

Public Function GetVolume(Index As Integer) As Long

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Function
    
    'Check to make sure that the buffer has the capability of altering its volume
    If Sound(Index).DSCaps.Volume = False Then Exit Function
    
    'Return the volume value
    GetVolume = Sound(Index).DSBuffer.GetVolume()

End Function

Public Function GetPan(Index As Integer) As Long

    'Check to make sure there is a sound loaded in the specified buffer
    If Sound(Index).DSEmpty Then Exit Function
    
    'Check to make sure that the buffer has the capability of altering its pan
    If Sound(Index).DSCaps.Pan = False Then Exit Function
    
    'Return the pan value
    GetPan = Sound(Index).DSBuffer.GetPan()

End Function

Public Function Get_DS_State(Index As Integer) As String

    'Returns the current state of the given sound
    Get_DS_State = Sound(Index).DSState

End Function

Public Sub Terminate_DSound()

Dim i As Integer

    'Delete all of the sounds created
    For i = 0 To UBound(Sound)
        RemoveSound i
    Next

End Sub

Public Sub DXInit()
'Direct Sound Initialization
    'Make the DirectSound object
    Set ds = DX.DirectSoundCreate("")
    
    'Set the DirectSound object's cooperative level (Priority gives us sole control)
    ds.SetCooperativeLevel Form1.hWnd, DSSCL_PRIORITY
    
    'Initialize our Sound array to zero
    ReDim Sound(0)
    Sound(0).DSEmpty = True
    Sound(0).DSState = "empty"

End Sub

