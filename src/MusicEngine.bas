Attribute VB_Name = "MusicEngine"
Option Explicit

'------ Plays OGG streams using StreamOGG and clsDLL classes ------
'------ Uses Direct Sound 8 interface and buffer streaming --------
'------ by davidgf. Thanks to DX9 SDK and [rm_code] ---------

Public Volume As Long
Public MusicStatus As Integer      '1 play, 0 stop, 2 pause

Private Const MYBUFFERSIZE As Long = 600000
Private Const MYBUFFERPARTSIZE As Long = 200000
Private Const B1_B2 As Long = 200000
Private Const B2_B3 As Long = 400000

Private MusicBuffer As DirectSoundSecondaryBuffer8
Private bytes(MYBUFFERSIZE) As Byte
Private ebytes(MYBUFFERPARTSIZE) As Byte

Private BufferCursor As DSCURSORS
Private Out As SND_RESULT
Private Read As Long
Private PauseCursor As Long
Private ContOut As Integer

Private OGGLib As StreamOGG

Private w1 As Boolean, w2 As Boolean, w3 As Boolean

Public Sub PlayMusic(ByVal filename As String)
OGGLib.StreamClose

Out = OGGLib.StreamOpen(filename)

Dim soundDESC As DSBUFFERDESC
soundDESC.lBufferBytes = MYBUFFERSIZE
soundDESC.fxFormat.nBitsPerSample = OGGLib.BitsPerSample
soundDESC.fxFormat.nChannels = OGGLib.Channels
soundDESC.fxFormat.lSamplesPerSec = OGGLib.SamplesPerSecond
soundDESC.fxFormat.nSize = 0
soundDESC.fxFormat.nBlockAlign = (OGGLib.BitsPerSample * OGGLib.Channels) / 8
soundDESC.fxFormat.lAvgBytesPerSec = OGGLib.SamplesPerSecond * soundDESC.fxFormat.nBlockAlign
soundDESC.lFlags = DSBCAPS_CTRLVOLUME + DSBCAPS_CTRLPAN
soundDESC.fxFormat.nFormatTag = 1
Set MusicBuffer = Nothing
Set MusicBuffer = DirectSound.CreateSoundBuffer(soundDESC)

Out = OGGLib.StreamRead(VarPtr(bytes(0)), MYBUFFERSIZE, Read)
MusicBuffer.WriteBuffer 0, MYBUFFERSIZE, bytes(0), DSBLOCK_DEFAULT
MusicBuffer.SetCurrentPosition 0
MusicBuffer.Play DSBPLAY_LOOPING

MusicStatus = 1

w1 = False
w2 = True
w3 = True
ContOut = 0
End Sub

Public Sub StopMusic()
If Not MusicBuffer Is Nothing Then MusicBuffer.Stop
If Not OGGLib Is Nothing Then OGGLib.StreamClose
MusicStatus = 0
ContOut = 0
End Sub

Public Sub RenderTime()
'this function MUST be called at least once a second to allow the prog to refresh
'the sound buffer. if not some sound spikes or bad music looping will be heard
'       splits the buffer into 3 areas for fast music streaming
'       loads the music as it is played
Static lastVol As Long

If Not MusicBuffer Is Nothing And MusicStatus = 1 Then
    If Volume <> lastVol Then MusicBuffer.SetVolume Volume
    
    MusicBuffer.GetCurrentPosition BufferCursor
    
    If BufferCursor.lPlay > B1_B2 And BufferCursor.lPlay < B2_B3 And w1 = False Then
        If ContOut <> 0 Then
            ContOut = ContOut + 1
            MusicBuffer.WriteBuffer 0, MYBUFFERPARTSIZE, ebytes(0), DSBLOCK_DEFAULT
            If ContOut = 4 Then GoTo SoundEnd
        Else
            Out = OGGLib.StreamRead(VarPtr(bytes(0)), MYBUFFERPARTSIZE, Read)
            MusicBuffer.WriteBuffer 0, MYBUFFERPARTSIZE, bytes(0), DSBLOCK_DEFAULT
            If Read < MYBUFFERPARTSIZE Then
                MusicBuffer.WriteBuffer Read, (MYBUFFERPARTSIZE - Read), ebytes(0), DSBLOCK_DEFAULT
                ContOut = ContOut + 1
            End If
        End If
        w1 = True
        w2 = False
        w3 = False
    End If
    
    If BufferCursor.lPlay > B2_B3 And w2 = False Then
        If ContOut <> 0 Then
            ContOut = ContOut + 1
            MusicBuffer.WriteBuffer B1_B2, MYBUFFERPARTSIZE, ebytes(0), DSBLOCK_DEFAULT
            If ContOut = 4 Then GoTo SoundEnd
        Else
            Out = OGGLib.StreamRead(VarPtr(bytes(0)), MYBUFFERPARTSIZE, Read)
            MusicBuffer.WriteBuffer B1_B2, MYBUFFERPARTSIZE, bytes(0), DSBLOCK_DEFAULT
            If Read < MYBUFFERPARTSIZE Then
                MusicBuffer.WriteBuffer Read + B1_B2, (MYBUFFERPARTSIZE - Read), ebytes(0), DSBLOCK_DEFAULT
                ContOut = ContOut + 1
            End If
        End If
        
        w1 = False
        w2 = True
        w3 = False
    End If
    
    If BufferCursor.lPlay < B1_B2 And w3 = False Then
        If ContOut <> 0 Then
            ContOut = ContOut + 1
            MusicBuffer.WriteBuffer B2_B3, MYBUFFERPARTSIZE, ebytes(0), DSBLOCK_DEFAULT
            If ContOut = 4 Then GoTo SoundEnd
        Else
            Out = OGGLib.StreamRead(VarPtr(bytes(0)), MYBUFFERPARTSIZE, Read)
            MusicBuffer.WriteBuffer B2_B3, MYBUFFERPARTSIZE, bytes(0), DSBLOCK_DEFAULT
            If Read < MYBUFFERPARTSIZE Then
                MusicBuffer.WriteBuffer Read + B2_B3, (MYBUFFERPARTSIZE - Read), ebytes(0), DSBLOCK_DEFAULT
                ContOut = ContOut + 1
            End If
        End If
        w1 = False
        w2 = False
        w3 = True
    End If
End If
Exit Sub

SoundEnd:
MusicBuffer.Stop
MusicStatus = 0
End Sub

Public Sub InitMusic()
Dim soundDESC As DSBUFFERDESC
soundDESC.lBufferBytes = MYBUFFERSIZE
soundDESC.lFlags = DSBCAPS_CTRLVOLUME + DSBCAPS_CTRLFREQUENCY
Set MusicBuffer = DirectSound.CreateSoundBuffer(soundDESC)

Set OGGLib = New StreamOGG

'---------------- MUSIC ------------------
Dim mypath As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "music_" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
MusicDir = mypath

ExtractFile EXE & "sound\music.dat", MusicDir
End Sub

Public Sub Pause()
MusicBuffer.Stop
MusicStatus = 2
End Sub

Public Sub ResumePlay()
MusicBuffer.Play DSBPLAY_LOOPING
MusicStatus = 1
End Sub

Public Sub DestroyMusic()
On Local Error Resume Next
OGGLib.StreamClose
Set OGGLib = Nothing
Set MusicBuffer = Nothing
End Sub

Public Sub LoadOGGBuffer(ByVal fileogg As String, ByRef buffer As DirectSoundSecondaryBuffer8)
Dim SOgg As StreamOGG
Dim hr As SND_RESULT
Dim bbytes() As Byte
hr = SOgg.StreamOpen(fileogg)
ReDim bbytes(SOgg.BufferSizeBytes)
SOgg.StreamRead VarPtr(bbytes(0)), SOgg.BufferSizeBytes, 0

Dim soundDESC As DSBUFFERDESC
soundDESC.lBufferBytes = SOgg.BufferSizeBytes
soundDESC.lFlags = DSBCAPS_CTRLVOLUME + DSBCAPS_CTRLFREQUENCY
Set buffer = DirectSound.CreateSoundBuffer(soundDESC)

buffer.WriteBuffer 0, SOgg.BufferSizeBytes, bbytes(0), DSBLOCK_ENTIREBUFFER
buffer.SetFrequency SOgg.BitsPerSample
End Sub

Public Sub InitMusicBasic()
Dim soundDESC As DSBUFFERDESC
soundDESC.lBufferBytes = MYBUFFERSIZE
soundDESC.lFlags = DSBCAPS_CTRLVOLUME + DSBCAPS_CTRLFREQUENCY
Set MusicBuffer = DirectSound.CreateSoundBuffer(soundDESC)

Set OGGLib = New StreamOGG
End Sub
