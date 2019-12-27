Attribute VB_Name = "SoundEngine"
Option Explicit

'------ 3D SOUND ----------

Public Sub UpdateListenerSettings()
DSListener.SetPosition CharPos.x, CharPos.y, CharPos.z, DS3D_DEFERRED
DSListener.SetOrientation sin(CharAngleH * Pi / 180), 0, Cos(CharAngleH * Pi / 180), 0, 1, 0, DS3D_DEFERRED
DSListener.CommitDeferredSettings
End Sub

Public Sub CreateSound(ByVal file As String, ByVal soundname As String)
Dim index As Integer
index = UBound(DS_Sounds) + 1
ReDim Preserve DS_Sounds(index)

DS_Sounds(index).Desc.lFlags = DSBCAPS_CTRL3D + DSBCAPS_CTRLVOLUME
DS_Sounds(index).Desc.guid3DAlgorithm = GUID_DS3DALG_DEFAULT
Set DS_Sounds(index).buffersec = DirectSound.CreateSoundBufferFromFile(file, DS_Sounds(index).Desc)
DS_Sounds(index).buffersec.SetVolume 0

Set DS_Sounds(index).buffer3d = DS_Sounds(index).buffersec.GetDirectSound3DBuffer()
DS_Sounds(index).buffer3d.SetConeAngles 90, 160, DS3D_IMMEDIATE
DS_Sounds(index).buffer3d.SetConeOutsideVolume -50, DS3D_IMMEDIATE
DS_Sounds(index).buffer3d.SetMaxDistance 15, DS3D_IMMEDIATE
DS_Sounds(index).buffer3d.SetMinDistance 3, DS3D_IMMEDIATE

DS_Sounds(index).file = file
DS_Sounds(index).soundname = soundname
End Sub

Public Sub PlaySound(ByVal soundname As String, pos As D3DVECTOR, orient As D3DVECTOR, Optional ByVal nors As Boolean = False)
Dim x As Integer, orient2 As D3DVECTOR
orient2 = Normalize(orient)
For x = 1 To UBound(DS_Sounds)
    If DS_Sounds(x).soundname = soundname Then
        DS_Sounds(x).buffer3d.SetConeOrientation orient2.x, orient2.y, orient2.z, DS3D_IMMEDIATE
        DS_Sounds(x).buffer3d.SetPosition pos.x, pos.y, pos.z, DS3D_IMMEDIATE
        DS_Sounds(x).pos = pos
        DS_Sounds(x).orient = orient2
        
        If nors Then
            DS_Sounds(x).buffersec.Stop
            DS_Sounds(x).buffersec.SetCurrentPosition 0
        End If
        DS_Sounds(x).buffersec.Play DSBPLAY_DEFAULT
        Exit Sub
    End If
Next
End Sub

Public Sub StopSound(ByVal soundname As String)
Dim x As Integer
For x = 1 To UBound(DS_Sounds)
    If DS_Sounds(x).soundname = soundname Then
        DS_Sounds(x).buffersec.Stop
        Exit Sub
    End If
Next
End Sub

Public Function SoundStopped(ByVal soundname As String) As Boolean
Dim x As Integer, cur As DSCURSORS
For x = 1 To UBound(DS_Sounds)
    If DS_Sounds(x).soundname = soundname Then
        DS_Sounds(x).buffersec.GetCurrentPosition cur
        If cur.lPlay = 0 Then SoundStopped = True
        Exit Function
    End If
Next
End Function

Public Sub StopAllSounds()
Dim x As Integer
For x = 1 To UBound(DS_Sounds)
    DS_Sounds(x).buffersec.Stop
Next
End Sub

Public Sub DestroySounds()
On Local Error Resume Next
Call StopAllSounds
Dim x As Integer
For x = 1 To UBound(DS_Sounds)
    Set DS_Sounds(x).buffer3d = Nothing
    Set DS_Sounds(x).buffersec = Nothing
Next
ReDim DS_Sounds(0)
End Sub

Public Function IsPlayingSound(ByVal sndname As String) As Boolean
Dim x As Integer, st As Long
For x = 1 To UBound(DS_Sounds_Plain)
    If DS_Sounds(x).soundname = sndname Then
        st = DS_Sounds(x).buffersec.GetStatus()
        If (st And DSBSTATUS_PLAYING) <> 0 Then IsPlayingSound = True
        Exit Function
    End If
Next
End Function


'------ Standard SOUND ----------

Public Sub CreateStandardSound(ByVal file As String, ByVal soundname As String)
Dim index As Integer
index = UBound(DS_Sounds_Plain) + 1
ReDim Preserve DS_Sounds_Plain(index)

DS_Sounds_Plain(index).Desc.lFlags = DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME
Set DS_Sounds_Plain(index).buffer = DirectSound.CreateSoundBufferFromFile(file, DS_Sounds_Plain(index).Desc)
DS_Sounds_Plain(index).file = file
DS_Sounds_Plain(index).soundname = soundname
End Sub

Public Sub PlayStandardSound(ByVal soundname As String, ByVal Looping As Boolean)
Dim x As Integer
For x = 1 To UBound(DS_Sounds_Plain)
    If DS_Sounds_Plain(x).soundname = soundname Then
        DS_Sounds_Plain(x).loop = Looping
        If Looping = False Then
            DS_Sounds_Plain(x).buffer.Play DSBPLAY_DEFAULT
        Else
            DS_Sounds_Plain(x).buffer.Play DSBPLAY_LOOPING
        End If
        Exit Sub
    End If
Next
End Sub

Public Sub StopStandardSound(ByVal soundname As String)
Dim x As Integer
For x = 1 To UBound(DS_Sounds_Plain)
    If DS_Sounds_Plain(x).soundname = soundname Then
        DS_Sounds_Plain(x).buffer.Stop
    End If
Next
End Sub

Public Sub DestroyStandardSound()
Dim x As Integer
For x = 1 To UBound(DS_Sounds_Plain)
    Set DS_Sounds_Plain(x).buffer = Nothing
Next
Erase DS_Sounds_Plain
ReDim DS_Sounds_Plain(0)
End Sub

Public Function IsPlayingStSound(ByVal sndname As String) As Boolean
Dim x As Integer, st As Long
For x = 1 To UBound(DS_Sounds_Plain)
    If DS_Sounds_Plain(x).soundname = sndname Then
        st = DS_Sounds_Plain(x).buffer.GetStatus()
        If (st And DSBSTATUS_PLAYING) <> 0 Then IsPlayingStSound = True
        Exit Function
    End If
Next
End Function

Public Sub FreezeSounds()
Dim x As Long
For x = 1 To UBound(DS_Sounds_Plain)
    DS_Sounds_Plain(x).freezed = False
    If IsPlayingStSound(DS_Sounds_Plain(x).soundname) Then
        DS_Sounds_Plain(x).freezed = True
        StopStandardSound DS_Sounds_Plain(x).soundname
    End If
Next
For x = 1 To UBound(DS_Sounds)
    DS_Sounds(x).freezed = False
    If IsPlayingSound(DS_Sounds(x).soundname) Then
        DS_Sounds(x).freezed = True
        StopSound DS_Sounds(x).soundname
    End If
Next
End Sub

Public Sub UnFreezeSounds()
Dim x As Long
For x = 1 To UBound(DS_Sounds_Plain)
    If DS_Sounds_Plain(x).freezed Then
        PlayStandardSound DS_Sounds_Plain(x).soundname, DS_Sounds_Plain(x).loop
    End If
Next
For x = 1 To UBound(DS_Sounds)
    If DS_Sounds(x).freezed Then
        PlaySound DS_Sounds(x).soundname, DS_Sounds(x).pos, DS_Sounds(x).orient
    End If
Next
End Sub
