VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StreamOGG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------  Implementation of the OGG VORBIS MUSIC decoder for VB
'--------  Uses vorbisfile.dll vorbis.dll ogg.dll and msvcrt.dll
'--------  Uses clsDLL class

'adapted by davidgf from Monoton Audio Lib (on planet source code) by [rm_code]

Private Type vorbis_info
    version                 As Long
    Channels                As Long
    rate                    As Long
    bitrate_upper           As Long
    bitrate_nominal         As Long
    bitrate_lower           As Long
    bitrate_window          As Long
    codec_setup             As Long
End Type

Private Enum byte_packaging
    bp_little_endian = 0
    bp_big_endian = 1
End Enum

Private Enum word_size
    word_size_8_bit = 1
    word_size_16_bit = 2
End Enum

Private Const BUFFER_SIZE   As Long = 4096

Private clsMSVCRT           As clsDLL
Private clsOgg              As clsDLL
Private clsVorbis           As clsDLL
Private clsVorbisFile       As clsDLL
Private udtInfo             As vorbis_info
Private vf(64& * 1024&)     As Byte
Private btBuffer()          As Byte
Private lngPosInBuffer      As Long
Private lngBufferData       As Long
Private lngDuration         As Long
Private lngCurPos           As Long
Private hFile               As Long
Private blnEOS              As Boolean
Private blnReady            As Boolean
Public BufferSizeBytes As Long

Private Function ogg_init_libs() As Boolean
    Dim bln As Boolean

    Set clsMSVCRT = New clsDLL
    Set clsOgg = New clsDLL
    Set clsVorbis = New clsDLL
    Set clsVorbisFile = New clsDLL

    bln = True

    bln = bln And clsMSVCRT.LoadDLL("msvcrt.dll")
    bln = bln And clsOgg.LoadDLL("ogg.dll")
    bln = bln And clsVorbis.LoadDLL("vorbis.dll")
    bln = bln And clsVorbisFile.LoadDLL("vorbisfile.dll")

    If Not bln Then Exit Function

    clsMSVCRT.IsCDECL = True
    clsOgg.IsCDECL = True
    clsVorbis.IsCDECL = True
    clsVorbisFile.IsCDECL = True

    ogg_init_libs = True
End Function

Private Sub ogg_free_libs()
    clsVorbisFile.UnloadDLL
    clsVorbis.UnloadDLL
    clsOgg.UnloadDLL
    clsMSVCRT.UnloadDLL
End Sub

Public Function FillBuffer() As Boolean
    Dim lngCurrentSection   As Long
    Dim lngRet              As Long

    lngRet = ov_read(VarPtr(vf(0)), VarPtr(btBuffer(0)), BUFFER_SIZE, bp_little_endian, word_size_16_bit, True, lngCurrentSection)

    If lngRet = 0 Then
        blnEOS = True
    ElseIf lngRet > 0 Then
        lngCurPos = lngCurPos + lngRet * 1000 / udtInfo.rate / 2 / udtInfo.Channels
        lngBufferData = lngRet
        lngPosInBuffer = 0
        FillBuffer = True
    End If
End Function

Private Sub Class_Initialize()
    blnReady = ogg_init_libs
    ReDim btBuffer(BUFFER_SIZE - 1) As Byte
End Sub

Private Sub Class_Terminate()
    StreamClose
    ogg_free_libs
End Sub

Public Property Get BitsPerSample() As Integer
    BitsPerSample = 16
End Property

Public Property Get BitsPerSecond() As Long
    BitsPerSecond = udtInfo.bitrate_nominal
End Property

Public Property Get Channels() As Integer
    Channels = udtInfo.Channels
End Property

Public Property Get Duration() As Long
    Duration = lngDuration
End Property

Public Property Get Position() As Long
    Position = lngCurPos
End Property

Public Property Get SamplesPerSecond() As Long
    SamplesPerSecond = udtInfo.rate
End Property

Public Property Get EndOfStream() As Boolean
    EndOfStream = blnEOS
End Property

Public Function StreamClose() As SND_RESULT
    If hFile <> 0 Then
        ov_clear VarPtr(vf(0))
        clsMSVCRT.CallFunc "fclose", hFile
        hFile = 0
        blnEOS = False
        lngBufferData = 0
        lngPosInBuffer = 0

        StreamClose = SND_ERR_SUCCESS
    Else
        StreamClose = SND_ERR_INVALID_SOURCE
    End If
End Function

Public Function StreamOpen(ByVal Source As String) As SND_RESULT
    Dim btFile()    As Byte
    Dim btMode(2)   As Byte
    Dim lngRet      As Long

    If Not blnReady Then
        StreamOpen = SND_ERR_INTERNAL
        Exit Function
    End If

    StreamClose
    
    btFile = StrConv(Source & Chr$(0), vbFromUnicode)

                            ' file...
    btMode(0) = Asc("r")    ' read
    btMode(1) = Asc("b")    ' binary

    hFile = clsMSVCRT.CallFunc("fopen", VarPtr(btFile(0)), VarPtr(btMode(0)))

    If hFile = 0 Then
        StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If ov_open(hFile, VarPtr(vf(0)), 0, 0) < 0 Then
        StreamOpen = SND_ERR_INVALID_SOURCE
        clsMSVCRT.CallFunc "fclose", hFile
        Exit Function
    End If

    udtInfo = ov_info(VarPtr(vf(0)), -1)

    If udtInfo.rate < 1 Then
        ov_clear VarPtr(vf(0))
        clsMSVCRT.CallFunc "fclose", hFile
        StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    If udtInfo.Channels > 2 Then
        ov_clear VarPtr(vf(0))
        clsMSVCRT.CallFunc "fclose", hFile
        StreamOpen = SND_ERR_INVALID_SOURCE
        Exit Function
    End If

    lngRet = ov_pcm_total(VarPtr(vf(0)), -1)
    BufferSizeBytes = lngRet

    lngDuration = lngRet / udtInfo.rate * 1000
    lngCurPos = 0

    StreamOpen = SND_ERR_SUCCESS
End Function

Public Function StreamRead(ByVal buffer_ptr As Long, ByVal buffer_len As Long, buffer_read As Long) As SND_RESULT
    If Not blnReady Then
        StreamRead = SND_ERR_INTERNAL
        Exit Function
    End If

    StreamRead = SND_ERR_SUCCESS

    buffer_read = 0

    Do While buffer_read < buffer_len

         If lngBufferData = 0 Then
            If Not FillBuffer Then
                StreamRead = SND_ERR_END_OF_STREAM
                Exit Function
            End If

        ElseIf (lngBufferData - lngPosInBuffer) < (buffer_len - buffer_read) Then
            If 0 < (lngBufferData - lngPosInBuffer) Then

                If 0 = IsBadReadPtr(ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, _
                                    lngBufferData - lngPosInBuffer) Then

                    If 0 = IsBadWritePtr(ByVal buffer_ptr + buffer_read, _
                                         lngBufferData - lngPosInBuffer) Then

                        CpyMem ByVal buffer_ptr + buffer_read, _
                               ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, _
                               lngBufferData - lngPosInBuffer

                    End If
                End If

                buffer_read = buffer_read + (lngBufferData - lngPosInBuffer)
            End If

            If Not FillBuffer Then
                StreamRead = SND_ERR_END_OF_STREAM
                Exit Function
            End If

        Else
            If 0 = IsBadReadPtr(ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, _
                                buffer_len - buffer_read) Then

                If 0 = IsBadWritePtr(ByVal buffer_ptr + buffer_read, _
                                     buffer_len - buffer_read) Then

                    CpyMem ByVal buffer_ptr + buffer_read, _
                           ByVal VarPtr(btBuffer(0)) + lngPosInBuffer, _
                           buffer_len - buffer_read

                End If
            End If

            lngPosInBuffer = lngPosInBuffer + (buffer_len - buffer_read)
            buffer_read = buffer_read + (buffer_len - buffer_read)

        End If
    Loop
End Function

Public Function StreamSeek(ByVal Value As Long, ByVal seek_mode As SND_SEEK_MODE) As SND_RESULT
    Dim dblTime     As Double

    If Not blnReady Then
        StreamSeek = SND_ERR_INTERNAL
        Exit Function
    End If

    Select Case seek_mode

        Case SND_SEEK_PERCENT
            If Value < 0 Or Value > 99 Then
                StreamSeek = SND_ERR_OUT_OF_RANGE
                Exit Function
            End If

            dblTime = (Duration / 1000) / 100 * Value

        Case SND_SEEK_SECONDS
            If Value < 0 Or Value > (Duration / 1000) Then
                StreamSeek = SND_ERR_OUT_OF_RANGE
                Exit Function
            End If

            dblTime = Value

    End Select

    lngCurPos = dblTime * 1000
    ov_time_seek VarPtr(vf(0)), dblTime

    lngBufferData = 0
    lngPosInBuffer = 0

    blnEOS = False

    StreamSeek = SND_ERR_SUCCESS
End Function

Private Function ov_time_seek(ByVal ovf_struct As Long, ByVal seconds As Double) As Long
    Dim lngTimeLo   As Long
    Dim lngTimeHi   As Long

    CpyMem lngTimeLo, seconds, 4
    CpyMem lngTimeHi, ByVal VarPtr(seconds) + 4, 4

    ov_time_seek = clsVorbisFile.CallFunc("ov_time_seek", ovf_struct, lngTimeLo, lngTimeHi)
End Function

Private Function ov_read(ByVal ovf_struct As Long, ByVal output_ptr As Long, _
    ByVal Length As Long, ByVal bigendianp As byte_packaging, _
    ByVal word As word_size, ByVal signed As Boolean, _
    ByRef bitstream As Long) As Long

    ov_read = clsVorbisFile.CallFunc("ov_read", ovf_struct, output_ptr, Length, bigendianp, word, Abs(signed), VarPtr(bitstream))
End Function

Private Function ov_pcm_total(ByVal ovf_struct As Long, ByVal link As Long) As Long
    ov_pcm_total = clsVorbisFile.CallFunc("ov_pcm_total", ovf_struct, link)
End Function

Private Function ov_info(ByVal ovf_struct As Long, ByVal link As Long) As vorbis_info
    Dim ptr     As Long
    Dim udt     As vorbis_info

    ptr = clsVorbisFile.CallFunc("ov_info", ovf_struct, link)
    If ptr = 0 Then Exit Function

    CpyMem udt, ByVal ptr, Len(udt)

    ov_info = udt
End Function

Private Function ov_open(ByVal hFile As Long, ByVal ovf_struct As Long, ByVal initial_ptr As Long, ByVal ibytes As Long) As Long
ov_open = clsVorbisFile.CallFunc("ov_open", hFile, ovf_struct, initial_ptr, ibytes)
End Function

Private Sub ov_clear(ByVal ovf_struct As Long)
clsVorbisFile.CallFunc "ov_clear", ovf_struct
End Sub

