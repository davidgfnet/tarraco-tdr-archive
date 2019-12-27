Attribute VB_Name = "VideoEngine"
Option Explicit

' video engine by davidgf
' frame streaming into texture resource
' using Avi functions from Windows
' Uses DirectShow filters ;)

Private Started As Boolean
Private StartTime As Double

Public Declare Sub AVIFileInit Lib "avifil32.dll" ()
Public Declare Function AVIFileOpen Lib "avifil32.dll" (ByRef ppfile As Long, ByVal szFile As String, ByVal uMode As Long, ByVal pclsidHandler As Long) As Long  'HRESULT
Public Declare Function AVIFileInfo Lib "avifil32.dll" (ByVal pfile As Long, pfi As AVI_FILE_INFO, ByVal lSize As Long) As Long 'HRESULT
Public Declare Function AVIStreamGetFrameOpen Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef bih As Any) As Long                                                                 'returns pointer to GETFRAME object on success (or NULL on error)
Public Declare Function AVIStreamGetFrame Lib "avifil32.dll" (ByVal pGetFrameObj As Long, ByVal lPos As Long) As Long                                                                 'returns pointer to packed DIB on success (or NULL on error)
Public Declare Function AVIStreamGetFrameClose Lib "avifil32.dll" (ByVal pGetFrameObj As Long) As Long ' returns zero on success (error number) after calling this function the GETFRAME object pointer is invalid
Public Declare Function AVIFileGetStream Lib "avifil32.dll" (ByVal pfile As Long, ByRef ppaviStream As Long, ByVal fccType As Long, ByVal lParam As Long) As Long
Public Declare Function AVIStreamInfo Lib "avifil32.dll" (ByVal pAVIStream As Long, ByRef psi As AVI_STREAM_INFO, ByVal lSize As Long) As Long
Public Declare Function AVIStreamStart Lib "avifil32.dll" (ByVal pavi As Long) As Long
Public Declare Function AVIStreamLength Lib "avifil32.dll" (ByVal pavi As Long) As Long
Public Declare Function AVIStreamRelease Lib "avifil32.dll" (ByVal pavi As Long) As Long 'ULONG
Public Declare Function AVIFileRelease Lib "avifil32.dll" (ByVal pfile As Long) As Long
Public Declare Sub AVIFileExit Lib "avifil32.dll" ()

Public Type BITMAPINFOHEADER
   biSize As Long: biWidth As Long: biHeight As Long: biPlanes As Integer:
   biBitCount As Integer: biCompression As Long: biSizeImage As Long
   biXPelsPerMeter As Long: biYPelsPerMeter As Long: biClrUsed As Long
   biClrImportant As Long
End Type

Public Type AVI_RECT
    left As Long: top As Long: right As Long: bottom As Long
End Type

Public Type AVI_STREAM_INFO
    fccType As Long: fccHandler As Long: dwFlags As Long: dwCaps As Long:
    wPriority As Integer: wLanguage As Integer: dwScale As Long: dwRate As Long:
    dwStart As Long: dwLength As Long: dwInitialFrames As Long: dwSuggestedBufferSize As Long
    dwQuality As Long: dwSampleSize As Long: rcFrame As AVI_RECT: dwEditCount As Long
    dwFormatChangeCount As Long: szName As String * 64
End Type

Public Type AVI_FILE_INFO
    dwMaxBytesPerSecond As Long: dwFlags As Long: dwCaps As Long: dwStreams As Long
    dwSuggestedBufferSize As Long: dwWidth As Long: dwHeight As Long
    dwScale As Long: dwRate As Long: dwLength As Long: dwEditCount As Long
    szFileType As String * 64
End Type

Public Type AVI_COMPRESS_OPTIONS
    fccType As Long: fccHandler As Long: dwKeyFrameEvery As Long: dwQuality As Long
    dwBytesPerSecond As Long: dwFlags As Long: lpFormat As Long
    cbFormat As Long: lpParms As Long: cbParms As Long: dwInterleaveEvery As Long
End Type

Global Const AVIERR_OK             As Long = 0&
Private Const SEVERITY_ERROR       As Long = &H80000000
Private Const FACILITY_ITF         As Long = &H40000
Private Const AVIERR_BASE          As Long = &H4000
Global Const OF_SHARE_DENY_WRITE   As Long = &H20
Global Const BI_RGB                As Long = 0
Global Const streamtypeVIDEO       As Long = 1935960438

Private pAVIFile As Long 'pointer to the avi file in the mem
Private pAVIStream As Long
Private pGetFrameObj As Long

Private firstFrame As Long
Private NumFrames As Long

Private fileInfo As AVI_FILE_INFO
Private streamInfo As AVI_STREAM_INFO
Private bih As BITMAPINFOHEADER

Private AudioFile As String

Public Sub OpenVideo(ByVal file As String)
Dim res As Long

Call AVIFileInit

res = AVIFileOpen(pAVIFile, file, OF_SHARE_DENY_WRITE, 0&)

res = AVIFileGetStream(pAVIFile, pAVIStream, streamtypeVIDEO, 0)
    
firstFrame = AVIStreamStart(pAVIStream)
NumFrames = AVIStreamLength(pAVIStream)
   
res = AVIFileInfo(pAVIFile, fileInfo, Len(fileInfo))

With bih
    .biBitCount = 24
    .biClrImportant = 0
    .biClrUsed = 0
    .biCompression = BI_RGB
    .biHeight = streamInfo.rcFrame.bottom - streamInfo.rcFrame.top
    .biPlanes = 1
    .biSize = 40
    .biWidth = streamInfo.rcFrame.right - streamInfo.rcFrame.left
    .biXPelsPerMeter = 0
    .biYPelsPerMeter = 0
    .biSizeImage = (((.biWidth * 3) + 3) And &HFFFC) * .biHeight
End With

pGetFrameObj = AVIStreamGetFrameOpen(pAVIStream, bih)

AudioFile = Replace(file, ".avi", "_s.ogg")
End Sub

Public Sub PlayVideo()
Dim FrameTime As Double, pDIB As Long, frame As Long
Dim verts(3) As myVertex, PauseTimer As Double

verts(0) = AssignMV(0, 0, 0, 0, 0)
verts(1) = AssignMV(D3DM.width * 1280 / 900, 0, 0, 1, 0)
verts(2) = AssignMV(0, D3DM.height * 1024 / 600, 0, 0, 1)
verts(3) = AssignMV(D3DM.width * 1280 / 900, D3DM.height * 1024 / 600, 0, 1, 1)

MusicEngine.StopMusic

Dim FrameTex As Direct3DTexture8
Dim dib As FDIBPointer

PauseTimer = GetTickCount()
Call FreezeAll

MusicEngine.PlayMusic AudioFile
StartTime = GetTickCount()

Do While frame < NumFrames
    If frame = 0 Then GoTo LoadFrame

    Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    Device.BeginScene

    Device.SetTexture 0, FrameTex
    Device.SetVertexShader myVertexFVF
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, verts(0), Len(verts(0))

    Device.EndScene
    Device.Present ByVal 0, ByVal 0, 0, ByVal 0
    DoEvents
    
    FrameTime = GetTickCount()   ' the frame has been presented, load next while waiting to present again
    
LoadFrame:
    frame = (GetTickCount() - StartTime) / 40    '40 milisends a frame ;)
    If frame >= NumFrames Then GoTo LoopExit
    
    pDIB = AVIStreamGetFrame(pGetFrameObj, frame)
    Set dib = New FDIBPointer
    dib.CreateFromPackedDIBPointer (pDIB)
    
    Set FrameTex = Nothing
    dib.CreateTexture FrameTex
    
    Set dib = Nothing
    
    Call MusicEngine.RenderTime
Loop

LoopExit:
MusicEngine.StopMusic

Call SetCorrectRenderStates(RGB(WorldProperties(TheGameSlot.WorldID).AmbientValues.x, WorldProperties(TheGameSlot.WorldID).AmbientValues.y, WorldProperties(TheGameSlot.WorldID).AmbientValues.z))
Call UnFreezeAll
AccelerationTimer = AccelerationTimer + (GetTickCount() - PauseTimer)
End Sub

Public Sub CloseVideo()
On Local Error Resume Next

Call AVIStreamGetFrameClose(pGetFrameObj)
Call AVIStreamRelease(pAVIStream)
Call AVIFileRelease(pAVIFile)
Call AVIFileExit
End Sub
