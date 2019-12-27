Attribute VB_Name = "Declarations"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal dwLen As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

'--- FreeImage Lib
Public Declare Function FreeImage_Load Lib "freeimage.dll" Alias "_FreeImage_Load@12" ( _
           ByVal fif As FREE_IMAGE_FORMAT, _
           ByVal filename As String, _
  Optional ByVal Flags As Long = 0) As Long
  
Public Declare Function FreeImage_GetWidth Lib "freeimage.dll" Alias "_FreeImage_GetWidth@4" (ByVal dib As Long) As Long
Public Declare Function FreeImage_GetHeight Lib "freeimage.dll" Alias "_FreeImage_GetHeight@4" (ByVal dib As Long) As Long
Public Declare Sub FreeImage_Unload Lib "freeimage.dll" Alias "_FreeImage_Unload@4" (ByVal dib As Long)

Public Enum FREE_IMAGE_FORMAT
   FIF_UNKNOWN = -1: FIF_BMP = 0: FIF_ICO = 1: FIF_JPEG = 2: FIF_JNG = 3: FIF_KOALA = 4
   FIF_LBM = 5: FIF_IFF = FIF_LBM: FIF_MNG = 6: FIF_PBM = 7: FIF_PBMRAW = 8: FIF_PCD = 9
   FIF_PCX = 10: FIF_PGM = 11: FIF_PGMRAW = 12: FIF_PNG = 13: FIF_PPM = 14: FIF_PPMRAW = 15
   FIF_RAS = 16: FIF_TARGA = 17: FIF_TIFF = 18: FIF_WBMP = 19: FIF_PSD = 20: FIF_CUT = 21
   FIF_XBM = 22: FIF_XPM = 23: FIF_DDS = 24: FIF_GIF = 25: FIF_HDR = 26: FIF_FAXG3 = 27: FIF_SGI = 28
End Enum
'----------------

Public Const Pi As Double = 3.14159265358979    'or magic formula: atn(1) * 4
Public Const MaxCameraDistance As Single = 6
Public Const MinCameraDistance As Single = 4
Public Const MinVerticalAngle As Single = 20
Public Const MaxVerticalAngle As Single = 45
Public Const MainCharDimensions As Single = 0.2

Public Const WorldAcceleration As Single = 9.8

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Type Model3D
    MeshName As String
    
    Mesh As D3DXMesh
    Materials() As D3DMATERIAL8
    NumMaterials As Long
    TexturesNames() As String
    
    tmp_adj As D3DXBuffer
    tmp_mat As D3DXBuffer
End Type

Public Type CollisionFloat
    numverts As Long
    vertices() As D3DVECTOR
End Type

Public Type WorldProp
    Mesh As String
    SkyBox As Integer
    AmbientValues As D3DVECTOR  'color as d3dvector :P
    UnlitAmbientValues As D3DVECTOR  'the ambient values for the non-lightmapped polys
    State As String
End Type

Public Type myVertex        'texture  and transformed vert
    x As Single
    y As Single
    z As Single
    rhw As Single
    Color As Long
    tu As Single
    tv As Single
End Type
Public Const myVertexFVF = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_TEX1

Public Type myVertexSimple  'texture but not transformes
    x As Single
    y As Single
    z As Single
    Color As Long
    tu As Single
    tv As Single
End Type
Public Const myVertexFVFSimple = D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1

Public Type myVertexAlpha       'not texture, bu transformed and alpha
    x As Single
    y As Single
    z As Single
    rhw As Single
    Color As Long
End Type
Public Const myVertexAlphaFVF = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE

Public Type salida
    respuesta As Integer
    puntocolision As D3DVECTOR
End Type

Public Declare Sub hprocess Lib "engine.dll" (ByRef Tri As D3DVECTOR, seg As D3DVECTOR, ByVal numtri As Long, ByRef collide As salida)
Public Declare Sub process Lib "engine.dll" (ByRef cam As D3DVECTOR, Tri As D3DVECTOR, ByVal numtri As Long, ByRef pos As D3DVECTOR, ByRef Pos2 As D3DVECTOR, ByRef upv As D3DVECTOR, ByVal distance As Double, ByVal angle As Double, ByVal AngleH As Double, ByVal disfromcol As Single, ByVal interpolationspeed As Single, ByVal midh As Single, ByVal hih As Single, ByVal avrframe As Single, ByVal speed As Single, ByVal dimensions As Single, ByRef tris As Long, ByVal CameraType As Long)
Public Declare Sub segintersect Lib "engine.dll" (ByRef origin As D3DVECTOR, ByRef Direction As D3DVECTOR, ByRef Tri As D3DVECTOR, ByVal numtris As Long, ByVal numsegs As Long, ByRef collide As salida)
Public Declare Sub segintersectfast Lib "engine.dll" (ByRef origin As D3DVECTOR, ByRef Direction As D3DVECTOR, ByRef Tri As D3DVECTOR, ByVal numtris As Long, ByVal numsegs As Long, ByRef collide As salida)
Public Declare Sub lib_compute_normals Lib "engine.dll" Alias "computenormals" (ByRef verts As D3DVECTOR, ByVal numverts As Long, ByRef normals As D3DVECTOR)
Public Declare Sub Visible Lib "engine.dll" Alias "visible" (ByRef pointfrom As D3DVECTOR, ByRef pointto As D3DVECTOR, ByRef Tri As D3DVECTOR, ByVal numtris As Long, ByVal numsegs As Long, ByRef collide As salida)
Public Declare Sub formalisetex Lib "engine.dll" (ByRef entrada As Byte, ByVal width As Long, ByVal height As Long, ByVal pith As Long, ByRef salida As Byte)

Public Type Sphere
    center As D3DVECTOR
    radius As Single
End Type

Public Type DMode
    Ancho As Long
    Alto As Long
    Frecuencia As Long
    Formato As Long
End Type

Public Type Sound
    file As String
    soundname As String
    buffer3d As DirectSound3DBuffer8
    buffersec As DirectSoundSecondaryBuffer8
    Desc As DSBUFFERDESC
    freezed As Boolean
    pos As D3DVECTOR
    orient As D3DVECTOR
End Type

Public Type SoundStandard
    file As String
    soundname As String
    buffer As DirectSoundSecondaryBuffer8
    Desc As DSBUFFERDESC
    freezed As Boolean
    loop As Boolean
End Type

Public Type FixedObject
    MeshID As Long
    
    Position As D3DVECTOR
    RotationY As Single
    transformedBoundingSphere As Sphere
    unimesh As Byte
    WorldID As Integer
    
    id As String
End Type

Public Type gamecondition
    type As String
    varg As D3DVECTOR
    varg2 As D3DVECTOR
    varg3 As D3DVECTOR
    larg As Long
    larg2 As Long
    sarg As String
    sarg2 As String
End Type

Public Type gameeffect
    type As String
    varg As D3DVECTOR
    varg2 As D3DVECTOR
    varg3 As D3DVECTOR
    larg As Long
    larg2 As Long
    sarg As String
    sarg2 As String
End Type

Public Type Scene
    Conditions() As gamecondition
    Effects() As gameeffect
    Enabled As Boolean
    id As Long
End Type

Public Type MissionTarget
    id As String
    
    Visible As Boolean
    
    Position As D3DVECTOR
    WorldID As Integer
    radius As Single
    height As Single
    
    timer As Single
    fadeduration As Single
End Type

Public Type tSavePoint
    WorldID As Long
    Position As D3DVECTOR
    id As String
    Visible As Boolean
End Type

Public Type TextureEx
    filename As String
    texture As Direct3DTexture8
End Type

Public Enum LightType
    Directional
    Target
    Omni
End Enum

Public Type myLight
    type As LightType
    Position As D3DVECTOR
    Direction As D3DVECTOR
    'CastShadows As Boolean
    'ShadowPoint As D3DVECTOR
    Range As Single
    Color As D3DCOLORVALUE
    Phi As Single           'outer cone
    Theta As Single         'inner cone   (theta<phi)
    WorldID As Long
End Type

Public Type tDoor
    World As Integer
    
    id As Integer
    Position As D3DVECTOR
    RotationH As Single
End Type
Public Type tArea
    World As Integer

    DoorName As String
    pos As D3DVECTOR
    Pos2 As D3DVECTOR
    NewWorld As Integer
    DoorId As Integer
End Type

'-------------- SAVING GAMES --------------
Public Type GameSlot
    MissionLevel As Long
    Coins As Long
    
    Health As Long
    
    ObjectsID() As Long
    NumObjects As Long
    
    Position As D3DVECTOR   'position of the save point / char
    RotationH As Single     'best angle rotation for the save point         '0 camera abaix, 270 camera a l'esquerra 180 dal 90 dreta
    
    WorldID As Integer      'which world is the char
End Type

'--------- MUSIC!!! -------

Public Declare Function FreeLibrary Lib "kernel32" (ByVal hModule As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function IsBadReadPtr Lib "kernel32" (ptr As Any, ByVal ucb As Long) As Long
Public Declare Function IsBadWritePtr Lib "kernel32" (ptr As Any, ByVal ucb As Long) As Long
Public Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal cBytes As Long)

Public Enum SND_RESULT
    SND_ERR_SUCCESS = 0
    SND_ERR_INVALID_SOURCE
    SND_ERR_INTERNAL
    SND_ERR_OUT_OF_RANGE
    SND_ERR_END_OF_STREAM
End Enum

Public Enum SND_SEEK_MODE
    SND_SEEK_PERCENT = 0
    SND_SEEK_SECONDS
End Enum

Public Type DynamicObject
    Position As D3DVECTOR
    
End Type

'--------- VIDEO ENGINE ----------
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

