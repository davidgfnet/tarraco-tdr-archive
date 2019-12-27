Attribute VB_Name = "Variables"
Option Explicit

Public Const NumWorlds = 5

'---------------- GLOBALS --------------
Global EXE As String
Global Reg As clsWinReg

Global ExitToWin As Boolean
Global RestartLoop As Boolean

Public Const RegName = "tarraco"

'------------- FILE PARSING ------------
Global PNames() As String
Global PValues() As String

'------------------- DX ----------------
Global DirectX As DirectX8
Global Direct3D As Direct3D8
Global Direct3DX As D3DX8
Global DirectInput As DirectInput8
Global DirectSound As DirectSound8
Global caps As D3DCAPS8
Global CharShader As Long
Global VSCapable As Boolean

'---------------- DS ---------------------
Global DS_Sounds() As Sound
Global PendingSounds() As String
Global DS_Sounds_Plain() As SoundStandard

Global DisableSound As Boolean
Global MusicDir As String

Global DSPrimaryBuffer As DirectSoundPrimaryBuffer8
Global DSPrimaryBufferDesc As DSBUFFERDESC

Global DSListener As DirectSound3DListener8
Global DSEnum As DirectSoundEnum8
Global DSDeviceGUID As String, DSDeviceDesc As String

Global BackgroundMusic() As String
Global BackgroundMusicTrack As Integer
Global BgMusicRandom As Boolean

'----------------- DI -------------------
Global Device As Direct3DDevice8
Global MouseDevice As DirectInputDevice8
Global DI_hevent As Long

Global MouseX As Single, MouseY As Single, MouseZ As Single
Global MouseB0 As Boolean, MouseB1 As Boolean
Global MouseClick0 As Boolean, MouseClick1 As Boolean
Global MouseTexture As Direct3DTexture8, MouseVerts(3) As myVertex
Global CursorW As Single, CursorH As Single

'-------------------- D3D ---------------
Global MainRender As Boolean
Global AuxMenu As Boolean

Global D3DM As D3DDISPLAYMODE
Global AntiAliasLevel As Long
Global DModesArray() As D3DDISPLAYMODE
Global ResolutionArray As Variant
Global RenderStates(20) As Long

'----------------- DRAW SETTINGS ---------------
Global DrawDepth As Integer  '1, 2, 3
Global DistanceFOAppear As Single
Global DistanceFODetail As Single
Global FarViewPlane As Single
Global CursorSpeed As Integer  '1, 2, 3
Global CharDetail As Integer  '1, 2, 3
Global TexQuality As Integer    '1 Low, 2 Normal

'------------------- MATHS -----------------
Global projMatrix As D3DMATRIX
Global viewMatrix As D3DMATRIX
Global FrustumPlanes(7) As D3DPLANE

'---------------------- RESOURCES -----------------
Global GameTextures() As TextureEx
Global GameTexturesSec() As TextureEx
Global GameTexturesTer() As TextureEx

Global PendingTextures() As String
Global FontTexture As Direct3DTexture8
Global FontVertices(3) As myVertex
Global SkyBoxTextures(14) As Direct3DTexture8
Global SkyBoxVertices(19) As myVertexSimple
Global FireTex As Direct3DTexture8
Global MissionTargets() As MissionTarget
Global MissionTargetModel As Model3D
Global SavePoints() As tSavePoint
Global SavePointModel As Model3D

'---------------------- PREV -----------------------
Global LTextMain As Direct3DTexture8
Global LNum As Direct3DTexture8
Global LVertices(19) As myVertex

'-------------------- MENU ------------------
Global MainMenu As cGameMenu
Global ConfigMenu As cGameMenu
Global LoadGameMenu As cGameMenu
Global SaveGameMenu As cGameMenu
Global AuxiliarMenu As cGameMenu
Global GlobalMenuOption As Integer
Global fadeVertices(3) As myVertexAlpha

'-------------------- UI ----------------------
Global MessageState As Integer      '-1 no messages, 1 fading in, 2 showing, 3 fading out, 0 start message
Global MessageTime As Single, MessageFadeTime As Single     'showtime
Global MessageTimer As Double, MessageNum As Long
Global MessageString() As String

Global PaperTex As Direct3DTexture8
Global MessageFont As Direct3DTexture8
Global Coin As Direct3DTexture8
Global UIvertices(3) As myVertex
Global MiniMapVertices(35) As myVertex

Global MiniMapTex(2) As Direct3DTexture8
Global MiniMapBorder As Direct3DTexture8
Global MiniMapIcons As Direct3DTexture8

'------------------- VIDEO -------------------
Global VideoOn As String

'-------------- LOAD STAGE SCREEN --------------
Global LSBar As Direct3DTexture8, LSBar2 As Direct3DTexture8
Global LSVert(3) As myVertex
Global LoadingWorld As Direct3DTexture8

'-------------------- GAME ----------------------
    '-------- GAME FLOW ---------
    Global SlotID As Long
    Global TheGameSlot As GameSlot
    Global SavedGameSlot As GameSlot
    
    Global LevelScenes() As Scene
    Global SceneNumber As Long
    
    Global DisableDoors As Boolean
    Global LockedDoors() As String
    Global FadeState As Integer
    Global FadeTimeMS As Single
    
    Global DoorArray() As tDoor
    Global AreaArray() As tArea
    '------------------------ WORLD -----------------
    Global LastWorldLoaded As Integer
    Global RenderModelsLM(1 To 30) As cAdvMesh
    Global RenderModelsLMAux(1 To 10) As cAdvMesh
    Global RenderModelsLMAuxNum As Integer
    Global WorldProperties(1 To NumWorlds) As WorldProp                 'world descriptor
    Global CollisionFloats(1 To NumWorlds) As CollisionFloat            'collision triangles
    Global CollisionFloatsAux() As Long                         'long aux array for engine.dll
    'Global CollisionFloatsObj(1 To NumWorlds) As CollisionFloat         'temp array for fo
    
    Global MakeFade1 As Boolean, MakeFade2 As Boolean
    Global ResetAllMoves As Boolean
    
        'fixed objects arrays
        Global FixedObjects() As FixedObject        'at startup
        'Global FixedObjectsStage() As FixedObject   'each scene
        
        'array for the startup
        Global FOComplex() As Model3D
        Global FOSimple() As Model3D
        
        'array for the scene (celaned and reloaded each scene)
        'Global FOComplexStage() As Model3D
        'Global FOSimpleStage() As Model3D
        
        Global Fires() As cFire
        
    Global Lights() As myLight

    '-------------------- MAINCHAR --------------------
    Global CharPos As D3DVECTOR             'position now
    Global CharPosBefore As D3DVECTOR       'posistion last frame
    Global CharAngleH As Single             'rotation angle for the char
    Global CharAngleV As Single             'cam vertical angle
    Global CharDistance As Single           'camera - char distance
    Global CharState As Integer             'state animation (walk, run...)
    
    Global MainChar As Cal3DModel
    
    '--------------- OTHER CHARS ---------------
    Global Cal3DModelArray() As Cal3DModel
    Global CharSoldiers(5) As clsDynObj
    Global MovingCharacters() As clsDynObj
        
    '--------------- CAMERA -------------------
    Global CameraPos As D3DVECTOR, realCameraPos As D3DVECTOR, CameraLookAt As D3DVECTOR
    Global vectorUP As D3DVECTOR
    Global CameraType As Integer    '0 3rd person, 1 1st person
    
    '--------------- PHYSICS ---------------
    Global Movement As Single
    Global Jumping As Boolean
    Global InitY As Single
    Global AccelerationTimer As Double
    
    Global TMatrix As D3DMATRIX
    
    Global TimeCounter As Double
    Global FrameAverage As Double, StableFrameAverage As Double
    
    Global CharSpeed As Single
    Global Run As Boolean
    
    Global CurrentCharAnimation As String
    Global CharAnimationTimer As Double

    Global DynObjPositions() As D3DVECTOR
    Global NumDynObjPositions As Long
    
    Global KeyAdd As Boolean, KeySubtract As Boolean

Global rev As Long
