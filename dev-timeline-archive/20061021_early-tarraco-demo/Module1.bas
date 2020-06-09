Attribute VB_Name = "Module1"
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type


Global Zone As Integer, ALevel As Integer
Global hevent As Long
    
Option Explicit

Global ress As Single
Global tris1() As Long

Global FrustumPlanes(0 To 5) As D3DPLANE
    Dim temporizador As Double

Global counter As Double
Global p1 As Double, p2 As Double
Global AverageFrame As Double
Global CPUMode As String

Global SWidth As Single, SHeight As Single

Global YInicial As Double
Global VSalto As Double
Public sleeptimer As Double
Public ATimer2 As Double
Public ATimer1 As Double
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Saltar As Boolean
Public TimerSalto As Double, TimerSalto2 As Double
Public vvertical As Single
Public BJDO As Boolean
Dim v4Light As D3DVECTOR4

Global UpVector As D3DVECTOR
Global CameraCollision As Boolean

Global ACCTimer As tdouble

Public vbuf As Direct3DVertexBuffer8
Public Type CUSTOMVERTEX
    x As Single
    y As Single
    z  As Single
    rhw As Single
    color As Long
End Type
Public Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE)

Public Type myvert
    x As Single
    y As Single
    z  As Single
    tu As Single
    tv As Single
End Type
Public Const myvert_fvf = D3DFVF_XYZ Or D3DFVF_TEX1

Public VArray(20) As myvert
Public CubeTex(4) As Direct3DTexture8

Public Type salida
    respuesta As Integer
    puntocolision As D3DVECTOR
End Type

Public Type tdouble
    n As Double
End Type

Public Declare Sub colision3d Lib "dx4vbmathslib.dll" (ByRef tri1 As D3DVECTOR, ByRef tri2 As D3DVECTOR, ByVal numtri1 As Long, ByVal numtri2 As Long, ByRef collide As salida)
Public Declare Sub colision3dseg Lib "dx4vbmathslib.dll" (ByRef tri As D3DVECTOR, seg As D3DVECTOR, ByVal numtri As Long, ByRef collide As salida)
Public Declare Sub cameracoords Lib "dx4vbmathslib.dll" (ByRef cam As D3DVECTOR, tri As D3DVECTOR, ByVal numtri As Long, ByRef pos As D3DVECTOR, ByRef upv As D3DVECTOR, ByVal distance As Double, ByVal angle As Double, ByVal angleh As Double, ByVal disfromcol As Single)
Public Declare Sub hprocess Lib "dx4vbmathslib.dll" (ByRef tri As D3DVECTOR, seg As D3DVECTOR, ByVal numtri As Long, ByRef collide As salida)
Public Declare Sub smoothcamera Lib "dx4vbmathslib.dll" (ByRef cam As D3DVECTOR, tri As D3DVECTOR, ByVal numtri As Long, ByRef pos As D3DVECTOR, ByRef upv As D3DVECTOR, ByVal distance As Double, ByVal angle As Double, ByVal angleh As Double, ByVal disfromcol As Single, ByVal interpolation As Single)

Public Declare Sub process Lib "dx4vbmathslib.dll" (ByRef cam As D3DVECTOR, tri As D3DVECTOR, ByVal numtri As Long, ByRef pos As D3DVECTOR, ByRef pos2 As D3DVECTOR, ByRef upv As D3DVECTOR, ByVal distance As Double, ByVal angle As Double, ByVal angleh As Double, ByVal disfromcol As Single, ByVal interpolationspeed As Single, ByVal midh As Single, ByVal hih As Single, ByVal avrframe As Single, ByVal speed As Single, ByVal dimensions As Single, ByRef tris As Long)

Public Declare Sub colision3dsphere Lib "dx4vbmathslib.dll" (ByRef tri As D3DVECTOR, ByRef centre As D3DVECTOR, ByVal numtri As Long, ByVal Radius As Single, ByRef collide As salida)
'__declspec( dllexport ) _stdcall fisica(D3DVECTOR *tri, D3DVECTOR *coords, D3DVECTOR *cam, long numtri, float height, D3DVECTOR *upv, double distance, double angle, double angleh, float disfromcol, salida *collide)
'Public Declare Sub fisica Lib "dx4vbmathslib.dll" (ByRef tri As D3DVECTOR, coords As D3DVECTOR, cam As D3DVECTOR, ByVal numtri As Long, ByVal height As Single, ByRef upv As D3DVECTOR, ByVal distance As Double, ByVal angle As Double, ByVal angleh As Double, ByVal disfromcol As Single, ByVal timer As Single, ByVal bajado As Single, ByVal avance As Integer)

Global triangulos() As D3DVECTOR
Global tripersona() As D3DVECTOR
Global DModeSelected As Integer

Dim cox As Single, coy As Single, coz As Single
Dim cox2 As Single, coy2 As Single, coz2 As Single
Dim camx As Single, camy As Single, camz As Single
Dim camx2 As Single, camy2 As Single, camz2 As Single
Dim camx3 As Single, camy3 As Single, camz3 As Single

Dim Frame1() As D3DVERTEX
Dim Frame2() As D3DVERTEX
Dim Frame3() As D3DVERTEX

Global Frames() As D3DVERTEX
Global Meshes() As D3DXMesh

Global tween As Single
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public HayColision As Boolean
Public DX As DirectX8
Public Dx3D As Direct3D8
Public DInput As DirectInput8
Public Device As Direct3DDevice8
Public DispMode As D3DDISPLAYMODE
Public D3D As D3DX8
Public Objeto1 As D3DXMesh
Public Objeto1_C As D3DXMesh
Public Objeto2(2) As D3DXMesh
Public ObjetoCollide As D3DXMesh
Public Objeto2Render As D3DXMesh
Public MtrlBuffer1 As D3DXBuffer
Public MtrlBuffer2 As D3DXBuffer
Public MtrlBufferC As D3DXBuffer
Dim MeshMaterials() As D3DMATERIAL8
Dim MeshTextures() As Direct3DTexture8
Dim MeshTexturesOrd() As Direct3DBaseTexture8
Dim MeshMaterialsOrd() As D3DMATERIAL8
Dim MeshTexturesList() As String
Dim MeshTexturesListOrd() As String
Dim MeshMaterials1() As D3DMATERIAL8
Dim MeshMaterials2() As D3DMATERIAL8
Dim MeshTextures1() As Direct3DTexture8
Dim MeshTextures2() As Direct3DTexture8
Dim MeshMaterialsC() As D3DMATERIAL8
Dim MeshTexturesC() As Direct3DTexture8
Dim nMaterialesC As Long
Dim nMateriales As Long
Dim nMateriales2 As Long

Public VerticesCol() As D3DVECTOR
Public VerticesColCopy() As D3DVECTOR

Public DIDevice As DirectInputDevice8

Public matView As D3DMATRIX
Public matWorld As D3DMATRIX
Public matProj As D3DMATRIX

Public Mat2 As D3DMATRIX

Dim frames2 As Long

Public Fuente As D3DXFont
Public FuenteD As StdFont

Public Const PI As Single = 3.14159265358979

Public luz As D3DLIGHT8
Public Luz2 As D3DLIGHT8
Public Luz3 As D3DLIGHT8
Public luz4 As D3DLIGHT8

Global salir As Integer
Global Nivel As Single
Global RotY As Single
Global RotX As Single

Global VerticalAngle As Single
Global CameraDistance As Single
Global CameraDistanceProj As Single

Global AccelerationTimer As Double
Global LastFrameTime As Double
Public Const WorldAcceleration As Double = 9.8
Public Const UnitsPerMetre As Double = 1
Global BajadoPorCaida As Double

Global ok As Boolean
Global CameraRelX As Single
Global CameraRelZ As Single

Sub Main()
Set DX = New DirectX8
Set Dx3D = DX.Direct3DCreate
Set D3D = New D3DX8
Set DInput = DX.DirectInputCreate()

Set DIDevice = DInput.CreateDevice("guid_SysMouse")

Dim parametros As D3DPRESENT_PARAMETERS

Dx3D.GetAdapterDisplayMode 0, DispMode

Load resolucion

CameraDistance = 5

Dim x As Integer, DispModeList As D3DDISPLAYMODE, selec As Integer, bits As Integer

For x = 1 To Dx3D.GetAdapterModeCount(0) - 1
        Dx3D.EnumAdapterModes 0, x, DispModeList
        If DispModeList.Format = 22 Then
            bits = 32
        ElseIf DispModeList.Format = 23 Then
            bits = 16
        End If
        resolucion.combo.AddItem DispModeList.Width & " x " & DispModeList.Height & " / " & DispModeList.RefreshRate & " / " & bits & " bits"
        If DispModeList.Width = DispMode.Width And _
        DispModeList.Height = DispMode.Height And _
        DispModeList.RefreshRate = DispMode.RefreshRate Then
            selec = resolucion.combo.ListCount - 1
        End If
Next

resolucion.combo.ListIndex = x - 2

resolucion.antialias.AddItem "NO MULTISAMPLE"
For x = 2 To 16
    If Dx3D.CheckDeviceMultiSampleType(0, D3DDEVTYPE_HAL, D3DFMT_A8R8G8B8, True, x) = 0 Then
        resolucion.antialias.AddItem x
    End If
Next
resolucion.antialias.ListIndex = 0


resolucion.Show 1

If ok = False Then End

Dx3D.EnumAdapterModes 0, DModeSelected + 1, DispMode

'MsgBox Dx3D.CheckDeviceMultiSampleType(0, D3DDEVTYPE_HAL, D3DFMT_A8R8G8B8, True, D3DMULTISAMPLE_2_SAMPLES)

With parametros
    '.Windowed = 1
    .SwapEffect = D3DSWAPEFFECT_DISCARD
    '.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    .BackBufferFormat = DispMode.Format
    .AutoDepthStencilFormat = D3DFMT_D24X8
    .EnableAutoDepthStencil = 1
    .BackBufferCount = 1
    '.BackBufferWidth = Form1.ScaleWidth / Screen.TwipsPerPixelX
    '.BackBufferHeight = Form1.ScaleHeight / Screen.TwipsPerPixelY
    .BackBufferWidth = DispMode.Width
    .BackBufferHeight = DispMode.Height
    
    SWidth = DispMode.Width
    SHeight = DispMode.Height
    '.MultiSampleType = D3DMULTISAMPLE_NONE
    '.MultiSampleType = D3DMULTISAMPLE_6_SAMPLES
    .MultiSampleType = ALevel
End With

'MsgBox Dx3D.CheckDeviceMultiSampleType(0, D3DDEVTYPE_HAL, D3DFMT_D24X8, 0, D3DMULTISAMPLE_6_SAMPLES)

Load Form1

'Form1.WindowState = vbMaximized
Err.Number = 0
On Local Error Resume Next
Set Device = Dx3D.CreateDevice(0, D3DDEVTYPE_HAL, Form1.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, parametros)
CPUMode = "Hardware Vertex Processing"
If Err.Number <> 0 Then
    Set Device = Dx3D.CreateDevice(0, D3DDEVTYPE_HAL, Form1.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, parametros)
    CPUMode = "Software Vertex Processing"
End If
On Error GoTo 0
Form1.Show

'Form1.WindowState = vbMaximized

Call DIDevice.SetCommonDataFormat(DIFORMAT_MOUSE)
Call DIDevice.SetCooperativeLevel(Form1.hWnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE)

    Dim diProp As DIPROPLONG
    diProp.lHow = DIPH_DEVICE
    diProp.lObj = 0
    diProp.lData = 30
    
    Call DIDevice.SetProperty("DIPROP_BUFFERSIZE", diProp)
    
    hevent = DX.CreateEvent(Form1)
    DIDevice.SetEventNotification hevent



Device.SetRenderState D3DRS_LIGHTING, 1          'enable lighting

Device.SetRenderState D3DRS_ZENABLE, 1           'enable the z buffer

Device.SetRenderState D3DRS_AMBIENT, RGB(90, 90, 90)

Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1
Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

Device.SetRenderState D3DRS_ALPHATESTENABLE, 1

Device.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
Device.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

D3DXMatrixPerspectiveFovLH matProj, 45 * PI / 180, 1, 0.1, 250
Device.SetTransform D3DTS_PROJECTION, matProj

Set FuenteD = New StdFont
Dim FD As IFont

FuenteD.Name = "Arial"
FuenteD.Size = 12
Set FD = FuenteD
Set Fuente = D3D.CreateFont(Device, FD.hFont)
Dim y As Integer
Dim ad As D3DXBuffer

If Zone <> 4 Then
    Set Objeto1 = D3D.LoadMeshFromX(App.Path & "\world.x", D3DXMESH_MANAGED, Device, ad, MtrlBuffer1, nMateriales)
    Set Objeto1_C = D3D.LoadMeshFromX(App.Path & "\world.x", D3DXMESH_MANAGED, Device, Nothing, MtrlBuffer1, nMateriales)
    Set ObjetoCollide = D3D.LoadMeshFromX(App.Path & "\world.x", D3DXMESH_MANAGED + D3DXMESH_32BIT, Device, Nothing, MtrlBufferC, nMaterialesC)
Else
    Set Objeto1 = D3D.LoadMeshFromX(App.Path & "\world2.x", D3DXMESH_MANAGED, Device, ad, MtrlBuffer1, nMateriales)
    Set Objeto1_C = D3D.LoadMeshFromX(App.Path & "\world2.x", D3DXMESH_MANAGED, Device, Nothing, MtrlBuffer1, nMateriales)
    Set ObjetoCollide = D3D.LoadMeshFromX(App.Path & "\world2.x", D3DXMESH_MANAGED + D3DXMESH_32BIT, Device, Nothing, MtrlBufferC, nMaterialesC)
End If
optimize Objeto1, ad


'cropmesh

ReDim MeshMaterials1(nMateriales - 1) As D3DMATERIAL8
ReDim MeshTextures1(nMateriales - 1) As Direct3DTexture8
ReDim MeshTexturesList(nMateriales - 1)
Dim XMat() As D3DXMATERIAL
ReDim XMat(nMateriales - 1)
        Dim image As IPictureDisp, TransparentImage As Boolean
        Dim imagew As Long, imageh As Long, iminfo As BITMAP


On Local Error Resume Next
For x = 0 To nMateriales - 1
    D3D.BufferGetMaterial MtrlBuffer1, x, MeshMaterials1(x)
    MeshMaterials1(x).Ambient = MeshMaterials1(x).diffuse
    If D3D.BufferGetTextureName(MtrlBuffer1, x) <> "" Then
    
        If LCase(Right(D3D.BufferGetTextureName(MtrlBuffer1, x), 3)) = "tga" Then TransparentImage = True
        
        If TransparentImage = True Then
            imagew = 512
            imageh = 512
        Else
            Set image = LoadPicture(D3D.BufferGetTextureName(MtrlBuffer1, x))
            GetObject image.Handle, Len(iminfo), iminfo
            imagew = iminfo.bmWidth
            imageh = iminfo.bmHeight
        End If

        If LCase(Right(D3D.BufferGetTextureName(MtrlBuffer1, x), 3)) <> "tga" Then
            Err.Number = 0
            'Set MeshTextures1(X) = D3D.CreateTextureFromFile(Device, D3D.BufferGetTextureName(MtrlBuffer1, X))
            Set MeshTextures1(x) = D3D.CreateTextureFromFileEx(Device, D3D.BufferGetTextureName(MtrlBuffer1, x), imagew, imageh, D3DX_DEFAULT, 0, D3DFMT_DXT2, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
            If Err.Number <> 0 Then Set MeshTextures1(x) = D3D.CreateTextureFromFileEx(Device, App.Path & "\" & D3D.BufferGetTextureName(MtrlBuffer1, x), 256, 256, D3DX_DEFAULT, 0, D3DFMT_DXT2, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        Else
            Err.Number = 0
            Set MeshTextures1(x) = D3D.CreateTextureFromFileEx(Device, D3D.BufferGetTextureName(MtrlBuffer1, x), imagew, imageh, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
            If Err.Number <> 0 Then Set MeshTextures1(x) = D3D.CreateTextureFromFileEx(Device, App.Path & "\" & D3D.BufferGetTextureName(MtrlBuffer1, x), 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        End If
    End If
    XMat(x).MatD3D = MeshMaterials1(x)
    XMat(x).TextureFilename = D3D.BufferGetTextureName(MtrlBuffer1, x)
    MeshTexturesList(x) = D3D.BufferGetTextureName(MtrlBuffer1, x)
Next
On Local Error GoTo 0
Dim cont As Integer
ReDim MeshTexturesListOrd(0)
ReDim MeshMaterialsOrd(0)

Set CubeTex(0) = D3D.CreateTextureFromFileEx(Device, "front.bmp", 768, 400, D3DX_DEFAULT, 0, D3DFMT_DXT1, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
Set CubeTex(3) = D3D.CreateTextureFromFileEx(Device, "left.bmp", 768, 400, D3DX_DEFAULT, 0, D3DFMT_DXT1, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
Set CubeTex(2) = D3D.CreateTextureFromFileEx(Device, "back.jpg", 768, 400, D3DX_DEFAULT, 0, D3DFMT_DXT1, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
Set CubeTex(1) = D3D.CreateTextureFromFileEx(Device, "right.jpg", 768, 400, D3DX_DEFAULT, 0, D3DFMT_DXT1, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
Set CubeTex(4) = D3D.CreateTextureFromFileEx(Device, "up.jpg", 768, 768, D3DX_DEFAULT, 0, D3DFMT_DXT1, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)

For x = 0 To nMateriales - 1
    For y = 0 To cont - 1
        If MeshTexturesListOrd(y) = MeshTexturesList(x) Then GoTo 2
    Next
    ReDim Preserve MeshMaterialsOrd(UBound(MeshMaterialsOrd) + 1)
    ReDim Preserve MeshTexturesListOrd(UBound(MeshTexturesListOrd) + 1)
    MeshMaterialsOrd(cont) = MeshMaterials1(x)
    MeshTexturesListOrd(cont) = MeshTexturesList(x)
    cont = cont + 1
2
Next

For y = 0 To 2
    Set Objeto2(y) = D3D.LoadMeshFromX(App.Path & "\mod" & Trim(Str(y + 1)) & ".x", D3DXMESH_MANAGED, Device, Nothing, MtrlBuffer2, nMateriales2)
    ReDim MeshMaterials2(nMateriales2) As D3DMATERIAL8
    ReDim MeshTextures2(nMateriales2) As Direct3DTexture8

    For x = 0 To nMateriales2 - 1
        D3D.BufferGetMaterial MtrlBuffer2, x, MeshMaterials2(x)
        MeshMaterials2(x).Ambient = MeshMaterials2(x).diffuse
        If D3D.BufferGetTextureName(MtrlBuffer2, x) <> "" Then
            Set MeshTextures2(x) = D3D.CreateTextureFromFileEx(Device, App.Path & "\" & D3D.BufferGetTextureName(MtrlBuffer2, x), 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        End If
    Next
Next

ReDim Meshes(59)

For y = 0 To 59
    Set Meshes(y) = D3D.LoadMeshFromX(App.Path & "\mod1.x", D3DXMESH_MANAGED, Device, Nothing, MtrlBuffer2, nMateriales2)
Next

Set Objeto2Render = D3D.LoadMeshFromX(App.Path & "\mod3.x", D3DXMESH_MANAGED, Device, Nothing, MtrlBuffer2, nMateriales2)

Set MtrlBuffer1 = Nothing
Set MtrlBuffer2 = Nothing



luces

Dim RenderTempMat As D3DMATRIX, I As Long
Dim RenderTempMat2 As D3DMATRIX
Dim RenderTempMat3 As D3DMATRIX
Dim RenderTempMat4 As D3DMATRIX

Dim fps As Integer
Dim tpf As Single
fps = 60
tpf = 1000 / fps

Dim tempo As Single, tempo2 As Single
Dim Frames As Integer
tempo = Timer
Device.LightEnable 0, False
Device.LightEnable 1, 1
Device.LightEnable 2, False
Device.LightEnable 3, False

If Zone = 2 Then
    coz = 170
    cox = -150
ElseIf Zone = 1 Then
    coz = -3
    cox = -3
ElseIf Zone = 4 Then
    coz = -8
    cox = -6
Else
    coz = 360
    cox = -150
End If
RotX = 0
YInicial = 20
coy = 20
If Zone = 4 Then
    YInicial = 1
    coy = 1
End If
VerticalAngle = 40
PreAnim

    D3DXMatrixIdentity RenderTempMat
    D3DXMatrixTranslation RenderTempMat, 0, 0, 0
    
    'rotate the 3d object on the Z axis by 'rangle'

PreLoadColision

DoEvents
DoEvents
DIDevice.Acquire

Do While salir = 0
'    temporizador = GetTickCount()
    counter = GetTickCount()
    Static Timer_ As Double
    If Timer_ <> 0 Then
        Timer_ = (GetTickCount() - Timer_)
        Timer_ = (AverageFrame * 4 + Timer_) / 5
        AverageFrame = Timer_
    End If
    Timer_ = GetTickCount()

    If GetAsyncKeyState(vbKeyUp) Or GetAsyncKeyState(vbKeyW) Then
        RotY = -1
    ElseIf GetAsyncKeyState(vbKeyDown) Or GetAsyncKeyState(vbKeyS) Then
        RotY = 1
    Else
        RotY = 0
    End If
    If GetAsyncKeyState(vbKeyLeft) Then
        RotX = RotX - 0.1
    ElseIf GetAsyncKeyState(vbKeyRight) Then
        RotX = RotX + 0.1
    End If
    If GetAsyncKeyState(vbKeyPageDown) Then
        Nivel = Nivel - 1
    ElseIf GetAsyncKeyState(vbKeyPageUp) Then
        Nivel = Nivel + 1
    End If
    
    cox2 = cox: coy2 = coy: coz2 = coz
    
    animar
    colision
    
    
    Call Device.Clear(0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0)
    Device.BeginScene
    
    D3DXMatrixLookAtLH matView, v3(camx, camy - 0.2, camz), v3(cox, coy + 1, coz), UpVector
    Device.SetTransform D3DTS_VIEW, matView
    
    VArray(0) = Vert(-0.5, 0.5, 0.5, 0.001, 0.001)
    VArray(1) = Vert(0.5, 0.5, 0.5, 0.999, 0.001)
    VArray(2) = Vert(-0.5, 0, 0.5, 0.001, 0.999)
    VArray(3) = Vert(0.5, 0, 0.5, 0.999, 0.999)
    
    VArray(4) = Vert(-0.5, 0.5, -0.5, 0.001, 0.001)
    VArray(5) = Vert(-0.5, 0.5, 0.5, 1#, 0.001)
    VArray(6) = Vert(-0.5, 0, -0.5, 0.001, 0.999)
    VArray(7) = Vert(-0.5, 0, 0.5, 0.999, 0.999)
    
    VArray(8) = Vert(0.5, 0.5, -0.5, 0.001, 0.001)
    VArray(9) = Vert(-0.5, 0.5, -0.5, 1, 0.001)
    VArray(10) = Vert(0.5, 0, -0.5, 0.001, 0.999)
    VArray(11) = Vert(-0.5, 0, -0.5, 0.999, 0.999)
    
    VArray(12) = Vert(0.5, 0.5, 0.5, 0.001, 0.001)
    VArray(13) = Vert(0.5, 0.5, -0.5, 0.999, 0.001)
    VArray(14) = Vert(0.5, 0, 0.5, 0.001, 0.999)
    VArray(15) = Vert(0.5, 0, -0.5, 0.999, 0.999)
    
    VArray(16) = Vert(-0.5, 0.5, -0.5, 0.01, 0.01)
    VArray(17) = Vert(0.5, 0.5, -0.5, 0.999, 0.01)
    VArray(18) = Vert(-0.5, 0.5, 0.5, 0.01, 0.999)
    VArray(19) = Vert(0.5, 0.5, 0.5, 0.999, 0.999)
    
    D3DXMatrixTranslation RenderTempMat, camx, camy - 0.2 - 0.02, camz
    
    Device.SetRenderState D3DRS_ZWRITEENABLE, 0
    Device.SetRenderState D3DRS_LIGHTING, 0
    Device.SetTransform D3DTS_WORLD, RenderTempMat
    
    Device.SetVertexShader myvert_fvf
    
    Device.SetTexture 0, CubeTex(0)
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VArray(0), Len(VArray(0))
    
    Device.SetTexture 0, CubeTex(1)
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VArray(4), Len(VArray(0))
    
    Device.SetTexture 0, CubeTex(2)
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VArray(8), Len(VArray(0))
    
    Device.SetTexture 0, CubeTex(3)
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VArray(12), Len(VArray(0))
    
    'Device.SetTexture 0, CubeTex(4)
    'Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, VArray(16), Len(VArray(0))
    
    Device.SetRenderState D3DRS_LIGHTING, 1
    Device.SetRenderState D3DRS_ZWRITEENABLE, 1

    'ComputeClipPlanes matProj, matView
    
    D3DXMatrixTranslation RenderTempMat, cox, coy, coz
    D3DXMatrixRotationY RenderTempMat3, RotX * PI / 180 - PI / 2
    D3DXMatrixMultiply RenderTempMat, RenderTempMat3, RenderTempMat
    
    'rotate the 3d object on the Z axis by 'rangle'
    Device.SetTransform D3DTS_WORLD, RenderTempMat
    'for each of the 3d object, objects
    For I = 0 To nMateriales2 - 1
            'Device.SetMaterial spec
            Device.SetMaterial MeshMaterials2(I)    'set the object
            Device.SetTexture 0, MeshTextures2(I)  'set the texture
            If RotY = 0 Then
                Objeto2(0).DrawSubset I
            Else
                Objeto2Render.DrawSubset I
            End If
    Next
    
    D3DXMatrixTranslation RenderTempMat, 0, 0, 0
    'D3DXMatrixScaling RenderTempMat, 10, 10, 10
    
    'rotate the 3d object on the Z axis by 'rangle'
    Device.SetTransform D3DTS_WORLD, RenderTempMat
    'for each of the 3d object, objects
    
    
    For I = 0 To nMateriales - 1
        Device.SetMaterial MeshMaterials1(I)    'set the object
        Device.SetTexture 0, MeshTextures1(I)  'set the texture
        Objeto1.DrawSubset I               'draw object
    Next
    
  
    'For i = 0 To nMateriales - 1
    '    If InStr(1, MeshTexturesList(i), "tga") <> 0 Then
    '        Device.SetMaterial MeshMaterials1(i)    'set the object
    '        Device.SetTexture 0, MeshTextures1(i)  'set the texture
    '        Objeto1.DrawSubset i               'draw object
    '    End If
    'Next
    
    Frames = Frames + 1
    If tempo2 < Timer Then
        tempo2 = Timer + 1
        frames2 = Frames
        Frames = 0
        tempo = GetTickCount()
    End If
    Dim textrect As RECT
        textrect.Top = 0
        textrect.bottom = 50
        textrect.Right = 300
            counter = GetTickCount() - counter
'    temporizador = GetTickCount() - temporizador
'    Debug.Print temporizador
    D3D.DrawText Fuente, &HFFFFFFFF, frames2 & " - XYZ(" & cox & ", " & coy & ", " & coz & ")" & " // " & RotX, textrect, DT_TOP Or DT_LEFT
    'Debug.Print ress
    ress = 0

    Device.EndScene
        
    Call Device.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    DoEvents 'Le damos un respiro a windows para que haga sus cosas :)
Loop

DIDevice.Unacquire

End
End Sub

Public Function v3(x As Single, y As Single, z As Single) As D3DVECTOR
v3.x = x
v3.y = y
v3.z = z
End Function


Sub PreAnim()

Dim FrameFinal() As D3DVERTEX
ReDim Frame1(Objeto2(0).GetNumVertices)
ReDim Frame2(Objeto2(0).GetNumVertices)
ReDim Frame3(Objeto2(0).GetNumVertices)
ReDim FrameFinal(Objeto2(0).GetNumVertices)

Dim hresult As Long, vTemp3D As D3DVECTOR, vTemp2D As D3DVECTOR2

hresult = D3DXMeshVertexBuffer8GetData(Objeto2(0), 0, Len(Frame1(0)) * Objeto2(0).GetNumVertices, 0, Frame1(0))
hresult = D3DXMeshVertexBuffer8GetData(Objeto2(1), 0, Len(Frame2(0)) * Objeto2(1).GetNumVertices, 0, Frame2(0))
hresult = D3DXMeshVertexBuffer8GetData(Objeto2(2), 0, Len(Frame3(0)) * Objeto2(2).GetNumVertices, 0, Frame3(0))

Dim x As Integer, I As Integer

'ReDim Meshes(59)

For x = 30 To 59 Step 1
        For I = 0 To Objeto2(0).GetNumVertices - 1 '//Cycle through every vertex
            '//2a. Interpolate the Positions
                D3DXVec3Lerp vTemp3D, v3(Frame1(I).x, Frame1(I).y, Frame1(I).z), v3(Frame3(I).x, Frame3(I).y, Frame3(I).z), (x - 30) / 30
                FrameFinal(I).x = vTemp3D.x
                FrameFinal(I).y = vTemp3D.y
                FrameFinal(I).z = vTemp3D.z
                
            '//2b. Interpolate the Normals
                D3DXVec3Lerp vTemp3D, v3(Frame1(I).nx, Frame1(I).ny, Frame1(I).nz), v3(Frame3(I).nx, Frame3(I).ny, Frame3(I).nz), (x - 30) / 30
                FrameFinal(I).nx = vTemp3D.x
                FrameFinal(I).ny = vTemp3D.y
                FrameFinal(I).nz = vTemp3D.z
            
            '//2c. Interpolate the Texture Coordinates
                D3DXVec2Lerp vTemp2D, v2(Frame1(I).tu, Frame1(I).tv), v2(Frame3(I).tu, Frame3(I).tv), (x - 30) / 30
                FrameFinal(I).tu = vTemp2D.x
                FrameFinal(I).tv = vTemp2D.y
        Next I
        hresult = D3DXMeshVertexBuffer8SetData(Meshes(x), 0, Len(FrameFinal(0)) * Objeto2(0).GetNumVertices, 0, FrameFinal(0))
Next

For x = 29 To 0 Step -1
        For I = 0 To Objeto2(0).GetNumVertices - 1 '//Cycle through every vertex
            '//2a. Interpolate the Positions
                D3DXVec3Lerp vTemp3D, v3(Frame1(I).x, Frame1(I).y, Frame1(I).z), v3(Frame2(I).x, Frame2(I).y, Frame2(I).z), (30 - x) / 30
                FrameFinal(I).x = vTemp3D.x
                FrameFinal(I).y = vTemp3D.y
                FrameFinal(I).z = vTemp3D.z
                
            '//2b. Interpolate the Normals
                D3DXVec3Lerp vTemp3D, v3(Frame1(I).nx, Frame1(I).ny, Frame1(I).nz), v3(Frame2(I).nx, Frame2(I).ny, Frame2(I).nz), (30 - x) / 30
                FrameFinal(I).nx = vTemp3D.x
                FrameFinal(I).ny = vTemp3D.y
                FrameFinal(I).nz = vTemp3D.z
            
            '//2c. Interpolate the Texture Coordinates
                D3DXVec2Lerp vTemp2D, v2(Frame1(I).tu, Frame1(I).tv), v2(Frame2(I).tu, Frame2(I).tv), (30 - x) / 30
                FrameFinal(I).tu = vTemp2D.x
                FrameFinal(I).tv = vTemp2D.y

        Next I
        hresult = D3DXMeshVertexBuffer8SetData(Meshes(x), 0, Len(FrameFinal(0)) * Objeto2(0).GetNumVertices, 0, FrameFinal(0))
Next


End Sub

Sub animar()
Static direccion As Integer

If RotY = 0 Then Exit Sub

If direccion = 0 Then
tween = tween + 200 * AverageFrame / 1000
If tween >= 59 Then direccion = 1
Else
tween = tween - 200 * AverageFrame / 1000 ' 10 * Timer
If tween <= 0 Then direccion = 0
End If

Dim frame As Integer

If tween > 0 Then
    frame = tween
    If frame >= 60 Then frame = 59
    If frame <= 0 Then frame = 1
    Set Objeto2Render = Meshes(frame)
Else
    frame = tween
    If frame >= 60 Then frame = 59
    If frame <= 0 Then frame = 1
    Set Objeto2Render = Meshes(frame)
End If
End Sub

Public Function v2(x As Single, y As Single) As D3DVECTOR2
v2.x = x
v2.y = y
End Function

Sub PreLoadColision()
Dim hresult As Long
Dim vertexs() As D3DVERTEX

ReDim vertexs(ObjetoCollide.GetNumVertices)

hresult = D3DXMeshVertexBuffer8GetData(ObjetoCollide, 0, Len(vertexs(0)) * ObjetoCollide.GetNumVertices, 0, vertexs(0))

Dim I As Long

Dim midesc As D3DINDEXBUFFER_DESC
Dim IBuf As Direct3DIndexBuffer8
Dim tam As Long, tam2 As Long
Dim out As Long
Dim vector() As Long

Set IBuf = ObjetoCollide.GetIndexBuffer()
IBuf.Lock 0, 0, out, 16
IBuf.GetDesc midesc
IBuf.Unlock

tam = midesc.Size

ReDim vector(midesc.Size / 4)

'D3DIndexBuffer8GetData ObjetoCollide.GetIndexBuffer(), 0, midesc.Size, 0, vector(0)
D3DXMeshIndexBuffer8GetData ObjetoCollide, 0, midesc.Size, 0, vector(0)

ReDim triangulos(ObjetoCollide.GetNumFaces * 3)

On Local Error GoTo 0

For I = 0 To ObjetoCollide.GetNumFaces * 3 - 1 Step 3 '//Cycle through every vertex
    triangulos(I).x = vertexs(vector(I)).x
    triangulos(I).y = vertexs(vector(I)).y
    triangulos(I).z = vertexs(vector(I)).z
    triangulos(I + 1).x = vertexs(vector(I + 1)).x
    triangulos(I + 1).y = vertexs(vector(I + 1)).y
    triangulos(I + 1).z = vertexs(vector(I + 1)).z
    triangulos(I + 2).x = vertexs(vector(I + 2)).x
    triangulos(I + 2).y = vertexs(vector(I + 2)).y
    triangulos(I + 2).z = vertexs(vector(I + 2)).z
Next

ReDim tris1(UBound(triangulos) / 3)
End Sub

Public Sub colision()
On Local Error Resume Next
Dim vertices(5) As D3DVECTOR
Dim NextY As Double, NextY2 As Double
Dim temp As Double, bajado As Boolean
Dim res As salida, res2 As salida

vertices(0).x = cox
vertices(0).y = coy + 1
vertices(0).z = coz

hprocess triangulos(0), vertices(0), (UBound(triangulos) / 3), res

'Debug.Print res.puntocolision.Y

'calcular la posicion que le corresponde según el tiempo
If AccelerationTimer <> 0 Then
    temp = ((GetTickCount() - AccelerationTimer) / 1000)
    NextY = YInicial + (VSalto - 7) * temp + 0.5 * -WorldAcceleration * temp ^ 2
    
    bajado = True
    
    If res.respuesta = 1 And (NextY <= coy) Then    'siempre colisionará a menos que caiga al vacío!!!
    If res.puntocolision.y >= NextY Then
        'el punto esta por arriba, subir escaleras o salvar obstaculos
        'coy = coy + (res.puntocolision.Y - coy)
        temp = (res.puntocolision.y - coy) * 25 * AverageFrame / 1000
        If temp < 0.01 Then temp = (res.puntocolision.y - coy)
        coy = coy + temp
        bajado = False
    Else
        'el punto de colision esta por debajo
        If NextY > res.puntocolision.y Then
            If (NextY - res.puntocolision.y) < 0.1 And VSalto = 0 Then
                'el desnivel es muy pequeño, por tanto son unas escaleras. Bajar
                'para evitar un efecto malo de caída
                coy = res.puntocolision.y
                bajado = False
            Else
                coy = NextY
            End If
        Else
            'puede caer hasta el punto de colisión
            'para encontrar el decremento de altura : (coy - res.puntocolision.y)
            'si el descenso es muy pequeño no hacerlo, ya que provoca una
            'pequeña vibración de la imagen a causa del redondeo de los números
            'If (coy - res.puntocolision.Y) > 0.005 Then
                coy = res.puntocolision.y
                bajado = False
            'Else
                'aunque está un poco en el aire consideramos que está en el suelo
                'bajado = False
            'End If
        End If
    End If
    Else
        coy = NextY
    End If
End If

If coy < -0.2 Then coy = -0.2
'If coy < 8 Then coy = 8

If bajado = False Then
    VSalto = 0
    YInicial = coy
    AccelerationTimer = GetTickCount()
    If GetAsyncKeyState(vbKeySpace) Then VSalto = 11
End If

Dim camv As D3DVECTOR
camv = v3(camx, camy, camz)
Dim posv As D3DVECTOR
posv = v3(cox, coy, coz)



process camv, triangulos(0), (UBound(triangulos) / 3), posv, v3(cox2, coy2, coz2), UpVector, CameraDistance, VerticalAngle, RotX, 0.5, 5, 0.7, 1.2, AverageFrame, RotY * 4, 0.35, tris1(0)
'process camv, triangulos(0), (UBound(triangulos) / 3), posv, v3(cox2, coy2, coz2), UpVector, CameraDistance, VerticalAngle, RotX, 5, 50, 10, 12, AverageFrame, RotY * 40, 035

cox = posv.x
coz = posv.z

camx = camv.x
camy = camv.y
camz = camv.z
End Sub

Public Sub TransformarGeometria(triangulo() As D3DVECTOR, matriz As D3DMATRIX, firstv As Integer, Optional ByVal numvertices As Integer = 3)
Dim I As Long
Dim a As Double, b As Double, c As Double, d As Double
For I = 0 To numvertices - 1
    D3DXVec3TransformCoord triangulo(firstv + I), triangulo(firstv + I), matriz
Next
End Sub

Public Sub CalcularCoordsCamara(ByRef camarax As Single, ByRef camaray As Single, ByRef camaraz As Single, ByVal angle As Single, ByVal distance As Single)
CameraDistanceProj = (distance * Cos(angle * PI / 180))
CameraRelX = CameraDistanceProj * Sin(RotX * PI / 180)
CameraRelZ = CameraDistanceProj * Cos(RotX * PI / 180)
camarax = CameraRelX + cox
camaraz = CameraRelZ + coz
camaray = coy + (distance * Sin(angle * PI / 180))
End Sub

Public Function DistanceB2P(Point1 As D3DVECTOR, Point2 As D3DVECTOR) As Single
DistanceB2P = Sqr((Point1.x - Point2.x) ^ 2 + (Point1.y - Point2.y) ^ 2 + (Point1.z - Point2.z) ^ 2)
End Function

Public Function Normalise(vector As D3DVECTOR) As D3DVECTOR
Dim module As Single
module = Sqr(vector.x ^ 2 + vector.y ^ 2 + vector.z ^ 2)
Normalise.x = vector.x / module
Normalise.y = vector.y / module
Normalise.z = vector.z / module
End Function

Public Function color(a As Single, r As Single, g As Single, b As Single) As D3DCOLORVALUE
color.a = a
color.r = r
color.g = g
color.b = b
End Function

Public Sub luces()
Dim mat As D3DMATERIAL8
mat.diffuse = color(1, 1, 1, 1)
mat.Ambient = color(1, 1, 1, 1)
Device.SetMaterial mat


'*******************************************************************
luz.Type = D3DLIGHT_DIRECTIONAL
luz.diffuse.r = 1
luz.diffuse.g = 1
luz.diffuse.b = 1
luz.diffuse.a = 1
luz.specular = color(1, 1, 1, 1)
luz.Direction = v3(1, -1, 1)

Luz2.Type = D3DLIGHT_DIRECTIONAL
Luz2.diffuse.r = 1
Luz2.diffuse.g = 1
Luz2.diffuse.b = 1
Luz2.diffuse.a = 1
Luz2.Range = 500
Luz2.specular = color(1, 1, 1, 1)
Luz2.Direction = Normalize(v3(-1, -1, -1))

Luz3.Type = D3DLIGHT_DIRECTIONAL
Luz3.diffuse.r = 1
Luz3.diffuse.g = 1
Luz3.diffuse.b = 1
Luz3.diffuse.a = 1
Luz3.specular = color(1, 1, 1, 1)
Luz3.Direction = v3(1, -1, -1)

luz4.Type = D3DLIGHT_DIRECTIONAL
luz4.diffuse.r = 1
luz4.diffuse.g = 1
luz4.diffuse.b = 1
luz4.diffuse.a = 1
luz4.specular = color(1, 1, 1, 1)
luz4.Direction = v3(-1, -1, -1)

Device.SetLight 0, luz
Device.SetLight 1, Luz2
Device.SetLight 2, Luz3
Device.SetLight 3, luz4
'*******************************************************************

End Sub

Public Function Arccos(ByVal x As Double) As Double
Arccos = Atn(-x / Sqr(-x * x + 1)) + 2 * Atn(1)
End Function

Private Sub optimize(mesh As D3DXMesh, adjbuffer As D3DXBuffer)
    Dim s As Long
    Dim adjBuf1() As Long
    Dim adjBuf2() As Long
    Dim facemap() As Long
    Dim newmesh As D3DXMesh
    Dim vertexMap As D3DXBuffer
    
    s = adjbuffer.GetBufferSize
    ReDim adjBuf1(s / 4)
    ReDim adjBuf2(s / 4)
    
    s = mesh.GetNumFaces
    ReDim facemap(s)
    
    D3D.BufferGetData adjbuffer, 0, 4, s * 3, adjBuf1(0)
    mesh.OptimizeInplace D3DXMESHOPT_COMPACT Or D3DXMESHOPT_ATTRSORT Or D3DXMESHOPT_VERTEXCACHE, adjBuf1(0), adjBuf2(0), facemap(0), vertexMap
End Sub

Sub ComputeClipPlanes(proj As D3DMATRIX, view As D3DMATRIX)
    Dim vecsf(7) As D3DVECTOR, mat As D3DMATRIX, x As Byte

    D3DXMatrixMultiply mat, view, proj
    D3DXMatrixInverse mat, 0, mat
    
    vecsf(0) = v3(-1, -1, 0)
    vecsf(1) = v3(1, -1, 0)
    vecsf(2) = v3(-1, 1, 0)
    vecsf(3) = v3(1, 1, 0)
    vecsf(4) = v3(-1, -1, 1)
    vecsf(5) = v3(1, -1, 1)
    vecsf(6) = v3(-1, 1, 1)
    vecsf(7) = v3(1, 1, 1)
    
    For x = 0 To 7
        D3DXVec3TransformCoord vecsf(x), vecsf(x), mat
    Next
    
    D3DXPlaneFromPoints FrustumPlanes(0), vecsf(0), vecsf(1), vecsf(2)
    D3DXPlaneFromPoints FrustumPlanes(1), vecsf(6), vecsf(7), vecsf(5)
    D3DXPlaneFromPoints FrustumPlanes(2), vecsf(2), vecsf(6), vecsf(4)
    D3DXPlaneFromPoints FrustumPlanes(3), vecsf(7), vecsf(3), vecsf(5)
    D3DXPlaneFromPoints FrustumPlanes(4), vecsf(2), vecsf(3), vecsf(6)
    D3DXPlaneFromPoints FrustumPlanes(5), vecsf(1), vecsf(0), vecsf(4)
End Sub

Public Function CheckSphere(Center As D3DVECTOR, Radius As Single) As Boolean
Dim TCenter As D3DVECTOR
Dim Matrix As D3DMATRIX, I As Long
Dim Dist As Single

Device.GetTransform D3DTS_WORLD, Matrix

D3DXVec3TransformCoord TCenter, Center, Matrix


For I = 0 To 5
    Dist = D3DXPlaneDotCoord(FrustumPlanes(I), TCenter)
        If Dist < -Radius Then
            CheckSphere = False 'not visible
        Exit Function
    End If
Next I
CheckSphere = True 'visible
End Function

Public Function Normalize(v As D3DVECTOR) As D3DVECTOR
Dim m As Single
m = Sqr(v.x ^ 2 + v.y ^ 2 + v.z ^ 2)
Normalize.x = v.x / m
Normalize.y = v.y / m
Normalize.z = v.z / m
End Function

Public Function Vert(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal tu As Single, ByVal tv As Single) As myvert
Vert.x = x
Vert.y = y
Vert.z = z
Vert.tu = tu
Vert.tv = tv
End Function
