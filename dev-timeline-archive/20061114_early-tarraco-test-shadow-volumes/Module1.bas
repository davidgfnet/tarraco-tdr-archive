Attribute VB_Name = "Module1"
Dim vLightInWorldSpace As D3DVECTOR, vLightInObjectSpace As D3DVECTOR
Dim svb As Direct3DVertexBuffer8
   Dim svert(3) As svert

'Option Explicit

Global Sombra As shadow
Global counter As Double
Global p1 As Double, p2 As Double
Global AverageFrame As Double
Global CPUMode As String
Global YInicial As Double
Global VSalto As Double
Public sleeptimer As Double
Public ATimer2 As Double
Public ATimer1 As Double
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
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
    X As Single
    Y As Single
    z  As Single
    rhw As Single
    color As Long
End Type
Public Const D3DFVF_CUSTOMVERTEX = (D3DFVF_XYZRHW Or D3DFVF_DIFFUSE)

Public Type salida
    respuesta As Integer
    puntocolision As D3DVECTOR
End Type

Public Type tdouble
    n As Double
End Type

Public Type svert
    'p As D3DVECTOR4
    X As Single
    Y As Single
    z As Single
    rhw As Single
    color As Long
End Type

Const fl As Long = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE


Public Declare Sub colision3d Lib "dx4vbmathslib.dll" (ByRef tri1 As D3DVECTOR, ByRef tri2 As D3DVECTOR, ByVal numtri1 As Long, ByVal numtri2 As Long, ByRef collide As salida)
Public Declare Sub colision3dseg Lib "dx4vbmathslib.dll" (ByRef tri As D3DVECTOR, seg As D3DVECTOR, ByVal numtri As Long, ByRef collide As salida)
Public Declare Sub cameracoords Lib "dx4vbmathslib.dll" (ByRef cam As D3DVECTOR, tri As D3DVECTOR, ByVal numtri As Long, ByRef pos As D3DVECTOR, ByRef upv As D3DVECTOR, ByVal distance As Double, ByVal angle As Double, ByVal angleh As Double, ByVal disfromcol As Single)
Public Declare Sub hprocess Lib "dx4vbmathslib.dll" (ByRef tri As D3DVECTOR, seg As D3DVECTOR, ByVal numtri As Long, ByRef collide As salida)
Public Declare Sub smoothcamera Lib "dx4vbmathslib.dll" (ByRef cam As D3DVECTOR, tri As D3DVECTOR, ByVal numtri As Long, ByRef pos As D3DVECTOR, ByRef upv As D3DVECTOR, ByVal distance As Double, ByVal angle As Double, ByVal angleh As Double, ByVal disfromcol As Single, ByVal interpolation As Single)

Public Declare Sub process Lib "dx4vbmathslib.dll" (ByRef cam As D3DVECTOR, tri As D3DVECTOR, ByVal numtri As Long, ByRef pos As D3DVECTOR, ByRef pos2 As D3DVECTOR, ByRef upv As D3DVECTOR, ByVal distance As Double, ByVal angle As Double, ByVal angleh As Double, ByVal disfromcol As Single, ByVal interpolationspeed As Single, ByVal midh As Single, ByVal hih As Single, ByVal avrframe As Single, ByVal speed As Single, ByVal dimensions As Single)
'Public Declare Sub buildshadow Lib "dx4vbmathslib.dll" (ByRef cam As D3DVECTOR, tri As D3DVECTOR, ByVal numtri As Long, ByRef pos As D3DVECTOR, ByRef pos2 As D3DVECTOR, ByRef upv As D3DVECTOR, ByVal distance As Double, ByVal angle As Double, ByVal angleh As Double, ByVal disfromcol As Single, ByVal interpolationspeed As Single, ByVal midh As Single, ByVal hih As Single, ByVal avrframe As Single, ByVal speed As Single, ByVal dimensions As Single)

Public Declare Sub buildshadow3 Lib "engine.dll" (ByRef tris As D3DVECTOR, ByVal numtri As Long, ByRef light As D3DVECTOR, ByRef trisout As D3DVECTOR, ByRef trisout2 As D3DVECTOR, ByRef numtris As Long, ByRef Normals As D3DVECTOR, ByVal numn As Long)
Public Declare Sub buildshadow2 Lib "engine.dll" (ByRef tris As D3DVECTOR, ByVal numtri As Long, ByRef light As D3DVECTOR, ByRef trisout As D3DVECTOR, ByRef numtris As Long, ByRef Normals As D3DVECTOR, ByVal numn As Long)
Public Declare Sub buildshadow Lib "engine.dll" (ByRef tris As D3DVECTOR, ByVal numtri As Long, ByRef light As D3DVECTOR, ByRef trisout As D3DVECTOR, ByRef numtris As Long, ByRef Normals As D3DVECTOR, ByVal numn As Long)
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
Public Dx As DirectX8
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
Dim MeshTexturesList() As String
Dim MeshMaterials1() As D3DMATERIAL8
Dim MeshMaterials2() As D3DMATERIAL8
Dim MeshTextures2() As Direct3DTexture8
Dim MeshMaterialsC() As D3DMATERIAL8
Dim MeshTexturesC() As Direct3DTexture8
Dim nMaterialesC As Long
Dim nMateriales As Long
Dim nMateriales2 As Long

Public VerticesCol() As D3DVECTOR
Public VerticesColCopy() As D3DVECTOR

Public DIInstance As DirectInputDevice8

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

Global Salir As Integer
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
Set Dx = New DirectX8
Set Dx3D = Dx.Direct3DCreate
Set D3D = New D3DX8
Set DInput = Dx.DirectInputCreate()

Dim parametros As D3DPRESENT_PARAMETERS

Dx3D.GetAdapterDisplayMode 0, DispMode

Load resolucion

Dim X As Integer, DispModeList As D3DDISPLAYMODE, selec As Integer, bits As Integer

For X = 1 To Dx3D.GetAdapterModeCount(0) - 1
        Dx3D.EnumAdapterModes 0, X, DispModeList
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

Set Sombra = New shadow

resolucion.combo.ListIndex = X - 2

resolucion.Show 1

If ok = False Then End

Dx3D.EnumAdapterModes 0, DModeSelected + 1, DispMode

With parametros
    '.Windowed = 1
    '.SwapEffect = D3DSWAPEFFECT_FLIP
    .SwapEffect = D3DSWAPEFFECT_DISCARD
    '.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
    .BackBufferFormat = DispMode.Format
    .AutoDepthStencilFormat = D3DFMT_D24S8
    .EnableAutoDepthStencil = 1
    .BackBufferCount = 1
    '.flags = D3DPRESENTFLAG_DISCARD_DEPTHSTENCIL
    '.BackBufferWidth = Form1.ScaleWidth / Screen.TwipsPerPixelX
    '.BackBufferHeight = Form1.ScaleHeight / Screen.TwipsPerPixelY
    .BackBufferWidth = DispMode.Width
    .BackBufferHeight = DispMode.Height
End With

Load Form1
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

Device.SetRenderState D3DRS_LIGHTING, 1          'enable lighting
'Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1
'Device.SetRenderState D3DRS_DITHERENABLE, 1

'Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
'Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
'Device.SetRenderState D3DRS_ALPHABLENDENABLE, True
'Device.SetRenderState D3DRS_ALPHATESTENABLE, True

Device.SetRenderState D3DRS_ZENABLE, 1           'enable the z buffer
'Device.SetRenderState D3DRS_ZFUNC ,


'Device.SetRenderState D3DRS_MULTISAMPLE_ANTIALIAS, True
'Device.SetRenderState D3DRS_CULLMODE, 1

Device.SetRenderState D3DRS_AMBIENT, RGB(200, 200, 200)

'Device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
'Device.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
'Device.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
'Device.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'Device.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_CURRENT

'Device.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
'Device.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR

D3DXMatrixPerspectiveFovLH matProj, PI / 4, 45 * PI / 180, 0.1, 50
Device.SetTransform D3DTS_PROJECTION, matProj


Set FuenteD = New StdFont
Dim FD As IFont

FuenteD.Name = "Arial"
FuenteD.Size = 12
Set FD = FuenteD
Set Fuente = D3D.CreateFont(Device, FD.hFont)
Dim Y As Integer
Dim ad As D3DXBuffer

Set Objeto1 = D3D.LoadMeshFromX(App.Path & "\world.x", D3DXMESH_MANAGED, Device, ad, MtrlBuffer1, nMateriales)
Set Objeto1_C = D3D.LoadMeshFromX(App.Path & "\world.x", D3DXMESH_MANAGED, Device, Nothing, MtrlBuffer1, nMateriales)

ReDim MeshMaterials1(nMateriales - 1) As D3DMATERIAL8
ReDim MeshTextures1(nMateriales - 1) As Direct3DTexture8
ReDim MeshTexturesList(nMateriales - 1)
Dim XMat() As D3DXMATERIAL
ReDim XMat(nMateriales - 1)

On Local Error Resume Next
For X = 0 To nMateriales - 1
    D3D.BufferGetMaterial MtrlBuffer1, X, MeshMaterials1(X)
    MeshMaterials1(X).Ambient = MeshMaterials1(X).diffuse
    If D3D.BufferGetTextureName(MtrlBuffer1, X) <> "" Then
        Err.Number = 0
        Set MeshTextures1(X) = D3D.CreateTextureFromFileEx(Device, D3D.BufferGetTextureName(MtrlBuffer1, X), 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, D3DColorARGB(255, 0, 0, 0), ByVal 0, ByVal 0)
        If Err.Number <> 0 Then Set MeshTextures1(X) = D3D.CreateTextureFromFileEx(Device, App.Path & "\" & D3D.BufferGetTextureName(MtrlBuffer1, X), 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, D3DColorARGB(255, 0, 0, 0), ByVal 0, ByVal 0)
    End If
    XMat(X).MatD3D = MeshMaterials1(X)
    XMat(X).TextureFilename = D3D.BufferGetTextureName(MtrlBuffer1, X)
    MeshTexturesList(X) = D3D.BufferGetTextureName(MtrlBuffer1, X)
Next
On Local Error GoTo 0


Set ObjetoCollide = D3D.LoadMeshFromX(App.Path & "\world.x", D3DXMESH_MANAGED, Device, Nothing, MtrlBufferC, nMaterialesC)

For Y = 0 To 2
    Set Objeto2(Y) = D3D.LoadMeshFromX(App.Path & "\mod" & Trim(Str(Y + 1)) & ".x", D3DXMESH_MANAGED, Device, Nothing, MtrlBuffer2, nMateriales2)
    ReDim MeshMaterials2(nMateriales2) As D3DMATERIAL8
    ReDim MeshTextures2(nMateriales2) As Direct3DTexture8

    For X = 0 To nMateriales2 - 1
        D3D.BufferGetMaterial MtrlBuffer2, X, MeshMaterials2(X)
        MeshMaterials2(X).Ambient = MeshMaterials2(X).diffuse
        If D3D.BufferGetTextureName(MtrlBuffer2, X) <> "" Then
            Set MeshTextures2(X) = D3D.CreateTextureFromFileEx(Device, App.Path & "\" & D3D.BufferGetTextureName(MtrlBuffer2, X), 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
        End If
    Next
Next

ReDim Meshes(59)

For Y = 0 To 59
    Set Meshes(Y) = D3D.LoadMeshFromX(App.Path & "\mod1.x", D3DXMESH_MANAGED, Device, Nothing, MtrlBuffer2, nMateriales2)
Next

Set Objeto2Render = D3D.LoadMeshFromX(App.Path & "\mod3.x", D3DXMESH_MANAGED, Device, Nothing, MtrlBuffer2, nMateriales2)

Set MtrlBuffer1 = Nothing
Set MtrlBuffer2 = Nothing


   
Set svb = Device.CreateVertexBuffer(4 * Len(svert(0)), D3DUSAGE_WRITEONLY, fl, D3DPOOL_MANAGED)

Dim desc As D3DSURFACE_DESC, sur As Direct3DSurface8
Set sur = Device.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
sur.GetDesc desc


svert(0).X = 0
svert(0).Y = desc.Height
svert(1).X = 0
svert(1).Y = 0
svert(2).X = desc.Width
svert(2).Y = desc.Height
svert(3).X = desc.Width
svert(3).Y = 0

For Y = 0 To 3
    svert(Y).z = 0
    svert(Y).rhw = 1
    'svert(Y).color = D3DColorARGB(0, 0, 0, 0)
    'svert(Y).color.r = 1
    'svert(Y).color.g = 0
    svert(Y).color = D3DColorARGB(128, 0, 0, 0)
Next



D3DVertexBuffer8SetData svb, 0, 4 * Len(svert(0)), 0, svert(0)




luces

Dim RenderTempMat As D3DMATRIX, i As Long
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
Device.LightEnable 1, True
Device.LightEnable 2, False
Device.LightEnable 3, False

'cox = -26
'coz = 5
coz = -0
cox = 3
RotX = 0
YInicial = 1
coy = 1
VerticalAngle = 40
PreAnim

    D3DXMatrixIdentity RenderTempMat
    D3DXMatrixTranslation RenderTempMat, 0, 0, 0
    
    'rotate the 3d object on the Z axis by 'rangle'

PreLoadColision



'Sombra.Build Objeto1, vLightInObjectSpace


Do While Salir = 0
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
     
    Call Device.Clear(0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER Or D3DCLEAR_STENCIL, 0, 1#, 0)
    Device.BeginScene
    
    D3DXMatrixLookAtLH matView, v3(camx, camy - 0.5, camz), v3(cox, coy + 1, coz), UpVector
    Device.SetTransform D3DTS_VIEW, matView

    
    
        


    D3DXMatrixTranslation RenderTempMat, 0, 0, 0
    
    'rotate the 3d object on the Z axis by 'rangle'
    Device.SetTransform D3DTS_WORLD, RenderTempMat
    'for each of the 3d object, objects
    
    For i = 0 To nMateriales - 1
        If InStr(1, MeshTexturesList(i), "tga") = 0 Then
            Device.SetMaterial MeshMaterials1(i)    'set the object
            Device.SetTexture 0, MeshTextures1(i)  'set the texture
            Objeto1.DrawSubset i               'draw object
        End If
    Next
    
    For i = 0 To nMateriales - 1
        If InStr(1, MeshTexturesList(i), "tga") <> 0 Then
            Device.SetMaterial MeshMaterials1(i)    'set the object
            Device.SetTexture 0, MeshTextures1(i)  'set the texture
            Objeto1.DrawSubset i               'draw object
        End If
    Next
    
        D3DXMatrixTranslation RenderTempMat, cox, coy, coz
    D3DXMatrixRotationY RenderTempMat3, RotX * PI / 180 - PI / 2
    D3DXMatrixMultiply RenderTempMat, RenderTempMat3, RenderTempMat

Dim Inv As D3DMATRIX
vLightInWorldSpace = v3(0, 2, 2)
D3DXMatrixInverse Inv, 0, RenderTempMat
D3DXVec3TransformNormal vLightInObjectSpace, vLightInWorldSpace, Inv
       
'If RotY = 0 Then
'Dim a As Boolean
Sombra.Build Objeto2(0), vLightInObjectSpace
'a = True
'Else
'Sombra.Build Objeto2Render, vLightInObjectSpace
'End If





    D3DXMatrixTranslation RenderTempMat, cox, coy, coz
    D3DXMatrixRotationY RenderTempMat3, RotX * PI / 180 - PI / 2
    D3DXMatrixMultiply RenderTempMat, RenderTempMat3, RenderTempMat
    
    'rotate the 3d object on the Z axis by 'rangle'
    Device.SetTransform D3DTS_WORLD, RenderTempMat
    'for each of the 3d object, objects
    For i = 0 To nMateriales2 - 1
            'Device.SetMaterial spec
            Device.SetMaterial MeshMaterials2(i)    'set the object
            Device.SetTexture 0, MeshTextures2(i)  'set the texture
            If RotY = 0 Then
                Objeto2(0).DrawSubset i
            Else
                Objeto2Render.DrawSubset i
            End If
    Next


sten


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

        D3D.DrawText Fuente, &HFFFFFFFF, frames2 & " - " & CPUMode, textrect, DT_TOP Or DT_LEFT

    Device.EndScene
        
     Call Device.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    DoEvents 'Le damos un respiro a windows para que haga sus cosas :)
Loop

End
End Sub

Public Function v3(X As Single, Y As Single, z As Single) As D3DVECTOR
v3.X = X
v3.Y = Y
v3.z = z
End Function


Sub PreAnim()
Exit Sub
Dim FrameFinal() As D3DVERTEX
ReDim Frame1(Objeto2(0).GetNumVertices)
ReDim Frame2(Objeto2(0).GetNumVertices)
ReDim Frame3(Objeto2(0).GetNumVertices)
ReDim FrameFinal(Objeto2(0).GetNumVertices)

Dim hresult As Long, vTemp3D As D3DVECTOR, vTemp2D As D3DVECTOR2

hresult = D3DXMeshVertexBuffer8GetData(Objeto2(0), 0, Len(Frame1(0)) * Objeto2(0).GetNumVertices, 0, Frame1(0))
hresult = D3DXMeshVertexBuffer8GetData(Objeto2(1), 0, Len(Frame2(0)) * Objeto2(1).GetNumVertices, 0, Frame2(0))
hresult = D3DXMeshVertexBuffer8GetData(Objeto2(2), 0, Len(Frame3(0)) * Objeto2(2).GetNumVertices, 0, Frame3(0))

Dim X As Integer, i As Integer

'ReDim Meshes(59)

For X = 30 To 59 Step 1
        For i = 0 To Objeto2(0).GetNumVertices  '//Cycle through every vertex
            '//2a. Interpolate the Positions
                D3DXVec3Lerp vTemp3D, v3(Frame1(i).X, Frame1(i).Y, Frame1(i).z), v3(Frame3(i).X, Frame3(i).Y, Frame3(i).z), (X - 30) / 30
                FrameFinal(i).X = vTemp3D.X
                FrameFinal(i).Y = vTemp3D.Y
                FrameFinal(i).z = vTemp3D.z
                
            '//2b. Interpolate the Normals
                D3DXVec3Lerp vTemp3D, v3(Frame1(i).nx, Frame1(i).ny, Frame1(i).nz), v3(Frame3(i).nx, Frame3(i).ny, Frame3(i).nz), (X - 30) / 30
                FrameFinal(i).nx = vTemp3D.X
                FrameFinal(i).ny = vTemp3D.Y
                FrameFinal(i).nz = vTemp3D.z
            
            '//2c. Interpolate the Texture Coordinates
                D3DXVec2Lerp vTemp2D, vec2(Frame1(i).tu, Frame1(i).tv), vec2(Frame3(i).tu, Frame3(i).tv), (X - 30) / 30
                FrameFinal(i).tu = vTemp2D.X
                FrameFinal(i).tv = vTemp2D.Y
        Next i
        hresult = D3DXMeshVertexBuffer8SetData(Meshes(X), 0, Len(FrameFinal(0)) * Objeto2(0).GetNumVertices, 0, FrameFinal(0))
Next

For X = 29 To 0 Step -1
        For i = 0 To Objeto2(0).GetNumVertices  '//Cycle through every vertex
            '//2a. Interpolate the Positions
                D3DXVec3Lerp vTemp3D, v3(Frame1(i).X, Frame1(i).Y, Frame1(i).z), v3(Frame2(i).X, Frame2(i).Y, Frame2(i).z), (30 - X) / 30
                FrameFinal(i).X = vTemp3D.X
                FrameFinal(i).Y = vTemp3D.Y
                FrameFinal(i).z = vTemp3D.z
                
            '//2b. Interpolate the Normals
                D3DXVec3Lerp vTemp3D, v3(Frame1(i).nx, Frame1(i).ny, Frame1(i).nz), v3(Frame2(i).nx, Frame2(i).ny, Frame2(i).nz), (30 - X) / 30
                FrameFinal(i).nx = vTemp3D.X
                FrameFinal(i).ny = vTemp3D.Y
                FrameFinal(i).nz = vTemp3D.z
            
            '//2c. Interpolate the Texture Coordinates
                D3DXVec2Lerp vTemp2D, vec2(Frame1(i).tu, Frame1(i).tv), vec2(Frame2(i).tu, Frame2(i).tv), (30 - X) / 30
                FrameFinal(i).tu = vTemp2D.X
                FrameFinal(i).tv = vTemp2D.Y

        Next i
        hresult = D3DXMeshVertexBuffer8SetData(Meshes(X), 0, Len(FrameFinal(0)) * Objeto2(0).GetNumVertices, 0, FrameFinal(0))
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

Public Function vec2(X As Single, Y As Single) As D3DVECTOR2
vec2.X = X
vec2.Y = Y
End Function

Sub PreLoadColision()
Dim hresult As Long
Dim vertexs() As D3DVERTEX

ReDim vertexs(ObjetoCollide.GetNumVertices)

hresult = D3DXMeshVertexBuffer8GetData(ObjetoCollide, 0, Len(vertexs(0)) * ObjetoCollide.GetNumVertices, 0, vertexs(0))

Dim i As Long

Dim midesc As D3DINDEXBUFFER_DESC
Dim IBuf As Direct3DIndexBuffer8
Dim tam As Long, tam2 As Long
Dim out As Long
Dim vector() As Integer, vector2() As Integer

Set IBuf = ObjetoCollide.GetIndexBuffer()
IBuf.Lock 0, 0, out, 16
IBuf.GetDesc midesc
IBuf.Unlock

tam = midesc.Size
ReDim vector(midesc.Size / 2)

D3DIndexBuffer8GetData ObjetoCollide.GetIndexBuffer(), 0, midesc.Size, 0, vector(0)

ReDim triangulos(ObjetoCollide.GetNumFaces * 3)

For i = 0 To ObjetoCollide.GetNumFaces * 3 - 1 Step 3 '//Cycle through every vertex
    triangulos(i).X = vertexs(vector(i)).X
    triangulos(i).Y = vertexs(vector(i)).Y
    triangulos(i).z = vertexs(vector(i)).z
    triangulos(i + 1).X = vertexs(vector(i + 1)).X
    triangulos(i + 1).Y = vertexs(vector(i + 1)).Y
    triangulos(i + 1).z = vertexs(vector(i + 1)).z
    triangulos(i + 2).X = vertexs(vector(i + 2)).X
    triangulos(i + 2).Y = vertexs(vector(i + 2)).Y
    triangulos(i + 2).z = vertexs(vector(i + 2)).z
Next
End Sub

Public Sub colision()
On Local Error Resume Next
Dim Vertices(5) As D3DVECTOR
Dim NextY As Double, NextY2 As Double
Dim temp As Double, bajado As Boolean
Dim res As salida, res2 As salida

Vertices(0).X = cox
Vertices(0).Y = coy + 1
Vertices(0).z = coz

hprocess triangulos(0), Vertices(0), (UBound(triangulos) / 3), res

'calcular la posicion que le corresponde según el tiempo
If AccelerationTimer <> 0 Then
    temp = ((GetTickCount() - AccelerationTimer) / 1000)
    NextY = YInicial + (VSalto - 7) * temp + 0.5 * -WorldAcceleration * temp ^ 2
    
   
    bajado = True
    
    If res.respuesta = 1 And (NextY <= coy) Then    'siempre colisionará a menos que caiga al vacío!!!
    If res.puntocolision.Y >= NextY Then
        'el punto esta por arriba, subir escaleras o salvar obstaculos
        coy = coy + (res.puntocolision.Y - coy)
        bajado = False
    Else
        'el punto de colision esta por debajo
        If NextY > res.puntocolision.Y Then
            If (NextY - res.puntocolision.Y) < 0.1 And VSalto = 0 Then
                'el desnivel es muy pequeño, por tanto son unas escaleras. Bajar
                'para evitar un efecto malo de caída
                coy = res.puntocolision.Y
                bajado = False
            Else
                coy = NextY
            End If
        Else
            'puede caer hasta el punto de colisión
            'para encontrar el decremento de altura : (coy - res.puntocolision.y)
            'si el descenso es muy pequeño no hacerlo, ya que provoca una
            'pequeña vibración de la imagen a causa del redondeo de los números
            If (coy - res.puntocolision.Y) > 0.005 Then
                coy = res.puntocolision.Y
                bajado = False
            Else
                'aunque está un poco en el aire consideramos que está en el suelo
                bajado = False
            End If
        End If
    End If
    Else
        coy = NextY
    End If
End If

If coy < 0 Then coy = 0

If bajado = False Then
    VSalto = 0
    YInicial = coy
    AccelerationTimer = GetTickCount()
    If GetAsyncKeyState(vbKeySpace) Then VSalto = 11
End If

CameraDistance = 5
Dim camv As D3DVECTOR
camv = v3(camx, camy, camz)
Dim posv As D3DVECTOR
posv = v3(cox, coy, coz)

process camv, triangulos(0), (UBound(triangulos) / 3), posv, v3(cox2, coy2, coz2), UpVector, CameraDistance, VerticalAngle, RotX, 0.5, 5, 1, 1.2, AverageFrame, RotY * 4, 0.35

cox = posv.X
coz = posv.z

camx = camv.X
camy = camv.Y
camz = camv.z


End Sub

Public Sub TransformarGeometria(triangulo() As D3DVECTOR, matriz As D3DMATRIX, firstv As Integer, Optional ByVal numvertices As Integer = 3)
Dim i As Long
Dim a As Double, b As Double, c As Double, d As Double
For i = 0 To numvertices - 1
    D3DXVec3TransformCoord triangulo(firstv + i), triangulo(firstv + i), matriz
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
DistanceB2P = Sqr((Point1.X - Point2.X) ^ 2 + (Point1.Y - Point2.Y) ^ 2 + (Point1.z - Point2.z) ^ 2)
End Function

Public Function Normalise(vector As D3DVECTOR) As D3DVECTOR
Dim module As Single
module = Sqr(vector.X ^ 2 + vector.Y ^ 2 + vector.z ^ 2)
Normalise.X = vector.X / module
Normalise.Y = vector.Y / module
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
Luz2.specular = color(1, 1, 1, 1)
Luz2.Direction = v3(-1, -1, 1)

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

Public Function Arccos(ByVal X As Double) As Double
Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function


Public Sub sten()
Dim MiMat As D3DMATERIAL8
MiMat.diffuse.r = 1
MiMat.diffuse.g = 1
MiMat.diffuse.b = 1
Device.SetMaterial MiMat

Device.SetRenderState D3DRS_ZWRITEENABLE, False
Device.SetRenderState D3DRS_STENCILENABLE, True
Device.SetRenderState D3DRS_SHADEMODE, D3DSHADE_FLAT

Device.SetRenderState D3DRS_STENCILFUNC, D3DCMP_ALWAYS
Device.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_KEEP
Device.SetRenderState D3DRS_STENCILFAIL, D3DSTENCILOP_KEEP

Device.SetRenderState D3DRS_STENCILREF, &H1
Device.SetRenderState D3DRS_STENCILMASK, &HFFFFFFFF
Device.SetRenderState D3DRS_STENCILWRITEMASK, &HFFFFFFFF
Device.SetRenderState D3DRS_STENCILZFAIL, D3DSTENCILOP_INCR

Device.SetRenderState D3DRS_ALPHABLENDENABLE, True
Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

Dim RenderTempMat As D3DMATRIX, i As Long
Dim RenderTempMat2 As D3DMATRIX
Dim RenderTempMat3 As D3DMATRIX
Dim RenderTempMat4 As D3DMATRIX

Device.SetRenderState D3DRS_CULLMODE, D3DCULL_CW


D3DXMatrixTranslation RenderTempMat, cox, coy, coz
D3DXMatrixRotationY RenderTempMat3, RotX * PI / 180 - PI / 2
D3DXMatrixMultiply RenderTempMat, RenderTempMat3, RenderTempMat
    
Device.SetTransform D3DTS_WORLD, RenderTempMat
Sombra.Render
'Sombra.Render True

Device.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
Device.SetRenderState D3DRS_STENCILZFAIL, D3DSTENCILOP_DECR
Device.SetTransform D3DTS_WORLD, RenderTempMat
Sombra.Render

'------
Device.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD
Device.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
Device.SetRenderState D3DRS_ZWRITEENABLE, True
Device.SetRenderState D3DRS_STENCILENABLE, False
Device.SetRenderState D3DRS_ALPHABLENDENABLE, False


Device.SetRenderState D3DRS_ZENABLE, False
Device.SetRenderState D3DRS_STENCILENABLE, 1
Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1

Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
Device.SetTexture 0, Nothing

'Device.SetRenderState D3DRS_ALPHATESTENABLE, True

'Device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
'Device.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
'Device.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
'Device.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'Device.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
'Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE

Device.SetRenderState D3DRS_STENCILREF, &H1
Device.SetRenderState D3DRS_STENCILFUNC, D3DCMP_LESSEQUAL
Device.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_KEEP

Device.SetVertexShader fl
Device.SetStreamSource 0, svb, Len(svert(0))
Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2


Device.SetRenderState D3DRS_ZENABLE, 1

Device.SetRenderState D3DRS_STENCILENABLE, 0
'


End Sub

Public Sub sten2()
Dim MiMat As D3DMATERIAL8
MiMat.diffuse.r = 1
MiMat.diffuse.g = 1
MiMat.diffuse.b = 1
Device.SetMaterial MiMat

Device.SetRenderState D3DRS_ZWRITEENABLE, False
Device.SetRenderState D3DRS_STENCILENABLE, True
Device.SetRenderState D3DRS_SHADEMODE, D3DSHADE_FLAT

Device.SetRenderState D3DRS_STENCILFUNC, D3DCMP_ALWAYS
Device.SetRenderState D3DRS_STENCILZFAIL, D3DSTENCILOP_KEEP
Device.SetRenderState D3DRS_STENCILFAIL, D3DSTENCILOP_KEEP

Device.SetRenderState D3DRS_STENCILREF, &H1
Device.SetRenderState D3DRS_STENCILMASK, &HFFFFFFFF
Device.SetRenderState D3DRS_STENCILWRITEMASK, &HFFFFFFFF
Device.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_INCR

Device.SetRenderState D3DRS_ALPHABLENDENABLE, True
Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

Dim RenderTempMat As D3DMATRIX, i As Long
Dim RenderTempMat2 As D3DMATRIX
Dim RenderTempMat3 As D3DMATRIX
Dim RenderTempMat4 As D3DMATRIX



D3DXMatrixTranslation RenderTempMat, cox, coy, coz
D3DXMatrixRotationY RenderTempMat3, RotX * PI / 180 - PI / 2
D3DXMatrixMultiply RenderTempMat, RenderTempMat3, RenderTempMat
    
Device.SetTransform D3DTS_WORLD, RenderTempMat
Device.SetRenderState D3DRS_CULLMODE, D3DCULL_CW
Sombra.Render

Device.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
Device.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_DECR
Device.SetTransform D3DTS_WORLD, RenderTempMat
Sombra.Render

'------
Device.SetRenderState D3DRS_SHADEMODE, D3DSHADE_GOURAUD
Device.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW
Device.SetRenderState D3DRS_ZWRITEENABLE, True
Device.SetRenderState D3DRS_STENCILENABLE, False
Device.SetRenderState D3DRS_ALPHABLENDENABLE, False


Device.SetRenderState D3DRS_ZENABLE, False
Device.SetRenderState D3DRS_STENCILENABLE, 1
Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1

Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
Device.SetTexture 0, Nothing

'Device.SetRenderState D3DRS_ALPHATESTENABLE, True

'Device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
'Device.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
'Device.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
'Device.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
'Device.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE
'Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE

Device.SetRenderState D3DRS_STENCILREF, &H1
Device.SetRenderState D3DRS_STENCILFUNC, D3DCMP_LESSEQUAL
Device.SetRenderState D3DRS_STENCILPASS, D3DSTENCILOP_KEEP

Device.SetVertexShader fl
Device.SetStreamSource 0, svb, Len(svert(0))
Device.DrawPrimitive D3DPT_TRIANGLESTRIP, 0, 2


Device.SetRenderState D3DRS_ZENABLE, 1

Device.SetRenderState D3DRS_STENCILENABLE, 0
'


End Sub

