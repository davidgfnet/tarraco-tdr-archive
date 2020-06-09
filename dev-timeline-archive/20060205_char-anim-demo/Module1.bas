Attribute VB_Name = "Module1"

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long

Global tween As Single

Public Dx As DirectX8
Public Dx3D As Direct3D8
Public Device As Direct3DDevice8
Public DispMode As D3DDISPLAYMODE
Public D3D As D3DX8
Public Cubo As D3DXMesh
Public Cubo2 As D3DXMesh
Public Cubo3 As D3DXMesh
Public CuboF As D3DXMesh
Public MtrlCubo As D3DXBuffer
Public MtrlCubo2 As D3DXBuffer
Public MtrlCubo3 As D3DXBuffer
Dim MeshMaterials() As D3DMATERIAL8
Dim MeshTextures() As Direct3DTexture8
Dim MeshMaterials2() As D3DMATERIAL8
Dim MeshTextures2() As Direct3DTexture8

Public matView As D3DMATRIX
Public matWorld As D3DMATRIX
Public matProj As D3DMATRIX

Dim frames2 As Long

Public Fuente As D3DXFont
Public FuenteD As StdFont

Public Const PI As Single = 3.14159265358979

Public Luz As D3DLIGHT8
Public Luz2 As D3DLIGHT8

Global Salir As Integer
Global Nivel As Single
Global DespZ As Single
Global RotY As Single
Global RotX As Single

Sub Main()
Set Dx = New DirectX8
Set Dx3D = Dx.Direct3DCreate
Set D3D = New D3DX8

Dim parametros As D3DPRESENT_PARAMETERS
Dim nMateriales As Long

Dx3D.GetAdapterDisplayMode 0, DispMode

With parametros
    .Windowed = 1
    .SwapEffect = D3DSWAPEFFECT_FLIP
    .BackBufferFormat = DispMode.Format
    .AutoDepthStencilFormat = D3DFMT_D16
    .EnableAutoDepthStencil = 1
End With

Load Form1
Set Device = Dx3D.CreateDevice(0, D3DDEVTYPE_HAL, Form1.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, parametros)
Form1.Show

Device.SetRenderState D3DRS_LIGHTING, 1          'enable lighting
Device.SetRenderState D3DRS_ZENABLE, 1           'enable the z buffer
'Device.SetRenderState D3DRS_AMBIENT, RGB(255, 0, 0)

D3DXMatrixIdentity matView
D3DXMatrixLookAtLH matView, v3(0, 200, 0), v3(0, 0, 0), v3(0, 1, 0)
Device.SetTransform D3DTS_WORLD, matView

D3DXMatrixPerspectiveFovLH matProj, PI / 4, PI * 0.3, 5, 5000
Device.SetTransform D3DTS_PROJECTION, matProj

'*******************************************************************
Luz.Type = D3DLIGHT_DIRECTIONAL
Luz.Position = v3(0, 0, 0)
Luz.diffuse.r = 1
Luz.diffuse.g = 1
Luz.diffuse.b = 1
Luz.Direction = v3(0, -1, 1)
Luz.Theta = PI * 10
Luz.Phi = PI * 2
Luz.Attenuation1 = 0.0025    'atenuacion de la luz al alejarse
Luz.Range = 1000

Luz2.Type = D3DLIGHT_DIRECTIONAL
Luz2.Position = v3(0, 200, 0)
Luz2.diffuse.r = 1
Luz2.diffuse.g = 1
Luz2.diffuse.b = 1
Luz2.Direction = v3(0, -1, -1)
Luz2.Theta = PI * 10
Luz2.Phi = PI * 2
Luz2.Attenuation1 = 0.0025    'atenuacion de la luz al alejarse
Luz2.Range = 1000


Device.SetLight 0, Luz
Device.SetLight 1, Luz2
'*******************************************************************
Set FuenteD = New StdFont
Dim FD As IFont

FuenteD.Name = "Arial"
FuenteD.Size = 12
Set FD = FuenteD
Set Fuente = D3D.CreateFont(Device, FD.hFont)

Set Cubo = D3D.LoadMeshFromX(App.Path & "\mod1.x", D3DXMESH_MANAGED, Device, Nothing, MtrlCubo, nMateriales)
Set CuboF = D3D.LoadMeshFromX(App.Path & "\mod1.x", D3DXMESH_MANAGED, Device, Nothing, MtrlCubo, nMateriales)

ReDim MeshMaterials(nMateriales - 1) As D3DMATERIAL8
ReDim MeshTextures(nMateriales - 1) As Direct3DTexture8

Dim X As Integer
For X = 0 To nMateriales - 1
    D3D.BufferGetMaterial MtrlCubo, X, MeshMaterials(X)
    If D3D.BufferGetTextureName(MtrlCubo, X) <> "" Then
        Set MeshTextures(X) = D3D.CreateTextureFromFileEx(Device, App.Path & "\" & D3D.BufferGetTextureName(MtrlCubo, X), 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    End If
Next

nMateriales2 = nMateriales
Set MtrlCubo = Nothing

Set Cubo2 = D3D.LoadMeshFromX(App.Path & "\mod2.x", D3DXMESH_MANAGED, Device, Nothing, MtrlCubo2, nMateriales)

ReDim MeshMaterials2(nMateriales - 1) As D3DMATERIAL8
ReDim MeshTextures2(nMateriales - 1) As Direct3DTexture8

For X = 0 To nMateriales - 1
    D3D.BufferGetMaterial MtrlCubo2, X, MeshMaterials2(X)
    If D3D.BufferGetTextureName(MtrlCubo2, X) <> "" Then
        Set MeshTextures2(X) = D3D.CreateTextureFromFileEx(Device, App.Path & "\" & D3D.BufferGetTextureName(MtrlCubo2, X), 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    End If
Next

Set MtrlCubo = Nothing

Set Cubo3 = D3D.LoadMeshFromX(App.Path & "\mod3.x", D3DXMESH_MANAGED, Device, Nothing, MtrlCubo3, nMateriales)

ReDim MeshMaterials2(nMateriales - 1) As D3DMATERIAL8
ReDim MeshTextures2(nMateriales - 1) As Direct3DTexture8

For X = 0 To nMateriales - 1
    D3D.BufferGetMaterial MtrlCubo3, X, MeshMaterials2(X)
    If D3D.BufferGetTextureName(MtrlCubo3, X) <> "" Then
        Set MeshTextures2(X) = D3D.CreateTextureFromFileEx(Device, App.Path & "\" & D3D.BufferGetTextureName(MtrlCubo3, X), 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    End If
Next

Set MtrlCubo = Nothing

Dim RenderTempMat As D3DMATRIX, i As Long
Dim RenderTempMat2 As D3DMATRIX
Dim RenderTempMat3 As D3DMATRIX

Dim fps As Integer
Dim tpf As Single
fps = 60
tpf = 1000 / fps

Dim tempo As Single, tempo2 As Single
Dim frames As Integer
tempo = Timer
Device.LightEnable 0, True
Device.LightEnable 1, False
DespZ = -45
RotY = 10
RotX = -50

Do While Salir = 0
    tempo = Timer
    If GetAsyncKeyState(vbKeyAdd) Then
        DespZ = DespZ - 1
    ElseIf GetAsyncKeyState(vbKeySubtract) Then
        DespZ = DespZ + 1
    End If
    If GetAsyncKeyState(vbKeyUp) Then
        RotX = RotX + 1
    ElseIf GetAsyncKeyState(vbKeyDown) Then
        RotX = RotX - 1
    End If
    If GetAsyncKeyState(vbKeyLeft) Then
        RotY = RotY - 1
    ElseIf GetAsyncKeyState(vbKeyRight) Then
        RotY = RotY + 1
    End If
    If GetAsyncKeyState(vbKeyPageDown) Then
        Nivel = Nivel + 1
    ElseIf GetAsyncKeyState(vbKeyPageUp) Then
        Nivel = Nivel - 1
    End If
    
    If GetAsyncKeyState(vbKeyNumpad8) Then
        tween = tween + 0.005
    ElseIf GetAsyncKeyState(vbKeyNumpad2) Then
        tween = tween - 0.005
    End If
    If tween > 2 Then tween = 2
    If tween < 0 Then tween = 0

    Call Device.Clear(0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0)
    
    Animar
    
    Device.BeginScene
        
    D3DXMatrixIdentity RenderTempMat
    D3DXMatrixTranslation RenderTempMat, 0, 0, 0
    
    'rotate the 3d object on the Z axis by 'rangle'
    Device.SetTransform D3DTS_WORLD, RenderTempMat
    'for each of the 3d object, objects
    For i = 0 To nMateriales - 1
        Device.SetMaterial MeshMaterials2(i)    'set the object
        Device.SetTexture 0, MeshTextures2(i)  'set the texture
        CuboF.DrawSubset i               'draw object
    Next
    
    'D3DXMatrixIdentity RenderTempMat
    'D3DXMatrixTranslation RenderTempMat, RotY / 20, 0, RotX / 20
    
    'rotate the 3d object on the Z axis by 'rangle'
    'Device.SetTransform D3DTS_WORLD, RenderTempMat
    'for each of the 3d object, objects
    'For i = 0 To nMateriales2 - 1
    '    Device.SetMaterial MeshMaterials(i)    'set the object
    '    Device.SetTexture 0, MeshTextures(i)  'set the texture
    '    Cubo.DrawSubset i               'draw object
    'Next
    
    
    frames = frames + 1
    If tempo2 < Timer Then
        tempo2 = Timer + 1
        frames2 = frames
        frames = 0
    End If
    If DespZ > 360 Then DespZ = DespZ - 360
    Dim textrect As RECT
        textrect.Top = 0
        textrect.bottom = 20
        textrect.Right = 100
        D3D.DrawText Fuente, &HFFFFFFFF, frames2, textrect, DT_TOP Or DT_LEFT
        
    D3DXMatrixIdentity matView
    D3DXMatrixLookAtLH matView, v3(RotY / 20 - 20, 50, -30 + RotX / 20), v3(RotY / 50, 0, RotX / 20), v3(0, 1, 0)
    'D3DXMatrixTranslation RenderTempMat, 0, 200, 0
    'D3DXMatrixMultiply matView, matView, RenderTempMat
    Device.SetTransform D3DTS_VIEW, matView
    
   
   
    Device.EndScene
        
    Call Device.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    DoEvents 'Le damos un respiro a windows para que haga sus cosas :)
    If tpf - (Timer - tempo) > 0 And Timer - tempo < tpf Then
        'Sleep (Timer - tempo)
        'Debug.Print tpf - (Timer - tempo)
    End If
Loop

End
End Sub

Public Function v3(X As Single, Y As Single, z As Single) As D3DVECTOR
v3.X = X
v3.Y = Y
v3.z = z
End Function

Public Function v2(X As Single, Y As Single) As D3DVECTOR2
v2.X = X
v2.Y = Y
End Function


Sub Animar()
Dim hresult
Dim vTemp3D As D3DVECTOR
Dim vTemp2D As D3DVECTOR2
Dim frame1() As D3DVERTEX
Dim frame2() As D3DVERTEX
Dim framefinal() As D3DVERTEX


If tween > 1 Then
ReDim frame1(Cubo.GetNumVertices)
ReDim frame2(Cubo2.GetNumVertices)
ReDim framefinal(Cubo.GetNumVertices)

hresult = D3DXMeshVertexBuffer8GetData(Cubo, 0, Len(frame1(0)) * Cubo.GetNumVertices, 0, frame1(0))
hresult = D3DXMeshVertexBuffer8GetData(Cubo3, 0, Len(frame2(0)) * Cubo.GetNumVertices, 0, frame2(0))


'//Interpolate the Vertex Data
        For i = 0 To Cubo.GetNumVertices  '//Cycle through every vertex
            '//2a. Interpolate the Positions
                D3DXVec3Lerp vTemp3D, v3(frame1(i).X, frame1(i).Y, frame1(i).z), v3(frame2(i).X, frame2(i).Y, frame2(i).z), tween - 1
                framefinal(i).X = vTemp3D.X
                framefinal(i).Y = vTemp3D.Y
                framefinal(i).z = vTemp3D.z
                
            '//2b. Interpolate the Normals
                D3DXVec3Lerp vTemp3D, v3(frame1(i).nx, frame1(i).ny, frame1(i).nz), v3(frame2(i).nx, frame2(i).ny, frame2(i).nz), tween - 1
                framefinal(i).nx = vTemp3D.X
                framefinal(i).ny = vTemp3D.Y
                framefinal(i).nz = vTemp3D.z
            
            '//2c. Interpolate the Texture Coordinates
                D3DXVec2Lerp vTemp2D, v2(frame1(i).tu, frame1(i).tv), v2(frame2(i).tu, frame2(i).tv), tween - 1
                framefinal(i).tu = vTemp2D.X
                framefinal(i).tv = vTemp2D.Y

        Next i
        
'//Stick the vertex data back into the vertex buffer
hresult = D3DXMeshVertexBuffer8SetData(CuboF, 0, Len(framefinal(0)) * Cubo.GetNumVertices, 0, framefinal(0))

Else
ReDim frame1(Cubo.GetNumVertices)
ReDim frame2(Cubo2.GetNumVertices)
ReDim framefinal(Cubo.GetNumVertices)

hresult = D3DXMeshVertexBuffer8GetData(Cubo, 0, Len(frame1(0)) * Cubo.GetNumVertices, 0, frame2(0))
hresult = D3DXMeshVertexBuffer8GetData(Cubo2, 0, Len(frame2(0)) * Cubo.GetNumVertices, 0, frame1(0))


'//Interpolate the Vertex Data
        For i = 0 To Cubo.GetNumVertices  '//Cycle through every vertex
            '//2a. Interpolate the Positions
                D3DXVec3Lerp vTemp3D, v3(frame1(i).X, frame1(i).Y, frame1(i).z), v3(frame2(i).X, frame2(i).Y, frame2(i).z), tween
                framefinal(i).X = vTemp3D.X
                framefinal(i).Y = vTemp3D.Y
                framefinal(i).z = vTemp3D.z
                
            '//2b. Interpolate the Normals
                D3DXVec3Lerp vTemp3D, v3(frame1(i).nx, frame1(i).ny, frame1(i).nz), v3(frame2(i).nx, frame2(i).ny, frame2(i).nz), tween
                framefinal(i).nx = vTemp3D.X
                framefinal(i).ny = vTemp3D.Y
                framefinal(i).nz = vTemp3D.z
            
            '//2c. Interpolate the Texture Coordinates
                D3DXVec2Lerp vTemp2D, v2(frame1(i).tu, frame1(i).tv), v2(frame2(i).tu, frame2(i).tv), tween
                framefinal(i).tu = vTemp2D.X
                framefinal(i).tv = vTemp2D.Y

        Next i
        
'//Stick the vertex data back into the vertex buffer
hresult = D3DXMeshVertexBuffer8SetData(CuboF, 0, Len(framefinal(0)) * Cubo.GetNumVertices, 0, framefinal(0))

End If
End Sub
