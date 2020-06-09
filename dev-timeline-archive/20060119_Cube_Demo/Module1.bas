Attribute VB_Name = "Module1"

Option Explicit
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Dx As DirectX8
Public Dx3D As Direct3D8
Public Device As Direct3DDevice8
Public DispMode As D3DDISPLAYMODE
Public D3D As D3DX8
Public Cubo As D3DXMesh
Public MtrlCubo As D3DXBuffer
Dim MeshMaterials() As D3DMATERIAL8
Dim MeshTextures() As Direct3DTexture8

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
Global DespZ As Long
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

D3DXMatrixPerspectiveFovLH matProj, PI / 4, PI * 0.3, 5, 1000
Device.SetTransform D3DTS_PROJECTION, matProj

'*******************************************************************
Luz.Type = D3DLIGHT_SPOT
Luz.Position = v3(-100, 0, 0)
Luz.diffuse.r = 1
Luz.diffuse.g = 1
Luz.diffuse.b = 1
Luz.Direction = v3(0, 0, 0)
Luz.Theta = PI * 10
Luz.Phi = PI * 2
Luz.Attenuation1 = 0.0025    'atenuacion de la luz al alejarse
Luz.Range = 1000

Luz2.Type = D3DLIGHT_SPOT
Luz2.Position = v3(100, 0, 0)
Luz2.diffuse.r = 1
Luz2.diffuse.g = 1
Luz2.diffuse.b = 1
Luz2.Direction = v3(0, 0, 0)
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

Set Cubo = D3D.LoadMeshFromX(App.Path & "\cubo.x", D3DXMESH_MANAGED, Device, Nothing, MtrlCubo, nMateriales)

ReDim MeshMaterials(nMateriales - 1) As D3DMATERIAL8
ReDim MeshTextures(nMateriales - 1) As Direct3DTexture8

Dim X As Integer
For X = 0 To nMateriales - 1
    D3D.BufferGetMaterial MtrlCubo, X, MeshMaterials(X)
    If D3D.BufferGetTextureName(MtrlCubo, X) <> "" Then
        Set MeshTextures(X) = D3D.CreateTextureFromFileEx(Device, App.Path & "\" & D3D.BufferGetTextureName(MtrlCubo, X), 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    End If
Next

Set MtrlCubo = Nothing
Dim RenderTempMat As D3DMATRIX, i As Long
Dim RenderTempMat2 As D3DMATRIX
Dim RenderTempMat3 As D3DMATRIX

Dim tempo As Single
Dim frames As Integer
tempo = GetTickCount() + 1
Device.LightEnable 0, True
Device.LightEnable 1, True

Do While Salir = 0
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
        RotY = RotY + 1
    ElseIf GetAsyncKeyState(vbKeyRight) Then
        RotY = RotY - 1
    End If
    If GetAsyncKeyState(vbKeyPageDown) Then
        Nivel = Nivel + 1
    ElseIf GetAsyncKeyState(vbKeyPageUp) Then
        Nivel = Nivel - 1
    End If

    Call Device.Clear(0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0)
    
    Device.BeginScene
    
     
        D3DXMatrixIdentity matView
    D3DXMatrixLookAtLH matView, v3(0, 200, 0), v3(0, 0, 0), v3(0, 1, 0)
    Device.SetTransform D3DTS_WORLD, matView
    
    
    
        D3DXMatrixIdentity RenderTempMat
    D3DXMatrixIdentity RenderTempMat2
    D3DXMatrixIdentity RenderTempMat3
    
    D3DXMatrixTranslation RenderTempMat, 0, 0, DespZ + 50
    D3DXMatrixRotationY RenderTempMat2, RotY / 180
    D3DXMatrixRotationX RenderTempMat3, RotX / 180
    
    D3DXMatrixMultiply RenderTempMat2, RenderTempMat2, RenderTempMat
    D3DXMatrixMultiply RenderTempMat3, RenderTempMat3, RenderTempMat2
    
    'rotate the 3d object on the Z axis by 'rangle'
    Device.SetTransform D3DTS_WORLD, RenderTempMat3
    'for each of the 3d object, objects
    For i = 0 To nMateriales - 1
        Device.SetMaterial MeshMaterials(i)    'set the object
        Device.SetTexture 0, MeshTextures(i)  'set the texture
        Cubo.DrawSubset i               'draw object
    Next
    
        
    D3DXMatrixIdentity RenderTempMat
    D3DXMatrixIdentity RenderTempMat2
    D3DXMatrixIdentity RenderTempMat3
    
    D3DXMatrixTranslation RenderTempMat, -50, 0, DespZ + 150
    D3DXMatrixRotationY RenderTempMat2, RotY / 180
    D3DXMatrixRotationX RenderTempMat3, RotX / 180
    
    D3DXMatrixMultiply RenderTempMat2, RenderTempMat2, RenderTempMat
    D3DXMatrixMultiply RenderTempMat3, RenderTempMat3, RenderTempMat2
    
    
    Device.SetTransform D3DTS_WORLD, RenderTempMat3
    
    For i = 0 To nMateriales - 1
        Device.SetMaterial MeshMaterials(i)    'set the object
        Device.SetTexture 0, MeshTextures(i)  'set the texture
        Cubo.DrawSubset i               'draw object
    Next
    
    D3DXMatrixIdentity RenderTempMat
    D3DXMatrixIdentity RenderTempMat2
    D3DXMatrixIdentity RenderTempMat3
   
    D3DXMatrixTranslation RenderTempMat, -50, 0, DespZ + 50
    D3DXMatrixRotationY RenderTempMat2, RotY / 180
    D3DXMatrixRotationX RenderTempMat3, RotX / 180
    
    D3DXMatrixMultiply RenderTempMat2, RenderTempMat2, RenderTempMat
    D3DXMatrixMultiply RenderTempMat3, RenderTempMat3, RenderTempMat2
    
    
    Device.SetTransform D3DTS_WORLD, RenderTempMat3
    
    For i = 0 To nMateriales - 1
        Device.SetMaterial MeshMaterials(i)    'set the object
        Device.SetTexture 0, MeshTextures(i)  'set the texture
        Cubo.DrawSubset i               'draw object
    Next

    
    D3DXMatrixIdentity RenderTempMat
    D3DXMatrixIdentity RenderTempMat2
    D3DXMatrixIdentity RenderTempMat3
    
    D3DXMatrixTranslation RenderTempMat, 0, 0, DespZ + 150
    D3DXMatrixRotationY RenderTempMat2, RotY / 180
    D3DXMatrixRotationX RenderTempMat3, RotX / 180
    
    D3DXMatrixMultiply RenderTempMat2, RenderTempMat2, RenderTempMat
    D3DXMatrixMultiply RenderTempMat3, RenderTempMat3, RenderTempMat2
    
    
    Device.SetTransform D3DTS_WORLD, RenderTempMat3
    
    For i = 0 To nMateriales - 1
        Device.SetMaterial MeshMaterials(i)    'set the object
        Device.SetTexture 0, MeshTextures(i)  'set the texture
        Cubo.DrawSubset i               'draw object
    Next
    
    D3DXMatrixIdentity RenderTempMat
    D3DXMatrixIdentity RenderTempMat2
    D3DXMatrixIdentity RenderTempMat3
    
    D3DXMatrixTranslation RenderTempMat, 50, 0, DespZ + 150
    D3DXMatrixRotationY RenderTempMat2, RotY / 180
    D3DXMatrixRotationX RenderTempMat3, RotX / 180
    
    D3DXMatrixMultiply RenderTempMat2, RenderTempMat2, RenderTempMat
    D3DXMatrixMultiply RenderTempMat3, RenderTempMat3, RenderTempMat2
    
    
    Device.SetTransform D3DTS_WORLD, RenderTempMat3
    
    For i = 0 To nMateriales - 1
        Device.SetMaterial MeshMaterials(i)    'set the object
        Device.SetTexture 0, MeshTextures(i)  'set the texture
        Cubo.DrawSubset i               'draw object
    Next
    
    D3DXMatrixIdentity RenderTempMat
    D3DXMatrixIdentity RenderTempMat2
    D3DXMatrixIdentity RenderTempMat3
    
    D3DXMatrixTranslation RenderTempMat, 50, 0, DespZ + 50
    D3DXMatrixRotationY RenderTempMat2, RotY / 180
    D3DXMatrixRotationX RenderTempMat3, RotX / 180
    
    D3DXMatrixMultiply RenderTempMat2, RenderTempMat2, RenderTempMat
    D3DXMatrixMultiply RenderTempMat3, RenderTempMat3, RenderTempMat2
    
    
    Device.SetTransform D3DTS_WORLD, RenderTempMat3
    
    For i = 0 To nMateriales - 1
        Device.SetMaterial MeshMaterials(i)    'set the object
        Device.SetTexture 0, MeshTextures(i)  'set the texture
        Cubo.DrawSubset i               'draw object
    Next
    



    frames = frames + 1
    If tempo < GetTickCount() Then
        tempo = GetTickCount() + 1000
        frames2 = frames
        frames = 0
    End If
    Dim textrect As RECT
        textrect.Top = 0
        textrect.bottom = 20
        textrect.Right = 100
        D3D.DrawText Fuente, &HFFFFFFFF, frames2, textrect, DT_TOP Or DT_LEFT
        
   
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
