VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cal3DModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'############################################
'## CAL 3D IMPLEMENTATION CLASS            ##
'## USES Cal3DVb to access to Cal3D class  ##
'## Created by davidgf (davidgf@tinet.org) ##
'############################################

Private Type CalAnim
    AnimID As Long
    AnimIDString As String
    AnimLength As Single
End Type

Private Cal3DCore As Cal3DObject
Private Cal3DAnims() As CalAnim
Private CalTextures() As String

Private EmptyMaterial As D3DMATERIAL8
Private s As Long
Public ModelName As String
Public ModelDimensions As Single

Private Sub Class_Initialize()
Set Cal3DCore = New Cal3DObject
ReDim Cal3DAnims(0)
ReDim CadTextures(0)
EmptyMaterial.Ambient.r = 1: EmptyMaterial.Ambient.g = 1: EmptyMaterial.Ambient.b = 1: EmptyMaterial.Ambient.a = 1
EmptyMaterial.diffuse.r = 1: EmptyMaterial.diffuse.g = 1: EmptyMaterial.diffuse.b = 1: EmptyMaterial.diffuse.a = 1
EmptyMaterial.specular.r = 1: EmptyMaterial.specular.g = 1: EmptyMaterial.specular.b = 1: EmptyMaterial.specular.a = 1
EmptyMaterial.emissive.r = 1: EmptyMaterial.emissive.g = 1: EmptyMaterial.emissive.b = 1: EmptyMaterial.emissive.a = 1
End Sub

Private Sub Class_Terminate()
Set Cal3DCore = Nothing
ReDim Cal3DAnims(0)
ReDim CadTextures(0)
End Sub

'Loads all the data given a path and a file mash
Public Sub LoadData(ByVal path As String, ByVal Mask As String)
Dim file As String, result As Long
If right(path, 1) <> "\" Then path = path & "\"

If FileExists(path & Mask & ".csf") Then Cal3DCore.loadskeleton path & Mask & ".csf", result
If result <> 1 Then Debug.Print "--- ERROR LOADING CAL3D SKELETON ---"

file = Dir(path & Mask & "*.*", vbArchive)
Do While file <> ""
    If FileExists(path & file) Then
        Select Case UCase(right(file, 3))
        Case "CMF"
            Cal3DCore.loadmesh path & file, result
            If result = -1 Then Debug.Print "--- ERROR LOADING CAL3D MESH ---"
        Case "CRF"
            Cal3DCore.loadmaterial path & file, result
            If result = -1 Then Debug.Print "--- ERROR LOADING CAL3D MATERIAL ---"
        End Select
    End If
    file = Dir()
Loop
End Sub

'Load anim and assign AnimName
Public Sub LoadAnim(ByVal file As String, ByVal AnimID As String)
Dim result As Long
If FileExists(file) Then
    Cal3DCore.loadanimation file, result
    If result <> -1 Then
        ReDim Preserve Cal3DAnims(UBound(Cal3DAnims) + 1)
        Cal3DAnims(UBound(Cal3DAnims)).AnimID = result
        Cal3DAnims(UBound(Cal3DAnims)).AnimIDString = AnimID
        Cal3DCore.getanimationduration result, Cal3DAnims(UBound(Cal3DAnims)).AnimLength
        Cal3DAnims(UBound(Cal3DAnims)).AnimLength = Cal3DAnims(UBound(Cal3DAnims)).AnimLength * 1000
    End If
End If
End Sub

Public Function AnimLength(ByVal AnimID As String) As Single
Dim x As Long
For x = 1 To UBound(Cal3DAnims)
    If Cal3DAnims(x).AnimIDString = AnimID Then
        AnimLength = Cal3DAnims(x).AnimLength
        Exit For
    End If
Next
End Function

'Extract Texture Filenames before the creation of any model
Public Sub NowReady()
Dim TextureList(6400) As Long, NumTex As Long   'Max 100 Tex per Model
Dim bytes() As Byte, x As Long, MyString As String

Cal3DCore.ready TextureList(0), NumTex
ReDim CalTextures(NumTex - 1)

ReDim bytes(NumTex * 256 - 1)
CopyMemory ByVal VarPtr(bytes(0)), ByVal VarPtr(TextureList(0)), NumTex * 256
For x = 0 To NumTex * 256 - 1
    MyString = MyString & Chr(bytes(x))
Next
For x = 0 To NumTex - 1
    CalTextures(x) = UCase(Trim(mid(MyString, 256 * x + 1, 256)))
Next
End Sub

Public Sub LoadTexturesToSec(ByVal path As String)
Dim x As Long, y As Long, z As Long
y = UBound(GameTexturesSec)
For x = 0 To UBound(CalTextures)
    If CalTextures(x) <> "" Then
        z = z + 1
        ReDim Preserve GameTexturesSec(y + z)
        GameTexturesSec(y + z).filename = CalTextures(x)
        Set GameTexturesSec(y + z).texture = LoadTextureAndReturn(path & CalTextures(x))
    End If
Next
End Sub

'Sets the LOD (0.0 - 1.0) of the given ModelID
Public Sub SetLevelOfDetail(ByVal LOD As Single, ByVal ModelID As Long)
Cal3DCore.SetLOD ModelID, LOD
End Sub

'creates a model based on the loaded meshes and materials
'which can be animated with any animation loaded
Public Function CreateModel() As Long
Cal3DCore.CreateModel CreateModel
End Function

Public Sub ClearCycle(ByVal ModelID As Long, ByVal AnimIDString As String, ByVal Delay As Single)
Dim AnimID As Long
For s = 1 To UBound(Cal3DAnims)
    If Cal3DAnims(s).AnimIDString = AnimIDString Then
        AnimID = Cal3DAnims(s).AnimID
        Exit For
    End If
Next
Cal3DCore.ClearCycle ModelID, AnimID, Delay
End Sub

Public Sub BlendCycle(ByVal ModelID As Long, ByVal AnimIDString As String, ByVal Weight As Single, ByVal Delay As Single)
Dim AnimID As Long
For s = 1 To UBound(Cal3DAnims)
    If Cal3DAnims(s).AnimIDString = AnimIDString Then
        AnimID = Cal3DAnims(s).AnimID
        Exit For
    End If
Next
Cal3DCore.BlendCycle ModelID, AnimID, Weight, Delay
End Sub

Public Sub ExecuteAction(ByVal ModelID As Long, ByVal AnimIDString As String, ByVal Weight As Single, ByVal DelayIn As Single, ByVal DelayOut As Single)
Dim AnimID As Long
For s = 1 To UBound(Cal3DAnims)
    If Cal3DAnims(s).AnimIDString = AnimIDString Then
        AnimID = Cal3DAnims(s).AnimID
        Exit For
    End If
Next
Cal3DCore.ExecuteAction ModelID, AnimID, DelayIn, DelayOut, Weight
End Sub

Public Sub Update(ByVal ModelID As Long, ByVal Time As Double)
Cal3DCore.Update ModelID, Time
End Sub

Public Sub SetAnimationTime(ByVal ModelID As Long, ByVal Time As Single)
Cal3DCore.SetAnimationTime ModelID, Time
End Sub

Public Sub Render(ByVal ModelID)
'RENDERS THE MODEL USING D3D8.1
'Max 100000 Verts with 100000 faces and 100 materials and textures
Dim NumVertices As Long, NumFaces As Long, NumMaterials As Long
Dim vertices(500000) As Single, indices(300000) As Long, Textures(300) As Long, uvcoords(200000) As Single
Dim Atributes(300) As Long 'two int per material/subset
Dim VertexBuffer As Direct3DVertexBuffer8, Indexbuffer As Direct3DIndexBuffer8, UVBuffer As Direct3DVertexBuffer8
Dim y As Long, num As Long

If VSCapable Then
    Cal3DCore.Render ModelID, NumVertices, NumFaces, NumMaterials, vertices(0), uvcoords(0), indices(0), Textures(0), Atributes(0), 1
    
    Set VertexBuffer = Device.CreateVertexBuffer(NumVertices * 12, D3DUSAGE_WRITEONLY, D3DFVF_XYZ, D3DPOOL_MANAGED)
    Set UVBuffer = Device.CreateVertexBuffer(NumVertices * 8, D3DUSAGE_WRITEONLY, D3DFVF_TEX1, D3DPOOL_MANAGED)
    Set Indexbuffer = Device.CreateIndexBuffer(NumFaces * 12, D3DUSAGE_WRITEONLY, D3DFMT_INDEX32, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData VertexBuffer, 0, NumVertices * 12, 0, vertices(0)
    D3DVertexBuffer8SetData UVBuffer, 0, NumVertices * 8, 0, uvcoords(0)
    D3DIndexBuffer8SetData Indexbuffer, 0, NumFaces * 12, 0, indices(0)
    
    Device.SetStreamSource 0, VertexBuffer, 12
    Device.SetStreamSource 1, UVBuffer, 8
    Device.SetMaterial EmptyMaterial
    Device.SetVertexShader CharShader
    
    For s = 0 To NumMaterials - 1
        If CalTextures(Textures(s)) = "" Then
            Device.SetTexture 0, Nothing
        Else
            DeviceSetTexture CalTextures(Textures(s))
        End If
        
        Device.SetIndices Indexbuffer, Atributes(s * 3 + 2)
        Device.DrawIndexedPrimitive D3DPT_TRIANGLELIST, 0, Atributes(s * 3 + 1) * 3, Atributes(s * 3) * 3, Atributes(s * 3 + 1)
    Next
Else
    Cal3DCore.Render ModelID, NumVertices, NumFaces, NumMaterials, vertices(0), uvcoords(0), indices(0), Textures(0), Atributes(0), 0
        
    Set VertexBuffer = Device.CreateVertexBuffer(NumFaces * 3 * 20, D3DUSAGE_WRITEONLY, D3DFVF_XYZ Or D3DFVF_TEX1, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData VertexBuffer, 0, NumFaces * 3 * 20, 0, vertices(0)
        
    Device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
    Device.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
    Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    Device.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_DISABLE
        
    Device.SetStreamSource 0, VertexBuffer, 20
    Device.SetMaterial EmptyMaterial
    Device.SetVertexShader D3DFVF_XYZ Or D3DFVF_TEX1
    
    For s = 0 To NumMaterials - 1
        If CalTextures(Textures(s)) = "" Then
            Device.SetTexture 0, Nothing
        Else
            DeviceSetTexture CalTextures(Textures(s))
        End If
        
        If Atributes(s * 3 + 1) < caps.MaxPrimitiveCount Then
            Device.DrawPrimitive D3DPT_TRIANGLELIST, Atributes(s * 3) * 3, Atributes(s * 3 + 1)
        Else
            For y = 0 To Atributes(s * 3 + 1) Step caps.MaxPrimitiveCount
                If y + caps.MaxPrimitiveCount > Atributes(s * 3 + 1) Then
                    num = Atributes(s * 3 + 1) - y
                Else
                    num = caps.MaxPrimitiveCount
                End If
                Device.DrawPrimitive D3DPT_TRIANGLELIST, (Atributes(s * 3) + y) * 3, num
                
                Device.DrawPrimitive D3DPT_TRIANGLELIST, Atributes(s * 3) * 3, Atributes(s * 3 + 1)
            Next
        End If
    Next
End If

Set VertexBuffer = Nothing
Set Indexbuffer = Nothing
Set UVBuffer = Nothing
End Sub

'retrieves materials
Private Sub LongToMaterial(ByRef MyLong As Long, ByRef Material As D3DMATERIAL8)
Dim bytes(3) As Byte
With Material
    CopyMemory ByVal VarPtr(bytes(0)), ByVal VarPtr(MyLong), 4
    .Ambient.r = bytes(1) / 255: .Ambient.g = bytes(2) / 255: .Ambient.b = bytes(3) / 255: .Ambient.a = bytes(0) / 255
    CopyMemory ByVal VarPtr(bytes(0)), ByVal (VarPtr(MyLong) + 4), 4
    .diffuse.r = bytes(1) / 255: .diffuse.g = bytes(2) / 255: .diffuse.b = bytes(3) / 255: .diffuse.a = bytes(0) / 255
    CopyMemory ByVal VarPtr(bytes(0)), ByVal (VarPtr(MyLong) + 8), 4
    .specular.r = bytes(1) / 255: .specular.g = bytes(2) / 255: .specular.b = bytes(3) / 255: .specular.a = bytes(0) / 255
End With
End Sub
