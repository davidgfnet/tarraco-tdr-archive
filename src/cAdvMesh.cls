VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAdvMesh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Lightmapped meshes Class by davidgf. See DX8 SDK for more info about materials
' and attribute tables
' Lightmaps doesn't share any unique materials in the mesh
Public AlphaTest As Boolean

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647
Private Const OFFSET_2 = 65536
Private Const MAXINT_2 = 32767

Private Type LMVertex
    x As Single
    y As Single
    z As Single
    tu As Single
    tv As Single
    tu2 As Single
    tv2 As Single
End Type
Private Const LMVertex_FVF = D3DFVF_XYZ Or D3DFVF_TEX2     '28 bytes per vertex ;)

Private Type MyAttr
    FaceStart As Long
    FaceCount As Long
    TexIndex As Long
    LMIndex As Long
End Type
Private MyAttrs() As MyAttr

Dim x As Long

Private VertexBuffer As Direct3DVertexBuffer8
Private Indexbuffer As Direct3DIndexBuffer8

Private MaterialsArray() As D3DMATERIAL8
Private TexturesArray() As String
Private NumMaterials As Long, NumLMMaterials As Long
Private MaterialsArrayLM() As D3DMATERIAL8
Private TexturesArrayLM() As String
Private Attr() As D3DXATTRIBUTERANGE, AttrLM() As D3DXATTRIBUTERANGE

Private NoLM As Boolean

Public Sub Draw()
If NoLM Then
    'if this model doesn't have any LM then scape here
    Call DrawNoLM
    Exit Sub
End If

Device.SetVertexShader LMVertex_FVF
Dim ind As Integer, y As Integer, num As Long

Device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SELECTARG1
Device.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
Device.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
Device.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE

'Device.SetTextureStageState 0, D3DTSS_COLORARG1, D3DTA_TEXTURE
'Device.SetTextureStageState 0, D3DTSS_COLORARG2, D3DTA_DIFFUSE
'Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
'Device.SetTextureStageState 1, D3DTSS_COLORARG1, D3DTA_TEXTURE
'Device.SetTextureStageState 1, D3DTSS_COLORARG2, D3DTA_CURRENT
'Device.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_MODULATE

'Device.SetRenderState D3DRS_AMBIENT, RGB(255, 0, 0)

Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1

For x = 0 To NumMaterials - 1
    If UCase(right(TexturesArray(x), 4)) <> ".TGA" Then
        Device.SetMaterial MaterialsArray(x)
        If TexturesArray(x) = "" Then
            Device.SetTexture 0, Nothing
        Else
            DeviceSetTexture TexturesArray(x)
        End If
        
        For y = 0 To NumLMMaterials - 1
            If AttrLM(y).FaceStart >= Attr(x).FaceStart And (AttrLM(y).FaceStart + AttrLM(y).FaceCount) <= (Attr(x).FaceStart + Attr(x).FaceCount) Then
                ind = y
                Exit For
            End If
        Next
        
        DeviceSetTexture TexturesArrayLM(ind), 1
        
        If Attr(x).FaceCount < caps.MaxPrimitiveCount Then
            Device.SetStreamSource 0, VertexBuffer, 28
            Device.DrawPrimitive D3DPT_TRIANGLELIST, Attr(x).FaceStart * 3, Attr(x).FaceCount
        Else
            For y = 0 To Attr(x).FaceCount Step caps.MaxPrimitiveCount
                Device.SetStreamSource 0, VertexBuffer, 28
                If y + caps.MaxPrimitiveCount > Attr(x).FaceCount Then
                    num = Attr(x).FaceCount - y
                Else
                    num = caps.MaxPrimitiveCount
                End If
                Device.DrawPrimitive D3DPT_TRIANGLELIST, (Attr(x).FaceStart + y) * 3, num
            Next
        End If
    End If
Next

Device.SetTexture 1, Nothing        'draw primitives without LM but tansparency
If AlphaTest Then
    Device.SetRenderState D3DRS_ALPHAREF, &H1
    Device.SetRenderState D3DRS_ALPHATESTENABLE, 1
    Device.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
End If
For x = 0 To NumMaterials - 1
    If UCase(right(TexturesArray(x), 4)) = ".TGA" Then
        Device.SetMaterial MaterialsArray(x)
        DeviceSetTexture TexturesArray(x)
            
        If Attr(x).FaceCount < caps.MaxPrimitiveCount Then
            Device.SetStreamSource 0, VertexBuffer, 28
            Device.DrawPrimitive D3DPT_TRIANGLELIST, Attr(x).FaceStart * 3, Attr(x).FaceCount
        Else
            For y = 0 To Attr(x).FaceCount Step caps.MaxPrimitiveCount
                Device.SetStreamSource 0, VertexBuffer, 28
                If y + caps.MaxPrimitiveCount > Attr(x).FaceCount Then
                    num = Attr(x).FaceCount - y
                Else
                    num = caps.MaxPrimitiveCount
                End If
                Device.DrawPrimitive D3DPT_TRIANGLELIST, (Attr(x).FaceStart + y) * 3, num
            Next
        End If
    End If
Next
If AlphaTest Then Device.SetRenderState D3DRS_ALPHATESTENABLE, 0

Device.SetTexture 0, Nothing
Device.SetTextureStageState 1, D3DTSS_COLOROP, D3DTOP_DISABLE
Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
End Sub

Public Sub LoadFromFile(ByVal Mesh As String, ByVal MeshLM As String, ByVal path As String, ByVal Tertiary As Boolean)
Dim FinalVertices() As LMVertex
Dim hresult As Long
Dim MeshObj As D3DXMesh
Dim MeshObj2 As D3DXMesh
Dim AdjBuf As D3DXBuffer, MatBuf As D3DXBuffer

'---------------------- LOAD STANDARD MESH ------------------------
Set MeshObj = Direct3DX.LoadMeshFromX(Mesh, D3DXMESH_MANAGED Or D3DXMESH_32BIT, Device, AdjBuf, MatBuf, NumMaterials)
ReDim MaterialsArray(NumMaterials - 1)
ReDim TexturesArray(NumMaterials - 1)
ReDim Attr(NumMaterials - 1)
OptimizeMesh MeshObj, AdjBuf, True

For x = 0 To NumMaterials - 1
    Direct3DX.BufferGetMaterial MatBuf, x, MaterialsArray(x)
    MaterialsArray(x).Ambient = MaterialsArray(x).diffuse
    TexturesArray(x) = UCase(GetFileName(Direct3DX.BufferGetTextureName(MatBuf, x)))
    'If TexturesArray(x) <> "" Then Call AddTexture(TexturesArray(x))
Next
MeshObj.GetAttributeTable Attr(0), hresult
Set AdjBuf = Nothing: Set MatBuf = Nothing

If FileExists(MeshLM) = False Then NoLM = True

If NoLM = False Then
    '---------------------- LOAD LIGHTMAPPED MESH ------------------------
    Set MeshObj2 = Direct3DX.LoadMeshFromX(MeshLM, D3DXMESH_MANAGED Or D3DXMESH_32BIT, Device, AdjBuf, MatBuf, NumLMMaterials)
    ReDim MaterialsArrayLM(NumLMMaterials - 1)
    ReDim TexturesArrayLM(NumLMMaterials - 1)
    ReDim AttrLM(NumLMMaterials - 1)
    OptimizeMesh MeshObj2, AdjBuf, True
    
    For x = 0 To NumLMMaterials - 1
        Direct3DX.BufferGetMaterial MatBuf, x, MaterialsArrayLM(x)
        MaterialsArrayLM(x).Ambient = MaterialsArrayLM(x).diffuse
        TexturesArrayLM(x) = UCase(GetFileName(Direct3DX.BufferGetTextureName(MatBuf, x)))
        'If TexturesArrayLM(x) <> "" Then Call AddTexture(TexturesArrayLM(x))
    Next
    Set AdjBuf = Nothing: Set MatBuf = Nothing
    MeshObj2.GetAttributeTable AttrLM(0), hresult
End If

Set VertexBuffer = Device.CreateVertexBuffer(28 * MeshObj.GetNumFaces * 3, D3DUSAGE_WRITEONLY, LMVertex_FVF, D3DPOOL_MANAGED)


' -------------------------------------------------------
' ------ Copy Verts to VB with a couple of UVWs ---------
' -------------------------------------------------------

ReDim FinalVertices(MeshObj.GetNumFaces * 3 - 1)

Dim vertices() As D3DVERTEX
Dim Desc As D3DINDEXBUFFER_DESC
Dim IBuf As Direct3DIndexBuffer8
Dim y As Long, ignore As Long, indices32() As Long

ReDim vertices(MeshObj.GetNumVertices - 1)
ReDim tris(MeshObj.GetNumFaces * 3 - 1)

hresult = D3DXMeshVertexBuffer8GetData(MeshObj, 0, Len(vertices(0)) * MeshObj.GetNumVertices, 0, vertices(0))

Set IBuf = MeshObj.GetIndexBuffer()
IBuf.Lock 0, 0, ignore, 16
IBuf.GetDesc Desc
IBuf.Unlock

ReDim indices32(Desc.size / 4)   '4 as we use 32 bit mesh , 2 if we use 16 bit
D3DXMeshIndexBuffer8GetData MeshObj, 0, Desc.size, 0, indices32(0)

For y = 0 To MeshObj.GetNumFaces * 3 - 1 Step 3
    FinalVertices(y).x = vertices(indices32(y)).x
    FinalVertices(y).y = vertices(indices32(y)).y
    FinalVertices(y).z = vertices(indices32(y)).z
    FinalVertices(y + 1).x = vertices(indices32(y + 1)).x
    FinalVertices(y + 1).y = vertices(indices32(y + 1)).y
    FinalVertices(y + 1).z = vertices(indices32(y + 1)).z
    FinalVertices(y + 2).x = vertices(indices32(y + 2)).x
    FinalVertices(y + 2).y = vertices(indices32(y + 2)).y
    FinalVertices(y + 2).z = vertices(indices32(y + 2)).z
    
    FinalVertices(y).tu = vertices(indices32(y)).tu
    FinalVertices(y).tv = vertices(indices32(y)).tv
    FinalVertices(y + 1).tu = vertices(indices32(y + 1)).tu
    FinalVertices(y + 1).tv = vertices(indices32(y + 1)).tv
    FinalVertices(y + 2).tu = vertices(indices32(y + 2)).tu
    FinalVertices(y + 2).tv = vertices(indices32(y + 2)).tv
Next

ReDim indices32(0)
ReDim vertices(0)           'destroy all temp objects!!
Set IBuf = Nothing

If NoLM = False Then
    ReDim vertices(MeshObj2.GetNumVertices - 1)
    ReDim tris(MeshObj2.GetNumFaces * 3 - 1)

    hresult = D3DXMeshVertexBuffer8GetData(MeshObj2, 0, Len(vertices(0)) * MeshObj2.GetNumVertices, 0, vertices(0))
        
    Set IBuf = MeshObj2.GetIndexBuffer()
    IBuf.Lock 0, 0, ignore, 16
    IBuf.GetDesc Desc
    IBuf.Unlock
    
    ReDim indices32(Desc.size / 4)   '4 as we use 32 bit mesh , 2 if we use 16 bit
    D3DXMeshIndexBuffer8GetData MeshObj2, 0, Desc.size, 0, indices32(0)
        
    For y = 0 To MeshObj.GetNumFaces * 3 - 1 Step 3
        FinalVertices(y).tu2 = vertices(indices32(y)).tu
        FinalVertices(y).tv2 = vertices(indices32(y)).tv
        FinalVertices(y + 1).tu2 = vertices(indices32(y + 1)).tu
        FinalVertices(y + 1).tv2 = vertices(indices32(y + 1)).tv
        FinalVertices(y + 2).tu2 = vertices(indices32(y + 2)).tu
        FinalVertices(y + 2).tv2 = vertices(indices32(y + 2)).tv
    Next
    ReDim indices32(0)
    ReDim vertices(0)           'destroy all temp objects!!
    Set IBuf = Nothing
End If

D3DVertexBuffer8SetData VertexBuffer, 0, 28 * MeshObj.GetNumFaces * 3, 0, FinalVertices(0)
ReDim FinalVertices(0)
Set MeshObj2 = Nothing: Set MeshObj = Nothing

Dim z As Long, j As Long

If NoLM Then Exit Sub
If Tertiary Then
    y = UBound(GameTexturesTer)
    For x = 0 To UBound(TexturesArrayLM)
        If TexturesArrayLM(x) <> "" Then
            For j = 1 To UBound(GameTexturesTer)
                If GameTexturesTer(j).filename = TexturesArrayLM(x) Then GoTo AlreadyLoaded1
            Next
            z = z + 1
            ReDim Preserve GameTexturesTer(y + z)
            GameTexturesTer(y + z).filename = TexturesArrayLM(x)
            Set GameTexturesTer(y + z).texture = LoadTextureAndReturn(path & TexturesArrayLM(x))
        End If
AlreadyLoaded1:
    Next
Else
    y = UBound(GameTexturesSec)
    For x = 0 To UBound(TexturesArrayLM)
        If TexturesArrayLM(x) <> "" Then
            For j = 1 To UBound(GameTexturesSec)
                If GameTexturesSec(j).filename = TexturesArrayLM(x) Then GoTo AlreadyLoaded2
            Next
            z = z + 1
            ReDim Preserve GameTexturesSec(y + z)
            GameTexturesSec(y + z).filename = TexturesArrayLM(x)
            Set GameTexturesSec(y + z).texture = LoadTextureAndReturn(path & TexturesArrayLM(x))
        End If
AlreadyLoaded2:
    Next
End If
End Sub

Private Function CalculateDistance(v1 As LMVertex, v2 As LMVertex) As Double
CalculateDistance = Sqr((v1.x - v2.x) ^ 2 + (v1.y - v2.y) ^ 2 + (v1.z - v2.z) ^ 2)
End Function

Private Sub DrawNoLM()
Device.SetVertexShader LMVertex_FVF
Device.SetTexture 1, Nothing        'per si de cas!
Dim y As Integer, num As Long

Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1

For x = 0 To NumMaterials - 1
    If UCase(right(TexturesArray(x), 4)) <> ".TGA" Then
        Device.SetMaterial MaterialsArray(x)
        If TexturesArray(x) = "" Then
            Device.SetTexture 0, Nothing
        Else
            DeviceSetTexture TexturesArray(x)
        End If
        
        If Attr(x).FaceCount < caps.MaxPrimitiveCount Then
            Device.SetStreamSource 0, VertexBuffer, 28
            Device.DrawPrimitive D3DPT_TRIANGLELIST, Attr(x).FaceStart * 3, Attr(x).FaceCount
        Else
            For y = 0 To Attr(x).FaceCount Step caps.MaxPrimitiveCount
                Device.SetStreamSource 0, VertexBuffer, 28
                If y + caps.MaxPrimitiveCount > Attr(x).FaceCount Then
                    num = Attr(x).FaceCount - y
                Else
                    num = caps.MaxPrimitiveCount
                End If
                Device.DrawPrimitive D3DPT_TRIANGLELIST, (Attr(x).FaceStart + y) * 3, num
            Next
        End If
    End If
Next

If AlphaTest Then
    Device.SetRenderState D3DRS_ALPHAREF, &H1
    Device.SetRenderState D3DRS_ALPHATESTENABLE, 1
    Device.SetRenderState D3DRS_ALPHAFUNC, D3DCMP_GREATEREQUAL
End If
For x = 0 To NumMaterials - 1
    If UCase(right(TexturesArray(x), 4)) = ".TGA" Then
        Device.SetMaterial MaterialsArray(x)
        DeviceSetTexture TexturesArray(x)
            
        If Attr(x).FaceCount < caps.MaxPrimitiveCount Then
            Device.SetStreamSource 0, VertexBuffer, 28
            Device.DrawPrimitive D3DPT_TRIANGLELIST, Attr(x).FaceStart * 3, Attr(x).FaceCount
        Else
            For y = 0 To Attr(x).FaceCount Step caps.MaxPrimitiveCount
                Device.SetStreamSource 0, VertexBuffer, 28
                If y + caps.MaxPrimitiveCount > Attr(x).FaceCount Then
                    num = Attr(x).FaceCount - y
                Else
                    num = caps.MaxPrimitiveCount
                End If
                Device.DrawPrimitive D3DPT_TRIANGLELIST, (Attr(x).FaceStart + y) * 3, num
            Next
        End If
    End If
Next
If AlphaTest Then Device.SetRenderState D3DRS_ALPHATESTENABLE, 0
End Sub

