Attribute VB_Name = "AuxFunctions"
Option Explicit

'-----------------------------------------------------------------------------------------
'----------------------------------- Register Functions ----------------------------------
'-----------------------------------------------------------------------------------------

Public Function RegRead(ByVal ValueName As String) As String
RegRead = Reg.GetRegString("HKLM\Software\" & RegName & "\", ValueName)
End Function

Public Sub RegSave(ByVal ValueName As String, ByVal Value As String)
Reg.SetReg "HKLM\Software\" & RegName & "\", ValueName, Value
End Sub

Public Function ParseDM(ByVal DM As String) As D3DDISPLAYMODE
ParseDM.width = Val(left(DM, 4))
ParseDM.height = Val(mid(DM, 5, 4))
ParseDM.RefreshRate = Val(mid(DM, 9, 4))
ParseDM.Format = Val(right(DM, 4))

Dim x As Long, Ok As Boolean
For x = 0 To Direct3D.GetAdapterModeCount(0) - 1
    If ParseDM.RefreshRate <> 0 Then           'refresh rate defined, chek if it is valid
        If DModesArray(x).width = ParseDM.width And _
            DModesArray(x).height = ParseDM.height And _
            DModesArray(x).RefreshRate = ParseDM.RefreshRate And _
            DModesArray(x).Format = ParseDM.Format Then
            Ok = True
            Exit For
        End If
    Else         'refresh rate = 0, auto rr, no check
        If DModesArray(x).width = ParseDM.width And _
            DModesArray(x).height = ParseDM.height And _
            DModesArray(x).Format = ParseDM.Format Then
            Ok = True
            Exit For
        End If
    End If
Next

If Ok = False Then
    Direct3D.GetAdapterDisplayMode 0, ParseDM
End If
End Function

'-----------------------------------------------------------------------------------------
'----------------------------------- Maths Functions -------------------------------------
'-----------------------------------------------------------------------------------------

Public Function VecMultiply(v1 As D3DVECTOR, ByVal factor As Single) As D3DVECTOR
VecMultiply.x = v1.x * factor
VecMultiply.y = v1.y * factor
VecMultiply.z = v1.z * factor
End Function

Public Function VecDistFast(v1 As D3DVECTOR, v2 As D3DVECTOR) As Double
VecDistFast = (v1.x - v2.x) ^ 2 + (v1.y - v2.y) ^ 2 + (v1.z - v2.z) ^ 2
End Function

Public Function VecDistFastXZ(v1 As D3DVECTOR, v2 As D3DVECTOR) As Double
VecDistFastXZ = (v1.x - v2.x) ^ 2 + (v1.z - v2.z) ^ 2
End Function

Public Function CalculateDistance(v1 As D3DVECTOR, v2 As D3DVECTOR) As Single
CalculateDistance = Sqr((v1.x - v2.x) ^ 2 + (v1.y - v2.y) ^ 2 + (v1.z - v2.z) ^ 2)
End Function

Public Function CalculateDistanceXZ(v1 As D3DVECTOR, v2 As D3DVECTOR) As Single
CalculateDistanceXZ = Sqr((v1.x - v2.x) ^ 2 + (v1.z - v2.z) ^ 2)
End Function

Public Function v3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As D3DVECTOR
v3.x = x
v3.y = y
v3.z = z
End Function

Public Function Normalize(v As D3DVECTOR) As D3DVECTOR
On Local Error Resume Next
Dim module As Double
module = Sqr(v.x ^ 2 + v.y ^ 2 + v.z ^ 2)
If module = 0 Then Exit Function
Normalize.x = v.x / module
Normalize.y = v.y / module
Normalize.z = v.z / module
End Function

'----- Color function ------

Public Function ColorValue(ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single) As D3DCOLORVALUE
ColorValue.r = r
ColorValue.g = g
ColorValue.b = b
ColorValue.a = a
End Function

'-----------------------------------------------------------------------------
'---------------------------- Unpacking Functions ----------------------------
'-----------------------------------------------------------------------------

Public Function ExtractFile(ByVal file As String, ByVal Dir As String) As Boolean
'On Local Error GoTo Out
If FileExists(file) = False Then Exit Function
Dim filepath As String
filepath = Dir
If right(filepath, 1) <> "\" Then filepath = filepath & "\"
Dim ff As Integer, ff2 As Integer, x As Long
Dim cad As String, ln As Long, Max As Long, num As Long, cad2 As String
ff = FreeFile
Open file For Binary As #ff
    cad = Space(14)
    Get #ff, , cad
    Get #ff, , num
    
    For x = 1 To num
        Get #ff, , ln
        cad = Space(ln)
        Get #ff, , cad
        cad = Code(cad)
        ff2 = FreeFile
        Open filepath & cad For Binary As #ff2
            Get #ff, , ln
            If ln <= 1127 Then
                cad = Space(ln)
                Get #ff, , cad
                cad = Code(cad)
                Put #ff2, , cad
            Else
                cad2 = Space(ln - 1127)
                Get #ff, , cad2
                cad = Space(1127)
                Get #ff, , cad
                cad = Code(cad)
                Put #ff2, , cad
                Put #ff2, , cad2
            End If
        Close #ff2
    Next
Close #ff
ExtractFile = True
Out:
End Function

Public Function Code(var As String) As String
Dim n As Integer
Code = ""
For n = 1 To Len(var)
    Code = Code & Chr(255 - Asc(mid$(var, n, 1)))
Next
End Function

Public Function DecodeFile(ByVal file As String, ByVal file2 As String) As Boolean
On Local Error Resume Next
Kill file2
On Local Error GoTo Out:
Dim ff As Integer, ff2 As Integer
Dim tam As Long, x As Long
Dim by As Byte, by2 As Long
ff = FreeFile
Open file For Binary As #ff
ff2 = FreeFile
Open file2 For Binary As #ff2
tam = FileLen(file)
For x = 1 To tam
    If (x Mod 2) = 1 Then
        Get #ff, , by
        by = 255 - by
        Put #ff2, , by
    Else
        Get #ff, , by
        by2 = by
        by2 = by2 - 121
        If by2 < 0 Then by2 = by2 + 256
        by = by2
        Put #ff2, , by
    End If
Next
Close #ff
Close #ff2
DecodeFile = True
Out:
End Function


'---------------------------------------------------------------------------
'---------------------------- Frustum Culling ------------------------------
'---------------------------------------------------------------------------

Public Sub ComputeClipPlanes()
    Dim vecsf(7) As D3DVECTOR, mat As D3DMATRIX, x As Integer

    D3DXMatrixMultiply mat, viewMatrix, projMatrix
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


'----------------------------------------------------------------------
'---------------------- FIXED OBJECTS FUNCTIONS -----------------------
'----------------------------------------------------------------------

Public Sub DrawFixedObjects()
Dim x As Integer, y As Long, RMat As D3DMATRIX, TMat As D3DMATRIX
For x = 1 To UBound(FixedObjects)
    If FixedObjects(x).WorldID = TheGameSlot.WorldID Then
        If Abs(CharPos.x - FixedObjects(x).Position.x) < DistanceFOAppear And _
           Abs(CharPos.z - FixedObjects(x).Position.z) < DistanceFOAppear Then
            If CheckSphere(FixedObjects(x)) = True Then
                D3DXMatrixTranslation TMat, FixedObjects(x).Position.x, FixedObjects(x).Position.y, FixedObjects(x).Position.z
                D3DXMatrixRotationY RMat, FixedObjects(x).RotationY
                D3DXMatrixMultiply TMatrix, RMat, TMat
                Device.SetTransform D3DTS_WORLD, TMatrix
                If Abs(CharPos.x - FixedObjects(x).Position.x) < DistanceFODetail And _
                    Abs(CharPos.z - FixedObjects(x).Position.z) < DistanceFODetail Then
                    RenderModel FOComplex(FixedObjects(x).MeshID)
                Else
                    If FixedObjects(x).unimesh = 1 Then
                        RenderModel FOComplex(FixedObjects(x).MeshID)
                    Else
                        RenderModel FOSimple(FixedObjects(x).MeshID)
                    End If
                End If
            End If
        End If
    End If
Next
End Sub

'Public Sub DrawFixedObjectsStage()
'Dim x As Integer, y As Long, RMat As D3DMATRIX, TMat As D3DMATRIX
'For x = 1 To UBound(FixedObjectsStage)
'    If FixedObjectsStage(x).WorldID = TheGameSlot.WorldID Then
'        If Abs(CharPos.x - FixedObjectsStage(x).Position.x) < DistanceFOAppear And _
'           Abs(CharPos.z - FixedObjectsStage(x).Position.z) < DistanceFOAppear Then
'            If CheckSphere(FixedObjectsStage(x)) = True Then
'                D3DXMatrixTranslation TMat, FixedObjectsStage(x).Position.x, FixedObjectsStage(x).Position.y, FixedObjectsStage(x).Position.z
'                D3DXMatrixRotationY RMat, FixedObjectsStage(x).RotationY
'                D3DXMatrixMultiply TMatrix, RMat, TMat
'                Device.SetTransform D3DTS_WORLD, TMatrix
'                If Abs(CharPos.x - FixedObjectsStage(x).Position.x) < DistanceFODetail And _
'                    Abs(CharPos.z - FixedObjectsStage(x).Position.z) < DistanceFODetail Then
'                    RenderModel FOComplexStage(FixedObjectsStage(x).MeshID)
'                Else
'                    If FixedObjectsStage(x).unimesh = 1 Then
'                        RenderModel FOComplexStage(FixedObjectsStage(x).MeshID)
'                    Else
'                        RenderModel FOSimpleStage(FixedObjectsStage(x).MeshID)
'                    End If
'                End If
'            End If
'        End If
'    End If
'Next
'End Sub

Public Function CheckSphere(object As FixedObject) As Boolean
Dim i As Long

For i = 0 To 5
    If D3DXPlaneDotCoord(FrustumPlanes(i), object.transformedBoundingSphere.center) < -object.transformedBoundingSphere.radius Then
        CheckSphere = False
        Exit Function
    End If
Next
CheckSphere = True
End Function

'Public Function LoadFixedObject(Position As D3DVECTOR, ByVal meshfile As String, ByVal meshfilesimple As String, AutoY As Boolean, RotationY As Single, WorldID As Integer) As FixedObject
'Dim res As salida, x As Long, id As Long, y As Long
'LoadFixedObject.Position = Position
'
'For x = 1 To UBound(FOComplex)
'    If FOComplex(x).MeshName = UCase(GetFileName(meshfile)) And _
'       FOSimple(x).MeshName = UCase(GetFileName(meshfilesimple)) Then
'        id = x
'        GoTo SkipLoadingFO
'    End If
'Next
'
'ReDim Preserve FOComplex(UBound(FOComplex) + 1)
'ReDim Preserve FOSimple(UBound(FOSimple) + 1)
'id = UBound(FOSimple)
'
'Set FOComplex(id).Mesh = Direct3DX.LoadMeshFromX(meshfile, D3DXMESH_MANAGED, Device, FOComplex(id).tmp_adj, FOComplex(id).tmp_mat, FOComplex(id).NumMaterials)
'ReDim FOComplex(id).Materials(FOComplex(id).NumMaterials - 1)
'ReDim FOComplex(id).TexturesNames(FOComplex(id).NumMaterials - 1)
'OptimizeMesh FOComplex(id).Mesh, FOComplex(id).tmp_adj
'Set FOComplex(id).tmp_adj = Nothing
'
'For y = 0 To FOComplex(id).NumMaterials - 1
'    Direct3DX.BufferGetMaterial FOComplex(id).tmp_mat, y, FOComplex(id).Materials(y)
'    FOComplex(id).Materials(y).Ambient = FOComplex(id).Materials(y).diffuse     'aspect patch!
'    FOComplex(id).TexturesNames(y) = UCase(GetFileName(Direct3DX.BufferGetTextureName(FOComplex(id).tmp_mat, y)))
'    If FOComplex(id).TexturesNames(y) <> "" Then Call AddTexture(FOComplex(id).TexturesNames(y))
'Next
'Set FOComplex(id).tmp_mat = Nothing
'
'If FileExists(meshfilesimple) = True Then
'    Set FOSimple(id).Mesh = Direct3DX.LoadMeshFromX(meshfilesimple, D3DXMESH_MANAGED, Device, FOSimple(id).tmp_adj, FOSimple(id).tmp_mat, FOSimple(id).NumMaterials)
'    ReDim FOSimple(id).Materials(FOSimple(id).NumMaterials - 1)
'    ReDim FOSimple(id).TexturesNames(FOSimple(id).NumMaterials - 1)
'    OptimizeMesh FOSimple(id).Mesh, FOSimple(id).tmp_adj
'    Set FOSimple(id).tmp_adj = Nothing
'
'    For y = 0 To FOSimple(id).NumMaterials - 1
'        Direct3DX.BufferGetMaterial FOSimple(id).tmp_mat, y, FOSimple(id).Materials(y)
'        FOSimple(id).Materials(y).Ambient = FOSimple(id).Materials(y).diffuse     'aspect patch!
'        FOSimple(id).TexturesNames(y) = UCase(GetFileName(Direct3DX.BufferGetTextureName(FOSimple(id).tmp_mat, y)))
'        If FOSimple(id).TexturesNames(y) <> "" Then Call AddTexture(FOSimple(id).TexturesNames(y))
'    Next
'    Set FOSimple(id).tmp_mat = Nothing
'End If

'SkipLoadingFO:
'LoadFixedObject.MeshID = id
'LoadFixedObject.WorldID = WorldID
'If FileExists(meshfilesimple) Then FOSimple(id).MeshName = UCase(GetFileName(meshfilesimple))
'FOComplex(id).MeshName = UCase(GetFileName(meshfile))

'If FOSimple(id).MeshName = "" Then LoadFixedObject.unimesh = 1

'Dim Out As salida, center As D3DVECTOR, radius As Single

'If Position.y = -9999 Then
'    segintersectfast v3(Position.x, 500, Position.z), v3(0, -1, 0), CollisionFloats(WorldID).vertices(0), UBound(CollisionFloats(WorldID).vertices) / 3, 1, Out
        
'    LoadFixedObject.Position.y = Out.puntocolision.y
'End If
    
'Direct3DX.ComputeBoundingSphereFromMesh FOComplex(LoadFixedObject.MeshID).Mesh, center, radius
'D3DXMatrixTranslation TMatrix, LoadFixedObject.Position.x, LoadFixedObject.Position.y, LoadFixedObject.Position.z
'D3DXVec3TransformCoord center, center, TMatrix

'LoadFixedObject.RotationY = RotationY / 180 * Pi
'LoadFixedObject.transformedBoundingSphere.center = center
'LoadFixedObject.transformedBoundingSphere.radius = radius
'End Function

Public Function LoadFixedObject(Position As D3DVECTOR, ByVal meshfile As String, ByVal meshfilesimple As String, AutoY As Boolean, RotationY As Single, WorldID As Integer) As FixedObject
Dim res As salida, x As Long, id As Long, y As Long
LoadFixedObject.Position = Position

For x = 1 To UBound(FOComplex)
    If FOComplex(x).MeshName = UCase(GetFileName(meshfile)) And _
       FOSimple(x).MeshName = UCase(GetFileName(meshfilesimple)) Then
        id = x
        GoTo SkipLoadingFO
    End If
Next

ReDim Preserve FOComplex(UBound(FOComplex) + 1)
ReDim Preserve FOSimple(UBound(FOSimple) + 1)
id = UBound(FOSimple)

Set FOComplex(id).Mesh = Direct3DX.LoadMeshFromX(meshfile, D3DXMESH_MANAGED, Device, FOComplex(id).tmp_adj, FOComplex(id).tmp_mat, FOComplex(id).NumMaterials)
ReDim FOComplex(id).Materials(FOComplex(id).NumMaterials - 1)
ReDim FOComplex(id).TexturesNames(FOComplex(id).NumMaterials - 1)
OptimizeMesh FOComplex(id).Mesh, FOComplex(id).tmp_adj
Set FOComplex(id).tmp_adj = Nothing

For y = 0 To FOComplex(id).NumMaterials - 1
    Direct3DX.BufferGetMaterial FOComplex(id).tmp_mat, y, FOComplex(id).Materials(y)
    FOComplex(id).Materials(y).Ambient = FOComplex(id).Materials(y).diffuse     'aspect patch!
    FOComplex(id).TexturesNames(y) = UCase(GetFileName(Direct3DX.BufferGetTextureName(FOComplex(id).tmp_mat, y)))
    If FOComplex(id).TexturesNames(y) <> "" Then Call AddTexture(FOComplex(id).TexturesNames(y))
Next
Set FOComplex(id).tmp_mat = Nothing

If FileExists(meshfilesimple) = True Then
    Set FOSimple(id).Mesh = Direct3DX.LoadMeshFromX(meshfilesimple, D3DXMESH_MANAGED, Device, FOSimple(id).tmp_adj, FOSimple(id).tmp_mat, FOSimple(id).NumMaterials)
    ReDim FOSimple(id).Materials(FOSimple(id).NumMaterials - 1)
    ReDim FOSimple(id).TexturesNames(FOSimple(id).NumMaterials - 1)
    OptimizeMesh FOSimple(id).Mesh, FOSimple(id).tmp_adj
    Set FOSimple(id).tmp_adj = Nothing
    
    For y = 0 To FOSimple(id).NumMaterials - 1
        Direct3DX.BufferGetMaterial FOSimple(id).tmp_mat, y, FOSimple(id).Materials(y)
        FOSimple(id).Materials(y).Ambient = FOSimple(id).Materials(y).diffuse     'aspect patch!
        FOSimple(id).TexturesNames(y) = UCase(GetFileName(Direct3DX.BufferGetTextureName(FOSimple(id).tmp_mat, y)))
        If FOSimple(id).TexturesNames(y) <> "" Then Call AddTexture(FOSimple(id).TexturesNames(y))
    Next
    Set FOSimple(id).tmp_mat = Nothing
End If

SkipLoadingFO:
LoadFixedObject.MeshID = id
LoadFixedObject.WorldID = WorldID
If FileExists(meshfilesimple) Then FOSimple(id).MeshName = UCase(GetFileName(meshfilesimple))
FOComplex(id).MeshName = UCase(GetFileName(meshfile))

If FOSimple(id).MeshName = "" Then LoadFixedObject.unimesh = 1

Dim Out As salida, center As D3DVECTOR, radius As Single

If Position.y = -9999 Then
    segintersectfast v3(Position.x, 500, Position.z), v3(0, -1, 0), CollisionFloats(WorldID).vertices(0), UBound(CollisionFloats(WorldID).vertices) / 3, 1, Out
        
    LoadFixedObject.Position.y = Out.puntocolision.y
End If
    
Direct3DX.ComputeBoundingSphereFromMesh FOComplex(LoadFixedObject.MeshID).Mesh, center, radius
D3DXMatrixTranslation TMatrix, LoadFixedObject.Position.x, LoadFixedObject.Position.y, LoadFixedObject.Position.z
D3DXVec3TransformCoord center, center, TMatrix

LoadFixedObject.RotationY = RotationY / 180 * Pi
LoadFixedObject.transformedBoundingSphere.center = center
LoadFixedObject.transformedBoundingSphere.radius = radius
End Function

'----------------------------------------------------------------------
'--------------------- MESH EXTENDED FUNCTIONS ------------------------
'----------------------------------------------------------------------

Public Sub ExtractTriVec(Mesh As D3DXMesh, tris() As D3DVECTOR)
On Local Error Resume Next
Dim hresult As Long
Dim vertices() As D3DVERTEX
Dim Desc As D3DINDEXBUFFER_DESC
Dim IBuf As Direct3DIndexBuffer8
Dim indices() As Integer, x As Long, y As Long, ignore As Long, indices32() As Long

ReDim vertices(Mesh.GetNumVertices)
ReDim tris(Mesh.GetNumFaces * 3 - 1)

hresult = D3DXMeshVertexBuffer8GetData(Mesh, 0, Len(vertices(0)) * Mesh.GetNumVertices, 0, vertices(0))
    
Set IBuf = Mesh.GetIndexBuffer()
IBuf.Lock 0, 0, ignore, 16
IBuf.GetDesc Desc
IBuf.Unlock
If (Mesh.GetOptions And D3DXMESH_32BIT) Then
    '32 bit index mesh
    ReDim indices32(Desc.size / 4)   '4 as we use 32 bit mesh , 2 if we use 16 bit
    D3DXMeshIndexBuffer8GetData Mesh, 0, Desc.size, 0, indices32(0)
    
    For y = 0 To Mesh.GetNumFaces * 3 - 1 Step 3
        tris(y).x = vertices(indices32(y)).x
        tris(y).y = vertices(indices32(y)).y
        tris(y).z = vertices(indices32(y)).z
        tris(y + 1).x = vertices(indices32(y + 1)).x
        tris(y + 1).y = vertices(indices32(y + 1)).y
        tris(y + 1).z = vertices(indices32(y + 1)).z
        tris(y + 2).x = vertices(indices32(y + 2)).x
        tris(y + 2).y = vertices(indices32(y + 2)).y
        tris(y + 2).z = vertices(indices32(y + 2)).z
    Next
Else
    '16 bit index mesh
    ReDim indices(Desc.size / 2)
    D3DXMeshIndexBuffer8GetData Mesh, 0, Desc.size, 0, indices(0)
    
    For y = 0 To Mesh.GetNumFaces * 3 - 1 Step 3
        tris(y).x = vertices(indices(y)).x
        tris(y).y = vertices(indices(y)).y
        tris(y).z = vertices(indices(y)).z
        tris(y + 1).x = vertices(indices(y + 1)).x
        tris(y + 1).y = vertices(indices(y + 1)).y
        tris(y + 1).z = vertices(indices(y + 1)).z
        tris(y + 2).x = vertices(indices(y + 2)).x
        tris(y + 2).y = vertices(indices(y + 2)).y
        tris(y + 2).z = vertices(indices(y + 2)).z
    Next
End If

ReDim indices(0)
ReDim indices32(0)
ReDim vertices(0)           'destroy all temp objects!!
Set IBuf = Nothing
End Sub

Public Sub ComputeNormals(verts() As D3DVECTOR, normals() As D3DVECTOR)
ReDim normals((UBound(verts) + 1) / 3)
lib_compute_normals verts(0), UBound(verts) + 1, normals(0)
End Sub

Public Sub AddToArray(array1() As D3DVECTOR, outarray() As D3DVECTOR)
Dim x As Long, start As Long
If UBound(outarray) = 0 Then
    'the array where add is empty
    ReDim Preserve outarray(UBound(array1))
    
    For x = 0 To UBound(array1)
        outarray(x) = array1(x)
    Next
Else
    start = UBound(outarray) + 1
    ReDim Preserve outarray(UBound(outarray) + UBound(array1) + 1)
    
    For x = 0 To UBound(array1)
        outarray(start + x) = array1(x)
    Next
End If
End Sub

Public Function DetermineMaxX(array1() As D3DVECTOR) As Single
Dim Max As Long, x As Long
For x = 0 To UBound(array1)
    If array1(x).x > Max Then Max = array1(x).x
Next
DetermineMaxX = Max
End Function

Public Function DetermineMaxZ(array1() As D3DVECTOR) As Single
Dim Max As Long, x As Long
For x = 0 To UBound(array1)
    If array1(x).z > Max Then Max = array1(x).z
Next
DetermineMaxZ = Max
End Function

Public Function DetermineMinX(array1() As D3DVECTOR) As Single
Dim Min As Long, x As Long
For x = 0 To UBound(array1)
    If array1(x).x < Min Then Min = array1(x).x
Next
DetermineMinX = Min
End Function

Public Function DetermineMinZ(array1() As D3DVECTOR) As Single
Dim Min As Long, x As Long
For x = 0 To UBound(array1)
    If array1(x).z < Min Then Min = array1(x).z
Next
DetermineMinZ = Min
End Function

Public Sub TransformMesh(Mesh As D3DXMesh, mat As D3DMATRIX)
'-- extract all the vertices to form triangles ------
Dim hresult As Long
Dim vertices() As D3DVERTEX
Dim avector As D3DVECTOR, y As Long

ReDim vertices(Mesh.GetNumVertices)
hresult = D3DXMeshVertexBuffer8GetData(Mesh, 0, Len(vertices(0)) * Mesh.GetNumVertices, 0, vertices(0))
    
For y = 0 To Mesh.GetNumVertices - 1
    avector.x = vertices(y).x
    avector.y = vertices(y).y
    avector.z = vertices(y).z
    D3DXVec3TransformCoord avector, avector, mat
    vertices(y).x = avector.x
    vertices(y).y = avector.y
    vertices(y).z = avector.z
Next
    
D3DXMeshVertexBuffer8SetData Mesh, 0, Len(vertices(0)) * Mesh.GetNumVertices, 0, vertices(0)
End Sub

Public Sub TransformVerts(vertices() As D3DVECTOR, mat As D3DMATRIX)
'-- extract all the vertices to form triangles ------
Dim y As Long
For y = 0 To UBound(vertices)
    D3DXVec3TransformCoord vertices(y), vertices(y), mat
Next
End Sub

Public Sub OptimizeMesh(Mesh As D3DXMesh, adjbuffer As D3DXBuffer, Optional ByVal Simple As Boolean = False)
    Dim s As Long
    Dim adjBuf1() As Long
    Dim adjBuf2() As Long
    Dim facemap() As Long
    Dim vertexMap As D3DXBuffer
    
    s = adjbuffer.GetBufferSize
    ReDim adjBuf1(s / 4)
    ReDim adjBuf2(s / 4)
    
    s = Mesh.GetNumFaces
    ReDim facemap(s)
    
    Direct3DX.BufferGetData adjbuffer, 0, 4, s * 3, adjBuf1(0)
    If Simple Then
        Mesh.OptimizeInplace D3DXMESHOPT_ATTRSORT, adjBuf1(0), adjBuf2(0), facemap(0), vertexMap
    Else
        Mesh.OptimizeInplace D3DXMESHOPT_COMPACT Or D3DXMESHOPT_ATTRSORT Or D3DXMESHOPT_VERTEXCACHE, adjBuf1(0), adjBuf2(0), facemap(0), vertexMap
    End If
    
    ReDim adjBuf1(0)
    ReDim adjBuf2(0)
End Sub

Public Sub LoadModel3D(ByRef Model As Model3D, ByVal file As String)
Dim x As Long
Set Model.Mesh = Direct3DX.LoadMeshFromX(file, D3DXMESH_MANAGED, Device, Model.tmp_adj, Model.tmp_mat, Model.NumMaterials)
OptimizeMesh Model.Mesh, Model.tmp_adj

ReDim Model.Materials(Model.NumMaterials - 1)
ReDim Model.TexturesNames(Model.NumMaterials - 1)

For x = 0 To Model.NumMaterials - 1
    Direct3DX.BufferGetMaterial Model.tmp_mat, x, Model.Materials(x)
    Model.Materials(x).Ambient = Model.Materials(x).diffuse
    Model.TexturesNames(x) = UCase(GetFileName(Direct3DX.BufferGetTextureName(Model.tmp_mat, x)))
    If Model.TexturesNames(x) <> "" Then Call AddTexture(Model.TexturesNames(x))
Next
Set Model.tmp_mat = Nothing
Set Model.tmp_adj = Nothing
End Sub

'-----------------------------------------------------------------
'------------------------ TEXTURE MANAGING -----------------------
'-----------------------------------------------------------------

Public Sub AddTexture(ByVal file As String)
Dim x As Integer
For x = 1 To UBound(PendingTextures)
    If PendingTextures(x) = file Then Exit Sub
Next

ReDim Preserve PendingTextures(UBound(PendingTextures) + 1)
PendingTextures(UBound(PendingTextures)) = file
End Sub

Public Sub LoadTexture(ByVal file As String)
'--- adds a new texture into the texture array ----
Dim x As Integer
For x = 1 To UBound(GameTextures)
    If UCase(GetFileName(GameTextures(x).filename)) = UCase(GetFileName(file)) Then Exit Sub
Next
ReDim Preserve GameTextures(UBound(GameTextures) + 1)

Set GameTextures(UBound(GameTextures)).texture = LoadTextureAndReturn(file, False, 4)
GameTextures(UBound(GameTextures)).filename = UCase(GetFileName(file))
End Sub

Public Sub LoadTextureSec(ByVal file As String)
'--- adds a new texture into the secondary texture array ----
Dim x As Integer
For x = 1 To UBound(GameTexturesSec)
    If UCase(GetFileName(GameTexturesSec(x).filename)) = UCase(GetFileName(file)) Then Exit Sub
Next
ReDim Preserve GameTexturesSec(UBound(GameTexturesSec) + 1)

Set GameTexturesSec(UBound(GameTexturesSec)).texture = LoadTextureAndReturn(file)
GameTexturesSec(UBound(GameTexturesSec)).filename = UCase(GetFileName(file))
End Sub

Public Function LoadTextureAndReturn(ByVal file As String, Optional ByVal NC As Boolean = False, Optional ByVal MipMaps As Long = 0) As Direct3DTexture8
'--- loads a texture file and returnns a Direct3DTexture8 class ---
Dim TransparentImage As Boolean
Dim imagew As Long, imageh As Long

GetImageProperties file, imagew, imageh, TransparentImage

If caps.MaxTextureWidth < imagew Then imagew = caps.MaxTextureWidth
If caps.MaxTextureHeight < imageh Then imageh = caps.MaxTextureHeight
If caps.TextureCaps And D3DPTEXTURECAPS_SQUAREONLY Then
    If imagew > imageh Then
        imagew = imageh
    Else
        imageh = imagew
    End If
End If

CheckTextureDimensions imagew, imageh

If TransparentImage = False Then
    If Direct3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, D3DM.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_DXT1) = D3D_OK And NC = False Then
        Set LoadTextureAndReturn = Direct3DX.CreateTextureFromFileEx(Device, file, imagew, imageh, MipMaps, 0, D3DFMT_DXT1, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    Else
        Set LoadTextureAndReturn = Direct3DX.CreateTextureFromFileEx(Device, file, imagew, imageh, MipMaps, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    End If
Else
    If Direct3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, D3DM.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_DXT5) = D3D_OK And NC = False Then
        Set LoadTextureAndReturn = Direct3DX.CreateTextureFromFileEx(Device, file, imagew, imageh, MipMaps, 0, D3DFMT_DXT5, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    Else
        Set LoadTextureAndReturn = Direct3DX.CreateTextureFromFileEx(Device, file, imagew, imageh, MipMaps, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    End If
End If
End Function

Public Sub LoadMainTextures(Optional ByVal Pi As Integer, Optional ByVal PE As Integer)
On Local Error GoTo TexError
Dim mypath As String, x As Integer, FList() As String
ReDim GameTextures(0)
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "tex_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
MkDir mypath
If ExtractFile(EXE & "maps\main.dat", mypath) = True Then
    Call GetFileList(mypath, FList)
    For x = 1 To UBound(FList)
        Call LoadTexture(mypath & FList(x))
        If Pi <> 0 Then ShowLoadPercent Trim(Str(Int(((PE - Pi) * x / UBound(FList) + Pi))))
    Next
Else
    MsgBox "El programa no està correctament instal·lat. Falten un o més arxius o no són correctes", vbCritical, "Error intern"
End If

Set FireTex = LoadTextureAndReturn(mypath & "fire.dds", True)

DeleteDir mypath
ReDim PendingTextures(0)
Exit Sub

TexError:
DeleteDir mypath
On Local Error Resume Next
MouseDevice.Unacquire
Set Device = Nothing
frmGraphics.Hide
Unload frmGraphics
DoEvents

MsgBox "Error greu carregant el joc!", vbCritical, "Error intern"
End
End Sub

Public Sub LoadPendingTexturesSec(Optional ByVal Pi As Integer, Optional ByVal PE As Integer)
On Local Error GoTo TexError
Dim mypath As String, x As Integer
ReDim GameTexturesSec(0)
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "tex_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
MkDir mypath
If ExtractFile(EXE & "maps\sec.dat", mypath) = True Then
    For x = 1 To UBound(PendingTextures)
        Call LoadTextureSec(mypath & PendingTextures(x))
        If Pi <> 0 Then ShowLoadPercentStage Trim(Str(Int(((PE - Pi) * x / UBound(PendingTextures) + Pi))))
    Next
End If

DeleteDir mypath
ReDim PendingTextures(0)
Exit Sub

TexError:
DeleteDir mypath
On Local Error Resume Next
MouseDevice.Unacquire
Set Device = Nothing
frmGraphics.Hide
Unload frmGraphics
DoEvents

MsgBox "Error greu carregant el joc!", vbCritical, "Error intern"
End
End Sub

Public Sub DeviceSetTexture(ByVal file As String, Optional ByVal Stage As Integer = 0)
On Local Error Resume Next
Dim x As Integer
For x = 1 To UBound(GameTextures)
    If GameTextures(x).filename = file Then
        Device.SetTexture Stage, GameTextures(x).texture
        Exit Sub
    End If
Next
For x = 1 To UBound(GameTexturesSec)
    If GameTexturesSec(x).filename = file Then
        Device.SetTexture Stage, GameTexturesSec(x).texture
        Exit Sub
    End If
Next
For x = 1 To UBound(GameTexturesTer)
    If GameTexturesTer(x).filename = file Then
        Device.SetTexture Stage, GameTexturesTer(x).texture
        Exit Sub
    End If
Next
End Sub

'------------------------------------------------------------------------------------
'---------------------------------- MISSION TARGETS ---------------------------------
'------------------------------------------------------------------------------------

Public Sub AddMissionTarget(ByVal mid As String, pos As D3DVECTOR, ByVal height As Single, ByVal radius As Single, ByVal WID As Long)
Dim x As Long
x = UBound(MissionTargets) + 1
ReDim Preserve MissionTargets(x)
MissionTargets(x).Position = pos
MissionTargets(x).height = height
MissionTargets(x).radius = radius
MissionTargets(x).WorldID = WID
MissionTargets(x).id = mid

If pos.y = -9999 Then
    Dim Out As salida
    segintersectfast v3(pos.x, 500, pos.z), v3(0, -1, 0), CollisionFloats(WID).vertices(0), UBound(CollisionFloats(WID).vertices) / 3, 1, Out
    MissionTargets(x).Position.y = Out.puntocolision.y
End If
End Sub

'------------------------------------------------------------------------------------
'----------------------------------------- OTHERS -----------------------------------
'------------------------------------------------------------------------------------

Public Function CheckPointInCube(vert1 As D3DVECTOR, vert2 As D3DVECTOR, point As D3DVECTOR) As Boolean
If vert1.x < vert2.x Then
    If Not (point.x > vert1.x And point.x < vert2.x) Then Exit Function
Else
    If Not (point.x < vert1.x And point.x > vert2.x) Then Exit Function
End If
If vert1.z < vert2.z Then
    If Not (point.z > vert1.z And point.z < vert2.z) Then Exit Function
Else
    If Not (point.z < vert1.z And point.z > vert2.z) Then Exit Function
End If
If vert1.y < vert2.y Then
    If Not (point.y > vert1.y And point.y < vert2.y) Then Exit Function
Else
    If Not (point.y < vert1.y And point.y > vert2.y) Then Exit Function
End If
CheckPointInCube = True
End Function

Public Function TempPath() As String
Dim var As String
var = Space(512)
GetTempPath 512, var
TempPath = left(var, InStr(1, var, Chr(0)) - 1)
If right(TempPath, 1) <> "\" Then TempPath = TempPath & "\"
TempPath = TempPath & "TarTemp"
On Local Error Resume Next
MkDir TempPath
On Local Error GoTo 0

TempPath = TempPath & "\"
End Function

Public Sub DeleteTempDir()
Dim var As String
var = Space(512)
GetTempPath 512, var
var = left(var, InStr(1, var, Chr(0)) - 1)
If right(var, 1) <> "\" Then var = var & "\"
var = var & "TarTemp"

DeleteDir var
End Sub

Public Function AssignMV(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal tu As Single, ByVal tv As Single) As myVertex
AssignMV.x = x
AssignMV.y = y
AssignMV.z = z
AssignMV.tu = tu
AssignMV.tv = tv
AssignMV.rhw = 1
AssignMV.Color = D3DColorARGB(0, 255, 255, 255)
End Function

Public Function AssignMVAdv(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal tu As Single, ByVal tv As Single, ByVal Alpha As Single) As myVertex
AssignMVAdv.x = x
AssignMVAdv.y = y
AssignMVAdv.z = z
AssignMVAdv.tu = tu
AssignMVAdv.tv = tv
AssignMVAdv.rhw = 1
AssignMVAdv.Color = D3DColorARGB(ByVal Alpha, 255, 255, 255)
End Function

Public Function AssignMVS(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal tu As Single, ByVal tv As Single) As myVertexSimple
AssignMVS.x = x
AssignMVS.y = y
AssignMVS.z = z
AssignMVS.tu = tu
AssignMVS.tv = tv
AssignMVS.Color = D3DColorARGB(0, 255, 255, 255)
End Function

Public Function AssignMVA(ByVal x As Single, ByVal y As Single, ByVal z As Single, ByVal Color As Long) As myVertexAlpha
AssignMVA.x = x
AssignMVA.y = y
AssignMVA.z = z
AssignMVA.Color = Color
AssignMVA.rhw = 1
End Function

Public Sub MouseSetUp()
MouseDevice.Unacquire
Set MouseDevice = Nothing
Set MouseDevice = DirectInput.CreateDevice("guid_SysMouse")

MouseDevice.SetCommonDataFormat DIFORMAT_MOUSE
MouseDevice.SetCooperativeLevel frmGraphics.hWnd, DISCL_EXCLUSIVE Or DISCL_FOREGROUND

Dim diProp As DIPROPLONG
diProp.lHow = DIPH_DEVICE
diProp.lObj = 0
diProp.lData = 50      '50 Buffer for the mouse!!!

Call MouseDevice.SetProperty("DIPROP_BUFFERSIZE", diProp)
        
DI_hevent = DirectX.CreateEvent(frmGraphics)
MouseDevice.SetEventNotification DI_hevent

MouseDevice.Acquire
End Sub

Public Function GetAllResolutions(ByVal MinWidth As Single, ByVal MinHeight As Single, Out As Variant)
Dim x As Integer, y As Integer
ReDim Out(0)
For x = 0 To UBound(DModesArray) - 1
    If DModesArray(x).width >= MinWidth And DModesArray(x).height >= MinHeight Then
        For y = 0 To UBound(Out)
            If (Out(y) = DModesArray(x).width & " x " & DModesArray(x).height) Then GoTo MyOut
        Next
        ReDim Preserve Out(UBound(Out) + 1)
        Out(UBound(Out)) = DModesArray(x).width & " x " & DModesArray(x).height
    End If
MyOut:
Next
End Function

Public Sub DrawFont(ByVal Text As String, ByVal x As Single, ByVal y As Single, Optional ByVal totalwidth As Single = 0, Optional ByVal height As Single = 0, Optional ByVal FontWidth As Single = 0)
If Text = "" Then Exit Sub
Dim f As Integer, char As Byte, g As Integer
Dim xpos As Integer, ypos As Integer
Dim width As Single
Dim tu As Single, tu2 As Single, tv As Single, tv2 As Single

Device.SetTexture 0, FontTexture

For f = 1 To Len(Text)
    ypos = 0: xpos = 0
    char = Asc(mid(Text, f, 1))
    If (char >= 65 And char <= 90) Or (char >= 97 And char <= 122) Then    'letras
        If char < 90 Then char = char - 65
        If char > 90 Then char = char - 97
        Do While char >= 8
            ypos = ypos + 1
            char = char - 8
        Loop
        xpos = char
    ElseIf char >= 48 And char <= 57 Then
        char = char - 48
        ypos = 4
        If char > 7 Then
            ypos = ypos + 1
            char = char - 8
        End If
        xpos = char
    ElseIf char >= 44 And char <= 47 Then
        xpos = 2 + (char - 44)
        ypos = 3
    ElseIf char = 92 Then
        xpos = 6: ypos = 3
    ElseIf char = 59 Then
        xpos = 7: ypos = 3
    ElseIf char = 58 Then
        xpos = 2: ypos = 6
    ElseIf char = 63 Then
        xpos = 3: ypos = 6
    ElseIf char = 33 Then
        xpos = 2: ypos = 5
    ElseIf char = 34 Then
        xpos = 3: ypos = 5
    ElseIf char = 39 Then
        xpos = 4: ypos = 5
    ElseIf char = 186 Then            'º --->  flecha abajo
        xpos = 1: ypos = 6
    ElseIf char = 170 Then             'ª --->  flecha arriba
        xpos = 0: ypos = 6
    End If
    
    If char <> 32 Then
        If totalwidth = 0 Then
            width = FontWidth * Len(Text)
            If width = 0 Then width = Len(Text) * 32
        Else
            width = totalwidth
        End If
        If height = 0 Then height = 32
        
        tu = (xpos * 32) / 256
        tu2 = ((xpos + 1) * 32) / 256
        tv = (ypos * 32) / 256
        tv2 = ((ypos + 1) * 32) / 256

        FontVertices(0) = AssignMV(x + (f - 1) * width / Len(Text), y, 0, tu, tv)
        FontVertices(1) = AssignMV(x + f * width / Len(Text), y, 0, tu2, tv)
        FontVertices(2) = AssignMV(x + (f - 1) * width / Len(Text), y + height, 0, tu, tv2)
        FontVertices(3) = AssignMV(x + f * width / Len(Text), y + height, 0, tu2, tv2)
        
        Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, FontVertices(0), Len(FontVertices(0))
    End If
Next
End Sub

Public Function ParseResolution(ByVal var As String) As D3DDISPLAYMODE
Dim x As Integer
x = InStr(1, var, "x")
ParseResolution.width = Val(left(var, x - 2))
ParseResolution.height = Val(mid(var, x + 2))
End Function

Public Function Is32bit(DMode As D3DDISPLAYMODE) As Boolean
If DMode.Format = D3DFMT_X8R8G8B8 Then
    Is32bit = True
Else
    Is32bit = False
End If
End Function

Public Sub SaveRS()
RenderStates(0) = Device.GetRenderState(D3DRS_ALPHABLENDENABLE)
RenderStates(1) = Device.GetRenderState(D3DRS_ZENABLE)
RenderStates(2) = Device.GetRenderState(D3DRS_ZWRITEENABLE)
RenderStates(3) = Device.GetRenderState(D3DRS_LIGHTING)
RenderStates(4) = Device.GetRenderState(D3DRS_DESTBLEND)
RenderStates(5) = Device.GetRenderState(D3DRS_SRCBLEND)
RenderStates(6) = Device.GetRenderState(D3DRS_DIFFUSEMATERIALSOURCE)
RenderStates(7) = Device.GetTextureStageState(0, D3DTSS_MAGFILTER)
RenderStates(8) = Device.GetTextureStageState(0, D3DTSS_MINFILTER)
End Sub

Public Sub RestoreRS()
Device.SetRenderState D3DRS_ALPHABLENDENABLE, RenderStates(0)
Device.SetRenderState D3DRS_ZENABLE, RenderStates(1)
Device.SetRenderState D3DRS_ZWRITEENABLE, RenderStates(2)
Device.SetRenderState D3DRS_LIGHTING, RenderStates(3)
Device.SetRenderState D3DRS_DESTBLEND, RenderStates(4)
Device.SetRenderState D3DRS_SRCBLEND, RenderStates(5)
Device.SetRenderState D3DRS_DIFFUSEMATERIALSOURCE, RenderStates(6)
Device.SetTextureStageState 0, D3DTSS_MAGFILTER, RenderStates(7)
Device.SetTextureStageState 0, D3DTSS_MINFILTER, RenderStates(8)
End Sub

Public Sub RefreshDrawDepth()
Select Case DrawDepth
Case 1
    DistanceFOAppear = 30
    DistanceFODetail = 8
    FarViewPlane = 120
Case 2
    DistanceFOAppear = 50
    DistanceFODetail = 12
    FarViewPlane = 190
Case 3
    DistanceFOAppear = 80
    DistanceFODetail = 16
    FarViewPlane = 250
End Select
'DistanceFOAppear   '30 low,  50 medium,  80 high
'DistanceFODetail   '8 low, 12 medium, 16 high
'FarViewPlane       '120 low, 190 medium, 250 high
End Sub

Public Sub LoadSettings()
Dim var As String
var = RegRead("drawdepth")
Select Case Val(var)
Case 1, 2, 3
    DrawDepth = Val(var)
Case Else
    DrawDepth = 2
End Select

var = RegRead("cursorspeed")
Select Case Val(var)
Case 1, 2, 3, 4, 5
    CursorSpeed = Val(var)
Case Else
    CursorSpeed = 3
End Select

var = RegRead("chardetail")
Select Case Val(var)
Case 1, 2, 3
    CharDetail = Val(var)
Case Else
    CharDetail = 2
End Select

var = RegRead("texq")
Select Case Val(var)
Case 1, 2
    TexQuality = Val(var)
Case Else
    TexQuality = 2
End Select
End Sub

Public Sub SaveSettings()
RegSave "drawdepth", Trim(Str(DrawDepth))
RegSave "cursorspeed", Trim(Str(CursorSpeed))
RegSave "chardetail", Trim(Str(CharDetail))
RegSave "texq", Trim(Str(TexQuality))
End Sub

Public Sub CheckTextureDimensions(width As Long, height As Long)
If TexQuality = 1 Then
    If width > 256 Then width = 256
    If height > 256 Then height = 256
End If
End Sub

'---------------------------------------------------------------------------------
'-------------------------------- FILE FUNCTIONS ---------------------------------
'---------------------------------------------------------------------------------

Public Sub DeleteDir(ByVal directory As String, Optional ByVal timer As Single)
If timer = 0 Then timer = GetTickCount()
'---- Removes all the content of an existing folder (including subfolders) ----
On Local Error Resume Next
If directory = "" Then Exit Sub
Dim cad As String
If right(directory, 1) <> "\" Then directory = directory & "\"
Kill directory & "*.*"

Again:
cad = Dir(directory, vbDirectory)
If GetTickCount() - timer > 10000 Then Exit Sub
Do While cad <> ""
    If cad <> "." And cad <> ".." Then
        If GetTickCount() - timer > 10000 Then Exit Sub
        DeleteDir directory & cad, timer
        GoTo Again
    End If
    cad = Dir$
Loop
RmDir directory
End Sub

Public Sub GetImageProperties(ByVal file As String, width As Long, height As Long, transparent As Boolean)
Dim dib As Long
Select Case LCase(right(file, 3))
Case "jpg", "jpeg"
    dib = FreeImage_Load(FIF_JPEG, file)
    transparent = False
Case "bmp"
    dib = FreeImage_Load(FIF_BMP, file)
    transparent = False
Case "tga"
    dib = FreeImage_Load(FIF_TARGA, file)
    transparent = True
Case "dds"
    dib = FreeImage_Load(FIF_DDS, file)
    transparent = True
End Select
width = FreeImage_GetWidth(dib)
height = FreeImage_GetHeight(dib)
Call FreeImage_Unload(dib)
End Sub

Public Sub LoadAndParseFile(ByVal file As String)
ReDim PNames(0)
ReDim PValues(0)
Dim f As Integer, cad As String
f = FreeFile()
Open file For Input As #f
Do While Not EOF(f)
    Line Input #f, cad
    If left(cad, 1) <> "#" And cad <> "" And InStr(1, cad, "=") <> 0 Then
        ReDim Preserve PNames(UBound(PNames) + 1)
        ReDim Preserve PValues(UBound(PValues) + 1)
        PNames(UBound(PNames)) = left(cad, InStr(1, cad, "=") - 1)
        PValues(UBound(PValues)) = mid(cad, InStr(1, cad, "=") + 1)
    End If
Loop
Close #f
End Sub

Public Function PFGetProperty(ByVal PropertyName As String) As String
Dim x As Integer
For x = 1 To UBound(PNames)
    If UCase(PNames(x)) = UCase(PropertyName) Then
        PFGetProperty = PValues(x)
        Exit Function
    End If
Next
End Function

Public Function GetFileName(ByVal file As String) As String
If file = "" Then Exit Function
Dim xx As Long, x As Long
For x = Len(file) To 1 Step -1
    xx = InStr(x, file, "\")
    If xx <> 0 Then Exit For
Next
GetFileName = mid$(file, xx + 1, Len(file) - xx + 1)
End Function

Public Function FileExists(ByVal filename As String) As Boolean
'------ Returns if the file exists -------
On Local Error Resume Next
Dim temp As VbFileAttribute, fff As Integer
temp = GetAttr(filename)
fff = FreeFile()
Open filename For Input As #fff
Close #fff
If Err.number <> 0 Then
    FileExists = False
Else
    FileExists = True
End If
Err.Clear
Err.number = 0
End Function

Public Function GetModelCoreFromName(ByVal ModelName As String) As Cal3DModel
Dim x As Long
For x = 1 To UBound(Cal3DModelArray)
    If Cal3DModelArray(x).ModelName = ModelName Then
        Set GetModelCoreFromName = Cal3DModelArray(x)
        Exit For
    End If
Next
End Function


'----------------------------------------------------
'------------------ C++ Functions -------------------
'----------------------------------------------------

Public Function ProcessCameraCoords(Position As D3DVECTOR, ByVal AngleH As Single, ByVal AngleV As Single, ByVal distance As Single) As D3DVECTOR
Dim CameraDistanceProj As Single, CameraRelX As Single, CameraRelZ As Single
CameraDistanceProj = distance * Cos(AngleV * Pi / 180)
CameraRelX = CameraDistanceProj * sin(AngleH * Pi / 180)
CameraRelZ = CameraDistanceProj * Cos(AngleH * Pi / 180)
ProcessCameraCoords.x = CameraRelX + Position.x
ProcessCameraCoords.z = CameraRelZ + Position.z
ProcessCameraCoords.y = (distance * sin(AngleV * Pi / 180)) + Position.y
End Function

Public Function Process1stCameraCoords(Position As D3DVECTOR, ByVal AngleH As Single, ByVal AngleV As Single) As D3DVECTOR
Dim CameraDistanceProj As Single, CameraRelX As Single, CameraRelZ As Single
CameraRelX = sin(AngleH * Pi / 180 + Pi)
CameraRelZ = Cos(AngleH * Pi / 180 + Pi)
Process1stCameraCoords.x = CameraRelX + Position.x
Process1stCameraCoords.z = CameraRelZ + Position.z
    Process1stCameraCoords.y = Position.y + sin(-3 * ((AngleV - MinVerticalAngle) - (MaxVerticalAngle - MinVerticalAngle) / 2) * Pi / 180) + 1.6         '(distance * sin(AngleV * Pi / 180)) + Position.y
End Function


'--------------------------------------------------------
'-------------------- DS  Functions ---------------------
'--------------------------------------------------------

Public Function GetGUIDFromDesc(ByVal Desc As String) As String
Dim x As Long
For x = 1 To DSEnum.GetCount
    If Desc = DSEnum.GetDescription(x) Then
        GetGUIDFromDesc = DSEnum.GetGuid(x)
        Exit For
    End If
Next
End Function

Public Function GetDescFromGUID(ByVal guid As String) As String
Dim x As Long
For x = 1 To DSEnum.GetCount
    If guid = DSEnum.GetGuid(x) Then
        GetDescFromGUID = DSEnum.GetDescription(x)
        Exit For
    End If
Next
End Function


Public Function ArcCos(ByVal a As Double) As Double
On Error Resume Next

If a = 1 Or a = -1 Then
    ArcCos = 0
    Exit Function
End If

ArcCos = Atn(-a / Sqr(-a * a + 1)) + 2 * Atn(1)
On Error GoTo 0
End Function

Public Sub GetFileList(ByVal path As String, list() As String)
Dim file As String
ReDim list(0)
file = Dir(path, vbArchive)
Do While file <> ""
    ReDim Preserve list(UBound(list) + 1)
    list(UBound(list)) = file
    file = Dir()
Loop
End Sub

Public Sub CreateVertexSahders()
If VSCapable = False Then Exit Sub
Dim decl(5) As Long
decl(0) = D3DVSD_STREAM(0)
decl(1) = D3DVSD_REG(D3DVSDE_POSITION, D3DVSDT_FLOAT3)
decl(2) = D3DVSD_STREAM(1)
decl(3) = D3DVSD_REG(D3DVSDE_TEXCOORD0, D3DVSDT_FLOAT2)
decl(4) = D3DVSD_END()
Call Device.CreateVertexShader(decl(0), ByVal 0, CharShader, 0)
End Sub

Public Sub ReleaseVertexShaders()
If VSCapable = False Then Exit Sub
Device.DeleteVertexShader CharShader
End Sub

'------------------------------------------------------
'------------------- SAVE FUNCTIONS -------------------
'------------------------------------------------------

Public Sub SaveGame(ByVal Slot As Integer, SavePos As D3DVECTOR, SaveAngle As Single)
'--- TAR1                                   -->4bytes
'--- date HHnnSSddMMyyyy formated double    -->8bytes
'--- Mission Level (Long)                   -->4bytes
'--- Position (d3dvector)                   -->12bytes
'--- RotationH (single)                     -->4bytes
'--- WorldID (long)                         -->4bytes
'--- Health (Long)                          -->4bytes
'--- Coins (Long)                           -->4bytes

Dim SaveCamera As D3DVECTOR, FFile As Integer, MyDouble As Double
SaveCamera = ProcessCameraCoords(SavePos, SaveAngle, (MaxVerticalAngle + MinVerticalAngle) / 2, (MaxCameraDistance + MinCameraDistance) / 2)

With SavedGameSlot
    .Position = SavePos
    .WorldID = TheGameSlot.WorldID
    .RotationH = SaveAngle
    .MissionLevel = TheGameSlot.MissionLevel
End With

On Local Error Resume Next
MkDir EXE & "save"
Kill EXE & "save\game" & Trim(Str(Slot)) & ".sav"
On Local Error GoTo 0

FFile = FreeFile()
Open EXE & "save\game" & Trim(Str(Slot)) & ".sav" For Binary As #FFile
    Put #FFile, , "TAR1"
    MyDouble = Val(Format(Now(), "hh") & Format(Now(), "nn") & Format(Now(), "ss") & Format(Now(), "dd") & Format(Now(), "mm") & Format(Now(), "yyyy"))
    Put #FFile, , MyDouble
    Put #FFile, , SavedGameSlot.MissionLevel
    Put #FFile, , SavePos.x
    Put #FFile, , SavePos.y
    Put #FFile, , SavePos.z
    Put #FFile, , SavedGameSlot.RotationH
    Put #FFile, , SavedGameSlot.WorldID
    Put #FFile, , SavedGameSlot.Health
    Put #FFile, , SavedGameSlot.Coins
Close #FFile
End Sub

Public Function GameExists(ByVal Slot As Integer) As Boolean
GameExists = FileExists(EXE & "save\game" & Trim(Str(Slot)) & ".sav")
End Function

Public Function GetGameDate(ByVal Slot As Integer) As String
If GameExists(Slot) Then
    Dim FFile As Integer, buffer As String, MyDouble As Double
    FFile = FreeFile()
    Open EXE & "save\game" & Trim(Str(Slot)) & ".sav" For Binary As #FFile
        buffer = Space(4)
        Get #FFile, , buffer
        Get #FFile, , MyDouble
        buffer = Format(MyDouble, "00000000000000")
        buffer = mid(buffer, 7, 2) & "/" & mid(buffer, 9, 2) & "/" & mid(buffer, 13, 2) & " " & mid(buffer, 1, 2) & ":" & mid(buffer, 3, 2)
        GetGameDate = buffer
    Close #FFile
End If
End Function

Public Sub LoadSavedGame(ByVal Slot As Integer)

End Sub

Public Function ToLong(ByVal num As Single) As Long
Dim MyLong As Long
CopyMemory MyLong, num, 4
ToLong = MyLong
End Function
