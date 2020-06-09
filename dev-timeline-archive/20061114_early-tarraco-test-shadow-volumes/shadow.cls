VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "shadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Edges() As Long
Private NumberEdges As Long

'Private Type ShadowVertex
'    p As D3DVECTOR4
'    color As Long
'End Type

Private Vertices() As D3DVECTOR
Private Vertices2() As D3DVECTOR

Private NumberOfVertices

Public Sub Render()
Device.SetVertexShader D3DFVF_XYZ
Device.DrawPrimitiveUP D3DPT_TRIANGLELIST, NumberOfVertices / 3, Vertices(0), Len(Vertices(0))
End Sub


Private Function resta(v1 As D3DVECTOR, v2 As D3DVECTOR) As D3DVECTOR
resta.X = v1.X - v2.X
resta.Y = v1.Y - v2.Y
resta.z = v1.z - v2.z
End Function

Private Sub AddEdge(ByVal v1 As Long, ByVal v2 As Long)
Dim i As Long
For i = 0 To NumberEdges - 1
    If (Edges(2 * i) = v1 And Edges(2 * i + 1) = v2) _
    Or (Edges(2 * i) = v2 And Edges(2 * i + 1) = v1) Then
        If NumberEdges > 1 Then
            Edges(2 * i) = Edges(2 * (NumberEdges - 1))
            Edges(2 * i + 1) = Edges(2 * (NumberEdges - 1) + 1)
        End If
        
        NumberEdges = NumberEdges - 1
        Exit Sub
    End If
Next

Edges(2 * NumberEdges) = v1
Edges(2 * NumberEdges + 1) = v2
NumberEdges = NumberEdges + 1
End Sub



Public Sub Build(Mesh As D3DXMesh, light As D3DVECTOR)
Dim triangulos() As D3DVECTOR
NumberEdges = 0
Erase Vertices

ReDim Vertices(Mesh.GetNumFaces * 15)
ReDim Vertices2(Mesh.GetNumFaces * 3)

Dim hresult As Long
Dim vertexs() As D3DVERTEX

ReDim vertexs(Mesh.GetNumVertices)

hresult = D3DXMeshVertexBuffer8GetData(Mesh, 0, Len(vertexs(0)) * Mesh.GetNumVertices, 0, vertexs(0))

Dim i As Long



Dim midesc As D3DINDEXBUFFER_DESC
Dim IBuf As Direct3DIndexBuffer8
Dim tam As Long, tam2 As Long
Dim out As Long
Dim vector() As Integer, vector2() As Integer

Set IBuf = Mesh.GetIndexBuffer()
IBuf.Lock 0, 0, out, 16
IBuf.GetDesc midesc
IBuf.Unlock

tam = midesc.Size
ReDim vector(midesc.Size / 2)

D3DIndexBuffer8GetData Mesh.GetIndexBuffer(), 0, midesc.Size, 0, vector(0)

ReDim triangulos(Mesh.GetNumFaces * 3 - 1)
Dim Normals() As D3DVECTOR
ReDim Normals(Mesh.GetNumFaces - 1)

Dim v0 As D3DVECTOR, v1 As D3DVECTOR, v2 As D3DVECTOR
Dim vc1 As D3DVECTOR, vc2 As D3DVECTOR, vNormal As D3DVECTOR
ReDim Edges(Mesh.GetNumFaces() * 6)
Dim j As Long

'For i = 0 To Mesh.GetNumVertices - 1
For i = 0 To Mesh.GetNumFaces * 3 - 1 Step 3   '//Cycle through every vertex
    triangulos(i).X = vertexs(vector(i)).X
    triangulos(i).Y = vertexs(vector(i)).Y
    triangulos(i).z = vertexs(vector(i)).z
    
    
    triangulos(i + 1).X = vertexs(vector(i + 1)).X
    triangulos(i + 1).Y = vertexs(vector(i + 1)).Y
    triangulos(i + 1).z = vertexs(vector(i + 1)).z
    
    triangulos(i + 2).X = vertexs(vector(i + 2)).X
    triangulos(i + 2).Y = vertexs(vector(i + 2)).Y
    triangulos(i + 2).z = vertexs(vector(i + 2)).z
    
    'triangulos(i + 1).X = vertexs(i + 1).X
    'triangulos(i + 1).Y = vertexs(i + 1).Y
    'triangulos(i + 1).z = vertexs(i + 1).z
    'triangulos(i + 2).X = vertexs(i + 2).X
    'triangulos(i + 2).Y = vertexs(i + 2).Y
    'triangulos(i + 2).z = vertexs(i + 2).z
    'Normals(j) = Normal(triangulos(i), triangulos(i + 1), triangulos(i + 2))
    'j = j + 1
Next

ComputeNormals triangulos, Normals

buildshadow2 triangulos(0), Mesh.GetNumFaces, light, Vertices(0), NumberOfVertices, Normals(0), NumberOfVertices / 3
NumberOfVertices = UBound(Vertices)

Exit Sub

For i = 0 To Mesh.GetNumFaces * 3 - 1 Step 3   '//Cycle through every vertex
    'triangulos(i).X = vertexs(vector(i)).X
    'triangulos(i).Y = vertexs(vector(i)).Y
    'triangulos(i).z = vertexs(vector(i)).z
    'triangulos(i + 1).X = vertexs(vector(i + 1)).X
    'triangulos(i + 1).Y = vertexs(vector(i + 1)).Y
    'triangulos(i + 1).z = vertexs(vector(i + 1)).z
    'triangulos(i + 2).X = vertexs(vector(i + 2)).X
    'triangulos(i + 2).Y = vertexs(vector(i + 2)).Y
    'triangulos(i + 2).z = vertexs(vector(i + 2)).z
    
    v0 = v3(vertexs(vector(i)).X, vertexs(vector(i)).Y, vertexs(vector(i)).z)
    v1 = v3(vertexs(vector(i + 1)).X, vertexs(vector(i + 1)).Y, vertexs(vector(i + 1)).z)
    v2 = v3(vertexs(vector(i + 2)).X, vertexs(vector(i + 2)).Y, vertexs(vector(i + 2)).z)
    
    vc1 = resta(v2, v1)
    vc2 = resta(v1, v0)
    
    D3DXVec3Cross vNormal, vc1, vc2
    
    If D3DXVec3Dot(vNormal, light) >= 0 Then
        'Edges(j) = vector(i)
        'Edges(j + 1) = vector(i + 1)
        
        'Edges(j + 2) = vector(i + 1)
        'Edges(j + 3) = vector(i + 2)
        
        'Edges(j + 4) = vector(i + 2)
        'Edges(j + 5) = vector(i)
        'j = j + 6
        
        AddEdge vector(i), vector(i + 1)
        AddEdge vector(i + 1), vector(i + 2)
        AddEdge vector(i + 2), vector(i)
    End If
Next

Dim vect1 As D3DVECTOR, vect2 As D3DVECTOR, vect3 As D3DVECTOR, vect4 As D3DVECTOR
Dim LightGrande As D3DVECTOR
LightGrande.X = light.X * 100
LightGrande.Y = light.Y * 100
LightGrande.z = light.z * 100

Dim numV As Long

For i = 0 To NumberEdges - 1
    vect1 = triangulos(Edges(2 * i))
    vect2 = triangulos(Edges(2 * i + 1))
'For i = 0 To Mesh.GetNumFaces * 6 - 1 Step 2
'    vect1 = triangulos(Edges(i))
'    vect2 = triangulos(Edges(i + 1))
    vect3 = resta(vect1, LightGrande)
    vect4 = resta(vect2, LightGrande)
    
    Vertices(numV) = vect1
    Vertices(numV + 1) = vect2
    Vertices(numV + 2) = vect3
    
    Vertices(numV + 3) = vect2
    Vertices(numV + 4) = vect4
    Vertices(numV + 5) = vect3
    
    'Debug.Print vect1.X & " / " & vect1.Y & " / " & vect1.z & " //// " & vect2.X & " / " & vect2.Y & " / " & vect2.z

    numV = numV + 6
Next

NumberOfVertices = numV
End Sub



Public Sub ComputeNormals(Verts() As D3DVECTOR, Normals() As D3DVECTOR)
Dim X As Long, cont As Long
Dim v1 As D3DVECTOR, v2 As D3DVECTOR, vmodule As Double
ReDim Normals((UBound(Verts) + 1) / 3)
For X = 0 To UBound(Verts) Step 3
    v1.X = Verts(X + 2).X - Verts(X + 1).X
    v1.Y = Verts(X + 2).Y - Verts(X + 1).Y
    v1.z = Verts(X + 2).z - Verts(X + 1).z
    
    v2.X = Verts(X + 1).X - Verts(X).X
    v2.Y = Verts(X + 1).Y - Verts(X).Y
    v2.z = Verts(X + 1).z - Verts(X).z
    
    D3DXVec3Cross v1, v1, v2
    v1 = Normalize(v1)
    
    Normals(cont) = v1
    cont = cont + 1
Next
End Sub


Public Function Normalize(v As D3DVECTOR) As D3DVECTOR
Dim module As Double
module = Sqr(v.X ^ 2 + v.Y ^ 2 + v.z ^ 2)
Normalize.X = v.X / module
Normalize.Y = v.Y / module
Normalize.z = v.z / module
End Function
