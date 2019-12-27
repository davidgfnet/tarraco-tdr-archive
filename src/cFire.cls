VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFire"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public cfgPitch As Single

Public World As Long
Public radius As Single
Public NumberOfParticles As Long
Public FlameHeightInc As Single
Public FadeInc As Single
Public FlameCompression As Single

Private Type fireVert
    x As Single: y As Single: z As Single
    tu As Single: tv As Single
End Type
Private Const fireShader As Long = D3DFVF_XYZ Or D3DFVF_TEX1

Dim compression As Single
Dim vecData() As D3DVECTOR4, vecStream() As fireVert
Dim VertexBuffer As Direct3DVertexBuffer8
Dim q As Single
Dim pPosition As D3DVECTOR

Private Sub Class_Terminate()
Set VertexBuffer = Nothing
End Sub

Public Sub Initialize()
Dim x As Long
ReDim vecData(NumberOfParticles - 1)
ReDim vecStream((NumberOfParticles - 1) * 6 * 2)

compression = radius / (NumberOfParticles - 1)

Set VertexBuffer = Nothing
Set VertexBuffer = Device.CreateVertexBuffer(Len(vecStream(0)) * (UBound(vecStream()) + 1), 0, fireShader, D3DPOOL_MANAGED)

For x = 0 To NumberOfParticles - 1 Step 1
    vecData(x).x = (radius / 2) - Rnd * radius
    vecData(x).y = 0
    vecData(x).z = Rnd * radius
Next

Call process
End Sub

Public Sub process()
Dim Ca As Single, Sa As Single, x As Long, y As Single, comp As Single
'comp = Sqr(Sqr(Sqr(FlameCompression / StableFrameAverage * 1000)))

Ca = Cos(cfgPitch)
Sa = sin(cfgPitch)
  
For x = 0 To NumberOfParticles - 1 Step 1
    With vecData(x)
      If .z < 0.05 Then     'min size
        .y = 0
        .x = (radius / 2) - Rnd * radius
        .z = radius
        .w = 1 - Rnd * 2
      End If
      
      .x = .x '* comp
      .y = .y + FlameHeightInc * FrameAverage / 1000
      .z = .z - FadeInc * FrameAverage / 1000
      If .w < 0 Then
        .w = .w - Rnd * 0.05 - 0.05
      Else
        .w = .w + Rnd * 0.05 + 0.05
      End If
      q = .w
      
      y = x * compression - (radius / 2)
      
      fxTransform Ca, Sa, fxMakeVertex(.x - .z / 2, .y - .z / 2, y, 0, 0), vecStream(x * 6 + 0), .x, .y, y
      fxTransform Ca, Sa, fxMakeVertex(.x + .z / 2, .y - .z / 2, y, 1, 0), vecStream(x * 6 + 1), .x, .y, y
      fxTransform Ca, Sa, fxMakeVertex(.x - .z / 2, .y + .z / 2, y, 0, 1), vecStream(x * 6 + 2), .x, .y, y
      fxTransform Ca, Sa, fxMakeVertex(.x + .z / 2, .y - .z / 2, y, 1, 0), vecStream(x * 6 + 3), .x, .y, y
      fxTransform Ca, Sa, fxMakeVertex(.x + .z / 2, .y + .z / 2, y, 1, 1), vecStream(x * 6 + 4), .x, .y, y
      fxTransform Ca, Sa, fxMakeVertex(.x - .z / 2, .y + .z / 2, y, 0, 1), vecStream(x * 6 + 5), .x, .y, y
    End With
  Next x
  'reload vertex buffer
  D3DVertexBuffer8SetData VertexBuffer, 0, Len(vecStream(0)) * (UBound(vecStream()) + 1), 0, vecStream(0)

End Sub

Private Function fxTransform(Ca As Single, Sa As Single, src As fireVert, dst As fireVert, cx As Single, cy As Single, cz As Single)
  'Temp Variables
  Dim Yy As Single
  'transformation
  dst = src
  With dst
    'subtract center
    .x = .x - cx
    .y = .y - cy
    .z = .z - cz
    'rotate by Z axis
    Yy = .y * Ca - .z * Sa
    .z = .y * Sa + .z * Ca
    .y = Yy
    'add center
    .x = .x + cx
    .y = .y + cy
    .z = .z + cz
    'texture rotation (more realistic sometimes)
    .tu = (.tu - 0.5) * 2
    .tv = (.tv - 0.5) * 2
    If .tv = 1 And .tu = 1 Then
      .tu = sin(q + 45 * 3.14 / 180)
      .tv = Cos(q + 45 * 3.14 / 180)
    Else
      If .tv = -1 And .tu = -1 Then
        .tu = sin(q + 225 * 3.14 / 180)
        .tv = Cos(q + 225 * 3.14 / 180)
      Else
        If .tv = 1 And .tu = -1 Then
          .tu = sin(q + 135 * 3.14 / 180)
          .tv = Cos(q + 135 * 3.14 / 180)
        Else
          If .tv = -1 And .tu = 1 Then
            .tu = sin(q + 315 * 3.14 / 180)
            .tv = Cos(q + 315 * 3.14 / 180)
          End If
        End If
      End If
    End If
    .tu = .tu / 2 + 0.5
    .tv = .tv / 2 + 0.5
    .tu = .tu * 1.1
    .tv = .tv * 1.1
  End With
 End Function

Private Function fxMakeVertex(x As Single, y As Single, z As Single, U As Single, v As Single) As fireVert
  With fxMakeVertex
    .x = x
    .y = y
    .z = z
    .tu = U
    .tv = v
  End With
End Function

Public Property Get Position() As D3DVECTOR
Position = pPosition
End Property

Public Property Let Position(vNewValue As D3DVECTOR)
pPosition = vNewValue
End Property

Public Sub Render()
If World <> TheGameSlot.WorldID Then Exit Sub
If CheckSphere() = False Then Exit Sub
Dim WorldMat As D3DMATRIX
D3DXMatrixInverse WorldMat, 0, viewMatrix
WorldMat.m41 = pPosition.x: WorldMat.m42 = pPosition.y: WorldMat.m43 = pPosition.z
Device.SetTransform D3DTS_WORLD, WorldMat
Device.SetTexture 0, FireTex

Me.process

Device.SetVertexShader fireShader
Device.SetStreamSource 0, VertexBuffer, Len(vecStream(0))
Device.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
Device.DrawPrimitive D3DPT_TRIANGLELIST, 0, (UBound(vecStream()) + 1) / 3
Device.SetRenderState D3DRS_CULLMODE, D3DCULL_CCW

Device.SetTexture 0, Nothing
End Sub

Public Sub InitializeFast(ByVal radius As Single, ByVal compression As Single, ByVal height As Single, ByVal fade As Single, ByVal nump As Long, pos As D3DVECTOR, ByVal WorldID As Long)
Me.radius = radius
Me.FlameCompression = compression
Me.FadeInc = fade
Me.NumberOfParticles = nump
Me.FlameHeightInc = height
pPosition = pos
World = WorldID

Me.Initialize
End Sub

Private Function CheckSphere() As Boolean
Dim i As Long, rad As Single, height As Single
height = FlameHeightInc * radius / FadeInc
If radius > height Then
    rad = radius
Else
    rad = height
End If

For i = 0 To 5
    If D3DXPlaneDotCoord(FrustumPlanes(i), v3(pPosition.x, pPosition.y + (rad / 2), pPosition.z)) < -rad Then
        CheckSphere = False
        Exit Function
    End If
Next
CheckSphere = True
End Function

