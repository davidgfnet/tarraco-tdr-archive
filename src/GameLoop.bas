Attribute VB_Name = "GameLoop"
Option Explicit

'------- Game Loop Module -----------
'--- All the hard rendering here! ---

Public Sub GameLoop()
'----- debug info ------
Dim myfont As D3DXFont, sfont As StdFont, hfont As IFont, re As RECT
Set sfont = New StdFont
sfont.Name = "arial"
sfont.size = 12
Set hfont = sfont
Set myfont = Direct3DX.CreateFont(Device, hfont.hfont)
re.top = 0
re.bottom = 300
re.left = 0
re.right = 300
'------ debug info -----

StartGameLoop:

'----------- CLEAN UP VARS ----------
Call VarCleanUp

MusicEngine.StopMusic

'-------- SET UP PROJ MATRIX --------
D3DXMatrixPerspectiveFovLH projMatrix, 45 * Pi / 180, 1, 0.1, FarViewPlane
Device.SetTransform D3DTS_PROJECTION, projMatrix

MainRender = True


Do While MainRender
    TimeCounter = GetTickCount()
    
    Call UpdateListenerSettings
    
    Call MusicEngine.RenderTime
    
    Call UpdateScene            'check the argument flow
    If RestartLoop Then GoTo StartGameLoop       'check if we have lost this stage
    
    Call ComputeDoors           'be sure of the change of world
        
    '-------------------- INPUT --------------------
    ' - Mouse
        If MouseX > FrameAverage * 2.5 Then MouseX = FrameAverage * 2.5
        If MouseX < FrameAverage * -2.5 Then MouseX = FrameAverage * -2.5
        CharAngleH = CharAngleH + MouseX * CursorSpeed / 15
        If CharAngleH > 360 Then CharAngleH = CharAngleH - 360
        If CharAngleH < 0 Then CharAngleH = CharAngleH + 360
        
        CharAngleV = CharAngleV + MouseY * CursorSpeed / 25
        If CharAngleV > MaxVerticalAngle Then CharAngleV = MaxVerticalAngle
        If CharAngleV < MinVerticalAngle Then CharAngleV = MinVerticalAngle
        
        If KeySubtract Then MouseZ = MouseZ + FrameAverage
        If KeyAdd Then MouseZ = MouseZ - FrameAverage
        CharDistance = CharDistance + MouseZ / 2000 * CursorSpeed
        If CharDistance > MaxCameraDistance Then CharDistance = MaxCameraDistance
        If CharDistance < MinCameraDistance Then CharDistance = MinCameraDistance
        MouseX = 0: MouseY = 0: MouseZ = 0
    ' - Keyboard        --> configured in frmGraphics / ProcessCharMoves Sub to increase speed

    Call ProcessCharMoves

    Call PhysicsModule
    
    Call ProcessDynamicObjects
    
    '------------------- MATRIX VIEW ---------------
    If CameraType = 0 Then
        CameraLookAt = v3(CharPos.x, CharPos.y + 1.1, CharPos.z)
        realCameraPos = v3(CameraPos.x, CameraPos.y - 0.1, CameraPos.z)
    Else
        realCameraPos = v3(CharPos.x, CharPos.y + 1.6, CharPos.z)
        CameraPos = realCameraPos
        CameraPos.y = CameraPos.y + 0.1
        CameraLookAt = Process1stCameraCoords(CharPos, CharAngleH, CharAngleV)
        If CharSpeed <> 0 Then
            realCameraPos.y = realCameraPos.y + sin(GetTickCount() / 100) / 20
            CameraLookAt.y = CameraLookAt.y + sin(GetTickCount() / 100) / 20
            CameraPos.y = CameraPos.y + sin(GetTickCount() / 100) / 20
        End If
    End If
    
    D3DXMatrixLookAtLH viewMatrix, realCameraPos, CameraLookAt, vectorUP
    Device.SetTransform D3DTS_VIEW, viewMatrix
    Call ComputeClipPlanes
    
    '---------------- SOUND CONTROL --------------
    Call SoundCtrl
    
    Call BgMusic
    
    '------------------- RENDER! -------------------
    Call Device.Clear(0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0)
    Device.BeginScene
    
    Device.SetRenderState D3DRS_ALPHABLENDENABLE, 0
    Call DrawSkyBox
    
    Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1
    
    Call DrawFixedObjects
    'Call DrawFixedObjectsStage    'objects for this stage
    If CameraType = 0 Then Call DrawMainChar
    Call RenderDynamicObjects

    Call DrawMainMesh
    
    Call DrawTransparentObjects
    
    Call DrawUI
    
    Call MakeFade       'controls the fading options, must be the last drawing
    Call SceneFading    'controls the level / XML fading
    
    ' debug info
    'Direct3DX.DrawText myfont, &HFFFFFFFF, CurrentCharAnimation & vbCrLf & CharSpeed, re, DT_TOP Or DT_LEFT
    'Direct3DX.DrawText myfont, &HFFFFFFFF, "{" & CharPos.x & " / " & CharPos.y & " / " & CharPos.z & "}" & vbCrLf & Trim(Str(1000 / (FrameAverage + 0.1))), re, DT_TOP Or DT_LEFT
    
    Device.EndScene
    On Local Error Resume Next
    Call Device.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    DoEvents
    '------------------------------------------------
    
    TimeCounter = GetTickCount() - TimeCounter
    If FrameAverage = 0 Then FrameAverage = TimeCounter
    FrameAverage = (FrameAverage * 4 + TimeCounter) / 5
    StableFrameAverage = (StableFrameAverage * 199 + TimeCounter) / 200
    
    If AuxMenu = True Then Call AuxMenuSystem
    
    Call CheckSavePoints        'maybe we are in a save point?
    
    If VideoOn <> "" Then
        VideoEngine.OpenVideo VideoOn
        VideoEngine.PlayVideo
        VideoOn = ""
        VideoEngine.CloseVideo
    End If
Loop

MusicEngine.StopMusic
End Sub

Public Sub GameLoadLevel()
Call SetLevelAtributes(TheGameSlot.MissionLevel)
Call SoundEngine.DestroySounds: Call SoundEngine.DestroyStandardSound
Call SoundReady
Call LoadCommonSounds
Call LoadAllSounds
End Sub

Public Sub LoadAllSounds()
Dim mypath As String, x As Integer
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "sound_tmp_" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0

ExtractFile EXE & "sound\other.dat", mypath

For x = 1 To UBound(PendingSounds)
    SoundEngine.CreateSound mypath & PendingSounds(x), LCase(GetFileName(PendingSounds(x)))
Next

DeleteDir mypath
End Sub

'Public Function MeshOfWorld(ByVal MeshID As Integer, ByVal WorldID As Integer) As Boolean
'---- MeshID ----
    '-- 01 - Circ + Ciutat   |
    '-- 02 - Arcs            |   WorldID1                --> Cull this two
    '-- 03 - Circ + Pulvinar |                           -->
    
    '-- 04 - Forum           |   WorldID 2
    
    '-- 05 - Pretori         |   WorldID 3
    
'Select Case MeshID
'Case 1, 2, 3
'    If WorldID = 1 Then MeshOfWorld = True
'Case 4
'    If WorldID = 2 Then MeshOfWorld = True
'Case 5
'    If WorldID = 3 Then MeshOfWorld = True
'Case 6
'    If WorldID = 4 Then MeshOfWorld = True
'End Select
'End Function

Public Function GetMiniMapTex(ByVal WorldID As Integer) As Long
Select Case WorldID
Case 1, 2
    GetMiniMapTex = 1
Case 3
    GetMiniMapTex = 2
Case Else
    GetMiniMapTex = 0
End Select
End Function

Public Sub DrawSkyBox()
If WorldProperties(TheGameSlot.WorldID).SkyBox = 0 Then Exit Sub         'no sky box
D3DXMatrixTranslation TMatrix, CameraPos.x, CameraPos.y - 0.22 - 0.05, CameraPos.z
Device.SetTransform D3DTS_WORLD, TMatrix

Device.SetRenderState D3DRS_ZWRITEENABLE, 0
Device.SetRenderState D3DRS_LIGHTING, 0

Device.SetVertexShader myVertexFVFSimple
Device.SetTexture 0, SkyBoxTextures((WorldProperties(TheGameSlot.WorldID).SkyBox - 1) * 5)
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyBoxVertices(0), Len(SkyBoxVertices(0))
Device.SetTexture 0, SkyBoxTextures((WorldProperties(TheGameSlot.WorldID).SkyBox - 1) * 5 + 1)
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyBoxVertices(4), Len(SkyBoxVertices(0))
Device.SetTexture 0, SkyBoxTextures((WorldProperties(TheGameSlot.WorldID).SkyBox - 1) * 5 + 2)
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyBoxVertices(8), Len(SkyBoxVertices(0))
Device.SetTexture 0, SkyBoxTextures((WorldProperties(TheGameSlot.WorldID).SkyBox - 1) * 5 + 3)
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyBoxVertices(12), Len(SkyBoxVertices(0))
Device.SetTexture 0, SkyBoxTextures((WorldProperties(TheGameSlot.WorldID).SkyBox - 1) * 5 + 4)
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, SkyBoxVertices(16), Len(SkyBoxVertices(0))

Device.SetRenderState D3DRS_ZWRITEENABLE, 1
Device.SetRenderState D3DRS_LIGHTING, 1
End Sub

Public Sub DrawMainMesh()
Dim x As Long
D3DXMatrixIdentity TMatrix
Device.SetTransform D3DTS_WORLD, TMatrix
Select Case TheGameSlot.WorldID
Case 1
    RenderModelsLM(1).Draw
    RenderModelsLM(2).Draw
    RenderModelsLM(3).Draw
    RenderModelsLM(4).Draw
    RenderModelsLM(5).Draw

    If CameraPos.z < 0 Then
        RenderModelsLM(7).Draw
        RenderModelsLM(8).Draw
    Else
        RenderModelsLM(6).Draw
    End If
Case 2
    RenderModelsLM(11).Draw
    RenderModelsLM(12).Draw
Case 3
    RenderModelsLM(21).Draw
    RenderModelsLM(22).Draw
Case Else
    For x = 1 To RenderModelsLMAuxNum
        Call RenderModelsLMAux(x).Draw
    Next
End Select
End Sub

Public Sub RenderModel(Model As Model3D)
Dim x As Integer
For x = 0 To Model.NumMaterials - 1
    If Model.TexturesNames(x) = "" Then
        Device.SetTexture 0, Nothing
    Else
        DeviceSetTexture Model.TexturesNames(x)
    End If
    Device.SetMaterial Model.Materials(x)
    Model.Mesh.DrawSubset x
Next
End Sub

Public Sub DrawMainChar()
Dim mat As D3DMATRIX, Mat2 As D3DMATRIX
Dim RotationX As D3DMATRIX
D3DXMatrixRotationX RotationX, -Pi / 2
D3DXMatrixRotationY mat, CharAngleH / 180 * Pi + Pi
D3DXMatrixMultiply mat, RotationX, mat

D3DXMatrixTranslation Mat2, CharPos.x, CharPos.y, CharPos.z
D3DXMatrixMultiply mat, mat, Mat2
Device.SetTransform D3DTS_WORLD, mat

MainChar.Update 0, FrameAverage / 1000
MainChar.SetLevelOfDetail 1, 0
Device.SetRenderState D3DRS_AMBIENT, RGB(WorldProperties(TheGameSlot.WorldID).UnlitAmbientValues.x, WorldProperties(TheGameSlot.WorldID).UnlitAmbientValues.y, WorldProperties(TheGameSlot.WorldID).UnlitAmbientValues.z)

Device.SetRenderState D3DRS_LIGHTING, 1
Device.SetRenderState D3DRS_AMBIENT, RGB(255, 0, 0)
Device.SetRenderState D3DRS_AMBIENTMATERIALSOURCE, D3DMCS_MATERIAL

MainChar.Render 0
Device.SetRenderState D3DRS_AMBIENT, RGB(WorldProperties(TheGameSlot.WorldID).AmbientValues.x, WorldProperties(TheGameSlot.WorldID).AmbientValues.y, WorldProperties(TheGameSlot.WorldID).AmbientValues.z)
End Sub


Public Sub PhysicsModule()
Dim vertex1 As D3DVECTOR, Fallen As Boolean
Static VSpeed As Single
vertex1.x = CharPos.x
vertex1.y = CharPos.y + 1
vertex1.z = CharPos.z

Dim res As salida, temp As Double, NextY As Double

hprocess CollisionFloats(TheGameSlot.WorldID).vertices(0), vertex1, (UBound(CollisionFloats(TheGameSlot.WorldID).vertices) / 3), res

If AccelerationTimer <> 0 Then
    temp = ((GetTickCount() - AccelerationTimer) / 1000)
    NextY = InitY + (VSpeed - 7) * temp + 0.5 * -WorldAcceleration * temp ^ 2
    
    Fallen = True
    
    If res.respuesta = 1 And (NextY <= CharPos.y) Then
    If res.puntocolision.y >= NextY Then
        temp = (res.puntocolision.y - CharPos.y) * 25 * FrameAverage / 1000
        If temp < 0.01 Then temp = (res.puntocolision.y - CharPos.y)
        CharPos.y = CharPos.y + temp
        Fallen = False
    Else
        If NextY > res.puntocolision.y Then
            If (NextY - res.puntocolision.y) < 0.1 And VSpeed = 0 Then
                CharPos.y = res.puntocolision.y
                Fallen = False
            Else
                CharPos.y = NextY
            End If
        Else
            CharPos.y = res.puntocolision.y
            Fallen = False
        End If
    End If
    Else
        CharPos.y = NextY
    End If
End If

If Fallen = False Then
    VSpeed = 0
    InitY = CharPos.y
    AccelerationTimer = GetTickCount()
    If Jumping = True Then VSpeed = 11
Else
    Jumping = False
End If

CharPosBefore = CharPos

If MouseB0 = True Then
    process CameraPos, CollisionFloats(TheGameSlot.WorldID).vertices(0), (UBound(CollisionFloats(TheGameSlot.WorldID).vertices) / 3), CharPos, CharPosBefore, vectorUP, CharDistance, CharAngleV, CharAngleH, 0.5, 5, 0.7, 1.2, FrameAverage, CharSpeed * -1 * 10, 0.22, CollisionFloatsAux(0), CameraType + 1
Else
    process CameraPos, CollisionFloats(TheGameSlot.WorldID).vertices(0), (UBound(CollisionFloats(TheGameSlot.WorldID).vertices) / 3), CharPos, CharPosBefore, vectorUP, CharDistance, CharAngleV, CharAngleH, 0.5, 5, 0.7, 1.2, FrameAverage, CharSpeed * -1 * 1, 0.22, CollisionFloatsAux(0), CameraType + 1
End If

Dim x As Long, vec As D3DVECTOR
For x = 1 To UBound(MovingCharacters)
    If Not (MovingCharacters(x) Is Nothing) Then
        If MovingCharacters(x).Visible And MovingCharacters(x).WorldID = TheGameSlot.WorldID Then
            'check if paused too
            If VecDistFast(CharPos, MovingCharacters(x).Public_Position) < (MovingCharacters(x).dimensions + MainCharDimensions) ^ 2 Then
                'move the main char to the correct position
                vec.x = CharPos.x - MovingCharacters(x).Public_Position.x
                vec.z = CharPos.z - MovingCharacters(x).Public_Position.z
                vec = Normalize(vec)
                vec.x = vec.x * (MovingCharacters(x).dimensions + MainCharDimensions)
                vec.z = vec.z * (MovingCharacters(x).dimensions + MainCharDimensions)
                CharPos.x = MovingCharacters(x).Public_Position.x + vec.x
                CharPos.z = MovingCharacters(x).Public_Position.z + vec.z
            End If
        End If
    End If
Next
End Sub

Public Sub ProcessCharMoves()
If ResetAllMoves Then
    CurrentCharAnimation = ""
    CharAnimationTimer = 0
    Movement = 0
    ResetAllMoves = False
    CharSpeed = 0
    MainChar.ClearCycle 0, "walking", 0
    MainChar.ClearCycle 0, "running", 0
    MainChar.BlendCycle 0, "quiet", 1, 0
End If
If CurrentCharAnimation = "" Then CurrentCharAnimation = "quiet"

If GetAsyncKeyState(vbKeyShift) <> 0 Then
    Run = True
Else
    Run = False
End If

If Movement = 0 Then
    If GetTickCount() > CharAnimationTimer And CharAnimationTimer <> 0 Then
        CharAnimationTimer = 0
        CharSpeed = 0
        MainChar.ClearCycle 0, "walking", 0
        MainChar.ClearCycle 0, "running", 0
        MainChar.BlendCycle 0, "quiet", 1, 0
        CurrentCharAnimation = "quiet"
    End If
    If CurrentCharAnimation <> "quiet" And CharAnimationTimer = 0 Then
        MainChar.ClearCycle 0, "walking", 0.4
        MainChar.ClearCycle 0, "running", 0.4
        MainChar.BlendCycle 0, "quiet", 1, 0.5
        CharAnimationTimer = GetTickCount() + 400
    End If
Else
    If CurrentCharAnimation = "quiet" Or CharAnimationTimer <> 0 Then
        If Run Then
            CurrentCharAnimation = "running"
            MainChar.ClearCycle 0, "quiet", 0.3
            MainChar.BlendCycle 0, "running", 1, 0
            CurrentCharAnimation = "running"
            CharSpeed = 1.7
        Else
            CurrentCharAnimation = "walking"
            MainChar.ClearCycle 0, "quiet", 0.3
            MainChar.BlendCycle 0, "walking", 1, 0
            CurrentCharAnimation = "walking"
            CharSpeed = 1
        End If
    Else
        'we are running.... check run <-> walk change
        If CurrentCharAnimation = "walking" And Run Then
            MainChar.ClearCycle 0, "walking", 0.3
            MainChar.BlendCycle 0, "running", 1, 0.3
            CurrentCharAnimation = "running"
            CharSpeed = 1.7
        ElseIf CurrentCharAnimation = "running" And Run = False Then
            MainChar.ClearCycle 0, "running", 0.3
            MainChar.BlendCycle 0, "walking", 1, 0.3
            CurrentCharAnimation = "walking"
            CharSpeed = 1
        End If
    End If
End If

End Sub

Public Sub MakeFade()
'makes a fade between the worlds (see computedoors sub)
Static mTimer As Double
Dim Alpha As Single

InitAgain:

If MakeFade1 Then
    If mTimer = 0 Then mTimer = GetTickCount()
    If mTimer + 750 < GetTickCount() Then
        mTimer = 0
        EnterNewWorld 0, 0
        GoTo InitAgain
    End If

    Alpha = 255 * (GetTickCount() - mTimer) / 750
    If Alpha > 255 Then Alpha = 255
    If Alpha < 0 Then Alpha = 0
    fadeVertices(0) = AssignMVA(0, 0, 0, D3DColorARGB(Alpha, 0, 0, 0))
    fadeVertices(1) = AssignMVA(D3DM.width, 0, 0, D3DColorARGB(Alpha, 0, 0, 0))
    fadeVertices(2) = AssignMVA(0, D3DM.height, 0, D3DColorARGB(Alpha, 0, 0, 0))
    fadeVertices(3) = AssignMVA(D3DM.width, D3DM.height, 0, D3DColorARGB(Alpha, 0, 0, 0))
            
    Device.SetTexture 0, Nothing
    Device.SetVertexShader myVertexAlphaFVF
    Device.SetRenderState D3DRS_ZENABLE, 0
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, fadeVertices(0), Len(fadeVertices(0))
    Device.SetRenderState D3DRS_ZENABLE, 1
ElseIf MakeFade2 Then
    If mTimer = 0 Then mTimer = GetTickCount()
    If mTimer + 750 < GetTickCount() Then
        MakeFade2 = False
        MakeFade1 = False
        mTimer = 0
        Exit Sub
    End If

    Alpha = 255 - 255 * (GetTickCount() - mTimer) / 750
    If Alpha > 255 Then Alpha = 255
    If Alpha < 0 Then Alpha = 0
    fadeVertices(0) = AssignMVA(0, 0, 0, D3DColorARGB(Alpha, 0, 0, 0))
    fadeVertices(1) = AssignMVA(D3DM.width, 0, 0, D3DColorARGB(Alpha, 0, 0, 0))
    fadeVertices(2) = AssignMVA(0, D3DM.height, 0, D3DColorARGB(Alpha, 0, 0, 0))
    fadeVertices(3) = AssignMVA(D3DM.width, D3DM.height, 0, D3DColorARGB(Alpha, 0, 0, 0))
            
    Device.SetTexture 0, Nothing
    Device.SetVertexShader myVertexAlphaFVF
    Device.SetRenderState D3DRS_ZENABLE, 0
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, fadeVertices(0), Len(fadeVertices(0))
    Device.SetRenderState D3DRS_ZENABLE, 1
Else
    mTimer = 0
End If
End Sub

Public Sub SceneFading()
Static FTimer As Double
Dim Alpha As Single

SceneFading_Init:

Select Case FadeState
Case 0
    'no fade
    FTimer = 0
    Exit Sub
Case 1
    'fade in
    If FTimer = 0 Then FTimer = GetTickCount()
    If FTimer + FadeTimeMS < GetTickCount() Then
        FadeState = 0
        FTimer = 0
        GoTo SceneFading_Init
    End If
    Alpha = 255 - 255 * (GetTickCount() - FTimer) / FadeTimeMS
Case 2
    'fade out
    If FTimer = 0 Then FTimer = GetTickCount()
    If FTimer + FadeTimeMS < GetTickCount() Then
        FadeState = 3   'black
        FTimer = 0
        GoTo SceneFading_Init
    End If
    Alpha = 255 * (GetTickCount() - FTimer) / FadeTimeMS
Case 3
    Alpha = 255         'opac
    FTimer = 0
End Select

If Alpha > 255 Then Alpha = 255
If Alpha < 0 Then Alpha = 0
fadeVertices(0) = AssignMVA(0, 0, 0, D3DColorARGB(Alpha, 0, 0, 0))
fadeVertices(1) = AssignMVA(D3DM.width, 0, 0, D3DColorARGB(Alpha, 0, 0, 0))
fadeVertices(2) = AssignMVA(0, D3DM.height, 0, D3DColorARGB(Alpha, 0, 0, 0))
fadeVertices(3) = AssignMVA(D3DM.width, D3DM.height, 0, D3DColorARGB(Alpha, 0, 0, 0))

Device.SetTexture 0, Nothing
Device.SetVertexShader myVertexAlphaFVF
Device.SetRenderState D3DRS_ZENABLE, 0
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, fadeVertices(0), Len(fadeVertices(0))
Device.SetRenderState D3DRS_ZENABLE, 1
End Sub

Public Sub SoundCtrl()
Select Case CurrentCharAnimation
Case "standing"
    If IsPlayingStSound("pasos_mainchar") Then StopStandardSound "pasos_mainchar"
Case "walking"
    
Case "walking_end"
    
Case "walking_start"
    PlayStandardSound "pasos_mainchar", True
End Select
End Sub

Public Sub DrawTransparentObjects()
Dim x As Long, y As Long
Dim SMatrix As D3DMATRIX, TMatrix As D3DMATRIX, RMatrix As D3DMATRIX, TheMatrix As D3DMATRIX
Device.SetRenderState D3DRS_ALPHABLENDENABLE, 1
Device.SetRenderState D3DRS_ZWRITEENABLE, 0
Device.SetRenderState D3DRS_LIGHTING, 0

If UBound(MissionTargets) <> 0 Then
    For x = 1 To UBound(MissionTargets)
        If MissionTargets(x).Visible And MissionTargets(x).WorldID = TheGameSlot.WorldID Then
            D3DXMatrixScaling SMatrix, MissionTargets(x).radius + Cos(GetTickCount() / 250) / 20, MissionTargets(x).height + sin(GetTickCount() / 250) / 20, MissionTargets(x).radius + Cos(GetTickCount() / 250) / 20
            D3DXMatrixRotationY RMatrix, GetTickCount() / 1000
            D3DXMatrixTranslation TMatrix, MissionTargets(x).Position.x, MissionTargets(x).Position.y, MissionTargets(x).Position.z
            D3DXMatrixMultiply TheMatrix, SMatrix, RMatrix
            D3DXMatrixMultiply TheMatrix, TheMatrix, TMatrix
            Device.SetTransform D3DTS_WORLD, TheMatrix
            For y = 0 To MissionTargetModel.NumMaterials - 1
                Device.SetMaterial MissionTargetModel.Materials(y)
                DeviceSetTexture MissionTargetModel.TexturesNames(y)
                Call MissionTargetModel.Mesh.DrawSubset(y)
            Next
        End If
    Next
End If

If UBound(Fires) <> 0 Then
    Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_DESTALPHA
    Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    For x = 1 To UBound(Fires)
        Fires(x).Render
    Next
End If

Device.SetRenderState D3DRS_ZWRITEENABLE, 1
Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
Device.SetRenderState D3DRS_LIGHTING, 1

For x = 1 To UBound(SavePoints)
    If SavePoints(x).WorldID = TheGameSlot.WorldID Then
        D3DXMatrixRotationY TMatrix, GetTickCount() / 2000  'w = 90 graus per segon
        TMatrix.m41 = SavePoints(x).Position.x: TMatrix.m42 = SavePoints(x).Position.y + 0.75: TMatrix.m43 = SavePoints(x).Position.z
        Device.SetTransform D3DTS_WORLD, TMatrix
        RenderModel SavePointModel
    End If
Next
End Sub

Public Sub ProcessDynamicObjects()
Dim x As Long, y As Long
Dim Directions() As D3DVECTOR
Dim Salidas() As salida

ReDim DynObjPositions(0)
NumDynObjPositions = -1

For x = 1 To UBound(MovingCharacters)
    If Not (MovingCharacters(x) Is Nothing) Then
        If MovingCharacters(x).Visible And MovingCharacters(x).Paused = False And MovingCharacters(x).WorldID = TheGameSlot.WorldID Then
            Call MovingCharacters(x).RenderTime
        End If
    End If
Next

If NumDynObjPositions = -1 Then GoTo skip
ReDim Directions(NumDynObjPositions)
ReDim Salidas(NumDynObjPositions)
For x = 0 To NumDynObjPositions
    Directions(x).y = -1
Next
If NumDynObjPositions >= 0 Then segintersectfast DynObjPositions(0), Directions(0), CollisionFloats(TheGameSlot.WorldID).vertices(0), UBound(CollisionFloats(TheGameSlot.WorldID).vertices) / 3, NumDynObjPositions + 1, Salidas(0)

skip:

For x = 1 To UBound(MovingCharacters)
    If Not (MovingCharacters(x) Is Nothing) Then
        If MovingCharacters(x).Visible And MovingCharacters(x).Paused = False And MovingCharacters(x).WorldID = TheGameSlot.WorldID Then
            If MovingCharacters(x).AutoY = False Then
                Call MovingCharacters(x).RenderTime2
            Else
                MovingCharacters(x).CoordY = Salidas(y).puntocolision.y
                Call MovingCharacters(x).RenderTime2
                y = y + 1
            End If
        End If
    End If
    If Not (MovingCharacters(x) Is Nothing) Then
        If MovingCharacters(x).Visible And MovingCharacters(x).WorldID = TheGameSlot.WorldID Then
            'check if paused too
            Call MovingCharacters(x).RenderSound
        End If
    End If
Next
End Sub

Public Sub RenderDynamicObjects()
Dim x As Long
If UBound(MovingCharacters) = 0 Then Exit Sub
Device.SetRenderState D3DRS_AMBIENT, RGB(WorldProperties(TheGameSlot.WorldID).UnlitAmbientValues.x, WorldProperties(TheGameSlot.WorldID).UnlitAmbientValues.y, WorldProperties(TheGameSlot.WorldID).UnlitAmbientValues.z)
For x = 1 To UBound(MovingCharacters)
    If Not (MovingCharacters(x) Is Nothing) Then
        If MovingCharacters(x).WorldID = TheGameSlot.WorldID And MovingCharacters(x).Visible = True Then
            Call MovingCharacters(x).RenderNow
        End If
    End If
Next
Device.SetRenderState D3DRS_AMBIENT, RGB(WorldProperties(TheGameSlot.WorldID).AmbientValues.x, WorldProperties(TheGameSlot.WorldID).AmbientValues.y, WorldProperties(TheGameSlot.WorldID).AmbientValues.z)
End Sub

Public Sub CheckSavePoints()
If UBound(SavePoints) = 0 Then Exit Sub
Dim GoodPosition As D3DVECTOR, GoodAngle As Single
Dim x As Long
For x = 1 To UBound(SavePoints)
    If VecDistFastXZ(SavePoints(x).Position, CharPos) < 0.75 ^ 2 Then
    If Abs(SavePoints(x).Position.y - CharPos.y) < 0.4 Then
        GoodAngle = CharAngleH - 180
        If GoodAngle < 0 Then GoodAngle = GoodAngle + 360
        GoodPosition.x = CharPos.x - SavePoints(x).Position.x
        GoodPosition.y = CharPos.y - SavePoints(x).Position.y
        GoodPosition.z = CharPos.z - SavePoints(x).Position.z
        GoodPosition = Normalize(GoodPosition)   'distance = 1
        GoodPosition.x = GoodPosition.x + SavePoints(x).Position.x
        GoodPosition.y = GoodPosition.y + SavePoints(x).Position.y
        GoodPosition.z = GoodPosition.z + SavePoints(x).Position.z
        Call SaveMenuSystem(GoodPosition, GoodAngle)
        CharPos = GoodPosition
        CharAngleH = GoodAngle
        CameraPos = ProcessCameraCoords(CharPos, CharAngleH, CharAngleV, CharDistance)
        Exit Sub
    End If
    End If
Next
End Sub

Public Sub SaveMenuSystem(SavePos As D3DVECTOR, SaveAngle As Single)
Dim coordX As Single, CoordY As Single
Dim Out As Boolean, mTimer As Double, Alpha As Single, PauseTimer As Double
Device.SetRenderState D3DRS_ALPHABLENDENABLE, True
Device.SetVertexShader myVertexFVF
Device.SetRenderState D3DRS_ZENABLE, False
Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
Device.SetRenderState D3DRS_LIGHTING, 0
PauseTimer = GetTickCount()

Call FreezeAll
If MusicEngine.MusicStatus = 1 Then Call MusicEngine.Pause
Call SaveGameMenu.RefreshDM
coordX = D3DM.width / 2: coordX = D3DM.height / 2

Call RefreshSlots

Do While Out = False
    '----------- MOUSE COORDS PROCESS ---------------
    coordX = coordX + MouseX * (CursorSpeed / 3)
    CoordY = CoordY + MouseY * (CursorSpeed / 3)
    If coordX < 0 Then coordX = 0
    If coordX > D3DM.width Then coordX = D3DM.width
    If CoordY < 0 Then CoordY = 0
    If CoordY > D3DM.height Then CoordY = D3DM.height
    MouseVerts(0) = AssignMV(coordX, CoordY, 0, 0, 0)
    MouseVerts(1) = AssignMV(coordX + CursorW * D3DM.width / 800, CoordY, 0, 1, 0)
    MouseVerts(2) = AssignMV(coordX, CoordY + CursorH * D3DM.height / 600, 0, 0, 1)
    MouseVerts(3) = AssignMV(coordX + CursorW * D3DM.width / 800, CoordY + CursorH * D3DM.height / 600, 0, 1, 1)
    MouseX = 0: MouseY = 0
    '---------------- SEND EVENTS TO MENUS ------------
    SaveGameMenu.ProcessMouseMove coordX, CoordY
    If MouseClick0 = True Then
        SaveGameMenu.ProcessClick coordX, CoordY
        MouseClick0 = False
    End If
    
    If SaveGameMenu.isEvent Then
        Select Case SaveGameMenu.EventName
        Case "returngame"
            Out = True
        Case "slot1"
            Call SlotClick(1, SavePos, SaveAngle)
        Case "slot2"
            Call SlotClick(2, SavePos, SaveAngle)
        Case "slot3"
            Call SlotClick(3, SavePos, SaveAngle)
        Case "slot4"
            Call SlotClick(4, SavePos, SaveAngle)
        Case "slot5"
            Call SlotClick(5, SavePos, SaveAngle)
        Case "slot6"
            Call SlotClick(6, SavePos, SaveAngle)
        Case "_dialog_exit_yes"
            Call SlotClick(0, SavePos, SaveAngle)
        Case "_dialog_exit_no"
            
        End Select
    End If
    SaveGameMenu.isEvent = False
    
    Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    Device.BeginScene
    
    '------------ RENDER THE MENU --------
    SaveGameMenu.RenderMenu
    
    '----------- RENDER THE CURSOR OVER ALL -----------
    Device.SetTexture 0, MouseTexture
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MouseVerts(0), Len(MouseVerts(0))
    
    Device.EndScene
    Device.Present ByVal 0, ByVal 0, 0, ByVal 0
    
    Sleep 5
    DoEvents
Loop

'only return to the game
Call SetCorrectRenderStates(RGB(WorldProperties(TheGameSlot.WorldID).AmbientValues.x, WorldProperties(TheGameSlot.WorldID).AmbientValues.y, WorldProperties(TheGameSlot.WorldID).AmbientValues.z))
Call UnFreezeAll
If MusicEngine.MusicStatus = 2 Then Call MusicEngine.ResumePlay
AccelerationTimer = AccelerationTimer + (GetTickCount() - PauseTimer)
ResetAllMoves = True: Movement = 0: CharSpeed = 0
End Sub

Private Sub RefreshSlots()
Dim f As Integer
For f = 1 To 6
    If GameExists(f) Then
        SaveGameMenu.EditLabel "label_slot_" & Trim(Str(f)), GetGameDate(f)
    Else
        SaveGameMenu.EditLabel "label_slot_" & Trim(Str(f)), "  lliure"
    End If
Next
End Sub

Private Sub SlotClick(ByVal Slot As Integer, SavePos As D3DVECTOR, SaveAngle As Single)
Static LastSlot As Integer
If Slot = 0 Then
    MenuSaveGame LastSlot, SavePos, SaveAngle
Else
    If GameExists(Slot) Then
        SaveGameMenu.ShowDialogBox "ja existeix un\narxiu guardat\nvols sobreescriure'l?", 1
        LastSlot = Slot
    Else
        MenuSaveGame Slot, SavePos, SaveAngle
    End If
End If
Call RefreshSlots
End Sub

Private Sub MenuSaveGame(ByVal Slot As Integer, SavePos As D3DVECTOR, SaveAngle As Single)
SaveGame Slot, SavePos, SaveAngle
SaveGameMenu.ShowDialogBox "partida guardada", 0
End Sub

Public Sub AuxMenuSystem()
'the menu system which appears when pressing escape
Dim coordX As Single, CoordY As Single
Dim Out As Boolean, mTimer As Double, Alpha As Single, PauseTimer As Double
Device.SetRenderState D3DRS_ALPHABLENDENABLE, True
Device.SetVertexShader myVertexFVF
Device.SetRenderState D3DRS_ZENABLE, False
Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
Device.SetRenderState D3DRS_LIGHTING, 0
PauseTimer = GetTickCount()

Call FreezeAll
If MusicEngine.MusicStatus = 1 Then Call MusicEngine.Pause
Call AuxiliarMenu.RefreshDM

coordX = D3DM.width / 2
CoordY = D3DM.height / 2

Do While Out = False
    '----------- MOUSE COORDS PROCESS ---------------
    coordX = coordX + MouseX * (CursorSpeed / 3)
    CoordY = CoordY + MouseY * (CursorSpeed / 3)
    If coordX < 0 Then coordX = 0
    If coordX > D3DM.width Then coordX = D3DM.width
    If CoordY < 0 Then CoordY = 0
    If CoordY > D3DM.height Then CoordY = D3DM.height
    MouseVerts(0) = AssignMV(coordX, CoordY, 0, 0, 0)
    MouseVerts(1) = AssignMV(coordX + CursorW * D3DM.width / 800, CoordY, 0, 1, 0)
    MouseVerts(2) = AssignMV(coordX, CoordY + CursorH * D3DM.height / 600, 0, 0, 1)
    MouseVerts(3) = AssignMV(coordX + CursorW * D3DM.width / 800, CoordY + CursorH * D3DM.height / 600, 0, 1, 1)
    MouseX = 0: MouseY = 0
    '---------------- SEND EVENTS TO MENUS ------------
    AuxiliarMenu.ProcessMouseMove coordX, CoordY
    If MouseClick0 = True Then
        AuxiliarMenu.ProcessClick coordX, CoordY
        MouseClick0 = False
    End If
    
    If AuxiliarMenu.isEvent Then
        Select Case AuxiliarMenu.EventName
        Case "returngame"
            Out = True
        Case "returnmenu"
            Out = True
            MainRender = False
        End Select
    End If
    AuxiliarMenu.isEvent = False
    
    If Out And Not MainRender Then
        mTimer = GetTickCount()
        Do While (mTimer + 1000) > GetTickCount()
            Alpha = 255 * (GetTickCount() - mTimer) / 1000
            If Alpha > 255 Then Alpha = 255
            If Alpha < 0 Then Alpha = 0
            fadeVertices(0) = AssignMVA(0, 0, 0, D3DColorARGB(Alpha, 0, 0, 0))
            fadeVertices(1) = AssignMVA(D3DM.width, 0, 0, D3DColorARGB(Alpha, 0, 0, 0))
            fadeVertices(2) = AssignMVA(0, D3DM.height, 0, D3DColorARGB(Alpha, 0, 0, 0))
            fadeVertices(3) = AssignMVA(D3DM.width, D3DM.height, 0, D3DColorARGB(Alpha, 0, 0, 0))
            
            Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
            Device.BeginScene
            '------------ RENDER THE MENU --------
            Device.SetVertexShader myVertexFVF
            AuxiliarMenu.RenderMenu
            '----------- RENDER THE CURSOR OVER ALL -----------
            Device.SetTexture 0, MouseTexture
            Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MouseVerts(0), Len(MouseVerts(0))
            
            Device.SetTexture 0, Nothing
            Device.SetVertexShader myVertexAlphaFVF
            Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, fadeVertices(0), Len(fadeVertices(0))
            
            Device.EndScene
            Device.Present ByVal 0, ByVal 0, 0, ByVal 0
            
            Sleep 5
            DoEvents
        Loop
        GoTo OutOfAuxMenu
    End If
    
    Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    Device.BeginScene
    
    '------------ RENDER THE MENU --------
    AuxiliarMenu.RenderMenu
    
    '----------- RENDER THE CURSOR OVER ALL -----------
    Device.SetTexture 0, MouseTexture
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MouseVerts(0), Len(MouseVerts(0))
    
    Device.EndScene
    Device.Present ByVal 0, ByVal 0, 0, ByVal 0
    
    Sleep 5
    DoEvents
Loop

OutOfAuxMenu:

If MainRender = True Then
    'only return to the game
    Call SetCorrectRenderStates(RGB(WorldProperties(TheGameSlot.WorldID).AmbientValues.x, WorldProperties(TheGameSlot.WorldID).AmbientValues.y, WorldProperties(TheGameSlot.WorldID).AmbientValues.z))
    
    Call UnFreezeAll
    
    If MusicEngine.MusicStatus = 2 Then Call MusicEngine.ResumePlay
Else
    'exit to the menu
    
    MusicEngine.StopMusic
    Call DestroyThisScene
End If
AuxMenu = False
AccelerationTimer = AccelerationTimer + (GetTickCount() - PauseTimer)
End Sub

Public Sub FreezeAll()
Dim x As Long
For x = 1 To UBound(MovingCharacters)
    If Not (MovingCharacters(x) Is Nothing) Then
        MovingCharacters(x).FreezeTimer
    End If
Next
Call FreezeSounds
End Sub

Public Sub UnFreezeAll()
Dim x As Long
For x = 1 To UBound(MovingCharacters)
    If Not (MovingCharacters(x) Is Nothing) Then
        MovingCharacters(x).UnFreezeTimer
    End If
Next
Call UnFreezeSounds
End Sub

Public Sub DrawMessages()
Static num As Long
Dim Alpha As Single, Alpha2 As Single, Alpha3 As Single

Select Case MessageState
Case -1
    Exit Sub
Case 0
    MessageTimer = GetTickCount()
    MessageState = 1
    Exit Sub
Case 1
    If MessageFadeTime > (GetTickCount() - MessageTimer) Then
        'fading in
        num = 1
        Alpha = (GetTickCount() - MessageTimer) / MessageFadeTime * 255
    Else
        MessageTimer = GetTickCount()
        MessageState = 2
        Alpha = 255
    End If
Case 2
    Alpha = 255
    If Not (MessageTime > (GetTickCount() - MessageTimer)) Then
        MessageTimer = GetTickCount()
        If num = UBound(MessageString) Then
            MessageState = 5
        Else
            num = num + 1
            Alpha2 = 0: Alpha3 = 255
            MessageState = 3
        End If
    End If
Case 3
    Alpha = 255
    If (GetTickCount() - MessageTimer) < MessageTime / 5 Then
        Alpha2 = (GetTickCount() - MessageTimer) / (MessageTime / 5) * 255
        Alpha3 = 255 - (GetTickCount() - MessageTimer) / (MessageTime / 5) * 255
        If Alpha3 > 255 Then Alpha3 = 255
        If Alpha3 < 0 Then Alpha3 = 7
        If Alpha2 > 255 Then Alpha2 = 255
        If Alpha2 < 0 Then Alpha2 = 7
    Else
        MessageTimer = GetTickCount()
        MessageState = 2
        Alpha = 255
    End If
Case 5
    If MessageFadeTime > (GetTickCount() - MessageTimer) Then
        'fading out
        Alpha = 255 - (GetTickCount() - MessageTimer) / MessageFadeTime * 255
        If Alpha > 255 Then Alpha = 255
        If Alpha < 0 Then Alpha = 7
    Else
        MessageState = -1
        Alpha = 0
    End If
End Select

UIvertices(0) = AssignMVAdv(200 * D3DM.width / 800, 450 * D3DM.height / 600, 0, 0, 0, Alpha)
UIvertices(1) = AssignMVAdv(600 * D3DM.width / 800, 450 * D3DM.height / 600, 0, 1, 0, Alpha)
UIvertices(2) = AssignMVAdv(200 * D3DM.width / 800, 550 * D3DM.height / 600, 0, 0, 1, Alpha)
UIvertices(3) = AssignMVAdv(600 * D3DM.width / 800, 550 * D3DM.height / 600, 0, 1, 1, Alpha)

'multiblending!!! texture alhpa * vertex alhpa = final alhpa
Device.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
Device.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
Device.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_DIFFUSE

Device.SetTexture 0, PaperTex
Device.SetVertexShader myVertexFVF
Device.SetRenderState D3DRS_ZENABLE, 0
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, UIvertices(0), Len(UIvertices(0))

If MessageState = 3 Then
    Call DrawMessageText(MessageString(num), 230 * D3DM.width / 800, 570 * D3DM.width / 800, 470 * D3DM.height / 600, 530 * D3DM.height / 600, Alpha2)
    Call DrawMessageText(MessageString(num - 1), 230 * D3DM.width / 800, 570 * D3DM.width / 800, 470 * D3DM.height / 600, 530 * D3DM.height / 600, Alpha3)
Else
    Call DrawMessageText(MessageString(num), 230 * D3DM.width / 800, 570 * D3DM.width / 800, 470 * D3DM.height / 600, 530 * D3DM.height / 600, Alpha)
End If

Device.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
Device.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_TEXTURE
Device.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_CURRENT
Device.SetRenderState D3DRS_ZENABLE, 1
End Sub

Public Sub DrawMessageText(ByVal cad As String, ByVal x1 As Single, ByVal x2 As Single, ByVal y1 As Single, ByVal y2 As Single, ByVal Alpha As Single)
Dim cad1 As String, cad2 As String, car As String, acode As Byte
Dim x As Long, offset As Single, tu1 As Single, tu2 As Single, tv1 As Single, tv2 As Single
Dim width As Single, height As Single, offset2 As Single

If InStr(1, cad, "\n") <> 0 Then
    cad1 = left(cad, InStr(1, cad, "\n") - 1)
    cad2 = mid(cad, InStr(1, cad, "\n") + 2)
End If

Device.SetTexture 0, MessageFont

If cad1 = "" Then
    width = Len(cad) * 20 * D3DM.width / 800
    height = 24 * D3DM.height / 600
    offset = ((x2 - x1) - width) / 2
    offset2 = ((y2 - y1) - height) / 2
    For x = 1 To Len(cad)
        car = mid(cad, x, 1)
        acode = Asc(car)
        If acode > 128 Then acode = 0
        tu1 = (acode Mod 16) / 16 + (4 / 512)
        tu2 = (acode Mod 16 + 1) / 16 - (4 / 512)
        tv1 = Int(acode / 16) / 8
        tv2 = (Int(acode / 16) + 1) / 8
        
        UIvertices(0) = AssignMVAdv(x1 + offset + (x - 1) * 20 * D3DM.width / 800, y1 + offset2, 0, tu1, tv1, Alpha)
        UIvertices(1) = AssignMVAdv(x1 + offset + x * 20 * D3DM.width / 800, y1 + offset2, 0, tu2, tv1, Alpha)
        UIvertices(2) = AssignMVAdv(x1 + offset + (x - 1) * 20 * D3DM.width / 800, y1 + height + offset2, 0, tu1, tv2, Alpha)
        UIvertices(3) = AssignMVAdv(x1 + offset + x * 20 * D3DM.width / 800, y1 + height + offset2, 0, tu2, tv2, Alpha)
        Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, UIvertices(0), Len(UIvertices(0))
    Next
Else
    width = Len(cad1) * 20 * D3DM.width / 800
    height = 24 * D3DM.height / 600
    offset = ((x2 - x1) - width) / 2
    For x = 1 To Len(cad1)
        car = mid(cad1, x, 1)
        acode = Asc(car)
        If acode > 128 Then acode = 0
        tu1 = (acode Mod 16) / 16 + (4 / 512)
        tu2 = (acode Mod 16 + 1) / 16 - (4 / 512)
        tv1 = Int(acode / 16) / 8
        tv2 = (Int(acode / 16) + 1) / 8
        
        UIvertices(0) = AssignMVAdv(x1 + offset + (x - 1) * 20 * D3DM.width / 800, y1, 0, tu1, tv1, Alpha)
        UIvertices(1) = AssignMVAdv(x1 + offset + x * 20 * D3DM.width / 800, y1, 0, tu2, tv1, Alpha)
        UIvertices(2) = AssignMVAdv(x1 + offset + (x - 1) * 20 * D3DM.width / 800, y1 + height, 0, tu1, tv2, Alpha)
        UIvertices(3) = AssignMVAdv(x1 + offset + x * 20 * D3DM.width / 800, y1 + height, 0, tu2, tv2, Alpha)
        Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, UIvertices(0), Len(UIvertices(0))
    Next
    
    width = Len(cad2) * 20 * D3DM.width / 800
    height = 24 * D3DM.height / 600
    offset = ((x2 - x1) - width) / 2
    For x = 1 To Len(cad2)
        car = mid(cad2, x, 1)
        acode = Asc(car)
        If acode > 128 Then acode = 0
        tu1 = (acode Mod 16) / 16 + (4 / 512)
        tu2 = (acode Mod 16 + 1) / 16 - (4 / 512)
        tv1 = Int(acode / 16) / 8
        tv2 = (Int(acode / 16) + 1) / 8
        
        UIvertices(0) = AssignMVAdv(x1 + offset + (x - 1) * 20 * D3DM.width / 800, y2 - height, 0, tu1, tv1, Alpha)
        UIvertices(1) = AssignMVAdv(x1 + offset + x * 20 * D3DM.width / 800, y2 - height, 0, tu2, tv1, Alpha)
        UIvertices(2) = AssignMVAdv(x1 + offset + (x - 1) * 20 * D3DM.width / 800, y2, 0, tu1, tv2, Alpha)
        UIvertices(3) = AssignMVAdv(x1 + offset + x * 20 * D3DM.width / 800, y2, 0, tu2, tv2, Alpha)
        Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, UIvertices(0), Len(UIvertices(0))
    Next
End If
End Sub

Public Sub DrawMiniMap()
If GetMiniMapTex(TheGameSlot.WorldID) = 0 Then Exit Sub

Dim py As Single, px As Single, x As Integer
Dim tu As Single, tv As Single, temp As Single, vec As D3DVECTOR, mat As D3DMATRIX
Dim size As Single
size = 5 * D3DM.width / 800

Device.SetRenderState D3DRS_ZENABLE, 0

px = 70 * D3DM.width / 800
py = 530 * D3DM.height / 600

Select Case TheGameSlot.WorldID
Case 1, 2
    For x = 0 To -31 Step -1
        tu = (450 - (-CharPos.x + 50 + 20 * sin((x * 11.25 + 90) / 180 * Pi))) / 450
        tv = (600 - (CharPos.z + 100 + 20 * Cos((x * 11.25 + 90) / 180 * Pi))) / 600
        MiniMapVertices(-x) = AssignMVAdv(px + (50 * D3DM.width / 800) * sin((CharAngleH + x * 11.25 + 90) / 180 * Pi), py + (50 * D3DM.height / 600) * Cos((CharAngleH + x * 11.25 + 90) / 180 * Pi), 0, tu, tv, 0)
    Next
Case 3
    For x = 0 To -31 Step -1
        tu = (60 - (-CharPos.x + 30 + 20 * sin((x * 11.25 + 90) / 180 * Pi))) / 60
        tv = (60 - (CharPos.z + 30 + 20 * Cos((x * 11.25 + 90) / 180 * Pi))) / 60
        MiniMapVertices(-x) = AssignMVAdv(px + (50 * D3DM.width / 800) * sin((CharAngleH + x * 11.25 + 90) / 180 * Pi), py + (50 * D3DM.height / 600) * Cos((CharAngleH + x * 11.25 + 90) / 180 * Pi), 0, tu, tv, 0)
    Next
End Select
 
Device.SetVertexShader myVertexFVF

Device.SetTexture 0, MiniMapTex(GetMiniMapTex(TheGameSlot.WorldID))
Device.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 30, MiniMapVertices(0), Len(MiniMapVertices(0))

Device.SetTexture 0, MiniMapIcons

MiniMapVertices(32) = AssignMVAdv(px - size, py - size, 0, 0, 0, 0)
MiniMapVertices(33) = AssignMVAdv(px + size, py - size, 0, 0.25, 0, 0)
MiniMapVertices(34) = AssignMVAdv(px - size, py + size, 0, 0, 0.25, 0)
MiniMapVertices(35) = AssignMVAdv(px + size, py + size, 0, 0.25, 0.25, 0)

Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MiniMapVertices(32), Len(MiniMapVertices(0))

'Draw MiniMap Icons
For x = 1 To UBound(MovingCharacters)
    If Not (MovingCharacters(x) Is Nothing) Then
        If MovingCharacters(x).WorldID = TheGameSlot.WorldID And _
           MovingCharacters(x).Visible = True Then
           vec.x = MovingCharacters(x).Public_Position.x - CharPos.x
           vec.z = MovingCharacters(x).Public_Position.z - CharPos.z
           D3DXMatrixRotationAxis mat, v3(0, 1, 0), (-CharAngleH + 180) / 180 * Pi
           vec = Normalize(vec)
           D3DXVec3TransformCoord vec, vec, mat
           
            If VecDistFastXZ(CharPos, MovingCharacters(x).Public_Position) < 400 Then
                temp = CalculateDistanceXZ(CharPos, MovingCharacters(x).Public_Position) / 20
                vec.x = vec.x * temp * 50 * D3DM.width / 800
                vec.z = vec.z * temp * 50 * D3DM.height / 600
                
                If Abs(MovingCharacters(x).Public_Position.y - CharPos.y) < 1 Then
                    'same level
                    MiniMapVertices(32) = AssignMVAdv(px + vec.x - size, py - vec.z - size, 0, 0.25, 0, 0)
                    MiniMapVertices(33) = AssignMVAdv(px + vec.x + size, py - vec.z - size, 0, 0.5, 0, 0)
                    MiniMapVertices(34) = AssignMVAdv(px + vec.x - size, py - vec.z + size, 0, 0.25, 0.25, 0)
                    MiniMapVertices(35) = AssignMVAdv(px + vec.x + size, py - vec.z + size, 0, 0.5, 0.25, 0)
                ElseIf MovingCharacters(x).Public_Position.y - CharPos.y > 0 Then
                    'its up
                    MiniMapVertices(32) = AssignMVAdv(px + vec.x - size, py - vec.z - size, 0, 0.5, 0, 0)
                    MiniMapVertices(33) = AssignMVAdv(px + vec.x + size, py - vec.z - size, 0, 0.75, 0, 0)
                    MiniMapVertices(34) = AssignMVAdv(px + vec.x - size, py - vec.z + size, 0, 0.5, 0.25, 0)
                    MiniMapVertices(35) = AssignMVAdv(px + vec.x + size, py - vec.z + size, 0, 0.75, 0.25, 0)
                Else
                    'its down
                    MiniMapVertices(32) = AssignMVAdv(px + vec.x - size, py - vec.z - size, 0, 0.75, 0, 0)
                    MiniMapVertices(33) = AssignMVAdv(px + vec.x + size, py - vec.z - size, 0, 1, 0, 0)
                    MiniMapVertices(34) = AssignMVAdv(px + vec.x - size, py - vec.z + size, 0, 0.75, 0.25, 0)
                    MiniMapVertices(35) = AssignMVAdv(px + vec.x + size, py - vec.z + size, 0, 1, 0.25, 0)
                End If
                
                Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MiniMapVertices(32), Len(MiniMapVertices(0))
            End If
        End If
    End If
Next

For x = 1 To UBound(MissionTargets)
    If MissionTargets(x).Visible = True And MissionTargets(x).WorldID = TheGameSlot.WorldID Then
        vec.x = MissionTargets(x).Position.x - CharPos.x
        vec.z = MissionTargets(x).Position.z - CharPos.z
        D3DXMatrixRotationAxis mat, v3(0, 1, 0), (-CharAngleH + 180) / 180 * Pi
        vec = Normalize(vec)
        D3DXVec3TransformCoord vec, vec, mat

        If VecDistFastXZ(CharPos, MissionTargets(x).Position) < 400 Then
            temp = CalculateDistanceXZ(CharPos, MissionTargets(x).Position) / 20
            vec.x = vec.x * temp * 50 * D3DM.width / 800
            vec.z = vec.z * temp * 50 * D3DM.height / 600

            If Abs(MissionTargets(x).Position.y - CharPos.y) < 1 Then
                    'same level
                    MiniMapVertices(32) = AssignMVAdv(px + vec.x - size, py - vec.z - size, 0, 0.25, 0.25, 0)
                    MiniMapVertices(33) = AssignMVAdv(px + vec.x + size, py - vec.z - size, 0, 0.5, 0.25, 0)
                    MiniMapVertices(34) = AssignMVAdv(px + vec.x - size, py - vec.z + size, 0, 0.25, 0.5, 0)
                    MiniMapVertices(35) = AssignMVAdv(px + vec.x + size, py - vec.z + size, 0, 0.5, 0.5, 0)
                ElseIf MissionTargets(x).Position.y - CharPos.y > 0 Then
                    'its up
                    MiniMapVertices(32) = AssignMVAdv(px + vec.x - size, py - vec.z - size, 0, 0.5, 0.25, 0)
                    MiniMapVertices(33) = AssignMVAdv(px + vec.x + size, py - vec.z - size, 0, 0.75, 0.25, 0)
                    MiniMapVertices(34) = AssignMVAdv(px + vec.x - size, py - vec.z + size, 0, 0.5, 0.5, 0)
                    MiniMapVertices(35) = AssignMVAdv(px + vec.x + size, py - vec.z + size, 0, 0.75, 0.5, 0)
                Else
                    'its down
                    MiniMapVertices(32) = AssignMVAdv(px + vec.x - size, py - vec.z - size, 0, 0.75, 0.25, 0)
                    MiniMapVertices(33) = AssignMVAdv(px + vec.x + size, py - vec.z - size, 0, 1, 0.25, 0)
                    MiniMapVertices(34) = AssignMVAdv(px + vec.x - size, py - vec.z + size, 0, 0.75, 0.5, 0)
                    MiniMapVertices(35) = AssignMVAdv(px + vec.x + size, py - vec.z + size, 0, 1, 0.5, 0)
                End If

                Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MiniMapVertices(32), Len(MiniMapVertices(0))
        End If
    End If
Next

MiniMapVertices(32) = AssignMVAdv(px + (90 * D3DM.width / 800) * sin((CharAngleH + 135) / 180 * Pi), py + (90 * D3DM.height / 600) * Cos((CharAngleH + 135) / 180 * Pi), 0, 0, 0, 0)
MiniMapVertices(33) = AssignMVAdv(px + (90 * D3DM.width / 800) * sin((CharAngleH + 45) / 180 * Pi), py + (90 * D3DM.height / 600) * Cos((CharAngleH + 45) / 180 * Pi), 0, 1, 0, 0)
MiniMapVertices(34) = AssignMVAdv(px + (90 * D3DM.width / 800) * sin((CharAngleH - 135) / 180 * Pi), py + (90 * D3DM.height / 600) * Cos((CharAngleH - 135) / 180 * Pi), 0, 0, 1, 0)
MiniMapVertices(35) = AssignMVAdv(px + (90 * D3DM.width / 800) * sin((CharAngleH - 45) / 180 * Pi), py + (90 * D3DM.height / 600) * Cos((CharAngleH - 45) / 180 * Pi), 0, 1, 1, 0)

Device.SetTexture 0, MiniMapBorder
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MiniMapVertices(32), Len(MiniMapVertices(0))

Device.SetRenderState D3DRS_ZENABLE, 1
End Sub

Public Sub GetTUTV(postion As D3DVECTOR, ByVal WorldID As Integer, tu As Single, tv As Single)
Select Case WorldID
Case 1, 2
    tu = (450 - (-postion.x + 50)) / 450
    tv = (600 - (postion.z + 100)) / 600
Case 3
    tu = (60 - (-postion.x + 30)) / 60
    tv = (60 - (postion.z + 30)) / 60
End Select
End Sub

Public Sub BgMusic()
Dim num As Integer
If UBound(BackgroundMusic) = 0 Then Exit Sub
If MusicEngine.MusicStatus = 0 Then
    If BgMusicRandom Then
        If UBound(BackgroundMusic) <> 1 Then
            Do
                num = Int(Rnd() * UBound(BackgroundMusic)) + 1
            Loop While num <> UBound(BackgroundMusic)
        End If
    Else
        BackgroundMusicTrack = BackgroundMusicTrack + 1
        If BackgroundMusicTrack > UBound(BackgroundMusic) Then BackgroundMusicTrack = 1
    End If
    MusicEngine.StopMusic
    MusicEngine.PlayMusic MusicDir & BackgroundMusic(BackgroundMusicTrack)
End If
Call MusicEngine.RenderTime
End Sub

Public Sub DrawUI()
Call DrawMessages    'message goes here, before fading but in front of all the scene
Call DrawMiniMap

Device.SetVertexShader myVertexFVF

UIvertices(0) = AssignMVAdv(40 * D3DM.width / 800, 40 * D3DM.height / 600, 0, 0, 0, 0)
UIvertices(1) = AssignMVAdv(80 * D3DM.width / 800, 40 * D3DM.height / 600, 0, 1, 0, 0)
UIvertices(2) = AssignMVAdv(40 * D3DM.width / 800, 80 * D3DM.height / 600, 0, 0, 1, 0)
UIvertices(3) = AssignMVAdv(80 * D3DM.width / 800, 80 * D3DM.height / 600, 0, 1, 1, 0)

Device.SetTexture 0, Coin
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, UIvertices(0), Len(UIvertices(0))
End Sub

Public Sub VarCleanUp()
FrameAverage = 0: TimeCounter = 0: LastWorldLoaded = 0
AccelerationTimer = 0
MakeFade2 = True: MakeFade1 = False
Movement = 0: MessageState = -1
ReDim MessageString(0)
RestartLoop = False
vectorUP = v3(0, 1, 0)
CurrentCharAnimation = "": CharAnimationTimer = 0: Movement = 0
MainChar.ClearCycle 0, "quiet", 0: MainChar.ClearCycle 0, "walking", 0
MainChar.BlendCycle 0, "quiet", 1, 0
ResetAllMoves = True: CharSpeed = 0
CameraType = 0
End Sub
