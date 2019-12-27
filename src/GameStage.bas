Attribute VB_Name = "GameStage"
Option Explicit

'- This module controls the game missions and stages ---

'--- First mission stage -----
Public Sub ResetGS(gs As GameSlot)
gs.Health = 100
gs.MissionLevel = 1
gs.Coins = 0
gs.NumObjects = 0
ReDim gs.ObjectsID(0)
End Sub

Public Sub EnterNewWorld(ByVal WorldID As Integer, ByVal door As Integer)
Static MemWorldID As Integer, MemDoor As Integer
If MakeFade1 = False Then
    MemWorldID = WorldID
    MemDoor = door
    MakeFade1 = True
    Exit Sub
Else
    'load the new world if itisn'a a fixed world
    Call WorldChange(MemWorldID)
    
    WorldID = MemWorldID
    door = MemDoor
    MakeFade1 = False
    MakeFade2 = True
    ResetAllMoves = True
End If

Dim x As Long
For x = 1 To UBound(DoorArray)
    If DoorArray(x).World = WorldID And DoorArray(x).id = door Then
        TheGameSlot.Position = DoorArray(x).Position
        TheGameSlot.RotationH = DoorArray(x).RotationH
        Exit For
    End If
Next

Call CharsWorldChange(TheGameSlot.WorldID, WorldID)

TheGameSlot.WorldID = WorldID
CharAngleH = TheGameSlot.RotationH
CharAngleV = (MaxVerticalAngle + MinVerticalAngle) / 2
CharPos = TheGameSlot.Position
CharDistance = (MaxCameraDistance + MinCameraDistance) / 2
CameraPos = ProcessCameraCoords(CharPos, CharAngleH, CharAngleV, CharDistance)
InitY = CharPos.y
Movement = 0
Call SetCorrectRenderStates(RGB(WorldProperties(TheGameSlot.WorldID).AmbientValues.x, WorldProperties(TheGameSlot.WorldID).AmbientValues.y, WorldProperties(TheGameSlot.WorldID).AmbientValues.z))
Call UpdateLights
End Sub

Public Sub CharsWorldChange(ByVal LastW As Integer, ByVal NewW As Integer)
Dim x As Long
For x = 1 To UBound(MovingCharacters)
    If Not (MovingCharacters(x) Is Nothing) Then
        If MovingCharacters(x).Visible And MovingCharacters(x).Paused = False And MovingCharacters(x).WorldID = LastW Then
            Call MovingCharacters(x).FreezeTimer
        End If
        If MovingCharacters(x).Visible And MovingCharacters(x).Paused = False And MovingCharacters(x).WorldID = NewW Then
            Call MovingCharacters(x).UnFreezeTimer
        End If
    End If
Next
End Sub

Public Sub UpdateLights()
Dim x As Long, cont As Long
Dim light As D3DLIGHT8
For x = 0 To 7
    Device.LightEnable x, 0
Next
For x = 1 To UBound(Lights)
    If Lights(x).WorldID = TheGameSlot.WorldID Then
        Select Case Lights(x).type
        Case Directional
            light.type = D3DLIGHT_DIRECTIONAL
            light.Direction = Lights(x).Direction
        Case Omni
            light.type = D3DLIGHT_POINT
            light.Position = Lights(x).Position
            light.Attenuation0 = 0.1
            light.Range = Lights(x).Range
        Case Target
            light.type = D3DLIGHT_SPOT
            light.Position = Lights(x).Position
            light.Direction = Lights(x).Direction
            light.Range = Lights(x).Range
            light.Phi = Lights(x).Phi
            light.Theta = Lights(x).Theta
        End Select
        light.diffuse = Lights(x).Color
        Device.SetLight cont, light
        Device.LightEnable cont, 1
        cont = cont + 1
    End If
Next
End Sub

Public Sub ComputeDoors()
If MakeFade1 Then Exit Sub      'if we are changing do not test doors or we'll get errors or a inf loop
If DisableDoors Then Exit Sub

Dim x As Long
For x = 1 To UBound(AreaArray)
    If AreaArray(x).World = TheGameSlot.WorldID Then
        If CheckPointInCube(AreaArray(x).pos, AreaArray(x).Pos2, CharPos) And NoLock(AreaArray(x).DoorName) Then
            EnterNewWorld AreaArray(x).NewWorld, AreaArray(x).DoorId
            Exit Sub
        End If
    End If
Next

End Sub

Public Function NoLock(ByVal door As String) As Boolean
Dim x As Long
NoLock = True
For x = 1 To UBound(LockedDoors)
    If LockedDoors(x) = door Then
        NoLock = False
        Exit For
    End If
Next
End Function

Public Sub WorldChange(ByVal NewWorld As Integer)
If LastWorldLoaded = NewWorld Then Exit Sub
If NewWorld < 4 Then Exit Sub

Device.EndScene
Call LoadingWorldScreen

Call ClearSecondaryCollisionArray
Call ClearTertiaryTextures
Call ClearTertiaryModels        'diria k no hay Secondary...
Call ComputeWorldCollision(NewWorld)
Call ComputeCollisionFloats2ndPassWorld(NewWorld)

Call LoadingWorldScreen

Dim mypath As String, x As Long, key As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "models_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "graphics\" & Trim(Str(NewWorld)) & ".dat", mypath
key = WorldProperties(NewWorld).State

For x = 1 To 10
    If FileExists(mypath & key & "_" & Trim(Str(x)) & ".x") Then
        Call LoadingWorldScreen
        Set RenderModelsLMAux(x) = New cAdvMesh
        RenderModelsLMAux(x).LoadFromFile mypath & key & "_" & Trim(Str(x)) & ".x", mypath & key & "_" & Trim(Str(x)) & "l.x", mypath, True
        RenderModelsLMAuxNum = RenderModelsLMAuxNum + 1
    End If
Next

Call LoadingWorldScreen
DeleteDir mypath
End Sub

Public Sub LoadingWorldScreen()
LVertices(0) = AssignMV(0, 0, 0, 0, 0)
LVertices(1) = AssignMV(D3DM.width, 0, 0, 1, 0)
LVertices(2) = AssignMV(0, D3DM.height, 0, 0, 1)
LVertices(3) = AssignMV(D3DM.width, D3DM.height, 0, 1, 1)

DoEvents
Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
Device.BeginScene
Device.SetTexture 0, LoadingWorld
Device.SetRenderState D3DRS_ZENABLE, 0
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LVertices(0), Len(LVertices(0))
Device.SetRenderState D3DRS_ZENABLE, 1
Device.EndScene
Device.Present ByVal 0, ByVal 0, 0, ByVal 0
DoEvents
End Sub

Public Sub ClearTertiaryModels()
Dim x As Long
For x = 1 To 10
    Set RenderModelsLMAux(x) = Nothing
Next
RenderModelsLMAuxNum = 0
End Sub

Public Sub ClearTertiaryTextures()
Dim x As Long
For x = 1 To UBound(GameTexturesTer)
    GameTexturesTer(x).filename = ""
    Set GameTexturesTer(x).texture = Nothing
Next
ReDim GameTexturesTer(0)
End Sub
