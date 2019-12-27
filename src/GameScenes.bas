Attribute VB_Name = "GameScenes"
Option Explicit

Public Sub UpdateScene()
'check for objectives or changes in the missions / world
'es comproven les condicions i s'executen les accions corresponents
'escrit en XML a stages/
Dim x As Long, y As Long, Ok As Boolean, z As Long, strarray() As String, h As Long
Dim vec1() As D3DVECTOR, vec2() As D3DVECTOR, outs() As salida, cont As Long
For x = 1 To UBound(LevelScenes)
    If LevelScenes(x).Enabled Then
        Ok = True
        For y = 1 To UBound(LevelScenes(x).Conditions)
            'check for the condition
            Select Case LevelScenes(x).Conditions(y).type
            Case "isinarea"
                If CharPos.x > 0 Then
                    If CharPos.z > 0 Then
                        If Not (CharPos.x > LevelScenes(x).Conditions(y).varg.x And CharPos.x < LevelScenes(x).Conditions(y).varg2.x And _
                                CharPos.z > LevelScenes(x).Conditions(y).varg.z And CharPos.z < LevelScenes(x).Conditions(y).varg2.z) Then
                            Ok = False
                        End If
                    Else
                        If Not (CharPos.x > LevelScenes(x).Conditions(y).varg.x And CharPos.x < LevelScenes(x).Conditions(y).varg2.x And _
                                CharPos.z < LevelScenes(x).Conditions(y).varg.z And CharPos.z > LevelScenes(x).Conditions(y).varg2.z) Then
                            Ok = False
                        End If
                    End If
                Else
                    If CharPos.z > 0 Then
                        If Not (CharPos.x < LevelScenes(x).Conditions(y).varg.x And CharPos.x > LevelScenes(x).Conditions(y).varg2.x And _
                                CharPos.z > LevelScenes(x).Conditions(y).varg.z And CharPos.z < LevelScenes(x).Conditions(y).varg2.z) Then
                            Ok = False
                        End If
                    Else
                        If Not (CharPos.x < LevelScenes(x).Conditions(y).varg.x And CharPos.x > LevelScenes(x).Conditions(y).varg2.x And _
                                CharPos.z < LevelScenes(x).Conditions(y).varg.z And CharPos.z > LevelScenes(x).Conditions(y).varg2.z) Then
                            Ok = False
                        End If
                    End If
                End If
            Case "isintarget"
                For z = 1 To UBound(MissionTargets)
                    If MissionTargets(z).id = LevelScenes(x).Conditions(y).sarg Then
                        If Not VecDistFast(CharPos, MissionTargets(z).Position) < MissionTargets(z).radius ^ 2 Then
                            Ok = False
                        End If
                        Exit For
                    End If
                Next
            Case "fadestate"
                If LevelScenes(x).Conditions(y).sarg = "clear" Then
                    If FadeState <> 0 Then Ok = False
                Else    'black
                    If FadeState <> 3 Then Ok = False
                End If
            Case "messagestate"
                If LevelScenes(x).Conditions(y).sarg = "no" Then
                    If MessageState <> -1 Then Ok = False
                Else
                    If MessageState = -1 Then Ok = False
                End If
            Case "charvisible"
                strarray = Split(LevelScenes(x).Conditions(y).sarg, ",")
                For z = 0 To UBound(strarray)
                    For h = 1 To UBound(MovingCharacters)
                        If MovingCharacters(h).id = Trim(strarray(z)) And MovingCharacters(h).WorldID = TheGameSlot.WorldID Then
                            If MovingCharacters(h).CheckCharVisible() Then
                                cont = cont + 1
                                ReDim Preserve vec1(cont - 1): ReDim Preserve vec2(cont - 1): ReDim Preserve outs(cont - 1)
                                vec1(cont - 1) = MovingCharacters(h).Public_Position
                                vec2(cont - 1) = CharPos
                                vec1(cont - 1).y = vec1(cont - 1).y + 0.5: vec2(cont - 1).y = vec2(cont - 1).y + 0.5
                                Exit For
                            End If
                        End If
                    Next
                Next
                If cont <> 0 Then
                    Visible vec1(0), vec2(0), CollisionFloats(TheGameSlot.WorldID).vertices(0), UBound(CollisionFloats(TheGameSlot.WorldID).vertices) / 3, cont, outs(0)
                    For z = 0 To cont - 1
                        If outs(z).respuesta = 0 Then
                            'one visible, goto skip
                            GoTo VisibleSkip
                        End If
                    Next
                End If
                Ok = False
VisibleSkip:
                ReDim vec1(0): ReDim vec2(0): ReDim outs(0)
            End Select
        Next
        If Ok = True Then
            'conditions ok, so execute effects
            For y = 1 To UBound(LevelScenes(x).Effects)
                Select Case LevelScenes(x).Effects(y).type
                Case "setcharpos"
                    CharPos = LevelScenes(x).Effects(y).varg
                    ResetAllMoves = True
                    InitY = CharPos.y
                    Movement = 0
                Case "disablescene"
                    For z = 1 To UBound(LevelScenes)
                        If LevelScenes(z).id = LevelScenes(x).Effects(y).larg Then
                            LevelScenes(z).Enabled = False
                            Exit For
                        End If
                    Next
                Case "enablescene"
                    For z = 1 To UBound(LevelScenes)
                        If LevelScenes(z).id = LevelScenes(x).Effects(y).larg Then
                            LevelScenes(z).Enabled = True
                            Exit For
                        End If
                    Next
                Case "enableobject"
                    For z = 1 To UBound(MissionTargets)
                        If MissionTargets(z).id = LevelScenes(x).Effects(y).sarg Then
                            MissionTargets(z).Visible = True
                            Exit For
                        End If
                    Next
                    For z = 1 To UBound(MovingCharacters)
                        If MovingCharacters(z).id = LevelScenes(x).Effects(y).sarg Then
                            MovingCharacters(z).Visible = True
                            MovingCharacters(z).SetPos 0
                            Exit For
                        End If
                    Next
                    For z = 1 To UBound(SavePoints)
                        If SavePoints(z).id = LevelScenes(x).Effects(y).sarg Then
                            SavePoints(z).Visible = True
                            Exit For
                        End If
                    Next
                Case "disableobject"
                    For z = 1 To UBound(MissionTargets)
                        If MissionTargets(z).id = LevelScenes(x).Effects(y).sarg Then
                            MissionTargets(z).Visible = False
                            Exit For
                        End If
                    Next
                    For z = 1 To UBound(MovingCharacters)
                        If MovingCharacters(z).id = LevelScenes(x).Effects(y).sarg Then
                            MovingCharacters(z).Visible = False
                            Exit For
                        End If
                    Next
                    For z = 1 To UBound(SavePoints)
                        If SavePoints(z).id = LevelScenes(x).Effects(y).sarg Then
                            SavePoints(z).Visible = False
                            Exit For
                        End If
                    Next
                Case "doors"
                    If LevelScenes(x).Effects(y).larg = 0 Then
                        DisableDoors = True
                    Else
                        DisableDoors = False
                    End If
                Case "lockdoor"
                    ReDim LockedDoors(UBound(LockedDoors) + 1)
                    LockedDoors(UBound(LockedDoors)) = LevelScenes(x).Effects(y).sarg
                Case "unlockdoor"
                    For z = 1 To UBound(LockedDoors)
                        If LockedDoors(z) = LevelScenes(x).Effects(y).sarg Then
                            For h = z To UBound(LockedDoors) - 1
                                LockedDoors(h) = LockedDoors(h + 1)
                            Next
                            ReDim Preserve LockedDoors(UBound(LockedDoors) - 1)
                            Exit For
                        End If
                    Next
                Case "alldoors"
                    If LevelScenes(x).Effects(y).sarg = "lock" Then
                        Call FillDoorsArray
                    Else
                        ReDim LockedDoors(0)
                    End If
                Case "setfade"
                    If LevelScenes(x).Effects(y).sarg = "in" Then
                        FadeState = 1
                    ElseIf LevelScenes(x).Effects(y).sarg = "out" Then
                        FadeState = 2
                    Else
                        FadeState = 3
                    End If
                Case "setfadetime"
                    FadeTimeMS = LevelScenes(x).Effects(y).larg   'sets the duration of a fade
                Case "setworld"
                    TheGameSlot.WorldID = LevelScenes(x).Effects(y).larg
                    Call SetCorrectRenderStates(RGB(WorldProperties(TheGameSlot.WorldID).AmbientValues.x, WorldProperties(TheGameSlot.WorldID).AmbientValues.y, WorldProperties(TheGameSlot.WorldID).AmbientValues.z))
                    Call UpdateLights
                Case "setcamera"
                    'reposiciona la càmera donat l'angle horitzontal
                    CharAngleV = (MaxVerticalAngle + MinVerticalAngle) / 2
                    CharAngleH = LevelScenes(x).Effects(y).larg
                    CharDistance = (MaxCameraDistance + MinCameraDistance) / 2
                    CameraPos = ProcessCameraCoords(CharPos, CharAngleH, CharAngleV, CharDistance)
                Case "showmessage"
                    MessageTime = LevelScenes(x).Effects(y).larg2
                    MessageFadeTime = LevelScenes(x).Effects(y).larg
                    MessageString() = Split("\u" & LevelScenes(x).Effects(y).sarg, "\u")
                    MessageState = 0
                Case "playvideo"
                    VideoOn = EXE & "video\" & LevelScenes(x).Effects(y).sarg
                Case "endstage"
                    Select Case LevelScenes(x).Effects(y).sarg
                    Case "reset"
                        Call ResetThisStage
                    End Select
                Case "collision"
                    Call ComputeCollisionFloats2ndPass
                    If TheGameSlot.WorldID > 3 Then Call ComputeCollisionFloats2ndPassWorld(TheGameSlot.WorldID)
                End Select
            Next
            If x = 1 Then
                'first starter sequence
                'save into the savedgameslot all the info
                With SavedGameSlot
                    .Position = CharPos
                    .WorldID = TheGameSlot.WorldID
                    .RotationH = CharAngleH
                    .MissionLevel = SceneNumber
                    '.
                End With
            End If
            Exit Sub
        End If
    End If
Next
End Sub

Public Sub RenderScene()

End Sub

Public Sub ResetCurrentLevel()
On Local Error Resume Next
Dim x As Long
LevelScenes(1).Enabled = True
For x = 2 To UBound(LevelScenes)
    LevelScenes(x).Enabled = False
Next
For x = 1 To UBound(MissionTargets)
    MissionTargets(x).Visible = False
Next
For x = 1 To UBound(MovingCharacters)
    MovingCharacters(x).SetPos 0
    MovingCharacters(x).Visible = False
Next
ReDim LockedDoors(0)
FadeState = 0
MessageState = -1
ReDim MessageString(0)
End Sub

Public Sub SetLevelAtributes(ByVal Level As Long)
'load the current level
Dim mypath As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "stages_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
LoadingStageScreen 0
ExtractFile EXE & "stages\stage" & Trim(Str(Level)) & ".dat", mypath
LoadingStageScreen 20

ReDim FixedObjects(0)
SceneNumber = Level
LoadParseExecuteXML mypath & "stage" & Trim(Str(Level)) & ".dat", mypath, 10, 64

Dim x As Long
For x = 1 To 3
    LoadWorldModels x, WorldProperties(x).Mesh, 64 + 12 * (x - 1), 64 + 12 * x
Next

Call ResetCurrentLevel

DeleteDir mypath
End Sub

Public Sub LoadParseExecuteXML(ByVal file As String, ByVal ResPath As String, ByVal barmin As Integer, ByVal barmax As Integer)
Dim ff As Integer, cad As String, cad2 As String
Dim lines() As String, x As Long, y As Long, z As Long
Dim R1 As Long, R2 As Long

Dim Position As D3DVECTOR, RotationY As Single
Dim Mesh As String, Meshsimple As String, AutoY As Boolean
Dim WorldID As Integer
Dim tempfire As cFire
Set tempfire = New cFire
Dim TempPath As String

ff = FreeFile()

Open file For Binary As #ff
    cad = Space(LOF(ff))
    Get #ff, , cad
    lines = Split(cad, vbCrLf)
Close #ff

Call DestroyThisScene

Do While x <= UBound(lines)
    If left(lines(x), 1) = "<" Then
        Position = v3(0, 0, 0): RotationY = 0: Mesh = "": Meshsimple = "": AutoY = False
        Select Case lines(x)
        Case "<ambient>"
            WorldProperties(WorldID).AmbientValues = ParsePos(lines(x + 1))
            x = x + 2
        Case "<unlitambient>"
            WorldProperties(WorldID).UnlitAmbientValues = ParsePos(lines(x + 1))
            x = x + 2
        Case "<night>"
            WorldProperties(WorldID).SkyBox = 3
        Case "<morning>"
            WorldProperties(WorldID).SkyBox = 1
        Case "<afternoon>"
            WorldProperties(WorldID).SkyBox = 2
        Case "<nosky>"
            WorldProperties(WorldID).SkyBox = 0
            
        Case "<geo_night>"
            WorldProperties(WorldID).Mesh = "night"
            'LoadWorldModels WorldID, "night"
        Case "<geo_morning>"
            WorldProperties(WorldID).Mesh = "morning"
        Case "<geo_afternoon>"
            WorldProperties(WorldID).Mesh = "afternoon"
        Case "<geo_unique>"
            WorldProperties(WorldID).Mesh = "unique"
        
        Case "<geo_load_night>"
            WorldProperties(WorldID).State = "night"
        Case "<geo_load_morning>"
            WorldProperties(WorldID).State = "morning"
        Case "<geo_load_afternoon>"
            WorldProperties(WorldID).State = "afternoon"
        Case "<geo_load_unique>"
            WorldProperties(WorldID).State = "unique"
            
        Case "<target>"
            Do While lines(x) <> "</target>"
                x = x + 1
                If lines(x) <> "</target>" And lines(x) <> "" Then TargetFromString lines(x), WorldID
            Loop
        Case "<fixedobject>"
            ReDim Preserve FixedObjects(UBound(FixedObjects) + 1)
            Do While lines(x) <> "</fixedobject>"
                x = x + 1
                cad = lines(x)
                R1 = InStr(1, cad, "=")
                cad2 = mid(cad, R1 + 1)
                If R1 <> 0 Then
                Select Case LCase(left(cad, R1 - 1))
                Case "pos"
                    Position = ParsePos(cad2)
                Case "rotationy"
                    RotationY = Val(cad2)
                Case "mesh"
                    Mesh = cad2
                Case "meshsimple"
                    Meshsimple = cad2
                Case "autoy"
                    If Val(cad) <> 0 Then AutoY = True
                End Select
                End If
            Loop
            FixedObjects(UBound(FixedObjects)) = LoadFixedObject(Position, ResPath & Mesh, ResPath & Meshsimple, AutoY, RotationY, WorldID)
        Case "<fire>"
            Do While lines(x) <> "</fire>"
                x = x + 1
                cad = lines(x)
                R1 = InStr(1, cad, "=")
                cad2 = mid(cad, R1 + 1)
                If R1 <> 0 Then
                Select Case LCase(left(cad, R1 - 1))
                Case "radius"
                    tempfire.radius = Val(cad2)
                Case "nump"
                    tempfire.NumberOfParticles = Val(cad2)
                Case "heightspeed"
                    tempfire.FlameHeightInc = Val(cad2)
                Case "fade"
                    tempfire.FadeInc = Val(cad2)
                Case "compression"
                    tempfire.FlameCompression = Val(cad2)
                Case "pos"
                    Do While LCase(left(lines(x), 4)) = "pos="
                        ReDim Preserve Fires(UBound(Fires) + 1)
                        Set Fires(UBound(Fires)) = New cFire
                        Fires(UBound(Fires)).InitializeFast tempfire.radius, tempfire.FlameCompression, tempfire.FlameHeightInc, tempfire.FadeInc, tempfire.NumberOfParticles, ParsePos(cad2), WorldID
                        
                        x = x + 1
                        cad = lines(x)
                        R1 = InStr(1, cad, "=")
                        cad2 = mid(cad, R1 + 1)
                    Loop
                End Select
                End If
            Loop
        Case "<savepoint>"
            Do While lines(x) <> "</savepoint>"
                x = x + 1
                cad = lines(x)
                R1 = InStr(1, cad, "=")
                cad2 = mid(cad, R1 + 1)
                If R1 <> 0 Then
                Select Case LCase(left(cad, R1 - 1))
                Case "pos"
                    Position = ParsePos(cad2)
                Case "id"
                    Mesh = cad2
                End Select
                End If
            Loop
            ReDim Preserve SavePoints(UBound(SavePoints) + 1)
            SavePoints(UBound(SavePoints)).Position = Position
            SavePoints(UBound(SavePoints)).WorldID = WorldID
            SavePoints(UBound(SavePoints)).id = Mesh
        Case "<model>"
            ReDim Preserve Cal3DModelArray(UBound(Cal3DModelArray) + 1)
            Set Cal3DModelArray(UBound(Cal3DModelArray)) = New Cal3DModel
            Do While lines(x) <> "</model>"
                x = x + 1
                cad = lines(x)
                R1 = InStr(1, cad, "=")
                cad2 = mid(cad, R1 + 1)
                If R1 <> 0 Then
                Select Case LCase(left(cad, R1 - 1))
                Case "data"
                    TempPath = AuxFunctions.TempPath()
                    If right(TempPath, 1) <> "\" Then TempPath = TempPath & "\"
                    TempPath = TempPath & "char_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
                    On Local Error Resume Next: MkDir TempPath: On Local Error GoTo 0
                    ExtractFile EXE & "graphics\" & cad2, TempPath
                Case "name"
                    Cal3DModelArray(UBound(Cal3DModelArray)).ModelName = cad2
                    Cal3DModelArray(UBound(Cal3DModelArray)).LoadData TempPath, Cal3DModelArray(UBound(Cal3DModelArray)).ModelName
                Case "dimensions"
                    Cal3DModelArray(UBound(Cal3DModelArray)).ModelDimensions = Val(cad2) / 2
                Case "anim"
                    cad = mid(cad2, InStr(1, cad2, ",") + 1)
                    cad2 = left(cad2, InStr(1, cad2, ",") - 1)
                    Cal3DModelArray(UBound(Cal3DModelArray)).LoadAnim TempPath & cad2, cad
                End Select
                End If
                LoadingStageScreen barmin + x * (barmax - barmin) / UBound(lines)
            Loop
            Cal3DModelArray(UBound(Cal3DModelArray)).NowReady
            Cal3DModelArray(UBound(Cal3DModelArray)).LoadTexturesToSec TempPath
            DeleteDir TempPath
        Case "<char>"
            y = UBound(MovingCharacters) + 1
            ReDim Preserve MovingCharacters(y)
            Set MovingCharacters(y) = New clsDynObj
            MovingCharacters(y).WorldID = WorldID
            Do While lines(x) <> "</char>"
                x = x + 1
                cad = lines(x)
                R1 = InStr(1, cad, "=")
                cad2 = mid(cad, R1 + 1)
                If R1 <> 0 Then
                Select Case LCase(left(cad, R1 - 1))
                Case "model"
                    MovingCharacters(y).SetCore GetModelCoreFromName(cad2)
                Case "y"
                    If cad2 = "true" Then
                        MovingCharacters(y).AutoY = True
                    ElseIf cad2 = "false" Then
                        MovingCharacters(y).AutoY = False
                    Else
                        MovingCharacters(y).AutoY = False
                        MovingCharacters(y).CoordY = Val(cad2)
                    End If
                Case "dim"
                    MovingCharacters(y).dimensions = Val(cad2)
                Case "addpoint"
                    AddPointToObject MovingCharacters(y), cad2
                Case "id"
                    MovingCharacters(y).id = cad2
                Case "sound"
                    MovingCharacters(y).SetSpeak cad2
                Case "radius"
                    MovingCharacters(y).StopRadius = Val(cad2)
                Case "viewradius"
                    MovingCharacters(y).VisionRadius = Val(cad2)
                Case "foview"
                    MovingCharacters(y).FOView = Val(cad2)
                Case "qanim"
                    MovingCharacters(y).AddQuietAnimation cad2
                End Select
                End If
            Loop
            MovingCharacters(y).SetPos 0
            MovingCharacters(y).SetLimit -1
        Case "<light>"
            y = UBound(Lights) + 1
            ReDim Preserve Lights(y)
            Lights(y).WorldID = WorldID
            Do While lines(x) <> "</light>"
                x = x + 1
                cad = lines(x)
                R1 = InStr(1, cad, "=")
                cad2 = mid(cad, R1 + 1)
                If R1 <> 0 Then
                Select Case LCase(left(cad, R1 - 1))
                Case "type"
                    Select Case cad2
                    Case "omni"
                        Lights(y).type = Omni
                    Case "target"
                        Lights(y).type = Target
                    Case "directional"
                        Lights(y).type = Directional
                    End Select
                Case "pos"
                    Lights(y).Position = ParsePos(cad2)
                Case "dir"
                    Lights(y).Direction = ParsePos(cad2)
                Case "icone"
                    Lights(y).Theta = Val(cad2)
                Case "ocone"
                    Lights(y).Phi = Val(cad2)
                Case "range"
                    Lights(y).Range = Val(cad2)
                Case "color"
                    Lights(y).Color = ParseColor(cad2)
                End Select
                End If
            Loop
        Case "<music>"
            BgMusicRandom = True
            Do While lines(x) <> "</music>"
                x = x + 1
                If lines(x) <> "" And lines(x) <> "</music>" Then
                    If left(lines(x), 1) <> "#" Then
                        y = UBound(BackgroundMusic) + 1
                        ReDim Preserve BackgroundMusic(y)
                        BackgroundMusic(y) = lines(x)
                    ElseIf lines(x) = "#followorder" Then
                        BgMusicRandom = False
                    End If
                End If
            Loop
        Case "<sounds>"
            Do While lines(x) <> "</sounds>"
                x = x + 1
                If lines(x) <> "" And lines(x) <> "</sounds>" Then
                    y = UBound(PendingSounds) + 1
                    ReDim Preserve PendingSounds(y)
                    PendingSounds(y) = lines(x)
                End If
            Loop
        Case Else
            If left(lines(x), Len("<world")) = "<world" Then
                WorldID = Val(mid(lines(x), Len("<world=x")))
            ElseIf left(lines(x), Len("<scene")) = "<scene" Then
                y = UBound(LevelScenes) + 1
                ReDim Preserve LevelScenes(y)
                ReDim LevelScenes(y).Conditions(0)
                ReDim LevelScenes(y).Effects(0)
                If y = 1 Then LevelScenes(y).Enabled = True
                LevelScenes(y).id = Val(mid(lines(x), Len("<scene=x")))
                Do While lines(x) <> "</scene>"
                    x = x + 1
                    cad = lines(x)
                    R1 = InStr(1, cad, "=")
                    cad2 = mid(cad, R1 + 1)
                    If R1 <> 0 Then
                    Select Case LCase(left(cad, R1 - 1))
                    Case "isinarea"
                        z = UBound(LevelScenes(y).Conditions) + 1
                        ReDim Preserve LevelScenes(y).Conditions(z)
                        ParseArea cad2, LevelScenes(y).Conditions(z).varg, LevelScenes(y).Conditions(z).varg2
                        LevelScenes(y).Conditions(z).type = LCase(left(cad, R1 - 1))
                    Case "isintarget", "messagestate", "charvisible", "fadestate"
                        z = UBound(LevelScenes(y).Conditions) + 1
                        ReDim Preserve LevelScenes(y).Conditions(z)
                        LevelScenes(y).Conditions(z).sarg = cad2
                        LevelScenes(y).Conditions(z).type = LCase(left(cad, R1 - 1))
                    Case "disablescene", "enablescene", "setworld", "setcamera", "doors", "setfadetime"
                        z = UBound(LevelScenes(y).Effects) + 1
                        ReDim Preserve LevelScenes(y).Effects(z)
                        LevelScenes(y).Effects(z).larg = Val(cad2)
                        LevelScenes(y).Effects(z).type = LCase(left(cad, R1 - 1))
                    Case "setcharpos"
                        z = UBound(LevelScenes(y).Effects) + 1
                        ReDim Preserve LevelScenes(y).Effects(z)
                        LevelScenes(y).Effects(z).varg = ParsePos(cad2)
                        LevelScenes(y).Effects(z).type = LCase(left(cad, R1 - 1))
                    Case "enableobject", "disableobject", "playvideo", "lockdoor", "unlockdoor", "alldoors", "endstage", "collision"
                        z = UBound(LevelScenes(y).Effects) + 1
                        ReDim Preserve LevelScenes(y).Effects(z)
                        LevelScenes(y).Effects(z).sarg = cad2
                        LevelScenes(y).Effects(z).type = LCase(left(cad, R1 - 1))
                    Case "setfade"
                        z = UBound(LevelScenes(y).Effects) + 1
                        ReDim Preserve LevelScenes(y).Effects(z)
                        LevelScenes(y).Effects(z).sarg = cad2
                        LevelScenes(y).Effects(z).type = LCase(left(cad, R1 - 1))
                    Case "showmessage"
                        z = UBound(LevelScenes(y).Effects) + 1
                        ReDim Preserve LevelScenes(y).Effects(z)
                        LevelScenes(y).Effects(z).larg = Val(cad2)
                        LevelScenes(y).Effects(z).larg2 = Val(mid(cad2, InStr(1, cad2, ",") + 1))
                        LevelScenes(y).Effects(z).sarg = mid(cad2, InStr(InStr(1, cad2, ",") + 1, cad2, ",") + 1)
                        LevelScenes(y).Effects(z).type = LCase(left(cad, R1 - 1))
                    End Select
                    End If
                Loop
            End If
        End Select
    End If
    LoadingStageScreen barmin + x * (barmax - barmin) / UBound(lines)
    x = x + 1
Loop

Set tempfire = Nothing
'Call LoadPendingTexturesSec(0, 100)
End Sub

Public Sub TargetFromString(ByVal cad As String, ByVal World As Long)
'sintaxi: id, alçada, radi, posició
Dim id As String, pos As D3DVECTOR, height As Single, radius As Single
Dim R1 As Long, R2 As Long
R1 = InStr(1, cad, ",")
id = Trim(left(cad, R1 - 1))
R2 = InStr(R1 + 1, cad, ",")
height = Val(mid(cad, R1 + 1, R2 - R1 - 1))
R1 = InStr(R2 + 1, cad, ",")
radius = Val(mid(cad, R2 + 1, R1 - R2 - 1))
pos = ParsePos(mid(cad, R1 + 1))

AddMissionTarget id, pos, height, radius, World
End Sub

Public Function ParsePos(ByVal cad As String) As D3DVECTOR
Dim R1 As Long, R2 As Long
R1 = InStr(1, cad, "{")
R2 = InStr(1, cad, ",")
ParsePos.x = Val(Trim(mid(cad, R1 + 1, R2 - R1 - 1)))
R1 = InStr(1, cad, ",")
R2 = InStr(R1 + 1, cad, ",")
ParsePos.y = Val(Trim(mid(cad, R1 + 1, R2 - R1 - 1)))
ParsePos.z = Val(Trim(mid(cad, R2 + 1)))
End Function

Public Sub ParseArea(ByVal cad As String, v1 As D3DVECTOR, v2 As D3DVECTOR)
Dim R1 As Long
R1 = InStr(InStr(1, cad, "{") + 1, cad, "{")
v1 = ParsePos(left(cad, R1))
v2 = ParsePos(mid(cad, R1))
End Sub

Public Function ParseColor(ByVal cad As String) As D3DCOLORVALUE
Dim R1 As Long, R2 As Long
R1 = InStr(1, cad, "(")
R2 = InStr(1, cad, ",")
ParseColor.r = Val(Trim(mid(cad, R1 + 1, R2 - R1 - 1)))
R1 = InStr(1, cad, ",")
R2 = InStr(R1 + 1, cad, ",")
ParseColor.g = Val(Trim(mid(cad, R1 + 1, R2 - R1 - 1)))
ParseColor.b = Val(Trim(mid(cad, R2 + 1)))
End Function

Public Sub AddPointToObject(object As clsDynObj, ByVal st As String)
Dim l As Long
Dim speed As Long, sleept As Long, pos As D3DVECTOR
l = InStr(1, st, "}")
pos = ParsePos(left(st, l))
l = InStr(l, st, ",")
speed = Val(Trim(mid(st, l + 1)))
l = InStr(l + 1, st, ",")
sleept = Val(Trim(mid(st, l + 1)))
If Not (object Is Nothing) Then object.AddPoint pos, speed, sleept
End Sub

Public Sub ShowLoadPercentStage(ByVal Percent As String)
Dim myPercent As String
myPercent = Percent
If Len(myPercent) < 3 Then myPercent = Space(3 - Len(myPercent)) & myPercent
Device.SetRenderState D3DRS_ZENABLE, False
Device.SetRenderState D3DRS_LIGHTING, False
Device.SetVertexShader myVertexFVF
Device.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
Device.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

Dim width As Single, height As Single
Dim top As Single, left As Single
Dim tu As Single, tu2 As Single
Dim bottom As Single, right As Single
Dim char As Integer
width = 290 * D3DM.width / 800
height = 42 * D3DM.height / 600

LVertices(0) = AssignMV((D3DM.width - width) / 2, (D3DM.height - height) / 2, 0, 0, 0)     'arriba izq
LVertices(1) = AssignMV((D3DM.width + width) / 2, (D3DM.height - height) / 2, 0, 1, 0)     'arriba der
LVertices(2) = AssignMV((D3DM.width - width) / 2, (D3DM.height + height) / 2, 0, 0, 1)     'abajo izq
LVertices(3) = AssignMV((D3DM.width + width) / 2, (D3DM.height + height) / 2, 0, 1, 1)     'abajo der

width = 30 * D3DM.width / 800
height = 30 * D3DM.height / 600
top = D3DM.height - height * 2
left = D3DM.width - width * 5
bottom = D3DM.height - height
right = D3DM.width - width

Dim f As Integer
For f = 1 To 3
    char = Val(mid(myPercent, f, 1))
    If char = 0 Then char = 10
    tu = ((char - 1) * 30) / 330
    tu2 = (char * 30) / 330
    LVertices(f * 4) = AssignMV(left + (30 * D3DM.width / 800) * (f - 1), top, 0, tu, 0)
    LVertices(f * 4 + 1) = AssignMV(left + (30 * D3DM.width / 800) * f, top, 0, tu2, 0)
    LVertices(f * 4 + 2) = AssignMV(left + (30 * D3DM.width / 800) * (f - 1), bottom, 0, tu, 1)
    LVertices(f * 4 + 3) = AssignMV(left + (30 * D3DM.width / 800) * f, bottom, 0, tu2, 1)
Next

tu = ((10) * 30) / 330
LVertices(16) = AssignMV(left + (30 * D3DM.width / 800) * 3, top, 0, tu, 0)
LVertices(17) = AssignMV(right, top, 0, 1, 0)
LVertices(18) = AssignMV(left + (30 * D3DM.width / 800) * 3, bottom, 0, tu, 1)
LVertices(19) = AssignMV(right, bottom, 0, 1, 1)

DoEvents
Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
Device.BeginScene

Device.SetTexture 0, LTextMain
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LVertices(0), Len(LVertices(0))

Device.SetTexture 0, LNum
If mid(myPercent, 1, 1) <> " " Then Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LVertices(4), Len(LVertices(0))
If mid(myPercent, 2, 1) <> " " Then Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LVertices(8), Len(LVertices(0))
If mid(myPercent, 3, 1) <> " " Then Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LVertices(12), Len(LVertices(0))
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LVertices(16), Len(LVertices(0))

Device.EndScene
Device.Present ByVal 0, ByVal 0, 0, ByVal 0
DoEvents
End Sub

Public Sub SetCorrectRenderStates(ByVal AmbientRGB As Long)
'---------- RENDER STATES BEFORE RENDERING ---------
Device.SetRenderState D3DRS_LIGHTING, 1
Device.SetRenderState D3DRS_ZWRITEENABLE, 1
Device.SetRenderState D3DRS_ZENABLE, 1
Device.SetRenderState D3DRS_ALPHABLENDENABLE, 0
Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

Device.SetRenderState D3DRS_AMBIENT, AmbientRGB

Device.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
Device.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

If caps.TextureFilterCaps And D3DPTFILTERCAPS_MAGFGAUSSIANCUBIC = D3DPTFILTERCAPS_MAGFGAUSSIANCUBIC Then
    Device.SetTextureStageState 1, D3DTSS_MAGFILTER, D3DTEXF_GAUSSIANCUBIC
Else
    Device.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC
End If
Device.SetTextureStageState 1, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

Device.SetTextureStageState 0, D3DTSS_MIPFILTER, D3DTEXF_LINEAR

'Device.SetTextureStageState 0, D3DTSS_MIPMAPLODBIAS, -1

Device.SetRenderState D3DRS_FOGENABLE, True
Device.SetRenderState D3DRS_FOGCOLOR, RGB(200, 200, 200)

Device.SetRenderState D3DRS_FOGTABLEMODE, D3DFOG_LINEAR
Device.SetRenderState D3DRS_FOGSTART, ToLong(FarViewPlane * 0.9)
Device.SetRenderState D3DRS_FOGEND, ToLong(FarViewPlane)

End Sub

Public Sub LoadingStageScreen2(ByVal LoadValue As Integer)
If LoadValue > 0 Then LoadValue = 100
Dim width As Single, height As Single, top As Single, left As Single
Dim tu As Single, tv As Single
width = 256 * D3DM.width / 800
height = 50 * D3DM.height / 600
tv = Int((LoadValue / 100) * 8)
If tv > 4 Then
    tv = tv - 4
    tu = 0.5
Else
    tu = 0
End If
tv = tv * 16 / 64

Device.SetRenderState D3DRS_ZENABLE, False
Device.SetRenderState D3DRS_LIGHTING, False
Device.SetVertexShader myVertexFVF
Device.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
Device.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

DoEvents
Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
Device.BeginScene

top = D3DM.height - height * 1.75
left = (D3DM.width - width) / 2
LSVert(0) = AssignMV(left, top, 0, tu, tv)
LSVert(1) = AssignMV(left + width, top, 0, tu + 0.5, tv)
LSVert(2) = AssignMV(left, top + height, 0, tu, tv + 16 / 64)
LSVert(3) = AssignMV(left + width, top + height, 0, tu + 0.5, tv + 16 / 64)

Device.SetTexture 0, LSBar
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LSVert(0), Len(LSVert(0))

Device.EndScene
Device.Present ByVal 0, ByVal 0, 0, ByVal 0
DoEvents
End Sub

Public Sub LoadingStageScreen(ByVal LoadValue As Integer)
If LoadValue > 100 Then LoadValue = 100
If LoadValue < 0 Then LoadValue = 0

Dim width As Single, height As Single, top As Single, left As Single
Dim tu As Single, tv As Single

Device.SetRenderState D3DRS_ZENABLE, False
Device.SetRenderState D3DRS_LIGHTING, False
Device.SetVertexShader myVertexFVF
Device.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_ANISOTROPIC
Device.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_ANISOTROPIC

DoEvents
Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
Device.BeginScene

top = D3DM.height - (61 * D3DM.height / 600) * 2
left = (D3DM.width - 507 * D3DM.width / 800) / 2
width = 507 * D3DM.width / 800
height = 61 * D3DM.height / 600
LSVert(0) = AssignMV(left, top, 0, 0, 0)
LSVert(1) = AssignMV(left + width, top, 0, 1, 0)
LSVert(2) = AssignMV(left, top + height, 0, 0, 1)
LSVert(3) = AssignMV(left + width, top + height, 0, 1, 1)
Device.SetTexture 0, LSBar
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LSVert(0), Len(LSVert(0))

top = D3DM.height - (61 * D3DM.height / 600) * 2 + 20 * D3DM.height / 600
left = (D3DM.width - 507 * D3DM.width / 800) / 2 + 26 * D3DM.width / 800
width = (455 * D3DM.width / 800) * LoadValue / 100
height = 22 * D3DM.height / 600
tu = LoadValue / 100
LSVert(0) = AssignMV(left, top, 0, 0, 0)
LSVert(1) = AssignMV(left + width, top, 0, tu, 0)
LSVert(2) = AssignMV(left, top + height, 0, 0, 1)
LSVert(3) = AssignMV(left + width, top + height, 0, tu, 1)
Device.SetTexture 0, LSBar2
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, LSVert(0), Len(LSVert(0))

Device.EndScene
Device.Present ByVal 0, ByVal 0, 0, ByVal 0
DoEvents
End Sub

Public Sub DestroyThisScene()
Dim x As Long, y As Long
Static Init As Boolean
On Local Error Resume Next

If Init = True Then
    For x = 0 To UBound(FOComplex)
        Set FOComplex(x).Mesh = Nothing
        Set FOComplex(x).tmp_adj = Nothing
        Set FOComplex(x).tmp_mat = Nothing
    Next
    For x = 0 To UBound(FOSimple)
        Set FOSimple(x).Mesh = Nothing
        Set FOSimple(x).tmp_adj = Nothing
        Set FOSimple(x).tmp_mat = Nothing
    Next
    For x = 0 To UBound(MovingCharacters)
        Set MovingCharacters(x) = Nothing
    Next
    For x = 0 To UBound(GameTexturesSec)
        Set GameTexturesSec(x).texture = Nothing
    Next
    For x = 0 To UBound(GameTexturesTer)
        Set GameTexturesTer(x).texture = Nothing
    Next
    For x = 0 To UBound(Fires)
        Set Fires(x) = Nothing
    Next
    For x = 1 To UBound(Cal3DModelArray)
        Set Cal3DModelArray(x) = Nothing
    Next
End If
ReDim Cal3DModelArray(0)
ReDim FOSimple(0)
ReDim FOComplex(0)
ReDim FixedObjects(0)
ReDim MovingCharacters(0)
ReDim GameTexturesSec(0)
ReDim GameTexturesTer(0)
ReDim LockedDoors(0)
ReDim Fires(0)

ReDim LevelScenes(0)
ReDim MissionTargets(0)
ReDim SavePoints(0)
ReDim PendingSounds(0)
ReDim BackgroundMusic(0)
ReDim Lights(0)

For x = 1 To UBound(RenderModelsLM)
    Set RenderModelsLM(x) = Nothing
Next

Call ClearSecondaryCollisionArray
Call SoundEngine.DestroySounds
ReDim DS_Sounds(0)
On Local Error GoTo 0
Init = True
End Sub

Public Sub FillDoorsArray()
ReDim LockedDoors(8)
LockedDoors(1) = "circ_pret_up"
LockedDoors(2) = "circ_pret_down"
LockedDoors(3) = "circ_tavernae"
LockedDoors(4) = "forum_pret"
LockedDoors(5) = "forum_circ_left"
LockedDoors(6) = "pret_circ_down"
LockedDoors(7) = "pret_circ_up"
LockedDoors(8) = "pret_forum"
End Sub

Public Sub ResetThisStage()
Call ResetCurrentLevel

RestartLoop = True
End Sub
