VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDynObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WorldID As Long
Public AutoY As Boolean
Public CoordY As Single
Public Visible As Boolean
Public id As String
Public StopRadius As Single
Public VisionRadius As Single
Public FOView As Single
Public Paused  As Boolean
Public dimensions As Single

Private ModelID As Long
Private CalModel As Cal3DModel
Private QAnims() As String

'Private Const OrientationOffset = 0.5
Private Type point
    point As D3DVECTOR
    speed As Single
    mSleep As Single
    StopSeconds As Single
End Type

Private Points() As point
Private NumPoints As Long

Private Sounds() As String
Private CurrentSound As String

Public Enum animtype
    StartMoving
    EndMoving
    Moving
    Quiet
    QuietWithEffects
End Enum

Private Enum AnimEstatusEnum
    CharQuiet
    CharMoving
    StartingWalking
    EndWalking
End Enum

Private Orientation As D3DVECTOR
Private Position As D3DVECTOR
Private LastPosition As D3DVECTOR
Private MatrixTransform As D3DMATRIX
Private RotationY As D3DMATRIX
Private Translation As D3DMATRIX
Private LastPoint As Long
Private NextPoint As Long
Private LimPoint As Long

Private NoUpdate As Boolean

Private AnimStatus As AnimEstatusEnum
Private PointTimer As Double
Private SleepingTime As Double
Private PauseTime As Double
Private QuietFXTimer As Double
Private FreezeTime As Double

Private CRadius As Single
Private CCenter As D3DVECTOR
Private Do2ndRender As Boolean

Public Sub PauseAnim()
Paused = True
PauseTime = GetTickCount()
AnimStatus = CharQuiet
End Sub

Public Sub ResumeAnim()
PauseTime = GetTickCount() - PauseTime
PointTimer = PointTimer + PauseTime
If QuietFXTimer <> 0 Then QuietFXTimer = QuietFXTimer + PauseTime
If SleepingTime <> 0 Then SleepingTime = SleepingTime + PauseTime
PauseTime = 0
AnimStatus = StartingWalking
Paused = False
End Sub

Public Sub FreezeTimer()
FreezeTime = GetTickCount()
End Sub

Public Sub UnFreezeTimer()
If FreezeTime = 0 Then Exit Sub     'no need to unfreeze
FreezeTime = GetTickCount() - FreezeTime

If PauseTime <> 0 Then PauseTime = PauseTime + FreezeTime
If SleepingTime <> 0 Then SleepingTime = SleepingTime + FreezeTime
If PointTimer <> 0 Then PointTimer = PointTimer + FreezeTime
FreezeTime = 0
End Sub

Public Sub SetPos(ByVal NumPoint As Long)
Position = Points(NumPoint).point
LastPoint = NumPoint

NextPoint = NumPoint + 1
If NextPoint > NumPoints - 1 Then NextPoint = 0
    
PointTimer = 0
NoUpdate = False
AnimStatus = StartingWalking
End Sub

Public Sub SetLimit(ByVal NumPoint As Long)
LimPoint = NumPoint
PointTimer = 0
NoUpdate = False
End Sub

Public Sub AddPoint(ByRef v As D3DVECTOR, ByVal SpeedAtThisPoint As Single, Optional ByVal SleepTime As Single)
ReDim Preserve Points(NumPoints)
Points(NumPoints).point = v
Points(NumPoints).speed = SpeedAtThisPoint
Points(NumPoints).mSleep = SleepTime
NumPoints = NumPoints + 1
End Sub

Private Function CheckPointInTrajectory(point As D3DVECTOR, ByVal radius As Single) As Boolean
Dim plane As D3DPLANE
If VecDistFast(point, Position) < radius ^ 2 Then
    'we are in the affect radius, check if we ara in front of the char or back
    D3DXPlaneFromPointNormal plane, Position, Orientation
    If D3DXPlaneDotCoord(plane, point) > 0 Then
        CheckPointInTrajectory = True
    Else
        CheckPointInTrajectory = False
    End If
End If
End Function

Public Sub RenderSound()
Dim RanNum As Long
If UBound(Sounds) = 0 Then Exit Sub
If AnimStatus = CharQuiet Then
    'stoped: check point out
    If CheckPointInTrajectory(CharPos, StopRadius * 2) = False Then
        'continue anim
        Me.ResumeAnim
        Exit Sub
    End If
Else
    'walking
    If CheckPointInTrajectory(CharPos, StopRadius) Then
        'stop anim
        RanNum = Int(Rnd() * UBound(Sounds)) + 1
        PauseAnim
        SoundEngine.PlaySound Sounds(RanNum), Position, Orientation
        Exit Sub
    End If
End If
End Sub

Public Sub SetSpeak(ByVal soundfile As String)
Dim x As Long
x = UBound(Sounds) + 1
ReDim Preserve Sounds(x)
Sounds(x) = soundfile
End Sub

Private Sub Class_Initialize()
ReDim Points(0)
ReDim Speeds(0)
ReDim ObjQuietEffect(0)
ReDim Sounds(0)
ReDim QAnims(0)
End Sub

Private Sub Class_Terminate()
ReDim Points(0)
ReDim Speeds(0)
ReDim ObjQuietEffect(0)
End Sub

Public Sub AddQuietAnimation(ByVal AnimName As String)
ReDim Preserve QAnims(UBound(QAnims) + 1)
QAnims(UBound(QAnims)) = AnimName
End Sub

Public Sub SetCore(CModel As Cal3DModel)
Set CalModel = CModel
ModelID = CalModel.CreateModel()
End Sub

Public Sub RenderTime(Optional ByVal ForceRender As Boolean = False)
'calculate orientation, position and bounding box for collsion
If Visible = False Then Exit Sub
Dim OrientationOffset As Single

'if this isn't the actual world don't render. In case we force the rendering, continue.
Do2ndRender = False
If ForceRender = False And WorldID <> TheGameSlot.WorldID Then Exit Sub

If PointTimer = 0 Then PointTimer = GetTickCount()
Dim interpolation As Double
Dim desp As Double
Dim prev_prev As D3DVECTOR, next_next As D3DVECTOR, tempvec As D3DVECTOR
Dim tmp_long As Long
Static FirstTime As Boolean

If NoUpdate Then Exit Sub
If PauseTime <> 0 Then Exit Sub

If SleepingTime <> 0 Then
    If GetTickCount() - SleepingTime > Points(LastPoint).mSleep Then
        SleepingTime = 0
        PointTimer = GetTickCount()
        AnimStatus = StartingWalking
    Else
        AnimStatus = CharQuiet
        GoTo MyJump
    End If
End If


    desp = (GetTickCount() - PointTimer) * Points(LastPoint).speed / 1000

    If VecDistFast(Points(LastPoint).point, Points(NextPoint).point) <> 0 Then
        interpolation = desp / VecDist(Points(LastPoint).point, Points(NextPoint).point)
    Else
        interpolation = desp / 0.01
    End If
    If interpolation > 1 Or VecDistFast(Points(LastPoint).point, Points(NextPoint).point) = 0 Then
        'we have reached and passed the next point, change points
        If NextPoint = LimPoint Then
            NoUpdate = True
            AnimStatus = CharQuiet
            Exit Sub
        End If
        LastPoint = NextPoint
        NextPoint = NextPoint + 1
        If NextPoint > NumPoints - 1 Then NextPoint = 0
        PointTimer = 0
        interpolation = 0
        If Points(LastPoint).mSleep <> 0 Then
            SleepingTime = GetTickCount()
            AnimStatus = CharQuiet
            Exit Sub
        End If
    End If
    Position = LinearInterpolation(Points(LastPoint).point, Points(NextPoint).point, interpolation)
    
    '--------- Calculate Orientation of the model --------
    If PointTimer = 0 Then desp = 0

    OrientationOffset = Points(LastPoint).speed / 4
    If VecDistFast(Points(LastPoint).point, Points(NextPoint).point) <> 0 Then
        interpolation = (desp + OrientationOffset) / VecDist(Points(LastPoint).point, Points(NextPoint).point)
    Else
        interpolation = interpolation + 0.001
    End If
    If interpolation > 1 Then
        tmp_long = NextPoint
        Do While interpolation > 1
            If (tmp_long + 1) > (NumPoints - 1) Then
                prev_prev = Points(tmp_long).point
                next_next = Points(0).point
                tmp_long = 0
            Else
                prev_prev = Points(tmp_long).point
                next_next = Points(tmp_long + 1).point
                tmp_long = tmp_long + 1
            End If
            If tmp_long = 0 Then
                tempvec = Points(NumPoints - 2).point
            ElseIf tmp_long = 1 Then
                tempvec = Points(NumPoints - 1).point
            Else
                tempvec = Points(tmp_long - 2).point
            End If
            desp = desp - VecDist(tempvec, prev_prev)
            interpolation = (desp + OrientationOffset) / VecDist(prev_prev, next_next)
        Loop
        Orientation = LinearInterpolation(prev_prev, next_next, interpolation)
        Orientation.x = Orientation.x - Position.x
        Orientation.y = Orientation.y - Position.y
        Orientation.z = Orientation.z - Position.z
    Else
        Orientation = LinearInterpolation(Points(LastPoint).point, Points(NextPoint).point, interpolation)
        Orientation.x = Orientation.x - Points(LastPoint).point.x
        Orientation.y = Orientation.y - Points(LastPoint).point.y
        Orientation.z = Orientation.z - Points(LastPoint).point.z
    End If
    
    '---- Create Translation And RotationY Matrices from position and orientation ------
    '-- Combine them into a transforma Matrix
    
    ' RotationY matrix is static, so if the model is on the next point we can multiply
    ' basing on the last matrix and there's no problem
    If Orientation.x > 0 And Orientation.z > 0 Then
        D3DXMatrixRotationY RotationY, Atn(Orientation.x / Orientation.z) '+ Pi
    ElseIf Orientation.x < 0 And Orientation.z > 0 Then
        D3DXMatrixRotationY RotationY, Atn(Orientation.x / Orientation.z) '+ Pi
    ElseIf Orientation.x > 0 And Orientation.z < 0 Then
        D3DXMatrixRotationY RotationY, Atn(Orientation.x / Orientation.z) + Pi
    ElseIf Orientation.x < 0 And Orientation.z < 0 Then
        D3DXMatrixRotationY RotationY, Atn(Orientation.x / Orientation.z) + Pi
    ElseIf Orientation.z = 0 And Orientation.x > 0 Then
        D3DXMatrixRotationY RotationY, -Pi / 2 + Pi
    ElseIf Orientation.z = 0 And Orientation.x < 0 Then
        D3DXMatrixRotationY RotationY, Pi / 2 + Pi
    ElseIf Orientation.x = 0 And Orientation.z > 0 Then
        D3DXMatrixRotationY RotationY, 0
    ElseIf Orientation.x = 0 And Orientation.z < 0 Then
        D3DXMatrixRotationY RotationY, Pi
    End If
    
MyJump:
    If AutoY = True Then
        NumDynObjPositions = NumDynObjPositions + 1
        ReDim Preserve DynObjPositions(NumDynObjPositions)
        DynObjPositions(NumDynObjPositions).x = Position.x
        DynObjPositions(NumDynObjPositions).z = Position.z
        If FirstTime Then
            'DynObjPositions(NumDynObjPositions).y = DynObjPositions(NumDynObjPositions).y + 2
            DynObjPositions(NumDynObjPositions).y = LastPosition.y
        Else
            DynObjPositions(NumDynObjPositions).y = Position.y
        End If
    End If
    
    Do2ndRender = True
    FirstTime = True
    
    Exit Sub
End Sub

Public Sub RenderTime2()
If Do2ndRender = False Then Exit Sub
If Visible = False Then Exit Sub
Dim RotationX As D3DMATRIX
D3DXMatrixRotationX RotationX, -Pi / 2
D3DXMatrixMultiply MatrixTransform, RotationX, RotationY

If AutoY = False Then
    D3DXMatrixTranslation Translation, Position.x, Position.y, Position.z
Else
    D3DXMatrixTranslation Translation, Position.x, CoordY, Position.z
End If
D3DXMatrixMultiply MatrixTransform, MatrixTransform, Translation
End Sub

Public Sub RenderNow()
'draw the character on the screen
If Visible = False Then Exit Sub

Static QuietFXAnim As String, LastQuietFXAnim As String 'the time with is required to update

Select Case AnimStatus
Case CharMoving
    If Points(NextPoint).mSleep <> 0 Then
        If VecDistFast(Position, Points(NextPoint).point) < (0.2 * Points(LastPoint).speed) ^ 2 Then
            CalModel.BlendCycle ModelID, "quiet", 0.7, 0.3
            CalModel.ClearCycle ModelID, "walking", 0.3
            AnimStatus = CharQuiet
        End If
    End If
    QuietFXAnim = ""
Case CharQuiet
    If UBound(QAnims) <> 0 Then
QuietInit:
        If QuietFXAnim = "" Then
            QuietFXAnim = QAnims(Int(Rnd() * UBound(QAnims)) + 1)
            If VecDist(Points(NextPoint).point, CharPos) / Points(LastPoint).speed > CalModel.AnimLength(QuietFXAnim) / 1000 Then
                CalModel.ExecuteAction ModelID, QuietFXAnim, 1, 0, 0
                QuietFXTimer = GetTickCount() + 1000 + Rnd() * 1500
                GoTo QuietInit
            Else
                QuietFXAnim = ""
            End If
        Else
            If GetTickCount() > (QuietFXTimer + CalModel.AnimLength(QuietFXAnim) + 1000 + Rnd() * 1500) Then
                If UBound(QAnims) = 1 Then
                    QuietFXAnim = QAnims(1)
                Else
                    Do
                        QuietFXAnim = QAnims(Int(Rnd() * UBound(QAnims)) + 1)
                    Loop While QuietFXAnim = LastQuietFXAnim
                    LastQuietFXAnim = QuietFXAnim
                End If
                If VecDist(Points(LastPoint).point, CharPos) / Points(LastPoint).speed > CalModel.AnimLength(QuietFXAnim) Then
                    CalModel.ExecuteAction ModelID, QuietFXAnim, 1, 0, 0
                    QuietFXTimer = GetTickCount()
                    CalModel.ExecuteAction ModelID, QuietFXAnim, 1, 0, 0
                Else
                    QuietFXTimer = GetTickCount()
                End If
            End If
        End If
    End If
Case EndWalking
    
Case StartingWalking
    CalModel.ClearCycle ModelID, "quiet", 0.3
    CalModel.BlendCycle ModelID, "walking", 0.7, 0.3
    AnimStatus = CharMoving
End Select

CalModel.Update ModelID, FrameAverage / 1000
'If CheckSphere(v3(Position.x, Position.y + CalModel.ModelDimensions, Position.z), CalModel.ModelDimensions) Then
    Device.SetTransform D3DTS_WORLD, MatrixTransform
    CalModel.Render ModelID
'End If
LastPosition = Position
End Sub

'-------- PRIVATE AUXILIAR FUNCTIONS --------
'--------------------------------------------

Private Function VecDist(v1 As D3DVECTOR, v2 As D3DVECTOR) As Double
VecDist = Sqr((v1.x - v2.x) ^ 2 + (v1.y - v2.y) ^ 2 + (v1.z - v2.z) ^ 2)
End Function

Private Function LinearInterpolation(v1 As D3DVECTOR, v2 As D3DVECTOR, ByVal interpolation As Double) As D3DVECTOR
LinearInterpolation.x = v1.x + interpolation * (v2.x - v1.x)
LinearInterpolation.y = v1.y + interpolation * (v2.y - v1.y)
LinearInterpolation.z = v1.z + interpolation * (v2.z - v1.z)
End Function

Private Function CheckSphere(center As D3DVECTOR, radius As Single) As Boolean
Dim i As Long
For i = 0 To 5
    If D3DXPlaneDotCoord(FrustumPlanes(i), center) < -radius Then
        CheckSphere = False
        Exit Function
    End If
Next
CheckSphere = True
End Function

Public Property Get Public_Position() As D3DVECTOR
Public_Position = Position
End Property

Public Function CheckCharVisible() As Boolean
Dim vec As D3DVECTOR, dot As Single
vec.x = CharPos.x - Position.x
vec.y = CharPos.y - Position.y
vec.z = CharPos.z - Position.z

dot = D3DXVec3Dot(Normalize(Orientation), Normalize(vec))
If dot > 0 Then
    If (ArcCos(dot) / Pi * 180) < (FOView / 2) Then
        If VecDistFast(Position, CharPos) < VisionRadius ^ 2 Then
            CheckCharVisible = True
        End If
    End If
End If
End Function

