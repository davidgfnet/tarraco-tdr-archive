Attribute VB_Name = "MainModule"
Option Explicit

Sub Main()
frmLoading.Show
'------------ Aplication essentials --------------
EXE = App.path
If right(EXE, 1) <> "\" Then EXE = EXE & "\"
Randomize timer
'------------ Program initial comprobations ----------
Call InitialComprobations
'------------ Init DirectX ------------------
Call InitDx
'------------ Load resolution modes ----------
Call LoadDM
'------------ Create Device -----------------
Unload frmLoading
Call CreateDev
'-- Load Loading screen and percent numbers --
Call LoadPrev
'-------------- Load Menus & Interface -------------
Call LoadMenu
Call LoadUI
'---------------- LoadSettings --------
Call LoadSettings
Call LoadDoors
ShowLoadPercent "20"
'-------------- Load Models & Obj & Col --------
ReDim PendingTextures(0)
ReDim LockedDoors(0)
Call LoadSpecialModels

Call RefreshDrawDepth
Call ComputeCollisionBasic
ShowLoadPercent "25"
ReDim CollisionFloatsAux(4000000)
ReDim FOComplex(0): ReDim FOSimple(0)
'-------------- Load Chars & Anims --------
Call LoadMainChar

'-------------- Start Music ----------
Call InitMusic
ShowLoadPercent "35"
'------------ Load Textures ----------
Call LoadMiniMaps
ShowLoadPercent "38"
Call LoadSkyBox
Call LoadMainTextures(55, 100)
'--------------- Capture Mouse ---------
Call MouseSetUp
'--------------- MenuShow ---------------
'(and destroy load screen as we wont use it any more)
Set LTextMain = Nothing
Set LNum = Nothing

MenuAgain:
Call MenuSystem
'----- After the Menu Start Game or Exit -----
If ExitToWin = False Then
    '--- GET INTO THE GAME LOOP ---
    Select Case GlobalMenuOption
    Case 1  'start new game
        ResetGS TheGameSlot
        ResetGS SavedGameSlot
        Call GameStart    'Start the game, the function will retun when exiting
        Call GameEnd      'The game has exited, delete world objects such as sound
    Case 2  'Load saved game
        
    End Select
End If
'---- End of the game --- Show menu or quit
If ExitToWin = False Then
    GoTo MenuAgain
End If

'--------- SAVE SETTINGS BEFOR EXITING -------
Call SaveSettings

'------- END (al fin) ------
On Local Error Resume Next
MouseDevice.Unacquire
Set MouseDevice = Nothing
Call MusicEngine.DestroyMusic
Set Device = Nothing
Set Direct3D = Nothing
Set Direct3DX = Nothing
Set DirectSound = Nothing
Set DirectInput = Nothing
Set DirectX = Nothing
Call DeleteDir(MusicDir)
Call DeleteTempDir

Unload frmGraphics
End
End Sub

Public Sub InitialComprobations()
Set Reg = New clsWinReg

If FileExists(EXE & "engine.dll") = False Then GoTo Error
If FileExists(EXE & "vorbis.dll") = False Then GoTo Error
If FileExists(EXE & "ogg.dll") = False Then GoTo Error
If FileExists(EXE & "vorbisfile.dll") = False Then GoTo Error
If FileExists(EXE & "freeimage.dll") = False Then GoTo Error

Reg.CreateKey "HKLM\Software\" & RegName & "\"

Exit Sub
Error:
MsgBox "El programa no està correctament instal·lat. Falten un o més arxius", vbCritical, "Error intern"
End
End Sub

Public Sub InitDx()
On Local Error GoTo ErrExit
Dim ret As Long
Set DirectX = New DirectX8
Set Direct3D = DirectX.Direct3DCreate()
Set DirectInput = DirectX.DirectInputCreate()
Set MouseDevice = DirectInput.CreateDevice("guid_SysMouse")
Set Direct3DX = New D3DX8
Direct3D.GetDeviceCaps 0, D3DDEVTYPE_HAL, caps

If D3DSHADER_VERSION_MAJOR(caps.VertexShaderVersion) < 1 Then
    VSCapable = False
    MsgBox "La targeta gràfica no és capaç d'utilitzar VertexShader v1" & vbCrLf & "Això provocarà que el joc vagi una mica més lent del normal", vbExclamation, "Error de Direct3D"
Else
    VSCapable = True
End If

On Local Error Resume Next

Set DSEnum = DirectX.GetDSEnum()
Dim SoundDevice As String, x As Long
SoundDevice = RegRead("Sound_GUID")
For x = 1 To DSEnum.GetCount
    If DSEnum.GetGuid(x) = SoundDevice Then
        GoTo AudioOK
    End If
Next
SoundDevice = ""

AudioOK:
Set DirectSound = DirectX.DirectSoundCreate(SoundDevice)
RegSave "Sound_GUID", SoundDevice
If Err.number <> 0 Then
    'err, try with the default adapter
    Err.Clear: Err.number = 0
    SoundDevice = ""
    Set DirectSound = DirectX.DirectSoundCreate(SoundDevice)
    RegSave "Sound_GUID", SoundDevice
End If
DSDeviceDesc = GetDescFromGUID(SoundDevice)

If Err.number <> 0 Then
    ret = MsgBox("Els controladors de so no estan correctament instalats. Vols continuar sense so?", vbExclamation + vbYesNo, "Error de so")
    If ret = vbYes Then
        DisableSound = True
    Else
        End
    End If
End If
ReDim DS_Sounds_Plain(0)
ReDim DS_Sounds(0)

Exit Sub
ErrExit:
    
MsgBox "El programa no ha pogut iniciar DirectX correctament. Comproba la versió instal·lada. Ha de ser 8.1 o superior", vbCritical, "Error intern"
End
End Sub

Public Sub LoadDM()
Dim x As Long, DModeTemp As D3DDISPLAYMODE
ReDim DModesArray(Direct3D.GetAdapterModeCount(0) - 1)
For x = 0 To Direct3D.GetAdapterModeCount(0) - 1
    Direct3D.EnumAdapterModes 0, x, DModeTemp
    DModesArray(x) = DModeTemp
Next
GetAllResolutions 800, 600, ResolutionArray
End Sub

Public Sub CreateDev()
Dim Param As D3DPRESENT_PARAMETERS
Dim LastDM As String

LastDM = RegRead("Resolution")
AntiAliasLevel = Val(RegRead("antialias"))

If LastDM = "" Then
    D3DM = CreateDM(800, 600, True)
Else
    D3DM = ParseDM(LastDM)
End If

If Direct3D.CheckDeviceMultiSampleType(0, D3DDEVTYPE_HAL, D3DM.Format, 0, AntiAliasLevel) <> D3D_OK Then
    'incorrect multisample!!!
    AntiAliasLevel = 0
End If

SaveResolution D3DM, AntiAliasLevel

frmGraphics.Show

With Param
    '.Windowed = 1
    .SwapEffect = D3DSWAPEFFECT_DISCARD
    .BackBufferFormat = D3DM.Format
    .EnableAutoDepthStencil = 1
    .BackBufferCount = 1
    .BackBufferWidth = D3DM.width
    .BackBufferHeight = D3DM.height
    .AutoDepthStencilFormat = D3DFMT_D24S8
    .MultiSampleType = AntiAliasLevel
End With

On Local Error Resume Next
'------- 24 bit depth + 8 bit stencil  ---- HARDWARE ACCELERATION --------
Err.number = 0
Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Param)
'------- 24 bit depth + 4 bit stencil + 4 unused bits ---- HARDWARE ACCELERATION --------
If Err.number <> 0 Then
    Param.AutoDepthStencilFormat = D3DFMT_D24X4S4
    Err.Clear: Err.number = 0
    Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Param)
End If

'------- 24 bit depth + 8 unused bits ---- HARDWARE ACCELERATION --------
If Err.number <> 0 Then
    Param.AutoDepthStencilFormat = D3DFMT_D24X8
    Err.Clear: Err.number = 0
    Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Param)
End If
'------- 16 bit depth ---- HARDWARE ACCELERATION --------
If Err.number <> 0 Then
    Param.AutoDepthStencilFormat = D3DFMT_D16
    Err.Clear: Err.number = 0
    Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Param)
End If
'------- 24 bit depth ---- SOFTWARE ACCELERATION --------
If Err.number <> 0 Then
    Param.AutoDepthStencilFormat = D3DFMT_D24X8
    Err.Clear: Err.number = 0
    Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Param)
End If
'------- 16 bit depth ---- SOFTWARE ACCELERATION --------
If Err.number <> 0 Then
    Param.AutoDepthStencilFormat = D3DFMT_D16
    Err.Clear: Err.number = 0
    Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Param)
End If
On Local Error GoTo 0

Call CreateVertexSahders
End Sub

Public Sub SaveResolution(DM As D3DDISPLAYMODE, ByVal AntiAlias As Long)
RegSave "Resolution", Format(DM.width, "0000") & Format(DM.height, "0000") & Format(DM.RefreshRate, "0000") & Format(DM.Format, "0000")
RegSave "AntiAlias", Trim(Str(AntiAlias))
End Sub

Public Sub CursorV(ByVal IsVisible As Boolean)
ShowCursor IsVisible
End Sub

Public Sub LoadPrev()
Dim mypath As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "tex_tmp_prev" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "maps\prev.dat", mypath

Set LTextMain = LoadTextureAndReturn(mypath & "loading.bmp")

Set LNum = LoadTextureAndReturn(mypath & "num.bmp")

Set MouseTexture = Direct3DX.CreateTextureFromFileEx(Device, mypath & "cursor.tga", 64, 64, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
CursorW = 40: CursorH = 40

'Set FontTexture = Direct3DX.CreateTextureFromFileEx(Device, mypath & "font.tga", 256, 256, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)

DeleteDir mypath

Call ShowLoadPercent("0")
End Sub

Public Sub ShowLoadPercent(ByVal Percent As String)
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

'------------------------------------------------------------------------------------------
'------- Here I create the menus for the game. All the textures included in ---------------
'------- Menu.dat inside graphics folder. For the events watch MenuSystem   ---------------
'------------------------------------------------------------------------------------------

Public Sub LoadMenu()
Dim mypath As String, x As Integer
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "tex_tmp_prev" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "maps\menu.dat", mypath
ShowLoadPercent "1"

Set AuxiliarMenu = New cGameMenu
AuxiliarMenu.TexPath = mypath
AuxiliarMenu.BackGround = "auxmenu.bmp"
AuxiliarMenu.AddButton 275, 250, 250, 75, "returngame.tga", "returngame_over.tga", "returngame", 0
AuxiliarMenu.AddButton 275, 350, 250, 75, "returnmenu.tga", "returnmenu_over.tga", "returnmenu", 0
AuxiliarMenu.BuildMenu

Set SaveGameMenu = New cGameMenu
SaveGameMenu.TexPath = mypath
SaveGameMenu.BackGround = "savemenu.bmp"
SaveGameMenu.AddButton 66, 202, 202, 102, "gameslot.tga", "gameslot_over.tga", "slot1", 0
SaveGameMenu.AddButton 288, 202, 202, 102, "gameslot.tga", "gameslot_over.tga", "slot2", 0
SaveGameMenu.AddButton 520, 202, 202, 102, "gameslot.tga", "gameslot_over.tga", "slot3", 0
SaveGameMenu.AddButton 66, 328, 202, 102, "gameslot.tga", "gameslot_over.tga", "slot4", 0
SaveGameMenu.AddButton 288, 328, 202, 102, "gameslot.tga", "gameslot_over.tga", "slot5", 0
SaveGameMenu.AddButton 520, 328, 202, 102, "gameslot.tga", "gameslot_over.tga", "slot6", 0
SaveGameMenu.AddLabel "partida 1", 95, 225, 16
SaveGameMenu.AddLabel "partida 2", 317, 225, 16
SaveGameMenu.AddLabel "partida 3", 549, 225, 16
SaveGameMenu.AddLabel "partida 4", 95, 351, 16
SaveGameMenu.AddLabel "partida 5", 317, 351, 16
SaveGameMenu.AddLabel "partida 6", 549, 351, 16

SaveGameMenu.AddLabel "lliure", 85, 260, 15, 12, "label_slot_1"
SaveGameMenu.AddLabel "lliure", 305, 260, 15, 12, "label_slot_2"
SaveGameMenu.AddLabel "lliure", 540, 260, 15, 12, "label_slot_3"
SaveGameMenu.AddLabel "lliure", 85, 386, 15, 12, "label_slot_4"
SaveGameMenu.AddLabel "lliure", 305, 386, 15, 12, "label_slot_5"
SaveGameMenu.AddLabel "lliure", 540, 386, 15, 12, "label_slot_6"

SaveGameMenu.DialogTexturef = "dialog.bmp"
SaveGameMenu.DialogTextureButtonsf = "dialog_buttons.bmp"

SaveGameMenu.AddButton 460, 490, 300, 75, "returngame.tga", "returngame_over.tga", "returngame", 0
SaveGameMenu.BuildMenu
ShowLoadPercent "5"

Set MainMenu = New cGameMenu
MainMenu.TexPath = mypath
MainMenu.BackGround = "mainmenu.bmp"
MainMenu.AddButton 600, 500, 150, 50, "exit.tga", "exit_over.tga", "exitwin", 0
MainMenu.AddButton 500, 150, 200, 60, "newgame.tga", "newgame_over.tga", "newgame", 0
MainMenu.AddButton 500, 250, 200, 60, "loadgame.tga", "loadgame_over.tga", "loadgame", 0
MainMenu.AddButton 500, 350, 200, 60, "options.tga", "options_over.tga", "options", 0
MainMenu.BuildMenu
ShowLoadPercent "7"

Set ConfigMenu = New cGameMenu
ConfigMenu.TexPath = mypath
ConfigMenu.BackGround = "optionsmenu.bmp"
ConfigMenu.AddButton 600, 500, 150, 50, "return.tga", "return_over.tga", "applyconfig", 0
ConfigMenu.AddList 450, 180, 250, 30, 16, 4, "resolution", 1

For x = 1 To UBound(ResolutionArray)
    ConfigMenu.AddElementToList "resolution", ResolutionArray(x)
Next
 
ConfigMenu.AddLabel "Resolucio de pantalla", 100, 187, 16, 15
ConfigMenu.AddLabel "profunditat de color", 100, 227, 16, 15
ConfigMenu.AddLabel "suavitzar imatge", 100, 267, 16, 15
ConfigMenu.AddLabel "dispositiu de so", 100, 307, 16, 15
ConfigMenu.AddLabel "Produnditat de dibuix", 100, 347, 16, 15
ConfigMenu.AddLabel "Velocitat del cursor", 100, 387, 16, 15

ConfigMenu.AddList 450, 220, 250, 30, 16, 4, "bits", 1
ConfigMenu.AddElementToList "bits", "16"
ConfigMenu.AddElementToList "bits", "32"
ConfigMenu.AddList 450, 260, 250, 30, 16, 4, "alias", 1
ConfigMenu.AddElementToList "alias", "no"
For x = 2 To 16
    If Direct3D.CheckDeviceMultiSampleType(0, D3DDEVTYPE_HAL, D3DM.Format, 0, x) = D3D_OK Then ConfigMenu.AddElementToList "alias", "x" & Trim(Str(x))
Next

ConfigMenu.AddList 450, 300, 250, 30, 12, 4, "sound", 1, 12
ConfigMenu.AddElementToList "sound", "automatic"
For x = 1 To DSEnum.GetCount
    ConfigMenu.AddElementToList "sound", DSEnum.GetDescription(x)
Next
ConfigMenu.AddList 450, 340, 250, 30, 16, 4, "zdetail", 1
ConfigMenu.AddElementToList "zdetail", "baixa"
ConfigMenu.AddElementToList "zdetail", "mitjana"
ConfigMenu.AddElementToList "zdetail", "alta"
ConfigMenu.AddList 450, 380, 250, 30, 16, 2, "vcursor", 1
ConfigMenu.AddElementToList "vcursor", "molt baixa"
ConfigMenu.AddElementToList "vcursor", "baixa"
ConfigMenu.AddElementToList "vcursor", "mitjana"
ConfigMenu.AddElementToList "vcursor", "alta"
ConfigMenu.AddElementToList "vcursor", "molt alta"

ConfigMenu.ListTextureOpen = "combo2.tga"
ConfigMenu.ListTextureOpenHover = "combo3.tga"
ConfigMenu.ListTexture = "combo1.tga"

ConfigMenu.BuildMenu
ShowLoadPercent "10"

DeleteDir mypath
End Sub

Public Sub MenuSystem()
Device.SetRenderState D3DRS_ALPHABLENDENABLE, True
Device.SetVertexShader myVertexFVF
Device.SetRenderState D3DRS_ZENABLE, False
Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA

Dim Out As Boolean
Dim MenuIndex As Integer, ActualMenu As cGameMenu, LastMenuIndex As Integer
Dim mTimer As Double
Dim coordX As Single, CoordY As Single
Dim Alpha As Single
Dim DMode As D3DDISPLAYMODE, NewAlias As Long
Dim FirstFade As Boolean
FirstFade = True

Call SoundReady
MusicEngine.PlayMusic MusicDir & "intro.ogg"

coordX = D3DM.width / 2
CoordY = D3DM.height / 2

GlobalMenuOption = 0
Do While Out = False
    Set ActualMenu = MenuFromIndex(MenuIndex)
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
    ActualMenu.ProcessMouseMove coordX, CoordY
    If MouseClick0 = True Then
        ActualMenu.ProcessClick coordX, CoordY
        MouseClick0 = False
    End If
    '--------------- Look for events ---------------
    If ActualMenu.isEvent Then
        Select Case ActualMenu.EventName
        Case "exitwin"
            ExitToWin = True
            Out = True
        Case "newgame"
            Out = True
            GlobalMenuOption = 1
        Case "loadgame"
            Out = True
            GlobalMenuOption = 2
        Case "options"
            If Is32bit(D3DM) Then
                ConfigMenu.SetSelectedListItem "bits", 2
            Else
                ConfigMenu.SetSelectedListItem "bits", 1
            End If

            ConfigMenu.SetSelectedListItemByText "resolution", D3DM.width & " x " & D3DM.height
            ConfigMenu.SetSelectedListItem "zdetail", DrawDepth
            ConfigMenu.SetSelectedListItem "vcursor", CursorSpeed
            If DSDeviceDesc = "" Then
                ConfigMenu.SetSelectedListItem "sound", 1
            Else
                ConfigMenu.SetSelectedListItemByText "sound", DSDeviceDesc
            End If
            If AntiAliasLevel <> 0 Then
                ConfigMenu.SetSelectedListItemByText "alias", "x" & Trim(Str(AntiAliasLevel))
            Else
                ConfigMenu.SetSelectedListItem "alias", 1
            End If
            'ConfigMenu.SetSelectedListItem "shadows", EnableShadows + 1
            MenuIndex = 1
        Case "applyconfig"
            MenuIndex = 0
            DMode = ParseResolution(ConfigMenu.GetListValue("resolution"))
            NewAlias = Val(mid(ConfigMenu.GetListValue("alias"), 2))
            If ConfigMenu.GetListValue("bits") = "32" Then
                DMode = CreateDM(DMode.width, DMode.height, True)
            Else
                DMode = CreateDM(DMode.width, DMode.height, False)
            End If
            If DMode.Format <> D3DM.Format Or _
                DMode.width <> D3DM.width Or _
                DMode.height <> D3DM.height Or _
                NewAlias <> AntiAliasLevel Then
                'the user has changed de DM
                    SaveRS
                    If ResetDevice(DMode, NewAlias) Then
                        D3DM = DMode
                        AntiAliasLevel = NewAlias
                        MainMenu.RefreshDM
                        ConfigMenu.RefreshDM
                    End If
                    SaveResolution D3DM, AntiAliasLevel
                    RestoreRS
            End If
            DrawDepth = ConfigMenu.GetListIndex("zdetail")
            CursorSpeed = ConfigMenu.GetListIndex("vcursor")
            CharDetail = ConfigMenu.GetListIndex("pdetail")
            TexQuality = ConfigMenu.GetListIndex("texq")
            If ConfigMenu.GetListValue("sound") <> DSDeviceDesc Then
                'changed device
                MusicEngine.DestroyMusic
                DSDeviceDesc = ConfigMenu.GetListValue("sound")
                If DSDeviceDesc = "automatic" Then DSDeviceDesc = ""
                Set DirectSound = Nothing
                If DSDeviceDesc = "" Then
                    Set DirectSound = DirectX.DirectSoundCreate("")
                Else
                    Set DirectSound = DirectX.DirectSoundCreate(GetGUIDFromDesc(DSDeviceDesc))
                End If
                Call SoundReady
                MusicEngine.InitMusicBasic
                MusicEngine.PlayMusic MusicDir & "intro.ogg"
                RegSave "Sound_GUID", GetGUIDFromDesc(DSDeviceDesc)
            End If
            'If ShadowsAvaliable Then EnableShadows = ConfigMenu.GetListIndex("shadows") - 1
        End Select
        ActualMenu.isEvent = False
    End If
    Set ActualMenu = MenuFromIndex(MenuIndex)
    
    '----------------------- MAKE A FADE OUT / IN BETWEEN MENUS --------------------
    If LastMenuIndex <> MenuIndex Or (Out = True And GlobalMenuOption <> 0) Then
        Device.SetVertexShader myVertexAlphaFVF
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
            
            Device.SetVertexShader myVertexFVF
            MenuFromIndex(LastMenuIndex).RenderMenu
            
            Device.SetTexture 0, Nothing
            Device.SetVertexShader myVertexAlphaFVF
            Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, fadeVertices(0), Len(fadeVertices(0))
            
            Device.EndScene
            Device.Present ByVal 0, ByVal 0, 0, ByVal 0
            MusicEngine.RenderTime
            DoEvents
        Loop
    End If
    If LastMenuIndex <> MenuIndex Or FirstFade Then
        FirstFade = False
        mTimer = GetTickCount()
        Do While (mTimer + 1000) > GetTickCount()
            Alpha = 255 - (255 * (GetTickCount() - mTimer) / 1000)
            If Alpha > 255 Then Alpha = 255
            If Alpha < 0 Then Alpha = 0
            fadeVertices(0) = AssignMVA(0, 0, 0, D3DColorARGB(Alpha, 0, 0, 0))
            fadeVertices(1) = AssignMVA(D3DM.width, 0, 0, D3DColorARGB(Alpha, 0, 0, 0))
            fadeVertices(2) = AssignMVA(0, D3DM.height, 0, D3DColorARGB(Alpha, 0, 0, 0))
            fadeVertices(3) = AssignMVA(D3DM.width, D3DM.height, 0, D3DColorARGB(Alpha, 0, 0, 0))
            
            Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
            Device.BeginScene
            
            Device.SetVertexShader myVertexFVF
            MenuFromIndex(MenuIndex).RenderMenu
            
            Device.SetTexture 0, Nothing
            Device.SetVertexShader myVertexAlphaFVF
            Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, fadeVertices(0), Len(fadeVertices(0))
            
            Device.EndScene
            Device.Present ByVal 0, ByVal 0, 0, ByVal 0
            MusicEngine.RenderTime
            DoEvents
        Loop
        LastMenuIndex = MenuIndex
        Device.SetVertexShader myVertexFVF
        MouseClick0 = False: MouseClick1 = False
        MouseX = 0: MouseY = 0
    End If
    '----------------------------------------------------------------------------------------------
    
    If GlobalMenuOption <> 0 Then
        MusicEngine.StopMusic
        Exit Sub
    End If

    Device.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
    Device.BeginScene
    
    '------------ RENDER THE MENU --------
    ActualMenu.RenderMenu           'the class does all the work
    
    '----------- RENDER THE CURSOR OVER ALL -----------
    Device.SetTexture 0, MouseTexture
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, MouseVerts(0), Len(MouseVerts(0))
    
    Device.EndScene
    Device.Present ByVal 0, ByVal 0, 0, ByVal 0
    MusicEngine.RenderTime
    
    Sleep 5
    DoEvents
Loop

MusicEngine.StopMusic
End Sub


Public Function MenuFromIndex(ByVal index As Integer) As cGameMenu
Select Case index
Case 0
    Set MenuFromIndex = MainMenu
Case 1
    Set MenuFromIndex = ConfigMenu
End Select
End Function

Public Function CreateDM(ByVal width As Single, ByVal height As Single, ByVal bit32 As Boolean) As D3DDISPLAYMODE
Dim x As Integer
CreateDM.width = width
CreateDM.height = height
CreateDM.RefreshRate = 0
If bit32 Then
    'chek for 32 bits format in fullscreen  (D3DFMT_X8R8G8B8)
    For x = 0 To UBound(DModesArray)
        If DModesArray(x).width = width And DModesArray(x).height = height Then
        If (DModesArray(x).Format And D3DFMT_X8R8G8B8) = D3DFMT_X8R8G8B8 Then
            CreateDM.Format = D3DFMT_X8R8G8B8
            Exit Function
        End If
        End If
    Next
    'if we arrive here it means that there isnt any DMode in 32 bits
    'so we have to try with 16 bits.....
    GoTo Try32Bits
Else
Try32Bits:
    'check for 16 bits format in fullscreen  (D3DFMT_R5G6B5)
    For x = 0 To UBound(DModesArray)
        If DModesArray(x).width = width And DModesArray(x).height = height Then
        If (DModesArray(x).Format And D3DFMT_R5G6B5) = D3DFMT_R5G6B5 Then
            CreateDM.Format = D3DFMT_R5G6B5
            Exit Function
        End If
        End If
    Next
    'if not, check for another possible format   ('chek for 16 bits format in fullscreen  (D3DFMT_X1R5G5B5)
    For x = 0 To UBound(DModesArray)
        If DModesArray(x).width = width And DModesArray(x).height = height Then
        If (DModesArray(x).Format And D3DFMT_X1R5G5B5) = D3DFMT_X1R5G5B5 Then
            CreateDM.Format = D3DFMT_X1R5G5B5
            Exit Function
        End If
        End If
    Next
End If
End Function

Public Function ResetDevice(DMode As D3DDISPLAYMODE, ByVal AntiAlias As Long) As Boolean
Dim Param As D3DPRESENT_PARAMETERS, OriginalDM As D3DDISPLAYMODE, failedonce As Boolean
Dim TestDM As D3DDISPLAYMODE
Dim Create As Boolean

If Device Is Nothing Then
    Create = True
End If

If Not Create Then Device.GetDisplayMode OriginalDM
TestDM = DMode
With Param
    '.Windowed = 1
    .SwapEffect = D3DSWAPEFFECT_DISCARD
    .BackBufferFormat = TestDM.Format
    .EnableAutoDepthStencil = 1
    .BackBufferCount = 1
    .BackBufferWidth = TestDM.width
    .BackBufferHeight = TestDM.height
    .AutoDepthStencilFormat = D3DFMT_D24S8
    .MultiSampleType = AntiAlias
End With

On Local Error Resume Next
'------- 24 bit depth + 8 bit stencil ---- HARDWARE ACCELERATION --------
Err.number = 0
If Create Then
    Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Param)
Else
    Device.Reset Param
End If
'------- 24 bit depth + 4 bit stencil ---- HARDWARE ACCELERATION --------
If Err.number <> 0 Then
    Param.AutoDepthStencilFormat = D3DFMT_D24X4S4
    Err.Clear: Err.number = 0
    If Create Then
        Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Param)
    Else
        Device.Reset Param
    End If
End If

'------- 24 bit depth + 8 unused bits ---- HARDWARE ACCELERATION --------
If Err.number <> 0 Then
    Param.AutoDepthStencilFormat = D3DFMT_D24X8
    Err.Clear: Err.number = 0
    If Create Then
        Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Param)
    Else
        Device.Reset Param
    End If
End If
'------- 16 bit depth  ---- HARDWARE ACCELERATION --------
If Err.number <> 0 Then
    Param.AutoDepthStencilFormat = D3DFMT_D16
    Err.Clear: Err.number = 0
    If Create Then
        Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Param)
    Else
        Device.Reset Param
    End If
End If

'------- 24 bit depth ---- SOFTWARE ACCELERATION --------
If Err.number <> 0 Then
    Param.AutoDepthStencilFormat = D3DFMT_D24X8
    Err.Clear: Err.number = 0
    If Create Then
        Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Param)
    Else
        Device.Reset Param
    End If
End If
'------- 16 bit depth ---- SOFTWARE ACCELERATION --------
If Err.number <> 0 Then
    Param.AutoDepthStencilFormat = D3DFMT_D16
    Err.Clear: Err.number = 0
    If Create Then
        Set Device = Direct3D.CreateDevice(0, D3DDEVTYPE_HAL, frmGraphics.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Param)
    Else
        Device.Reset Param
    End If
End If
If Err.number <> 0 And failedonce = False Then
    failedonce = True
    TestDM = OriginalDM
ElseIf failedonce = True And Err.number <> 0 Then
    ResetDevice = False
    Exit Function
End If
ResetDevice = True
On Local Error GoTo 0
Call CreateVertexSahders
End Function

Public Function TestStencilCaps() As Boolean
If caps.StencilCaps & D3DSTENCILOP_KEEP Then
If caps.StencilCaps & D3DSTENCILOP_ZERO Then
If caps.StencilCaps & D3DSTENCILOP_REPLACE Then
If caps.StencilCaps & D3DSTENCILOP_INCR Then
If caps.StencilCaps & D3DSTENCILOP_DECR Then
    TestStencilCaps = True
    Exit Function
End If
End If
End If
End If
End If
TestStencilCaps = False
End Function

Public Sub GameStart()
'------ Start the Game --------
'-- 1. Load the current sounds / animations / ambients and other stuff for the current part
'-- 2. Jump to GameLoop Module
'------------------------------------------------------------------------------------------------

Call GameLoadLevel
Call GameLoop.GameLoop

End Sub

Public Sub GameEnd()
Call DestroyStandardSound
End Sub

Public Sub LoadWorldModels(ByVal WorldID As Integer, ByVal key As String, Optional ByVal MinBar As Integer, Optional ByVal MaxBar As Integer)
Dim mypath As String, x As Long
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "models_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "graphics\" & Trim(Str(WorldID)) & ".dat", mypath

Select Case WorldID
Case 1
    For x = 1 To 8
        Set RenderModelsLM(x) = New cAdvMesh
        RenderModelsLM(x).LoadFromFile mypath & key & "_" & Trim(Str(x)) & ".x", mypath & key & "_" & Trim(Str(x)) & "l.x", mypath, False
        If MinBar <> 0 Then LoadingStageScreen MinBar + (MaxBar - MinBar) * x / 4
    Next
    RenderModelsLM(8).AlphaTest = True
Case 2
    For x = 1 To 2
        Set RenderModelsLM(10 + x) = New cAdvMesh
        RenderModelsLM(10 + x).LoadFromFile mypath & key & "_" & Trim(Str(x)) & ".x", mypath & key & "_" & Trim(Str(x)) & "l.x", mypath, False
        If MinBar <> 0 Then LoadingStageScreen MinBar + (MaxBar - MinBar) * x / 2
    Next
    RenderModelsLM(12).AlphaTest = True
Case 3
    For x = 1 To 2
        Set RenderModelsLM(20 + x) = New cAdvMesh
        RenderModelsLM(20 + x).LoadFromFile mypath & key & "_" & Trim(Str(x)) & ".x", mypath & key & "_" & Trim(Str(x)) & "l.x", mypath, False
        If MinBar <> 0 Then LoadingStageScreen MinBar + (MaxBar - MinBar) * x / 2
    Next
End Select

DeleteDir mypath
End Sub

Public Sub ComputeCollisionBasic()
Dim mypath As String, x As Long, y As Long, ignore As Long
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "geo_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "geometry\main.dat", mypath

Dim CMesh As D3DXMesh, hresult As Long, vertices() As D3DVERTEX, Desc As D3DINDEXBUFFER_DESC
Dim IBuf As Direct3DIndexBuffer8, indices() As Long

For x = 1 To 3
    Set CMesh = Direct3DX.LoadMeshFromX(mypath & "collision_" & Trim(Str(x)) & ".x", D3DXMESH_MANAGED + D3DXMESH_32BIT, Device, Nothing, Nothing, ignore)
    
    ReDim vertices(CMesh.GetNumVertices)
    hresult = D3DXMeshVertexBuffer8GetData(CMesh, 0, Len(vertices(0)) * CMesh.GetNumVertices, 0, vertices(0))
    
    Set IBuf = CMesh.GetIndexBuffer()
    IBuf.Lock 0, 0, ignore, 16
    IBuf.GetDesc Desc
    IBuf.Unlock
    ReDim indices(Desc.size / 4)   '4 as we use 32 bit mesh , 2 if we use 16 bit
    
    D3DXMeshIndexBuffer8GetData CMesh, 0, Desc.size, 0, indices(0)
    
    ReDim CollisionFloats(x).vertices(CMesh.GetNumFaces * 3 - 1)
    CollisionFloats(x).numverts = CMesh.GetNumFaces * 3

    For y = 0 To CMesh.GetNumFaces * 3 - 1 Step 3
        CollisionFloats(x).vertices(y).x = vertices(indices(y)).x
        CollisionFloats(x).vertices(y).y = vertices(indices(y)).y
        CollisionFloats(x).vertices(y).z = vertices(indices(y)).z
        CollisionFloats(x).vertices(y + 1).x = vertices(indices(y + 1)).x
        CollisionFloats(x).vertices(y + 1).y = vertices(indices(y + 1)).y
        CollisionFloats(x).vertices(y + 1).z = vertices(indices(y + 1)).z
        CollisionFloats(x).vertices(y + 2).x = vertices(indices(y + 2)).x
        CollisionFloats(x).vertices(y + 2).y = vertices(indices(y + 2)).y
        CollisionFloats(x).vertices(y + 2).z = vertices(indices(y + 2)).z
    Next
    
    ReDim indices(0)
    ReDim vertices(0)           'destroy all temp objects!!
    Set IBuf = Nothing
    Set CMesh = Nothing
Next
End Sub

Public Sub ComputeWorldCollision(ByVal WorldID As Integer)
Dim mypath As String, x As Long, y As Long, ignore As Long
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "geo_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "geometry\secondary.dat", mypath

Dim CMesh As D3DXMesh, hresult As Long, vertices() As D3DVERTEX, Desc As D3DINDEXBUFFER_DESC
Dim IBuf As Direct3DIndexBuffer8, indices() As Long

Set CMesh = Direct3DX.LoadMeshFromX(mypath & "collision_" & Trim(Str(WorldID)) & ".x", D3DXMESH_MANAGED + D3DXMESH_32BIT, Device, Nothing, Nothing, ignore)
    
ReDim vertices(CMesh.GetNumVertices)
hresult = D3DXMeshVertexBuffer8GetData(CMesh, 0, Len(vertices(0)) * CMesh.GetNumVertices, 0, vertices(0))
    
Set IBuf = CMesh.GetIndexBuffer()
IBuf.Lock 0, 0, ignore, 16
IBuf.GetDesc Desc
IBuf.Unlock
ReDim indices(Desc.size / 4)   '4 as we use 32 bit mesh , 2 if we use 16 bit

D3DXMeshIndexBuffer8GetData CMesh, 0, Desc.size, 0, indices(0)

ReDim CollisionFloats(WorldID).vertices(CMesh.GetNumFaces * 3 - 1)
CollisionFloats(WorldID).numverts = CMesh.GetNumFaces * 3

For y = 0 To CMesh.GetNumFaces * 3 - 1 Step 3
    CollisionFloats(WorldID).vertices(y).x = vertices(indices(y)).x
    CollisionFloats(WorldID).vertices(y).y = vertices(indices(y)).y
    CollisionFloats(WorldID).vertices(y).z = vertices(indices(y)).z
    CollisionFloats(WorldID).vertices(y + 1).x = vertices(indices(y + 1)).x
    CollisionFloats(WorldID).vertices(y + 1).y = vertices(indices(y + 1)).y
    CollisionFloats(WorldID).vertices(y + 1).z = vertices(indices(y + 1)).z
    CollisionFloats(WorldID).vertices(y + 2).x = vertices(indices(y + 2)).x
    CollisionFloats(WorldID).vertices(y + 2).y = vertices(indices(y + 2)).y
    CollisionFloats(WorldID).vertices(y + 2).z = vertices(indices(y + 2)).z
Next

CollisionFloats(WorldID).numverts = CMesh.GetNumFaces * 3
    
ReDim indices(0)
ReDim vertices(0)           'destroy all temp objects!!
Set IBuf = Nothing
Set CMesh = Nothing
End Sub

Public Sub ClearSecondaryCollisionArray()
Dim x As Long
For x = 4 To UBound(CollisionFloats)
    ReDim CollisionFloats(x).vertices(0)
    CollisionFloats(x).numverts = 0
Next
End Sub

Public Sub ComputeCollisionFloats2ndPass()
Dim x As Long, y As Long, TempArray() As D3DVECTOR, TransM As D3DMATRIX, RotM As D3DMATRIX, result As D3DMATRIX
Dim Max As Long

For x = 1 To 3  'NumWorlds
    If UBound(CollisionFloats(x).vertices) = CollisionFloats(x).numverts - 1 Then
        'array already constructed, so skip that point
    Else
        'truncate array in order to join the new FO verts
        ReDim Preserve CollisionFloats(x).vertices(CollisionFloats(x).numverts - 1)
    End If
    
    'add the FO verts NOW!
    For y = 1 To UBound(FixedObjects)
        If FixedObjects(y).WorldID = x Then
            If FixedObjects(y).unimesh = 1 Then
                ExtractTriVec FOComplex(FixedObjects(y).MeshID).Mesh, TempArray
            Else
                ExtractTriVec FOSimple(FixedObjects(y).MeshID).Mesh, TempArray
            End If
            D3DXMatrixTranslation TransM, FixedObjects(y).Position.x, FixedObjects(y).Position.y, FixedObjects(y).Position.z
            D3DXMatrixRotationY RotM, FixedObjects(y).RotationY
            D3DXMatrixMultiply result, RotM, TransM
            TransformVerts TempArray, result
            
            AddToArray TempArray, CollisionFloats(x).vertices
            If UBound(CollisionFloats(x).vertices) > Max Then Max = UBound(CollisionFloats(x).vertices)
        End If
    Next
Next
End Sub

Public Sub ComputeCollisionFloats2ndPassWorld(ByVal WorldID As Integer)
Dim x As Long, y As Long, TempArray() As D3DVECTOR, TransM As D3DMATRIX, RotM As D3DMATRIX, result As D3DMATRIX
Dim Max As Long

If UBound(CollisionFloats(WorldID).vertices) = CollisionFloats(WorldID).numverts - 1 Then
    'array already constructed, so skip that point
Else
    'truncate array in order to join the new FO verts
    ReDim Preserve CollisionFloats(WorldID).vertices(CollisionFloats(WorldID).numverts - 1)
End If

'add the FO verts NOW!
For y = 1 To UBound(FixedObjects)
    If FixedObjects(y).WorldID = WorldID Then
        If FixedObjects(y).unimesh = 1 Then
            ExtractTriVec FOComplex(FixedObjects(y).MeshID).Mesh, TempArray
        Else
            ExtractTriVec FOSimple(FixedObjects(y).MeshID).Mesh, TempArray
        End If
        D3DXMatrixTranslation TransM, FixedObjects(y).Position.x, FixedObjects(y).Position.y, FixedObjects(y).Position.z
        D3DXMatrixRotationY RotM, FixedObjects(y).RotationY
        D3DXMatrixMultiply result, RotM, TransM
        TransformVerts TempArray, result
        
        AddToArray TempArray, CollisionFloats(WorldID).vertices
        If UBound(CollisionFloats(x).vertices) > Max Then Max = UBound(CollisionFloats(x).vertices)
    End If
Next
End Sub

Public Sub LoadMainChar()
Dim mypath As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "mainchar_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "graphics\mainchar.dat", mypath
ShowLoadPercent "26"

Set MainChar = New Cal3DModel
MainChar.LoadData mypath, "mainchar"
MainChar.LoadAnim mypath & "mainchar_q.caf", "quiet"
MainChar.LoadAnim mypath & "mainchar_w.caf", "walking"
MainChar.LoadAnim mypath & "mainchar_r.caf", "running"
MainChar.NowReady
MainChar.CreateModel

ShowLoadPercent "30"

DeleteDir mypath
End Sub

Public Sub LoadSkyBox()
Dim mypath As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "sky_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "maps\sky.dat", mypath

'assume no compression is activated in order to avoid some view issues in the sky
Set SkyBoxTextures(0) = LoadTextureAndReturn(mypath & "day_front.bmp", False)
Set SkyBoxTextures(1) = LoadTextureAndReturn(mypath & "day_right.bmp", False)
Set SkyBoxTextures(2) = LoadTextureAndReturn(mypath & "day_back.bmp", False)
Set SkyBoxTextures(3) = LoadTextureAndReturn(mypath & "day_left.bmp", False)
Set SkyBoxTextures(4) = LoadTextureAndReturn(mypath & "day_up.bmp", False)

ShowLoadPercent "43"

Set SkyBoxTextures(5) = LoadTextureAndReturn(mypath & "aft_front.bmp", True)
Set SkyBoxTextures(6) = LoadTextureAndReturn(mypath & "aft_right.bmp", True)
Set SkyBoxTextures(7) = LoadTextureAndReturn(mypath & "aft_back.bmp", True)
Set SkyBoxTextures(8) = LoadTextureAndReturn(mypath & "aft_left.bmp", True)
Set SkyBoxTextures(9) = LoadTextureAndReturn(mypath & "aft_up.bmp ", True)

ShowLoadPercent "47"

Set SkyBoxTextures(10) = LoadTextureAndReturn(mypath & "night_front.bmp", True)
Set SkyBoxTextures(11) = LoadTextureAndReturn(mypath & "night_right.bmp", True)
Set SkyBoxTextures(12) = LoadTextureAndReturn(mypath & "night_back.bmp", True)
Set SkyBoxTextures(13) = LoadTextureAndReturn(mypath & "night_left.bmp", True)
Set SkyBoxTextures(14) = LoadTextureAndReturn(mypath & "night_up.bmp", True)

DeleteDir mypath

ShowLoadPercent "53"

SkyBoxVertices(0) = AssignMVS(-0.5, 0.6, 0.5, 0.001, 0.001)
SkyBoxVertices(1) = AssignMVS(0.5, 0.6, 0.5, 0.999, 0.001)
SkyBoxVertices(2) = AssignMVS(-0.5, 0, 0.5, 0.001, 0.999)
SkyBoxVertices(3) = AssignMVS(0.5, 0, 0.5, 0.999, 0.999)

SkyBoxVertices(12) = AssignMVS(-0.5, 0.6, -0.5, 0.001, 0.001)
SkyBoxVertices(13) = AssignMVS(-0.5, 0.6, 0.5, 1, 0.001)
SkyBoxVertices(14) = AssignMVS(-0.5, 0, -0.5, 0.001, 0.999)
SkyBoxVertices(15) = AssignMVS(-0.5, 0, 0.5, 0.999, 0.999)

SkyBoxVertices(8) = AssignMVS(0.5, 0.6, -0.5, 0.001, 0.001)
SkyBoxVertices(9) = AssignMVS(-0.5, 0.6, -0.5, 1, 0.001)
SkyBoxVertices(10) = AssignMVS(0.5, 0, -0.5, 0.001, 0.999)
SkyBoxVertices(11) = AssignMVS(-0.5, 0, -0.5, 0.999, 0.999)

SkyBoxVertices(4) = AssignMVS(0.5, 0.6, 0.5, 0.001, 0.001)
SkyBoxVertices(5) = AssignMVS(0.5, 0.6, -0.5, 0.999, 0.001)
SkyBoxVertices(6) = AssignMVS(0.5, 0, 0.5, 0.001, 0.999)
SkyBoxVertices(7) = AssignMVS(0.5, 0, -0.5, 0.999, 0.999)

SkyBoxVertices(16) = AssignMVS(-0.5, 0.6, -0.5, 0.01, 0.01)
SkyBoxVertices(17) = AssignMVS(0.5, 0.6, -0.5, 0.999, 0.01)
SkyBoxVertices(18) = AssignMVS(-0.5, 0.6, 0.5, 0.01, 0.999)
SkyBoxVertices(19) = AssignMVS(0.5, 0.6, 0.5, 0.999, 0.999)
End Sub

Public Sub LoadFixedObjects()
Dim mypath As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "fobj_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "graphics\obj.dat", mypath

Dim f As Integer, cad As String, lon As Long, lon2 As Long, sin As Single, x As Long, WorldID As Integer
Dim posv As D3DVECTOR, file1 As String, file2 As String, AutoY As Boolean, roty As Single
f = FreeFile
Open mypath & "obj.dat" For Binary As #f
    cad = Space(Len("FODATAFILE"))
    Get #f, , cad
    Get #f, , lon
    ReDim FixedObjects(lon)
    For x = 1 To lon
        Get #f, , sin
        posv.x = sin
        Get #f, , sin
        posv.y = sin
        If sin = -9999 Then
            AutoY = True
        Else
            AutoY = False
        End If
        Get #f, , sin
        posv.z = sin
        Get #f, , sin
        roty = sin
        Get #f, , lon2
        cad = Space(lon2)
        Get #f, , cad
        file1 = mypath & cad
        Get #f, , lon2
        cad = Space(lon2)
        Get #f, , cad
        file2 = mypath & cad
        Get #f, , WorldID
        FixedObjects(x) = LoadFixedObject(posv, file1, file2, AutoY, roty, WorldID)
    Next
Close #f

DeleteDir mypath
End Sub

Public Sub LoadSpecialModels()
Dim mypath As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "fobj_tmp" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "graphics\special.dat", mypath

LoadModel3D MissionTargetModel, mypath & "target.x"
LoadModel3D SavePointModel, mypath & "save.x"

DeleteDir mypath
End Sub

Public Sub SoundReady()
'-------------- INIT DIRECT SOUND ----------
DirectSound.SetCooperativeLevel frmGraphics.hWnd, DSSCL_PRIORITY
DSPrimaryBufferDesc.lFlags = DSBCAPS_CTRL3D Or DSBCAPS_PRIMARYBUFFER
Set DSPrimaryBuffer = DirectSound.CreatePrimarySoundBuffer(DSPrimaryBufferDesc)
Set DSListener = DSPrimaryBuffer.GetDirectSound3DListener()
End Sub

Public Sub LoadCommonSounds()
'-------------------------------------------
Dim mypath As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "sound_tmp_" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "sound\main.dat", mypath

'----- Global Sounds ------
CreateStandardSound mypath & "pasos.wav", "pasos_mainchar"

DeleteDir mypath
End Sub

Public Sub LoadUI()
Dim mypath As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "tex_tmp_prev" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "maps\ui.dat", mypath

ShowLoadPercent "12"
Set PaperTex = LoadTextureAndReturn(mypath & "paper.tga")
Set MessageFont = LoadTextureAndReturn(mypath & "messages.tga")
Set Coin = LoadTextureAndReturn(mypath & "coin.tga")
Set LSBar = LoadTextureAndReturn(mypath & "bar.bmp")
Set LSBar2 = LoadTextureAndReturn(mypath & "bar2.bmp")
Set LoadingWorld = LoadTextureAndReturn(mypath & "loading_world.bmp")
Set FontTexture = LoadTextureAndReturn(mypath & "font.tga")
ShowLoadPercent "16"

DeleteDir mypath
End Sub

Public Sub LoadMiniMaps()
Dim mypath As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "tex_tmp_prev" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "maps\navigate.dat", mypath

Dim x As Long

For x = 1 To 2
    Set MiniMapTex(x) = LoadTextureAndReturn(mypath & "minimap" & Trim(Str(x)) & ".jpg")
Next
Set MiniMapBorder = LoadTextureAndReturn(mypath & "border.tga")
Set MiniMapIcons = LoadTextureAndReturn(mypath & "icons.tga")

DeleteDir mypath
End Sub

Public Sub LoadDoors()
Dim mypath As String, x As Long, FFile As Integer, cad As String, cad2 As String, lastworld As Integer, Atributes() As String
mypath = TempPath()
If right(mypath, 1) <> "\" Then mypath = mypath & "\"
mypath = mypath & "tex_tmp_prev" & Format(Int(Rnd * 5000) + 1, "0000") & "\"
On Local Error Resume Next
MkDir mypath
On Local Error GoTo 0
ExtractFile EXE & "stages\doors.dat", mypath

ReDim DoorArray(0): ReDim AreaArray(0)

FFile = FreeFile()
Open mypath & "doors.dat" For Input As #FFile
Do While Not EOF(FFile)
    Line Input #FFile, cad
    If InStr(1, cad, "#") <> 0 Then cad = left(cad, InStr(1, cad, "#") - 1)
    x = InStr(1, cad, "=")
    If x <> 0 Then
        cad2 = mid(cad, x + 1)
        cad = left(cad, x - 1)
    End If
    Select Case cad
    Case "<world"
        lastworld = Val(cad2)
    Case "door"
        ReDim Preserve DoorArray(UBound(DoorArray) + 1)
        DoorArray(UBound(DoorArray)).World = lastworld
        Atributes = Split(cad2, ",")
        DoorArray(UBound(DoorArray)).Position = ParsePos(Atributes(0) & "," & Atributes(1) & "," & Atributes(2))
        DoorArray(UBound(DoorArray)).id = Val(Atributes(3))
        DoorArray(UBound(DoorArray)).RotationH = Val(Atributes(4))
    Case "area"
        ReDim Preserve AreaArray(UBound(AreaArray) + 1)
        AreaArray(UBound(AreaArray)).World = lastworld
        Atributes = Split(cad2, ",")
        AreaArray(UBound(AreaArray)).pos = ParsePos(Atributes(0) & "," & Atributes(1) & "," & Atributes(2))
        AreaArray(UBound(AreaArray)).Pos2 = ParsePos(Atributes(3) & "," & Atributes(4) & "," & Atributes(5))
        AreaArray(UBound(AreaArray)).DoorName = Trim(Atributes(6))
        AreaArray(UBound(AreaArray)).NewWorld = Val(Atributes(7))
        AreaArray(UBound(AreaArray)).DoorId = Val(Atributes(8))
    End Select
Loop
Close #FFile

DeleteDir mypath
End Sub
