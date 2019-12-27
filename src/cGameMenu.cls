VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGameMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'-------- All coordinates in a menu base of 800 x 600 ---------

Public isEvent As Boolean
Public EventName As String
'Public EventValue As String
Public BackGround As String
Public ListTexture As String
Public ListTextureOpen As String
Public ListTextureOpenHover As String
Public TexPath As String
Public DialogTexturef As String
Public DialogTextureButtonsf As String

Private Textures() As Direct3DTexture8
Private Textures2() As Direct3DTexture8
Private Textures3() As Direct3DTexture8
Private BGTex As Direct3DTexture8
Private ListTex As Direct3DTexture8
Private ListTexOpen As Direct3DTexture8
Private ListTexOpenHover As Direct3DTexture8
Private DialogTexture As Direct3DTexture8
Private DialogTextureButtons As Direct3DTexture8

Private Textures3Names() As String

Private Type myButton
    orX As Single
    orY As Single
    orWidth As Single
    orHeight As Single
    '---- transformed ----
    x As Single
    y As Single
    width As Single
    height As Single
    '--- other ----
    texture As String
    texturehover As String
    ColorK As Long
    b_name As String
    isover As Boolean
End Type

Private Type myList
    orX As Single
    orY As Single
    orWidth As Single
    orHeight As Single
    orFontHeight As Single
    orFontWidth As Single
    '---- transformed ----
    x As Single
    y As Single
    width As Single
    height As Single
    FontHeight As Single
    FontWidth As Single
    '----- other ------
    list() As String
    listover() As Byte
    l_name As String
    maxrows As Integer
    startelement As Integer
    selected As Integer
    scroll As Boolean
End Type

Private Type myLabel
    orX As Single
    orY As Single
    orFontWidth As Single
    orFontHeight As Single
    '---- transformed ----
    x As Single
    y As Single
    FontWidth As Single
    FontHeight As Single
    '------------------
    Text As String
    LabelID As String
End Type

Private Type myImage
    orX As Single
    orY As Single
    orWidth As Single
    orHeight As Single
    '---- transformed ----
    x As Single
    y As Single
    width As Single
    height As Single
    '---------------------
    texture As String
    ColorK As Long
End Type

Private bgvertices(3) As myVertex
Private numLists As Integer
Private arrayLists() As myList
Private numButtons As Integer
Private arrayButtons() As myButton
Private numLabels As Integer
Private arrayLabels() As myLabel
Private numImages As Integer
Private arrayImages() As myImage

Private aListIsOpened As Boolean
Private ListOpenedIndex As Integer

Private DialogText As String
Private DialogType As Byte

Private PublicVertices(3) As myVertex

Public Sub AddButton(ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal texture As String, ByVal texturehover As String, ByVal ButtonName As String, ByVal ColorK As Long)
numButtons = numButtons + 1
Dim current As Integer
current = numButtons - 1
ReDim Preserve arrayButtons(current)
arrayButtons(current).orX = x
arrayButtons(current).orY = y
arrayButtons(current).orWidth = width
arrayButtons(current).orHeight = height
arrayButtons(current).texture = texture
arrayButtons(current).texturehover = texturehover
arrayButtons(current).b_name = ButtonName
arrayButtons(current).ColorK = ColorK
End Sub

Public Sub AddList(ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal FontHeight As Single, ByVal maxrows As Integer, ByVal listname As String, ByVal selected As Integer, Optional ByVal FontWidth As Single = 0)
numLists = numLists + 1
Dim current As Integer
current = numLists - 1
ReDim Preserve arrayLists(current)
arrayLists(current).orX = x
arrayLists(current).orY = y
arrayLists(current).orWidth = width
arrayLists(current).orHeight = height
arrayLists(current).orFontHeight = FontHeight
If FontWidth = 0 Then FontWidth = FontHeight
arrayLists(current).orFontWidth = FontWidth
arrayLists(current).l_name = listname
arrayLists(current).maxrows = maxrows
arrayLists(current).selected = selected

ReDim arrayLists(current).listover(arrayLists(current).maxrows + 2)

ReDim arrayLists(current).list(0)
End Sub

Public Sub AddLabel(ByVal Text As String, ByVal x As Single, ByVal y As Single, ByVal FontHeight As Single, Optional ByVal FontWidth As Single = 0, Optional ByVal LabelID As String)
numLabels = numLabels + 1
ReDim Preserve arrayLabels(numLabels - 1)
arrayLabels(numLabels - 1).Text = Text
arrayLabels(numLabels - 1).orX = x
arrayLabels(numLabels - 1).orY = y
arrayLabels(numLabels - 1).LabelID = LabelID
arrayLabels(numLabels - 1).orFontHeight = FontHeight
If FontWidth = 0 Then
    arrayLabels(numLabels - 1).orFontWidth = FontHeight
Else
    arrayLabels(numLabels - 1).orFontWidth = FontWidth
End If
End Sub

Public Sub EditLabel(ByVal LabelID As String, ByVal Text As String)
Dim f As Long
For f = 0 To numLabels - 1
    If arrayLabels(f).LabelID = LabelID Then
        arrayLabels(f).Text = Text
        Exit Sub
    End If
Next
End Sub

Public Sub AddImage(ByVal file As String, ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal ColorK As Long)
numImages = numImages + 1
ReDim Preserve arrayImages(numImages - 1)
arrayImages(numImages - 1).orX = x
arrayImages(numImages - 1).orY = y
arrayImages(numImages - 1).orWidth = width
arrayImages(numImages - 1).orHeight = height
arrayImages(numImages - 1).texture = file
arrayImages(numImages - 1).ColorK = ColorK
End Sub

Public Sub ProcessClick(ByVal x As Single, ByVal y As Single)
'-------------- THIS COORDINATES ARE IN D3DM resolution!!!!!! --------------
Dim rows As Integer, number As Integer, bottom  As Boolean
If DialogText <> "" Then
    'check for the dialog buttons ONLY
    If DialogType = 0 Then
        If x > 350 * D3DM.width / 800 And x < 450 * D3DM.width / 800 Then
        If y > 364 * D3DM.height / 600 And y < 412 * D3DM.height / 600 Then
            'Ok BUTTON
            DialogText = ""
            Me.isEvent = True
            Me.EventName = "_dialog_exit"
        End If
        End If
    Else
        If x > 275 * D3DM.width / 800 And x < 350 * D3DM.width / 800 Then
        If y > 364 * D3DM.height / 600 And y < 412 * D3DM.height / 600 Then
            'Yes BUTTON
            DialogText = ""
            Me.isEvent = True
            Me.EventName = "_dialog_exit_yes"
        End If
        End If
        
        If x > 450 * D3DM.width / 800 And x < 525 * D3DM.width / 800 Then
        If y > 364 * D3DM.height / 600 And y < 412 * D3DM.height / 600 Then
            'Yes BUTTON
            DialogText = ""
            Me.isEvent = True
            Me.EventName = "_dialog_exit_no"
        End If
        End If

    End If
    Exit Sub
End If
If Not aListIsOpened Then
    Dim f As Integer
    For f = 0 To numButtons - 1
    With arrayButtons(f)
        If x > (.x) And x < (.x + .width) Then
        If y > (.y) And y < (.y + .height) Then
            isEvent = True
            EventName = .b_name
            Exit For
        End If
        End If
    End With
    Next
    
    For f = 0 To numLists - 1
    With arrayLists(f)
        If x > (.x) And x < (.x + .width) Then
        If y > (.y) And y < (.y + .height) Then
            'Open the list!
            aListIsOpened = True
            ListOpenedIndex = f
            Exit For
        End If
        End If
    End With
    Next
Else
    With arrayLists(ListOpenedIndex)
        rows = UBound(.list)
        If rows > .maxrows Then rows = .maxrows
        rows = rows + 2
        
        If x > (.x) And x < (.x + .width) And _
           y > (.y + .height) And y < (.y + .height + rows * .height) Then
            'The user has chosen a list element
            number = (y - .y - .height) \ .height + 1
            If .scroll = False Then
                If number = 1 Or number = UBound(.list) + 2 Then Exit Sub
                number = number - 1
                .selected = number
                aListIsOpened = False
            Else
                If number = 1 And .startelement <> 1 Then .startelement = .startelement - 1
                If number = .maxrows + 2 And .startelement <> (UBound(.list) - .maxrows + 1) Then .startelement = .startelement + 1
                If number <> 1 And number <> .maxrows + 2 Then
                    .selected = .startelement + number - 2
                    aListIsOpened = False
                End If
            End If
        Else
            aListIsOpened = False
        End If
    End With
End If
End Sub

Public Sub ProcessMouseMove(ByVal x As Single, ByVal y As Single)
'-------------- THIS COORDINATES ARE IN D3DM resolution!!!!!! --------------
Dim f As Integer, rows As Integer

If Not aListIsOpened Then
    For f = 0 To numButtons - 1
        arrayButtons(f).isover = False
    Next
    For f = 0 To numButtons - 1
    With arrayButtons(f)
        If x > (.x) And x < (.x + .width) Then
        If y > (.y) And y < (.y + .height) Then
            .isover = True
            Exit For
        End If
        End If
    End With
    Next
Else
    With arrayLists(ListOpenedIndex)
        For f = 1 To UBound(.listover)
            .listover(f) = 0
        Next
        
        rows = UBound(.list)
        If rows > .maxrows Then rows = .maxrows
        rows = rows + 2
        
        If x > (.x) And x < (.x + .width) And _
           y > (.y + .height) And y < (.y + .height + rows * .height) Then
            rows = (y - .y - .height) \ .height + 1
            If rows > .maxrows + 2 Then rows = .maxrows
            .listover(rows) = 1
        End If
    End With
End If
End Sub

'--------------- LOADS THE TEXTURES AND CALCULATES THE COORDS. ------------------
Public Sub BuildMenu()
ReDim Textures(numButtons - 1)
ReDim Textures2(numButtons - 1)
Dim f As Integer, g As Integer, number As Integer
For f = 0 To numButtons - 1
With arrayButtons(f)
    Set Textures(f) = LoadTex(.texture, .ColorK)
    Set Textures2(f) = LoadTex(.texturehover, .ColorK)
End With
Next

Set BGTex = LoadTex(BackGround, 0)

If DialogTexturef <> "" And DialogTextureButtonsf <> "" Then
    Set DialogTexture = LoadTex(DialogTexturef, 0)
    Set DialogTextureButtons = LoadTex(DialogTextureButtonsf, 0)
End If

For f = 0 To numLists - 1
With arrayLists(f)
    .startelement = 1
    .scroll = .maxrows < UBound(.list)
End With
Next

If ListTexture <> "" Then Set ListTex = LoadTex(ListTexture, 0)
If ListTextureOpen <> "" Then Set ListTexOpen = LoadTex(ListTextureOpen, 0)
If ListTextureOpenHover <> "" Then Set ListTexOpenHover = LoadTex(ListTextureOpenHover, 0)

ReDim Textures3Names(0)
ReDim Textures3(0)
For f = 0 To numImages - 1
    For g = 0 To UBound(Textures3Names) - 1
        If Textures3Names(g) = arrayImages(f).texture Then GoTo Out
    Next
    ReDim Preserve Textures3Names(UBound(Textures3Names) + 1)
    ReDim Preserve Textures3(UBound(Textures3) + 1)
    Textures3Names(UBound(Textures3Names) - 1) = arrayImages(f).texture
    Set Textures3(UBound(Textures3) - 1) = LoadTex(arrayImages(f).texture, arrayImages(f).ColorK)
Out:
Next

Call RefreshDM
End Sub

'----------------- RECALCULATES ALL THE COORDS. FOR THE NEW DM ----------------
Public Sub RefreshDM()
Dim f As Integer
For f = 0 To numButtons - 1
With arrayButtons(f)
    .x = .orX * D3DM.width / 800
    .y = .orY * D3DM.height / 600
    .width = .orWidth * D3DM.width / 800
    .height = .orHeight * D3DM.height / 600
End With
Next

For f = 0 To numLists - 1
With arrayLists(f)
    .x = .orX * D3DM.width / 800
    .y = .orY * D3DM.height / 600
    .width = .orWidth * D3DM.width / 800
    .height = .orHeight * D3DM.height / 600
    .FontHeight = .orFontHeight * D3DM.height / 600
    .FontWidth = .orFontWidth * D3DM.width / 800
End With
Next

For f = 0 To numImages - 1
With arrayImages(f)
    .x = .orX * D3DM.width / 800
    .y = .orY * D3DM.height / 600
    .width = .orWidth * D3DM.width / 800
    .height = .orHeight * D3DM.height / 600
End With
Next

For f = 0 To numLabels - 1
With arrayLabels(f)
    .x = .orX * D3DM.width / 800
    .y = .orY * D3DM.height / 600
    .FontWidth = .orFontWidth * D3DM.width / 800
    .FontHeight = .orFontHeight * D3DM.height / 600
End With
Next
End Sub

'------------------ SHOWS THE MENU WITH ALL THE EFECTS ---------------
Public Sub RenderMenu()
bgvertices(0) = AssignMV(0, 0, 0, 0.001, 0.001)
bgvertices(1) = AssignMV(D3DM.width, 0, 0, 0.999, 0.001)
bgvertices(2) = AssignMV(0, D3DM.height, 0, 0.001, 0.999)
bgvertices(3) = AssignMV(D3DM.width, D3DM.height, 0, 0.999, 0.999)
Device.SetTexture 0, BGTex
Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, bgvertices(0), Len(bgvertices(0))
Dim f As Integer, g As Integer, offset As Single, offset2 As Single

For f = 0 To numImages - 1
With arrayImages(f)
    PublicVertices(0) = AssignMV(.x, .y, 0, 0.001, 0.001)
    PublicVertices(1) = AssignMV(.x + .width, .y, 0, 0.999, 0.001)
    PublicVertices(2) = AssignMV(.x, .y + .height, 0, 0, 0.999)
    PublicVertices(3) = AssignMV(.x + .width, .y + .height, 0, 0.999, 0.999)
    Device.SetTexture 0, GetTex3(.texture)
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, PublicVertices(0), Len(PublicVertices(0))
End With
Next

For f = 0 To numButtons - 1
With arrayButtons(f)
    If .isover = False Or DialogText <> "" Then
        Device.SetTexture 0, Textures(f)
    Else
        Device.SetTexture 0, Textures2(f)
    End If
    PublicVertices(0) = AssignMV(.x, .y, 0, 0.001, 0.001)
    PublicVertices(1) = AssignMV(.x + .width, .y, 0, 0.999, 0.001)
    PublicVertices(2) = AssignMV(.x, .y + .height, 0, 0.001, 0.999)
    PublicVertices(3) = AssignMV(.x + .width, .y + .height, 0, 0.999, 0.999)
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, PublicVertices(0), Len(PublicVertices(0))
End With
Next

For f = 0 To numLabels - 1
With arrayLabels(f)
    DrawFont .Text, .x, .y, 0, .FontHeight, .FontWidth
End With
Next

For f = 0 To numLists - 1
With arrayLists(f)
    offset = (.width - Len(.list(.selected)) * .FontWidth) / 2
    offset2 = (.height - .FontHeight) / 2
    PublicVertices(0) = AssignMV(.x, .y, 0, 0.001, 0.001)
    PublicVertices(1) = AssignMV(.x + .width, .y, 0, 0.999, 0.001)
    PublicVertices(2) = AssignMV(.x, .y + .height, 0, 0.001, 0.999)
    PublicVertices(3) = AssignMV(.x + .width, .y + .height, 0, 0.999, 0.999)
    Device.SetTexture 0, ListTex
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, PublicVertices(0), Len(PublicVertices(0))
    If offset >= 0 Then
        DrawFont .list(.selected), .x + offset, .y + offset2, 0, .FontHeight, .FontWidth
    Else
        offset = Int(.width / .FontWidth)
        DrawFont left(.list(.selected), offset - 1), .x + .FontWidth / 2, .y + offset2, 0, .FontHeight, .FontWidth
    End If
End With
Next

If aListIsOpened And DialogText = "" Then
    With arrayLists(ListOpenedIndex)
    offset2 = (.height - .FontHeight) / 2
    If .scroll = True Then
        For f = 0 To .maxrows + 1
            PublicVertices(0) = AssignMV(.x, .y + .height * (f + 1), 0, 0.001, 0.001)
            PublicVertices(1) = AssignMV(.x + .width, .y + .height * (f + 1), 0, 0.999, 0.001)
            PublicVertices(2) = AssignMV(.x, .y + .height * (f + 2), 0, 0.001, 0.999)
            PublicVertices(3) = AssignMV(.x + .width, .y + .height * (f + 2), 0, 0.999, 0.999)
            If .listover(f + 1) = 1 Then
                If (f <> 0 And f <> .maxrows + 1) Or _
                   ((f = 0) And (.startelement <> 1)) Or _
                   ((f = .maxrows + 1) And .startelement <> (UBound(.list) - .maxrows + 1)) Then
                    Device.SetTexture 0, ListTexOpenHover
                Else
                    Device.SetTexture 0, ListTexOpen
                End If
            Else
                Device.SetTexture 0, ListTexOpen
            End If
            Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, PublicVertices(0), Len(PublicVertices(0))
    
            If f <> .maxrows + 1 And f <> 0 Then
                offset = (.width - Len(.list(.startelement + f - 1)) * .FontWidth) / 2#
                If offset >= 0 Then
                    DrawFont .list(.startelement + f - 1), .x + offset, .y + (f + 1) * .height + offset2, 0, .FontHeight, .FontWidth
                Else
                    offset = Int(.width / .FontWidth)
                    DrawFont left(.list(.startelement + f - 1), offset - 1), .x + .FontWidth / 2, .y + (f + 1) * .height + offset2, 0, .FontHeight, .FontWidth
                End If
            ElseIf f = 0 And .startelement <> 1 Then
                offset = (.width - .FontWidth) / 2#
                DrawFont "ª", .x + offset, .y + (f + 1) * .height + offset2, 0, .FontHeight, .FontWidth
            ElseIf (.startelement + .maxrows - 1) <> UBound(.list) And f = .maxrows + 1 Then
                offset = (.width - .FontWidth) / 2#
                DrawFont "º", .x + offset, .y + (f + 1) * .height + offset2, 0, .FontHeight, .FontWidth
            End If
        Next
    Else
        For f = 0 To UBound(.list) + 1
            PublicVertices(0) = AssignMV(.x, .y + .height * (f + 1), 0, 0.001, 0.001)
            PublicVertices(1) = AssignMV(.x + .width, .y + .height * (f + 1), 0, 0.999, 0.001)
            PublicVertices(2) = AssignMV(.x, .y + .height * (f + 2), 0, 0.001, 0.999)
            PublicVertices(3) = AssignMV(.x + .width, .y + .height * (f + 2), 0.001, 0.999, 0.999)
            If .listover(f + 1) = 1 And f <> UBound(.list) + 1 And f <> 0 Then
                Device.SetTexture 0, ListTexOpenHover
            Else
                Device.SetTexture 0, ListTexOpen
            End If
            Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, PublicVertices(0), Len(PublicVertices(0))
            
            If f <> UBound(.list) + 1 And f <> 0 Then
                offset = (.width - Len(.list(f)) * .FontWidth) / 2#
                If offset >= 0 Then
                    DrawFont .list(f), .x + offset, .y + (f + 1) * .height + offset2, 0, .FontHeight, .FontWidth
                Else
                    offset = Int(.width / .FontWidth)
                    DrawFont left(.list(f), offset - 1), .x + .FontWidth / 2, .y + (f + 1) * .height + offset2, 0, .FontHeight, .FontWidth
                End If
            End If
        Next
    End If
    End With
End If

If DialogText <> "" Then
    PublicVertices(0) = AssignMV((D3DM.width - D3DM.width * 512 / 800) / 2, (D3DM.height - D3DM.height * 256 / 600) / 2, 0, 0.001, 0.001)
    PublicVertices(1) = AssignMV((D3DM.width + D3DM.width * 512 / 800) / 2, (D3DM.height - D3DM.height * 256 / 600) / 2, 0, 0.999, 0.001)
    PublicVertices(2) = AssignMV((D3DM.width - D3DM.width * 512 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2, 0, 0.001, 0.999)
    PublicVertices(3) = AssignMV((D3DM.width + D3DM.width * 512 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2, 0, 0.999, 0.999)
    Device.SetTexture 0, DialogTexture
    Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, PublicVertices(0), Len(PublicVertices(0))
    
    If DialogType = 0 Then
        PublicVertices(0) = AssignMV((D3DM.width - D3DM.width * 100 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 64 * D3DM.height / 600, 0, 0.666, 0.001)
        PublicVertices(1) = AssignMV((D3DM.width + D3DM.width * 100 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 64 * D3DM.height / 600, 0, 0.999, 0.001)
        PublicVertices(2) = AssignMV((D3DM.width - D3DM.width * 100 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 16 * D3DM.height / 600, 0, 0.666, 0.999)
        PublicVertices(3) = AssignMV((D3DM.width + D3DM.width * 100 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 16 * D3DM.height / 600, 0, 0.999, 0.999)
        Device.SetTexture 0, DialogTextureButtons
        Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, PublicVertices(0), Len(PublicVertices(0))
    Else
        Device.SetTexture 0, DialogTextureButtons
        PublicVertices(0) = AssignMV((D3DM.width - D3DM.width * 250 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 64 * D3DM.height / 600, 0, 0.001, 0.001)
        PublicVertices(1) = AssignMV((D3DM.width - D3DM.width * 100 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 64 * D3DM.height / 600, 0, 0.333, 0.001)
        PublicVertices(2) = AssignMV((D3DM.width - D3DM.width * 250 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 16 * D3DM.height / 600, 0, 0.001, 0.999)
        PublicVertices(3) = AssignMV((D3DM.width - D3DM.width * 100 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 16 * D3DM.height / 600, 0, 0.333, 0.999)
        Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, PublicVertices(0), Len(PublicVertices(0))
        PublicVertices(0) = AssignMV((D3DM.width + D3DM.width * 100 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 64 * D3DM.height / 600, 0, 0.333, 0.001)
        PublicVertices(1) = AssignMV((D3DM.width + D3DM.width * 250 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 64 * D3DM.height / 600, 0, 0.666, 0.001)
        PublicVertices(2) = AssignMV((D3DM.width + D3DM.width * 100 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 16 * D3DM.height / 600, 0, 0.333, 0.999)
        PublicVertices(3) = AssignMV((D3DM.width + D3DM.width * 250 / 800) / 2, (D3DM.height + D3DM.height * 256 / 600) / 2 - 16 * D3DM.height / 600, 0, 0.666, 0.999)
        Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, PublicVertices(0), Len(PublicVertices(0))
    End If
    
    DrawParagraph DialogText, (D3DM.width - D3DM.width * 485 / 800) / 2, (D3DM.height - D3DM.height * 256 / 600) / 2, D3DM.width * 485 / 800, 160 * D3DM.height / 600

End If
End Sub

Private Function LoadTex(ByVal file As String, ByVal ColorK As Long) As Direct3DTexture8
Dim transparent As Boolean
Dim imagew As Long, imageh As Long

GetImageProperties file, imagew, imageh, transparent

If caps.MaxTextureWidth < imagew Then imagew = caps.MaxTextureWidth
If caps.MaxTextureHeight < imageh Then imageh = caps.MaxTextureHeight
If caps.TextureCaps And D3DPTEXTURECAPS_SQUAREONLY Then
    If imagew > imageh Then
        imagew = imageh
    Else
        imageh = imagew
    End If
End If

If transparent = False Then
    If Direct3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, D3DM.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_DXT1) = D3D_OK Then
        Set LoadTex = Direct3DX.CreateTextureFromFileEx(Device, TexPath & file, imagew, imageh, D3DX_DEFAULT, 0, D3DFMT_DXT1, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorK, ByVal 0, ByVal 0)
    Else
        Set LoadTex = Direct3DX.CreateTextureFromFileEx(Device, TexPath & file, imagew, imageh, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, ColorK, ByVal 0, ByVal 0)
    End If
Else
    If Direct3D.CheckDeviceFormat(0, D3DDEVTYPE_HAL, D3DM.Format, 0, D3DRTYPE_TEXTURE, D3DFMT_DXT5) = D3D_OK Then
        Set LoadTex = Direct3DX.CreateTextureFromFileEx(Device, TexPath & file, imagew, imageh, D3DX_DEFAULT, 0, D3DFMT_DXT5, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    Else
        Set LoadTex = Direct3DX.CreateTextureFromFileEx(Device, TexPath & file, imagew, imageh, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_LINEAR, D3DX_FILTER_LINEAR, 0, ByVal 0, ByVal 0)
    End If
End If
End Function

Public Sub AddElementToList(ByVal listname As String, ByVal element As String)
Dim f As Integer
For f = 0 To numLists
    If arrayLists(f).l_name = listname Then GoTo Found:
Next
Exit Sub
Found:
ReDim Preserve arrayLists(f).list(UBound(arrayLists(f).list) + 1)
arrayLists(f).list(UBound(arrayLists(f).list)) = element
End Sub

Public Sub ReInitialize()
Dim f As Integer
For f = 0 To numLists - 1
    arrayLists(f).startelement = 1
Next
End Sub

Private Function GetTex3(ByVal file As String) As Direct3DTexture8
Dim f As Integer
For f = 0 To UBound(Textures3Names) - 1
    If Textures3Names(f) = file Then
        Set GetTex3 = Textures3(f)
        Exit For
    End If
Next
End Function

Public Function GetListValue(ByVal listname As String) As String
Dim f As Integer
For f = 0 To numLists - 1
    If arrayLists(f).l_name = listname Then
        GetListValue = arrayLists(f).list(arrayLists(f).selected)
        Exit Function
    End If
Next
End Function

Public Function GetListIndex(ByVal listname As String) As Integer
Dim f As Integer
For f = 0 To numLists - 1
    If arrayLists(f).l_name = listname Then
        GetListIndex = arrayLists(f).selected
        Exit Function
    End If
Next
End Function

Public Sub SetSelectedListItem(ByVal listname As String, ByVal selected As Integer)
Dim f As Integer
For f = 0 To numLists - 1
    If arrayLists(f).l_name = listname Then
        arrayLists(f).selected = selected
        Exit Sub
    End If
Next
End Sub

Public Sub SetSelectedListItemByText(ByVal listname As String, ByVal selected As String)
Dim f As Integer, x As Integer
For f = 0 To numLists - 1
    If arrayLists(f).l_name = listname Then
        For x = 1 To UBound(arrayLists(f).list)
            If arrayLists(f).list(x) = selected Then
                arrayLists(f).selected = x
                Exit Sub
            End If
        Next
    End If
Next
End Sub

Public Sub ShowDialogBox(ByVal Message As String, ByVal dType As Integer)
'type 0 "ok", 1 "yes / no"
' EventName retunrs the value (yes/no)
'isevent (false, dialog present, true, the dialog has been closed ;))
DialogText = Message
DialogType = dType
End Sub

Public Sub DrawParagraph(ByVal Text As String, x As Single, y As Single, width As Single, height As Single)
Dim numl As Integer, z As Long, linearray() As String, cx As Single, cy As Single
Dim charwidth As Single, charheight As Single, charheight2 As Single
charwidth = 18 * D3DM.width / 800
charheight = 18 * D3DM.height / 600
charheight2 = 22 * D3DM.height / 600    'height for the chars with the space between lines
linearray = Split(Text, "\n")

For z = 1 To UBound(linearray) + 1
    cx = (width - charwidth * Len(linearray(z - 1))) / 2 + x
    cy = (height - charheight2 * UBound(linearray)) / 2 + y + charheight2 * (z - 1)
    DrawFont linearray(z - 1), cx, cy, , charheight, charwidth
Next
End Sub
