VERSION 5.00
Begin VB.Form frmGraphics 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Tarraco"
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   8505
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmGraphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements DirectXEvent8

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)

If Not (eventid = DI_hevent) Then Exit Sub

    Dim DevData(1 To 50) As DIDEVICEOBJECTDATA
    Dim nEvents As Long
    Dim i As Long

    On Local Error Resume Next    'handle errors???
    Err.number = 0
    nEvents = MouseDevice.GetDeviceData(DevData, DIGDD_DEFAULT)
    If Err.number <> 0 Then
        'the Mousedevice has been Lost!!!, try to recatch
        Call MouseSetUp
    End If

    For i = 1 To nEvents
        Select Case DevData(i).lOfs
            Case DIMOFS_X
                MouseX = MouseX + DevData(i).lData
            Case DIMOFS_Y
                MouseY = MouseY + DevData(i).lData
            Case DIMOFS_Z
                MouseZ = MouseZ + DevData(i).lData
            Case DIMOFS_BUTTON0
                If DevData(i).lData = 0 Then
                    If MouseB0 = True Then MouseClick0 = True
                    MouseB0 = False
                Else
                    MouseB0 = True
                End If
            Case DIMOFS_BUTTON1
                If DevData(i).lData = 0 Then
                    If MouseB1 = True Then MouseClick1 = True
                    MouseB1 = False
                    If MainRender Then Movement = 0
                Else
                    MouseB1 = True
                    If MainRender Then Movement = 1
                End If
        End Select
    Next

DoEvents
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If MainRender Then          'we are in the game loop
    If KeyCode = vbKeySpace And Jumping = False Then Jumping = True
    If KeyCode = vbKeyW Or KeyCode = vbKeyUp Then Movement = 1
    If KeyCode = vbKeyEscape Then AuxMenu = True
    If KeyCode = vbKeyH Then
        Open "log.txt" For Append As #45
            Print #45, CharPos.x & " / " & CharPos.y & " / " & CharPos.z
        Close #45
    End If
    If KeyCode = vbKeyAdd Or KeyCode = 187 Then
        KeyAdd = True
    End If
    If KeyCode = vbKeySubtract Or KeyCode = 189 Then
        KeySubtract = True
    End If
    If Shift = 1 And KeyCode = vbKeyEnd Then
        ReDim LockedDoors(0)
    End If
    If KeyCode = vbKeyC Then
        If CameraType = 0 Then
            CameraType = 1
        Else
            CameraType = 0
        End If
        CharAngleV = (MaxVerticalAngle + MinVerticalAngle) / 2
        CharDistance = (MaxCameraDistance + MinCameraDistance) / 2
    End If
End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If MainRender Then          'we are in the game loop
    'If Shift = 0 Then       'NO shifts keys
        If KeyCode = vbKeyW Or KeyCode = vbKeyUp Then Movement = 0
    'End If
    If KeyCode = vbKeyAdd Or KeyCode = 187 Then
        KeyAdd = False
    End If
    If KeyCode = vbKeySubtract Or KeyCode = 189 Then
        KeySubtract = False
    End If
End If
End Sub
