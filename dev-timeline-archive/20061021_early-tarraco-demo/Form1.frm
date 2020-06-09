VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Implements DirectXEvent8

Private Sub Form_Click()
salir = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then CameraCollision = Not CameraCollision
End Sub


Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
'On Local Error Resume Next
    If Not (eventid = hevent) Then Exit Sub 'bailout if the msg isn't for us
    
    '//At this point, we know something has happened - and we want to know what it is!!
    
    '0. Any Variables
        Dim DevData(1 To 30) As DIDEVICEOBJECTDATA 'storage for the event data
        Dim nEvents As Long 'how many events have just happened (usually 1)
        Dim I As Long 'looping variables
        
    '1. retrieve the data from the device.
        nEvents = DIDevice.GetDeviceData(DevData, DIGDD_DEFAULT)
        
    '2. loop through all the events
        For I = 1 To nEvents
            Select Case DevData(I).lOfs
                Case DIMOFS_X
                    RotX = RotX + DevData(I).lData * 0.25
                Case DIMOFS_Y
                    VerticalAngle = VerticalAngle + DevData(I).lData * 0.1
                    If VerticalAngle > 45 Then VerticalAngle = 45
                    If VerticalAngle < 20 Then VerticalAngle = 20
                Case DIMOFS_Z
                    CameraDistance = CameraDistance + DevData(I).lData * 0.001
                    If CameraDistance > 6 Then CameraDistance = 6
                    If CameraDistance < 4 Then CameraDistance = 4
                Case DIMOFS_BUTTON0
                    salir = 1
            End Select
        Next
        
        DoEvents
End Sub

