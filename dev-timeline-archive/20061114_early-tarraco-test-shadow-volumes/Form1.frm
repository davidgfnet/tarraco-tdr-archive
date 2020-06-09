VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   8520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10575
   LinkTopic       =   "Form1"
   ScaleHeight     =   8520
   ScaleWidth      =   10575
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Salir = 1
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then CameraCollision = Not CameraCollision
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim ln As Long
'RotX = 360 * x / (Me.ScaleWidth - Me.ScaleWidth / 100 * 2)
'If x < Me.ScaleWidth / 100 Then
'    ln = SetCursorPos((Me.Left + Me.ScaleWidth / 100 * 99) / Screen.TwipsPerPixelX, (Y + Me.Top + Me.Height - Me.ScaleHeight) / Screen.TwipsPerPixelY)
'ElseIf x > Me.ScaleWidth / 100 * 99 Then
'    ln = SetCursorPos(((Me.Left + Me.ScaleWidth / 100 * 2) / Screen.TwipsPerPixelX), (Y + Me.Top + Me.Height - Me.ScaleHeight) / Screen.TwipsPerPixelY)
'End If

'VerticalAngle = 25 * Y / (Me.ScaleHeight - Me.ScaleHeight / 100 * 2) + 20

RotX = 360 * X / (Screen.Width - Screen.Width / 100 * 2)
If X < Screen.Width / 100 Then
    ln = SetCursorPos((Screen.Width / 100 * 99) / Screen.TwipsPerPixelX, Y / Screen.TwipsPerPixelY)
ElseIf X > Me.ScaleWidth / 100 * 99 Then
    ln = SetCursorPos(((Screen.Width / 100 + 1) / Screen.TwipsPerPixelX), Y / Screen.TwipsPerPixelY)
End If

VerticalAngle = 30 * Y / (Screen.Height - Screen.Height / 100 * 2) + 25

'.Print Me.ScaleWidth
End Sub
