VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Tàrraco"
   ClientHeight    =   7965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9315
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Main.frx":08CA
   ScaleHeight     =   507.729
   ScaleMode       =   0  'User
   ScaleWidth      =   826.622
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   " X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9135
      TabIndex        =   0
      Top             =   0
      Width           =   195
   End
   Begin VB.Image l3 
      Height          =   465
      Left            =   2565
      Top             =   6255
      Width           =   4155
   End
   Begin VB.Image l2 
      Height          =   600
      Left            =   2610
      Top             =   5310
      Width           =   4065
   End
   Begin VB.Image l1 
      Height          =   825
      Left            =   1530
      Top             =   3465
      Width           =   6180
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Dim Exe As String

Private Sub Form_Load()
Dim ln As Long
Exe = App.Path
If Right(Exe, 1) = "\" Then
    Exe = Mid(Exe, 1, Len(Exe) - 1)
End If
If InStr(1, Exe, "\") <> 0 Then
    For ln = Len(Exe) To 1 Step -1
        If Mid(Exe, ln, 1) = "\" Then
        Exe = Left(Exe, ln)
        Exit For
        End If
    Next
End If
End Sub

Private Sub l1_Click()
ShellExecute Me.hwnd, vbNullString, Exe & "Joc i codi\", vbNullString, vbNullString, 1
End Sub

Private Sub l2_Click()
ShellExecute Me.hwnd, vbNullString, Exe & "PDF\Tàrraco - Creació d'un joc d'ordinador.pdf", vbNullString, vbNullString, 1
End Sub

Private Sub l3_Click()
ShellExecute Me.hwnd, vbNullString, Exe & "Acrobat Reader\Acrobat Reader 7.0 Español.exe", vbNullString, vbNullString, 1
End Sub

Private Sub Label1_Click()
End
End Sub
