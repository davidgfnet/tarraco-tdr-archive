VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FILE PACKER"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "decript"
      Height          =   465
      Left            =   225
      TabIndex        =   7
      Top             =   1935
      Width           =   1860
   End
   Begin VB.TextBox origin 
      Height          =   330
      Left            =   1845
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   180
      Width           =   6630
   End
   Begin VB.PictureBox bm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   180
      ScaleHeight     =   345
      ScaleWidth      =   8310
      TabIndex        =   4
      Top             =   1350
      Width           =   8340
      Begin VB.Label b 
         BackColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.TextBox SaveF 
      Height          =   330
      Left            =   180
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   810
      Width           =   5685
   End
   Begin VB.CommandButton LookFor 
      Caption         =   "..."
      Height          =   420
      Left            =   5985
      TabIndex        =   2
      Top             =   765
      Width           =   420
   End
   Begin VB.CommandButton Create 
      Caption         =   "CREATE FILE"
      Height          =   420
      Left            =   6570
      TabIndex        =   1
      Top             =   765
      Width           =   1950
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   6975
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Add 
      Caption         =   "SELECT FILE"
      Height          =   420
      Left            =   180
      TabIndex        =   0
      Top             =   135
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Add_Click()
dialog.DialogTitle = "Add file"
dialog.CancelError = True
dialog.FileName = ""
dialog.Filter = "All files|*.*"
dialog.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
dialog.MaxFileSize = 32000

On Local Error Resume Next
dialog.ShowOpen

If Err.Number <> 0 Then Exit Sub

origin.Text = dialog.FileName
End Sub

Private Sub clear_Click()
Lista.clear
End Sub

Private Sub Command1_Click()
dialog.DialogTitle = "Decode file"
dialog.CancelError = True
dialog.FileName = ""
dialog.Filter = "Data files|*.dat"
dialog.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
dialog.MaxFileSize = 32000

On Local Error Resume Next
dialog.ShowOpen

If Err.Number <> 0 Then Exit Sub

DecodeFile dialog.FileName, APath(dialog.FileName) & "\original_" & NombreR(dialog.FileName)
End Sub

Private Sub Create_Click()
Dim ff As Integer, x As Long, ff2 As Integer, cadena As String
Dim y As Long, largo As Long, car As Byte
Dim cars() As Byte, max As Integer, cad1 As String, cad2 As String, tam As Long
Dim by As Byte, by2 As Long
ff = FreeFile
On Local Error Resume Next
Kill SaveF.Text
On Local Error GoTo 0
Open SaveF.Text For Binary As #ff

ff2 = FreeFile
    tam = FileLen(origin.Text)
    Open origin.Text For Binary As #ff2
            For x = 1 To tam
                If (x Mod 2) = 1 Then
                    Get #ff2, , by
                    by = 255 - by
                    Put #ff, , by
                Else
                    Get #ff2, , by
                    by2 = by
                    by2 = by2 + 121
                    If by2 > 255 Then by2 = by2 - 256
                    by = by2
                    Put #ff, , by
                End If
                b.Width = bm.Width / tam * x
            Next
    Close #ff2
    
    DoEvents

Close #ff

MsgBox "OK", vbInformation
End Sub

Private Sub LookFor_Click()
dialog.DialogTitle = "Save file"
dialog.CancelError = True
dialog.FileName = ""
dialog.Filter = "Data files|*.dat"
dialog.Flags = cdlOFNPathMustExist Or cdlOFNOverwritePrompt
dialog.MaxFileSize = 32000

On Local Error Resume Next
dialog.ShowSave

If Err.Number <> 0 Then Exit Sub

SaveF.Text = dialog.FileName
End Sub

Private Sub xt_Click()
dialog.DialogTitle = "Xtract file"
dialog.CancelError = True
dialog.FileName = ""
dialog.Filter = "Data files|*.dat"
dialog.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
dialog.MaxFileSize = 32000

On Local Error Resume Next
dialog.ShowOpen

If Err.Number <> 0 Then Exit Sub

ExtractFile dialog.FileName, ""
End Sub
