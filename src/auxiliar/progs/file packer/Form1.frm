VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FILE PACKER"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton xt 
      Caption         =   "Xtract"
      Height          =   420
      Left            =   315
      TabIndex        =   8
      Top             =   3735
      Width           =   2220
   End
   Begin VB.CommandButton clear 
      Caption         =   "CLEAR"
      Height          =   420
      Left            =   1710
      TabIndex        =   7
      Top             =   2475
      Width           =   1095
   End
   Begin VB.PictureBox bm 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   225
      ScaleHeight     =   345
      ScaleWidth      =   11055
      TabIndex        =   5
      Top             =   3015
      Width           =   11085
      Begin VB.Label b 
         BackColor       =   &H00FF0000&
         Height          =   375
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   15
      End
   End
   Begin VB.TextBox SaveF 
      Height          =   330
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   5685
   End
   Begin VB.CommandButton LookFor 
      Caption         =   "..."
      Height          =   420
      Left            =   8775
      TabIndex        =   3
      Top             =   2475
      Width           =   420
   End
   Begin VB.CommandButton Create 
      Caption         =   "CREATE FILE"
      Height          =   420
      Left            =   9360
      TabIndex        =   2
      Top             =   2475
      Width           =   1950
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   10800
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Add 
      Caption         =   "ADD FILES"
      Height          =   420
      Left            =   180
      TabIndex        =   1
      Top             =   2475
      Width           =   1455
   End
   Begin VB.ListBox Lista 
      Height          =   2205
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   11220
   End
   Begin VB.Line Line1 
      X1              =   135
      X2              =   11340
      Y1              =   3555
      Y2              =   3555
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
dialog.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly Or cdlOFNAllowMultiselect Or cdlOFNExplorer
dialog.MaxFileSize = 32000

On Local Error Resume Next
dialog.ShowOpen

If Err.Number <> 0 Then Exit Sub

Dim tss As String
Dim tss2 As String
Dim i As Integer, j As Integer, d As Integer
Dim Matriz()

If InStr(1, dialog.FileName, Chr(0)) = 0 Then
    ReDim Matriz(1 To 1)
    Matriz(1) = dialog.FileName
    d = 1
    tss2 = APath(dialog.FileName)
    GoTo 4
End If

Dim ant As Long
For i = Len(dialog.FileName) To 1 Step -1
    DoEvents
    If Mid(dialog.FileName, i, 1) = "\" Then
        tss = Left(dialog.FileName, i - 1)
        d = i
        Exit For
    End If
Next
For i = d To Len(dialog.FileName)
    DoEvents
    If Mid(dialog.FileName, i, 1) = Chr(0) Then
        tss2 = tss & Mid(dialog.FileName, d, i - d)
        Exit For
    End If
Next
If Right(tss2, 1) <> "\" Then tss2 = tss2 & "\"
ant = i + 1
For d = (i + 1) To Len(dialog.FileName)
    DoEvents
    If Mid(dialog.FileName, d, 1) = Chr(0) Or d = Len(dialog.FileName) Then
        j = j + 1
        ReDim Preserve Matriz(1 To j)
        If d <> Len(dialog.FileName) Then
            Matriz(j) = tss2 & Mid(dialog.FileName, ant, d - ant)
        Else
            Matriz(j) = tss2 & Mid(dialog.FileName, ant)
        End If
        ant = d + 1
    End If
Next
d = j

4
For j = 1 To d
    If NombreR(Matriz(j)) = "" Then GoTo 30
    DoEvents
    If FileLen(Matriz(j)) = 0 Then GoTo 30
    Lista.AddItem Matriz(j)
30
Next

End Sub

Private Sub clear_Click()
Lista.clear
End Sub

Private Sub Create_Click()
Dim ff As Integer, x As Integer, ff2 As Integer, cadena As String
Dim y As Long, largo As Long, car As Byte
Dim cars() As Byte, max As Integer, cad1 As String, cad2 As String, tam As Long
ff = FreeFile
On Local Error Resume Next
Kill SaveF.Text
On Local Error GoTo 0
Open SaveF.Text For Binary As #ff
Put #ff, , "PACKEDDATAFILE"   '14 chars
largo = Lista.ListCount

Put #ff, , largo
For x = 0 To Lista.ListCount - 1
    ff2 = FreeFile
    largo = Len(NombreArchivo(Lista.List(x)))
    Put #ff, , largo
    Put #ff, , Cifrar(NombreArchivo(Lista.List(x)))
    tam = FileLen(Lista.List(x))
    Open Lista.List(x) For Binary As #ff2
        Put #ff, , tam
        max = 1127
        If max > tam Then max = tam
        
        cad1 = Space(max)
        Get #ff2, , cad1
        
        If (tam - 1127) > 0 Then
            cad2 = Space(tam - 1127)
            Get #ff2, , cad2
        Else
            cad2 = ""
        End If
        cad1 = Cifrar(cad1)
        Put #ff, , cad2
        Put #ff, , cad1
    Close #ff2
    b.Width = bm.Width / Lista.ListCount * (x + 1)
    DoEvents
Next
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

MkDir "C:\tmp\"
ExtractFile dialog.FileName, "C:\tmp"
End Sub
