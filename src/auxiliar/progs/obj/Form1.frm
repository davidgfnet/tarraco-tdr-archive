VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8820
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   ScaleHeight     =   8820
   ScaleWidth      =   9555
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtrep 
      Height          =   375
      Left            =   6705
      TabIndex        =   22
      Top             =   7155
      Width           =   2535
   End
   Begin VB.TextBox find 
      Height          =   375
      Left            =   6705
      TabIndex        =   20
      Top             =   6660
      Width           =   2535
   End
   Begin VB.ComboBox what 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   5670
      List            =   "Form1.frx":0016
      TabIndex        =   19
      Text            =   "Combo1"
      Top             =   6075
      Width           =   3570
   End
   Begin VB.CommandButton rep 
      Caption         =   "replace"
      Height          =   420
      Left            =   8100
      TabIndex        =   18
      Top             =   7740
      Width           =   1185
   End
   Begin VB.CommandButton copyc 
      Caption         =   "COPY"
      Height          =   510
      Left            =   5670
      TabIndex        =   17
      Top             =   4500
      Width           =   1680
   End
   Begin VB.TextBox roty 
      Height          =   375
      Left            =   6795
      TabIndex        =   15
      Top             =   3870
      Width           =   2535
   End
   Begin VB.CommandButton del 
      Caption         =   "DEL"
      Height          =   510
      Left            =   7515
      TabIndex        =   14
      Top             =   4500
      Width           =   1680
   End
   Begin MSComDlg.CommonDialog dialog 
      Left            =   3915
      Top             =   3915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton save 
      Caption         =   "save"
      Height          =   600
      Left            =   7515
      TabIndex        =   13
      Top             =   5130
      Width           =   1635
   End
   Begin VB.CommandButton open 
      Caption         =   "open"
      Height          =   600
      Left            =   5760
      TabIndex        =   12
      Top             =   5130
      Width           =   1590
   End
   Begin VB.TextBox simple 
      Height          =   375
      Left            =   6795
      TabIndex        =   10
      Top             =   3330
      Width           =   2535
   End
   Begin VB.TextBox complex 
      Height          =   375
      Left            =   6795
      TabIndex        =   8
      Top             =   2790
      Width           =   2535
   End
   Begin VB.TextBox CoZ 
      Height          =   375
      Left            =   6795
      TabIndex        =   6
      Top             =   2025
      Width           =   2535
   End
   Begin VB.TextBox CoY 
      Height          =   375
      Left            =   6795
      TabIndex        =   4
      Top             =   1530
      Width           =   2535
   End
   Begin VB.TextBox CoX 
      Height          =   375
      Left            =   6795
      TabIndex        =   2
      Top             =   1035
      Width           =   2535
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   555
      Left            =   5535
      TabIndex        =   1
      Top             =   225
      Width           =   3795
   End
   Begin MSComctlLib.TreeView tree 
      Height          =   8385
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   14790
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label Label8 
      Caption         =   "Replace:"
      Height          =   285
      Left            =   5490
      TabIndex        =   23
      Top             =   7200
      Width           =   1185
   End
   Begin VB.Label Label7 
      Caption         =   "Find:"
      Height          =   285
      Left            =   5490
      TabIndex        =   21
      Top             =   6705
      Width           =   1185
   End
   Begin VB.Label Label6 
      Caption         =   "Rotation Y:"
      Height          =   285
      Left            =   5580
      TabIndex        =   16
      Top             =   3915
      Width           =   1185
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   5580
      X2              =   9270
      Y1              =   2610
      Y2              =   2610
   End
   Begin VB.Label Label5 
      Caption         =   "Simple file:"
      Height          =   285
      Left            =   5580
      TabIndex        =   11
      Top             =   3375
      Width           =   1185
   End
   Begin VB.Label Label4 
      Caption         =   "Complex file:"
      Height          =   285
      Left            =   5580
      TabIndex        =   9
      Top             =   2835
      Width           =   1185
   End
   Begin VB.Label Label3 
      Caption         =   "Coord Z:"
      Height          =   285
      Left            =   5580
      TabIndex        =   7
      Top             =   2070
      Width           =   1185
   End
   Begin VB.Label Label2 
      Caption         =   "Coord Y:"
      Height          =   285
      Left            =   5580
      TabIndex        =   5
      Top             =   1575
      Width           =   1185
   End
   Begin VB.Label Label1 
      Caption         =   "Coord X:"
      Height          =   285
      Left            =   5580
      TabIndex        =   3
      Top             =   1080
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Obj
    complex As String
    simple As String
    x As Single
    Y As Single
    z As Single
    roty As Single
    used As Boolean
End Type

Dim Objects() As Obj

Private Sub Add_Click()
Dim x As Integer
For x = 1 To UBound(Objects)
    If Objects(x).used = False Then
        Objects(x).used = True
        Objects(x).x = 0
        Objects(x).Y = 0
        Objects(x).z = 0
        Objects(x).simple = ""
        Objects(x).complex = ""
        Objects(x).roty = 0
        tree.Nodes.Add , , "k" & Trim(Str(x)), "NEW OBJ"
        Exit Sub
    End If
Next
ReDim Preserve Objects(UBound(Objects) + 1)
tree.Nodes.Add , , "k" & Trim(Str(UBound(Objects))), "NEW OBJ"
Objects(UBound(Objects)).used = True
End Sub

Private Sub complex_GotFocus()
complex.SelStart = 0
complex.SelLength = Len(complex.Text)
End Sub

Private Sub complex_Validate(Cancel As Boolean)
CH
End Sub

Private Sub copyc_Click()
Dim x As Integer
For x = 1 To UBound(Objects)
    If Objects(x).used = False Then
        Objects(x).used = True
        Objects(x).x = Val(CoX.Text)
        Objects(x).Y = Val(CoY.Text)
        Objects(x).z = Val(coxz.Text)
        Objects(x).simple = simple.Text
        Objects(x).complex = complex.Text
        Objects(x).roty = Val(roty.Text)
        tree.Nodes.Add , , "k" & Trim(Str(x)), "NEW OBJ"
        Exit Sub
    End If
Next
ReDim Preserve Objects(UBound(Objects) + 1)
tree.Nodes.Add , , "k" & Trim(Str(UBound(Objects))), "NEW OBJ"
Objects(UBound(Objects)).used = True
x = UBound(Objects)
Objects(x).used = True
Objects(x).x = Val(CoX.Text)
Objects(x).Y = Val(CoY.Text)
Objects(x).z = Val(CoZ.Text)
Objects(x).simple = simple.Text
Objects(x).complex = complex.Text
Objects(x).roty = Val(roty.Text)
End Sub

Private Sub CoX_GotFocus()
CoX.SelStart = 0
CoX.SelLength = Len(CoX.Text)
End Sub

Private Sub CoX_Validate(Cancel As Boolean)
CH
End Sub

Private Sub CoY_GotFocus()
CoY.SelStart = 0
CoY.SelLength = Len(CoY.Text)
End Sub

Private Sub CoY_Validate(Cancel As Boolean)
CH
End Sub

Private Sub CoZ_GotFocus()
CoZ.SelStart = 0
CoZ.SelLength = Len(CoZ.Text)
End Sub

Private Sub CoZ_Validate(Cancel As Boolean)
CH
End Sub

Private Sub del_Click()
Objects(Val(Mid(tree.SelectedItem.Key, 2))).x = 0
Objects(Val(Mid(tree.SelectedItem.Key, 2))).Y = 0
Objects(Val(Mid(tree.SelectedItem.Key, 2))).z = 0
Objects(Val(Mid(tree.SelectedItem.Key, 2))).roty = 0

Objects(Val(Mid(tree.SelectedItem.Key, 2))).complex = ""
Objects(Val(Mid(tree.SelectedItem.Key, 2))).simple = ""

Objects(Val(Mid(tree.SelectedItem.Key, 2))).used = False

tree.Nodes.Remove tree.SelectedItem.Index
End Sub

Private Sub Form_Load()
ReDim Objects(0)
dialog.Filter = "Data files|*.dat"
End Sub

Private Sub open_Click()
On Local Error Resume Next
dialog.DialogTitle = "Open Fixed Object Dat File"
dialog.ShowOpen

If Err.Number <> 0 Then Exit Sub

Dim f As Integer
Dim sin As Single, cad As String, lon As Long, x As Long, lon2 As Long
f = FreeFile
tree.Nodes.Clear
Open dialog.FileName For Binary As #f
    cad = Space(Len("FODATAFILE"))
    Get #f, , cad
    Get #f, , lon
    ReDim Objects(lon)
    For x = 1 To lon
        Get #f, , sin
        Objects(x).x = sin
        Get #f, , sin
        Objects(x).Y = sin
        Get #f, , sin
        Objects(x).z = sin
        Get #f, , sin
        Objects(x).roty = sin
        Get #f, , lon2
        cad = Space(lon2)
        Get #f, , cad
        Objects(x).complex = cad
        Get #f, , lon2
        cad = Space(lon2)
        Get #f, , cad
        Objects(x).simple = cad
        Objects(x).used = True
        tree.Nodes.Add , , "k" & Trim(Str(x)), Objects(x).complex & " - " & Objects(x).simple & " //" & Objects(x).x & "/" & Objects(x).Y & "/" & Objects(x).z
    Next
Close #f
End Sub


Private Sub rep_Click()
Dim x As Long
For x = 1 To UBound(Objects)
    Select Case what
    Case "x"
        If Objects(x).x = Val(find.Text) Then Objects(x).x = Val(txtrep.Text)
    Case "y"
        If Objects(x).Y = Val(find.Text) Then Objects(x).Y = Val(txtrep.Text)
    Case "z"
        If Objects(x).z = Val(find.Text) Then Objects(x).x = Val(txtrep.Text)
    Case "roty"
        If Objects(x).roty = Val(find.Text) Then Objects(x).roty = Val(txtrep.Text)
    Case "simple"
        If Objects(x).simple = find.Text Then Objects(x).simple = txtrep.Text
    Case "complex"
        If Objects(x).complex = find.Text Then Objects(x).complex = txtrep.Text
    End Select
Next
End Sub

Private Sub roty_GotFocus()
roty.SelStart = 0
roty.SelLength = Len(roty.Text)
End Sub

Private Sub roty_Validate(Cancel As Boolean)
CH
End Sub

Private Sub save_Click()
dialog.DialogTitle = "Save FIXED OBJ DATA FILE"

On Local Error Resume Next
dialog.ShowSave

Dim f As Integer, x As Long
f = FreeFile

Open dialog.FileName For Binary As #f
    Put #f, , "FODATAFILE"
    Put #f, , CLng(tree.Nodes.Count)
    For x = 1 To UBound(Objects)
        If Objects(x).used = True Then
            Put #f, , CSng(Objects(x).x)
            Put #f, , CSng(Objects(x).Y)
            Put #f, , CSng(Objects(x).z)
            Put #f, , CSng(Objects(x).roty)
            
            Put #f, , CLng(Len(Objects(x).complex))
            Put #f, , Objects(x).complex
            
            Put #f, , CLng(Len(Objects(x).simple))
            Put #f, , Objects(x).simple
        End If
    Next
Close #f
End Sub

Private Sub simple_GotFocus()
simple.SelStart = 0
simple.SelLength = Len(simple.Text)
End Sub

Private Sub simple_Validate(Cancel As Boolean)
CH
End Sub

Private Sub tree_GotFocus()
CH
End Sub

Private Sub tree_NodeClick(ByVal Node As MSComctlLib.Node)
CoX.Text = Replace(Objects(Val(Mid(Node.Key, 2))).x, ",", ".")
CoY.Text = Replace(Objects(Val(Mid(Node.Key, 2))).Y, ",", ".")
CoZ.Text = Replace(Objects(Val(Mid(Node.Key, 2))).z, ",", ".")

roty.Text = Objects(Val(Mid(Node.Key, 2))).roty

simple.Text = Objects(Val(Mid(Node.Key, 2))).simple
complex.Text = Objects(Val(Mid(Node.Key, 2))).complex
CH
End Sub

Private Sub CH()
Dim x As Long
If tree.SelectedItem Is Nothing Then
Else
x = Val(Mid(tree.SelectedItem.Key, 2))
Objects(Val(Mid(tree.SelectedItem.Key, 2))).x = Val(CoX.Text)
Objects(Val(Mid(tree.SelectedItem.Key, 2))).Y = Val(CoY.Text)
Objects(Val(Mid(tree.SelectedItem.Key, 2))).z = Val(CoZ.Text)

Objects(Val(Mid(tree.SelectedItem.Key, 2))).roty = Val(roty.Text)

Objects(Val(Mid(tree.SelectedItem.Key, 2))).complex = complex.Text
Objects(Val(Mid(tree.SelectedItem.Key, 2))).simple = simple.Text

tree.SelectedItem.Text = Objects(x).complex & " - " & Objects(x).simple & " //" & Objects(x).x & "/" & Objects(x).Y & "/" & Objects(x).z
End If
End Sub
