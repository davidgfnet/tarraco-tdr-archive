VERSION 5.00
Begin VB.Form resolucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resolución"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5280
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton f4 
      Caption         =   "Pretori"
      Height          =   285
      Left            =   4140
      TabIndex        =   7
      Top             =   1350
      Width           =   870
   End
   Begin VB.OptionButton f3 
      Caption         =   "Recinte de culte"
      Height          =   285
      Left            =   2385
      TabIndex        =   6
      Top             =   1350
      Width           =   1635
   End
   Begin VB.OptionButton f2 
      Caption         =   "Forum"
      Height          =   285
      Left            =   1305
      TabIndex        =   5
      Top             =   1350
      Width           =   960
   End
   Begin VB.OptionButton f1 
      Caption         =   "Circ"
      Height          =   285
      Left            =   270
      TabIndex        =   4
      Top             =   1350
      Value           =   -1  'True
      Width           =   870
   End
   Begin VB.ComboBox antialias 
      Height          =   315
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   765
      Width           =   3300
   End
   Begin VB.CommandButton ok1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   315
      TabIndex        =   1
      Top             =   1935
      Width           =   4695
   End
   Begin VB.ComboBox combo 
      Height          =   315
      Left            =   270
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   270
      Width           =   4740
   End
   Begin VB.Label Label1 
      Caption         =   "Antialiasing level:"
      Height          =   240
      Left            =   270
      TabIndex        =   3
      Top             =   810
      Width           =   2265
   End
End
Attribute VB_Name = "resolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Ok1_Click()
ok = True
If f1 = True Then
    Zone = 1
ElseIf f2 = True Then
    Zone = 2
ElseIf f3 = True Then
    Zone = 3
Else
    Zone = 4
End If
ALevel = Val(antialias.Text)
DModeSelected = combo.ListIndex
Unload Me
End Sub

