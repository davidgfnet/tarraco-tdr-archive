VERSION 5.00
Begin VB.Form resolucion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resolución"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5280
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ok1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   465
      Left            =   315
      TabIndex        =   1
      Top             =   900
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
End
Attribute VB_Name = "resolucion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Ok1_Click()
ok = True
DModeSelected = combo.ListIndex
Unload Me
End Sub

