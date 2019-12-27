VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Edit Collection"
   ClientHeight    =   5040
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   7575
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.Toolbar tbrMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imlToolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cut"
            Object.ToolTipText     =   "Cut (Ctrl+X)"
            ImageKey        =   "cut"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copy"
            Object.ToolTipText     =   "Copy (Ctrl+C)"
            ImageKey        =   "copy"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Paste"
            Object.ToolTipText     =   "Paste (Ctrl+V)"
            ImageKey        =   "paste"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete (Del)"
            ImageKey        =   "delete"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Reload"
            Object.ToolTipText     =   "Reload Collection"
            ImageKey        =   "undo"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "InsertRow"
            Object.ToolTipText     =   "Insert Row"
            ImageKey        =   "insert"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "DelRow"
            Object.ToolTipText     =   "Remove Row"
            ImageKey        =   "delrow"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   3732
      Left            =   0
      ScaleHeight     =   3735
      ScaleWidth      =   7575
      TabIndex        =   1
      Top             =   360
      Width           =   7575
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4920
         TabIndex        =   2
         Top             =   2040
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2172
         Left            =   120
         TabIndex        =   3
         ToolTipText     =   "For Input: Doubleclick, Enter, or just start typing..."
         Top             =   120
         Width           =   3612
         _ExtentX        =   6376
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   12
      End
      Begin MSComDlg.CommonDialog dialog 
         Left            =   5040
         Top             =   855
         _ExtentX        =   688
         _ExtentY        =   688
         _Version        =   393216
         Filter          =   "Data files|*.dat"
      End
      Begin MSComctlLib.ImageList imlToolbar 
         Left            =   5640
         Top             =   840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0000
               Key             =   "insert"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0552
               Key             =   "delrow"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0664
               Key             =   "copy"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":0BB6
               Key             =   "paste"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1108
               Key             =   "cut"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":165A
               Key             =   "delete"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form1.frx":1BAC
               Key             =   "undo"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditcut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSelectAll 
         Caption         =   "Select all"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditInsertRow 
         Caption         =   "Insert Row"
      End
      Begin VB.Menu mnuEditDelRow 
         Caption         =   "Remove Row"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Type Tip
    X As Single
    Y As Single
    Z As Single
    RotY As Single
    len1 As Long
    Complex As String
    len2 As Long
    Simple As String
End Type

Dim data As Tip

Private m_Col As Object

Public Property Get ThisCollection() As Object
    If m_Col Is Nothing Then
        Set ThisCollection = Nothing
    Else
        Set ThisCollection = m_Col
    End If
End Property

Public Property Set ThisCollection(ByVal vNewValue As Object)
    If vNewValue Is Nothing Then
        Set m_Col = Nothing
    Else
        Set m_Col = vNewValue
        Reload
    End If
End Property

Public Sub Reload()
    Dim X As Integer
  
    Dim TLI As TLIApplication
    Dim Interface As InterfaceInfo
    Dim oSR As SearchResults
    Dim oSI As SearchItem
    Dim iCol As Integer
    Dim iRow As Integer
    Dim sString As String
    Dim obj As Object
    Dim i, n As Long
    
    If m_Col.Count = 0 Then Exit Sub
    Set obj = m_Col.Item(1)
    Set TLI = New TLIApplication
    
    i = MSFlexGrid1.Row
    n = MSFlexGrid1.Col
    
 'HEADERS...
    If Not obj Is Nothing Then
        Set Interface = TLI.InterfaceInfoFromObject(obj)
        Set oSR = Interface.Members.GetFilteredMembers
        For Each oSI In oSR
            If oSI.InvokeKinds = INVOKE_PROPERTYPUT + INVOKE_PROPERTYGET Then
                iCol = iCol + 1
                MSFlexGrid1.Cols = iCol + 1
                sString = oSI.Name
                MSFlexGrid1.TextMatrix(0, iCol) = sString
                MSFlexGrid1.TextMatrix(0, 0) = "Item"
            End If
        Next oSI
    End If

'ITEMS...
    iRow = 0
    For Each obj In m_Col
        iRow = iRow + 1
        MSFlexGrid1.Rows = iRow + 1
        iCol = 0
        Set Interface = TLI.InterfaceInfoFromObject(obj)
        Set oSR = Interface.Members.GetFilteredMembers
        For Each oSI In oSR
            If oSI.InvokeKinds = INVOKE_PROPERTYPUT + INVOKE_PROPERTYGET Then
                iCol = iCol + 1
                sString = TLI.InvokeHook(obj, oSI.Name, INVOKE_PROPERTYGET)
                MSFlexGrid1.TextMatrix(iRow, iCol) = sString
                MSFlexGrid1.TextMatrix(iRow, 0) = iRow
            End If
        Next oSI
    Next obj

    MSFlexGrid1.Row = i
    MSFlexGrid1.Col = n
    
    Set TLI = Nothing
    Set Interface = Nothing
    Set oSR = Nothing
    Set oSI = Nothing
    
End Sub

Private Sub Form_Load()
    MSFlexGrid1.RowHeightMin = Text1.Height
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Picture2.Height = Me.ScaleHeight - Picture2.Top
    
End Sub

Private Sub mnuEditCopy_Click()
    Clipboard.Clear
    Clipboard.SetText MSFlexGrid1.Clip

End Sub

Private Sub mnuEditcut_Click()
    Clipboard.Clear
    Clipboard.SetText MSFlexGrid1.Clip
    Dim i As Integer
    Dim j As Integer
    Dim strClip As String
    With MSFlexGrid1
        For i = 1 To .RowSel
            For j = 1 To .ColSel
                strClip = strClip & "" & vbTab
            Next
            strClip = strClip & vbCr
        Next
        .Clip = strClip
    End With

End Sub

Private Sub mnuEditDelete_Click()
    Dim i As Integer
    Dim j As Integer
    Dim strClip As String
    With MSFlexGrid1
        For i = 1 To .RowSel
            For j = 1 To .ColSel
                strClip = strClip & "" & vbTab
            Next
            strClip = strClip & vbCr
        Next
        .Clip = strClip
    End With

End Sub

Private Sub mnuEditDelRow_Click()
    
    With MSFlexGrid1
 
        m_Col.Remove .Row
        
        If .Rows > 2 Then
            .RemoveItem .Row
        End If
        
        Reload
     
    End With
    
End Sub

Private Sub mnuEditInsertRow_Click()
    Dim NewObject As Object
    
    'add new object to collection
    Set NewObject = m_Col.CreateItem
    m_Col.Add NewObject, MSFlexGrid1.Row
    Set NewObject = Nothing
    
    'add row to grid
    MSFlexGrid1.AddItem "", MSFlexGrid1.Row
    
    Reload
    
End Sub

Private Sub mnuEditPaste_Click()
    If Len(Clipboard.GetText) Then
        MSFlexGrid1.Clip = Clipboard.GetText
    End If
    
    'update each property in selected area...
    Dim i, n  As Long
    With MSFlexGrid1
        For i = .Row To .RowSel
            For n = .Col To .ColSel
                Update i, n
            Next n
        Next i
    End With
    
End Sub

Private Sub mnuEditSelectAll_Click()
    With MSFlexGrid1
        .Visible = False
        .Row = 1
        .Col = 1
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
        .TopRow = 1
        .Visible = True
    End With

End Sub

Private Sub mnuEditUndo_Click()
    MsgBox "Undo"
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub Edit(KeyAscii As Integer)
    
    If MSFlexGrid1.Row > 0 Then
 
        Select Case KeyAscii
            Case 0 To Asc(" ")      'means we double clicked the cell to edit it...
                Text1 = MSFlexGrid1 'so, fill the Text1 box with the info to be edited
                Text1.SelStart = Len(Text1.Text) 'and move cursor to the end...
            Case vbKeyDelete
                MSFlexGrid1 = ""
                Exit Sub 'No need for the edit box...
            Case Else                   'means we just started typing to overwrite the cell contents...
                Text1.Text = Chr(KeyAscii)   'put in first character...
                Text1.SelStart = 1      'put cursor at end (after first character)...
        End Select
    
        'position the edit box
        With Text1
            .FontName = MSFlexGrid1.FontName
            .FontSize = MSFlexGrid1.FontSize
            .Top = MSFlexGrid1.CellTop + MSFlexGrid1.Top
            .Left = MSFlexGrid1.CellLeft + MSFlexGrid1.Left
            .Width = MSFlexGrid1.CellWidth
            .Visible = True
            .ZOrder
            .SetFocus
        End With
        
    End If

End Sub

Private Sub mnuFileOpen_Click()
Set ThisCollection = Nothing
Set ThisCollection = New TestCollection1
Dim m_TestItem As TestItem1

On Local Error Resume Next
dialog.DialogTitle = "Open Fixed Object Dat File"
dialog.ShowOpen

If Err.Number <> 0 Then Exit Sub

Dim f As Integer
Dim sin As Single, cad As String, lon As Long, X As Long, lon2 As Long, inte As Integer
f = FreeFile
Open dialog.FileName For Binary As #f
    cad = Space(Len("FODATAFILE"))
    Get #f, , cad
    Get #f, , lon
    For X = 1 To lon
        Set m_TestItem = Nothing
        'Set m_TestItem = New TestItem1
        Set m_TestItem = ThisCollection.CreateItem
        Get #f, , sin
        m_TestItem.X = sin
        Get #f, , sin
        m_TestItem.Y = sin
        Get #f, , sin
        m_TestItem.Z = sin
        Get #f, , sin
        m_TestItem.RotY = sin
        Get #f, , lon2
        cad = Space(lon2)
        Get #f, , cad
        m_TestItem.Complex = cad
        Get #f, , lon2
        cad = Space(lon2)
        Get #f, , cad
        m_TestItem.Simple = cad
        Get #f, , inte
        m_TestItem.World = inte
        ThisCollection.Add m_TestItem
    Next
Close #f
Reload
End Sub

Private Sub mnuFileSave_Click()
dialog.DialogTitle = "Save FIXED OBJ DATA FILE"

On Local Error Resume Next
dialog.ShowSave

Dim f As Integer, X As Long, cad As String, inte As Integer
f = FreeFile

Open dialog.FileName For Binary As #f
    Put #f, , "FODATAFILE"
    Put #f, , CLng(ThisCollection.Count)
    For X = 1 To ThisCollection.Count
            Put #f, , CSng(ThisCollection(X).X)
            Put #f, , CSng(ThisCollection(X).Y)
            Put #f, , CSng(ThisCollection(X).Z)
            Put #f, , CSng(ThisCollection(X).RotY)
            
            cad = ThisCollection(X).Complex
            Put #f, , CLng(Len(Trim(ThisCollection(X).Complex)))
            Put #f, , cad
            
            cad = ThisCollection(X).Simple
            Put #f, , CLng(Len(Trim(ThisCollection(X).Simple)))
            Put #f, , cad
            inte = ThisCollection(X).World
            Put #f, , inte
    Next
Close #f
End Sub

Private Sub MSFlexGrid1_dblClick()
    Edit Asc(" ")
End Sub

Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)
    'start editing the cell with the first key pressed...
    Edit KeyAscii
End Sub

Private Sub MSFlexGrid1_GotFocus()
    'MSFlexGrid got focus means that we are done with the text box...
    MSFlexGrid1_LeaveCell
End Sub

Private Sub Update(Optional ByVal iRow As Integer, Optional ByVal iCol As Integer)
    'leave cell... here is where we update the info in Col...

    Dim sHeader As String
    Dim NewValue As Variant
    
    If iRow = 0 Then iRow = MSFlexGrid1.Row
    If iCol = 0 Then iCol = MSFlexGrid1.Col
    
    If iRow = 0 Then Exit Sub
    If iCol = 0 Then Exit Sub
    
    If Text1.Visible = True Then
        'Write the Contents of the Textbox into the Grid and hide the Textbox
        MSFlexGrid1.TextMatrix(iRow, iCol) = Text1.Text
        Text1.Visible = False
    End If
    
    sHeader = MSFlexGrid1.TextMatrix(0, iCol)
    NewValue = MSFlexGrid1.TextMatrix(iRow, iCol)
    
    'update property
    Dim obj As Object
    Set obj = m_Col(iRow)
    TLI.InvokeHook obj, sHeader, INVOKE_PROPERTYPUT, NewValue
    

Exit Sub

ERR_ROUTINE:
    MsgBox "Update Error No. " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

Private Sub MSFlexGrid1_LeaveCell()
    Update
End Sub

Private Sub MSFlexGrid1_Scroll()
    'Scrolling MSFlexGrid means that we are done with the text box...
    MSFlexGrid1_LeaveCell
End Sub


Private Sub Picture2_Resize()
    On Error Resume Next
    With MSFlexGrid1
        .Left = 0
        .Top = 0
        .Width = Picture2.ScaleWidth
        .Height = Picture2.ScaleHeight
    End With
End Sub


Private Sub tbrMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Cut"
            mnuEditcut_Click
        Case "Copy"
            mnuEditCopy_Click
        Case "Paste"
            mnuEditPaste_Click
        Case "Delete"
            mnuEditDelete_Click
        Case "Reload"
            Reload
        Case "InsertRow"
            mnuEditInsertRow_Click
        Case "DelRow"
            mnuEditDelRow_Click
        Case "View"
            MsgBox "Print Preview"
        
    End Select
    
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyEscape
            Text1.Visible = False
            MSFlexGrid1.SetFocus
        
        Case vbKeyReturn
            MSFlexGrid1.SetFocus
            Text1_KeyDown vbKeyDown, 1
        
        Case vbKeyDown
            MSFlexGrid1.SetFocus
            DoEvents
            If MSFlexGrid1.Row < MSFlexGrid1.Rows - 1 Then
               MSFlexGrid1.Row = MSFlexGrid1.Row + 1
            End If
        
        Case vbKeyUp
            MSFlexGrid1.SetFocus
            DoEvents
            If MSFlexGrid1.Row > MSFlexGrid1.FixedRows Then
               MSFlexGrid1.Row = MSFlexGrid1.Row - 1
            End If
            
    End Select
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    'noise suppression
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub


