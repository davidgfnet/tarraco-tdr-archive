Attribute VB_Name = "Module1"
Option Explicit

Dim m_TestCol As TestCollection1
Dim m_TestItem As TestItem1

Public Sub Main()
    Set m_TestCol = New TestCollection1
    
    'add 5 test items to the test collection...
    Dim i As Integer
    For i = 1 To 1
        Set m_TestItem = m_TestCol.CreateItem
        'm_TestItem.X = i & "1"
        'm_TestItem.Y = i & "2"
        'm_TestItem.Z = i & "3"
        'm_TestItem.RotY = i & "4"
        'm_TestItem.Simple = i & "5"
        'm_TestItem.Complex = i & "6"
        
        m_TestCol.Add m_TestItem
    Next

    'edit the collection in a grid...
    Dim frmGrid As Form1
    Set frmGrid = New Form1
    Set frmGrid.ThisCollection = m_TestCol
    frmGrid.Show
    
End Sub
