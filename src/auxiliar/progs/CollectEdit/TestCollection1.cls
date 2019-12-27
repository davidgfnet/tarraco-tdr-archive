VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestCollection1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"TestItem"
Attribute VB_Ext_KEY = "Member0" ,"TestItem"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function CreateItem() As TestItem1
    
    'create a new object
    Dim objNewMember As TestItem1
    Set objNewMember = New TestItem1

    'return the object created
    Set CreateItem = objNewMember
    Set objNewMember = Nothing

End Function

Public Sub Add(TempObject As TestItem1, Optional lBefore As Long, Optional lAfter As Long)
  If lBefore = 0 And lAfter = 0 Then
    mCol.Add TempObject
  ElseIf lBefore > 0 And lAfter = 0 Then
    mCol.Add TempObject, , lBefore
  ElseIf lBefore = 0 And lAfter > 0 Then
    mCol.Add TempObject, , , lAfter
  End If
End Sub

Public Property Get Item(vntIndexKey As Variant) As TestItem1
Attribute Item.VB_UserMemId = 0
    Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub

