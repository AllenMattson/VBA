VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "ListBox Move Item Demo"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3360
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MoveUpButton_Click()
    Dim NumItems As Integer
    Dim i As Integer
    Dim ItemNum As Integer
    Dim TempItem As String
    Dim TempList()
    
    If ListBox1.ListIndex = 0 Then Exit Sub
        
    NumItems = ListBox1.ListCount
    ReDim TempList(0 To NumItems - 1)
    
'   Fill array with list box items
    For i = 0 To NumItems - 1
        TempList(i) = ListBox1.List(i)
    Next i
    
'   Selected item
    ItemNum = ListBox1.ListIndex

'   Exchange items
    TempItem = TempList(ItemNum)
    TempList(ItemNum) = TempList(ItemNum - 1)
    TempList(ItemNum - 1) = TempItem
    ListBox1.List = TempList
    
'   Change the list index
    ListBox1.ListIndex = ItemNum - 1
End Sub


Private Sub MoveDownButton_Click()
    Dim NumItems As Integer
    Dim i As Integer
    Dim ItemNum As Integer
    Dim TempItem As String
    Dim TempList()
    
    If ListBox1.ListIndex = ListBox1.ListCount - 1 Then Exit Sub
        
    NumItems = ListBox1.ListCount
    ReDim TempList(0 To NumItems - 1)
    
'   Fill array with list box items
    For i = 0 To NumItems - 1
        TempList(i) = ListBox1.List(i)
    Next i
    
'   Selected item
    ItemNum = ListBox1.ListIndex

'   Exchange items
    TempItem = TempList(ItemNum)
    TempList(ItemNum) = TempList(ItemNum + 1)
    TempList(ItemNum + 1) = TempItem
    ListBox1.List = TempList
    
'   Change the list index
    ListBox1.ListIndex = ItemNum + 1
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub

