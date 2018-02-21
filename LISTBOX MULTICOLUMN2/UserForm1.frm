VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Multicolumn ListBox Demo"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    Dim i As Integer
    Dim Data(1 To 12, 1 To 2) As Variant
    Dim TheYr As Integer
    TheYr = Year(Now)
'   Fill the list box
    For i = 1 To 12
        Data(i, 1) = Format(DateSerial(TheYr, i, 1), "mmmm")
    Next i
    For i = 1 To 12
        Data(i, 2) = Day(DateSerial(TheYr, i + 1, 1) - 1)
    Next i
    ListBox1.ColumnCount = 2
    ListBox1.List = Data
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub



