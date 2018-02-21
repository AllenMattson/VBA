VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Zoomable UserForm"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3915
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim StartW, StartH

Private Sub UserForm_Initialize()
    StartW = Me.Width
    StartH = Me.Height
End Sub

Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub NormalButton_Click()
    ScrollBar1.Value = 100
End Sub

Private Sub ScrollBar1_Change()
    Me.Zoom = ScrollBar1.Value
    Me.Width = StartW * (ScrollBar1.Value / 100)
    Me.Height = StartH * (ScrollBar1.Value / 100)
    LabelZoom.Caption = ScrollBar1.Value & "%"
End Sub

