VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "UserForm2"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4230
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#Else
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If

Const GWL_STYLE = -16
Const WS_CAPTION = &HC00000

'UserForm position
Dim FormX As Double, FormY As Double

Private Sub UserForm_Initialize()
    Dim lngWindow As Long, lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, Me.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    Call SetWindowLong(lFrmHdl, GWL_STYLE, lngWindow)
    Call DrawMenuBar(lFrmHdl)
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'   Starting position when button is pressed
    If Button = 1 Then
        FormX = X
        FormY = Y
    End If
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'   Move the form if the left button is pressed
    If Button = 1 Then
        Me.Left = Me.Left + (X - FormX)
        Me.Top = Me.Top + (Y - FormY)
    End If
End Sub

Private Sub OKButton_Click()
    Unload Me
End Sub


