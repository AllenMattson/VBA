VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm Posing As A Toolbar"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   7770
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
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

'UserForm position
Dim FormX As Double, FormY As Double

Private Sub UserForm_Initialize()
    Dim lngWindow As Long, lFrmHdl As Long
    lFrmHdl = FindWindowA(vbNullString, Me.Caption)
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    Call SetWindowLong(lFrmHdl, -20, 385)
    Call DrawMenuBar(lFrmHdl)
    Call SetupControls
End Sub

Sub SetupControls()
    Dim ctl As Control
    Dim leftPos As Integer
'   Adjust spacing
    leftPos = 2
    For Each ctl In Me.Controls
        ctl.Top = 2
        ctl.Left = leftPos
        leftPos = leftPos + ctl.Width + 3
    Next ctl
    With Me
        .Width = leftPos + 4
        .Height = 52
    End With
End Sub

Private Sub Image1_Click()
    Call Macro1
End Sub

Private Sub Image2_Click()
    Call Macro2
End Sub

Private Sub Image3_Click()
    Call Macro3
End Sub

Private Sub Image4_Click()
    Call Macro4
End Sub

Private Sub Image5_Click()
    Call Macro5
End Sub

Private Sub Image6_Click()
    Call Macro6
End Sub

Private Sub Image7_Click()
    Call Macro7
End Sub

Private Sub Image8_Click()
    Call Macro8
End Sub

''''''''''
'The event-handlers below are for the mouse-over effects

Private Sub NoRaise()
'   Remove the raised effect from all controls
    Dim ctl As Control
    For Each ctl In Controls
        ctl.SpecialEffect = fmSpecialEffectFlat
    Next ctl
End Sub

Private Sub Image1_MouseMove(ByVal Button As Integer, _
  ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoRaise
    Image1.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub Image2_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoRaise
    Image2.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub Image3_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoRaise
    Image3.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub Image4_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoRaise
    Image4.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub Image5_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoRaise
    Image5.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub Image6_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoRaise
    Image6.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub Image7_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoRaise
    Image7.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub Image8_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoRaise
    Image8.SpecialEffect = fmSpecialEffectRaised
End Sub
Private Sub Userform_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoRaise
End Sub
