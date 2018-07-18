VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UGetAColor 
   Caption         =   "Color Picker"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   OleObjectBlob   =   "UGetAColor.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UGetAColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ColorValue As Variant

Private Sub UserForm_Initialize()
'   Get user's previous colors
    Me.scbRed.Value = GetSetting(APPNAME, "Colors", "RedValue", 128)
    Me.scbGreen.Value = GetSetting(APPNAME, "Colors", "GreenValue", 128)
    Me.scbBlue.Value = GetSetting(APPNAME, "Colors", "BlueValue", 128)
    scbRed_Change
    scbGreen_Change
    scbBlue_Change
End Sub

Private Sub cmdCancel_Click()
    Me.ColorValue = False
    Me.Hide
End Sub

Private Sub chkWhiteText_Click()
    If Me.chkWhiteText.Value Then
        Me.lblSample.ForeColor = RGB(255, 255, 255)
    Else
        Me.lblSample.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub cmdSelect_Click()
'   Assign the value and close the dialog
    Me.ColorValue = Me.lblSample.BackColor
    SaveSetting APPNAME, "Colors", "RedValue", Me.scbRed.Value
    SaveSetting APPNAME, "Colors", "BlueValue", Me.scbBlue.Value
    SaveSetting APPNAME, "Colors", "GreenValue", Me.scbGreen.Value
    Me.Hide
End Sub

Private Sub scbRed_Change()
    Me.lblRed.BackColor = RGB(Me.scbRed.Value, 0, 0)
    UpdateColor
End Sub

Private Sub scbGreen_Change()
    Me.lblGreen.BackColor = RGB(0, Me.scbGreen.Value, 0)
    UpdateColor
End Sub

Private Sub scbBlue_Change()
    Me.lblBlue.BackColor = RGB(0, 0, Me.scbBlue.Value)
    UpdateColor
End Sub

Private Sub UpdateColor()
    Me.lblSample.BackColor = RGB(Me.scbRed.Value, Me.scbGreen.Value, Me.scbBlue.Value)
    Me.lblRGB.Caption = Me.scbRed.Value & ", " & Me.scbGreen.Value & ", " & Me.scbBlue.Value
End Sub


