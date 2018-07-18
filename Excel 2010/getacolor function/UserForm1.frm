VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Color Picker"
   ClientHeight    =   2325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
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
'   Get user's previous colors
    ScrollBarRed.Value = GetSetting(APPNAME, "Colors", "RedValue", 128)
    ScrollBarGreen.Value = GetSetting(APPNAME, "Colors", "GreenValue", 128)
    ScrollBarBlue.Value = GetSetting(APPNAME, "Colors", "BlueValue", 128)
    Call ScrollBarRed_Change
    Call ScrollBarGreen_Change
    Call ScrollBarBlue_Change
End Sub

Private Sub CancelButton_Click()
    ColorValue = False
    Unload Me
End Sub

Private Sub cbWhiteText_Click()
    If cbWhiteText Then
        SampleLabel.ForeColor = RGB(255, 255, 255)
    Else
        SampleLabel.ForeColor = RGB(0, 0, 0)
    End If
End Sub

Private Sub SelectColorButton_Click()
'   Assign the value and close the dialog
    ColorValue = SampleLabel.BackColor
    SaveSetting APPNAME, "Colors", "RedValue", ScrollBarRed.Value
    SaveSetting APPNAME, "Colors", "BlueValue", ScrollBarBlue.Value
    SaveSetting APPNAME, "Colors", "GreenValue", ScrollBarGreen.Value
    Unload Me
End Sub


Private Sub ScrollBarRed_Change()
    LabelRed.BackColor = RGB(ScrollBarRed.Value, 0, 0)
    Call UpdateColor
End Sub

Private Sub ScrollBarGreen_Change()
    LabelGreen.BackColor = RGB(0, ScrollBarGreen.Value, 0)
    Call UpdateColor
End Sub

Private Sub ScrollBarBlue_Change()
    LabelBlue.BackColor = RGB(0, 0, ScrollBarBlue.Value)
    Call UpdateColor
End Sub

Private Sub UpdateColor()
    SampleLabel.BackColor = RGB(ScrollBarRed.Value, ScrollBarGreen.Value, ScrollBarBlue.Value)
    LabelRGB.Caption = ScrollBarRed & ", " & ScrollBarGreen & ", " & ScrollBarBlue
End Sub


