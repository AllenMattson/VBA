VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UProgress 
   Caption         =   "Progress"
   ClientHeight    =   1155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4320
   OleObjectBlob   =   "UProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub UpdateProgress(PctDone As Double)
    With Me
        .frmProgress.Caption = Format(PctDone, "0%")
        .lblProgress.Width = PctDone * (.frmProgress.Width - 10)
        .Repaint
    End With
End Sub

Public Sub SetDescription(Description As String)
    Me.lblDescription.Caption = Description
End Sub

Private Sub UserForm_Initialize()
    With Me
        'Use a color from the workbook's theme
        .lblProgress.BackColor = ActiveWorkbook.Theme. _
            ThemeColorScheme.Colors(msoThemeAccent1)
        '.lblProgress.BackColor = vbRed
        .lblProgress.Width = 0
    End With
End Sub
