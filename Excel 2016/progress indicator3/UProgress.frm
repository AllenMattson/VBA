VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UProgress 
   Caption         =   "Show Steps"
   ClientHeight    =   2760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "UProgress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub SetDescription(Description As String)
    Me.lblDescription.Caption = Description
End Sub

Public Sub AddStep(sStep As String)
    With Me.lbxSteps
        .AddItem sStep
        .TopIndex = Application.Max(.ListCount, .ListCount - 6)
    End With
    Me.Repaint
End Sub

