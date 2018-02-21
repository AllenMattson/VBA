VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Resizable UserForm"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4575
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MResizer = "ResizeGrab"
Private WithEvents objResizer As MSForms.Label
Attribute objResizer.VB_VarHelpID = -1
Private LeftResizePos As Single
Private TopResizePos As Single

Private Sub UserForm_Initialize()
'   add a resizing control to bottom right corner of userform
    Set objResizer = Me.Controls.Add("Forms.label.1", MResizer, True)
    With objResizer
        With .Font
            .Name = "Marlett"
            .Charset = 2
            .Size = 16
            .Bold = True
        End With
        .BackStyle = fmBackStyleTransparent
        .AutoSize = True
        .BorderStyle = fmBorderStyleNone
        .Caption = "o"
        .MousePointer = fmMousePointerSizeNWSE
        .ForeColor = RGB(100, 100, 100)
        .ZOrder
        .Top = Me.InsideHeight - .Height
        .Left = Me.InsideWidth - .Width
    End With
End Sub

Private Sub objResizer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        LeftResizePos = X
        TopResizePos = Y
    End If
End Sub

Private Sub objResizer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then
        With objResizer
            .Move .Left + X - LeftResizePos, .Top + Y - TopResizePos
            Me.Width = Me.Width + X - LeftResizePos
            Me.Height = Me.Height + Y - TopResizePos
            .Left = Me.InsideWidth - .Width
            .Top = Me.InsideHeight - .Height
        End With

'       Adjust the ListBox
        On Error Resume Next
        With ListBox1
            .Width = Me.Width - 22
            .Height = Me.Height - 100
        End With
        On Error GoTo 0

'       Adjust the Close Button
        With CloseButton
            .Left = Me.Width - 70
            .Top = Me.Height - 54
        End With
    End If
End Sub
Private Sub CloseButton_Click()
    Unload Me
End Sub

