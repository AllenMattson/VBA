VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Application Event Monitor"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MResizer = "ResizeGrab"
Private WithEvents objResizer As MSForms.Label
Attribute objResizer.VB_VarHelpID = -1
Private LeftResizePos As Single
Private TopResizePos As Single

Private Sub CancelButton_Click()
    Call StopTrackingEvents
End Sub


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
'   For resizing form
    If Button = 1 Then
        LeftResizePos = X
        TopResizePos = Y
    End If
End Sub

Private Sub objResizer_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'   For resizing form
    If Button = 1 Then
        With objResizer
            .Move .Left + X - LeftResizePos, .Top + Y - TopResizePos
            Me.Height = Me.Height + Y - TopResizePos
            .Left = Me.InsideWidth - .Width
            .Top = Me.InsideHeight - .Height
        End With
'       Adjust the Frame
        On Error Resume Next
        FrameEvents.Height = Me.Height - 66
        
'       Adjust the buttons
        MarkButton.Top = FrameEvents.Height + 14
        CancelButton.Top = FrameEvents.Height + 14

        On Error GoTo 0
    End If
End Sub


Private Sub MarkButton_Click()
    EventNum = EventNum + 1
    With UserForm1
        .lblEvents.AutoSize = False
        .lblEvents.Caption = .lblEvents.Caption & vbCrLf & String(40, "-")
        .lblEvents.Width = .FrameEvents.Width - 20
        .lblEvents.AutoSize = True
        .FrameEvents.ScrollHeight = .lblEvents.Height + 20
        .FrameEvents.ScrollTop = EventNum * 20
    End With
End Sub
