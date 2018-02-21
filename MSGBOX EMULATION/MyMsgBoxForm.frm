VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} MyMsgBoxForm 
   Caption         =   "Microsoft Excel"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   OleObjectBlob   =   "MyMsgBoxForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "MyMsgBoxForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub UserForm_Initialize()
    Const ButtonWidth As Double = 56
    Const ButtonHeight As Double = 18
    Const ButtonGap As Double = 0
    Const TextTop As Double = 8
    
    Dim BinString As String * 10
    Dim MinWidth As Double

'   Convert Buttons argument to a binary string
    BinString = DECTOBIN(Buttons1, 10)

'   Do the Caption
    If Title1 <> "" Then caption = Title1 Else caption = Application.Name

'   Do the icon
'   VbCritical = 16  Display Critical Message icon.
'   VbQuestion = 32  Display Question icon.
'   VbExclamation = 48  Display Exclamation icon.
'   VbInformation   64  Display Information Message icon.

    LabelPrompt.Left = 50
    ImageCritical.Visible = False
    ImageExclamation.Visible = False
    ImageInformation.Visible = False
    ImageQuestion.Visible = False
    Select Case Mid(BinString, 4, 3) 'digits 4-6 of 10-digit binary string
        Case "001" 'Critical
            With ImageCritical
                .Visible = True
                .Left = 10
                .Top = 8
            End With
        Case "010" 'Question
            With ImageQuestion
                .Visible = True
                .Left = 10
                .Top = 8
            End With
        Case "011" 'Exclamation
            With ImageExclamation
                .Visible = True
                .Left = 10
                .Top = 8
            End With
        Case "100" 'Information
            With ImageInformation
                .Visible = True
                .Left = 10
                .Top = 8
            End With
        Case "000" 'No Image, move label to the left
            LabelPrompt.Left = 10
    End Select

'   Do the Prompt
    With LabelPrompt
        .Top = TextTop
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Font.Bold = True
        .caption = Prompt1
        .AutoSize = False
'       Adjust width, based on video resolution
        .Width = GetSystemMetrics(0) * 0.44
        .AutoSize = True
'       Adjust dialog box width
        Me.Width = .Left + .Width + 10
        If .Height < 16 Then .Height = 16
    End With
     
'   Make sure dialog box is wide enough for the buttons
    Select Case Right(BinString, 3) 'digits 8-10 of 10-digit binary string
        Case "000": MinWidth = 20 + ButtonWidth
        Case "001": MinWidth = 20 + (ButtonWidth * 2) + ButtonGap
        Case "010": MinWidth = 20 + (ButtonWidth * 3) + (ButtonGap * 2)
        Case "011": MinWidth = 20 + (ButtonWidth * 3) + (ButtonGap * 2)
        Case "100": MinWidth = 20 + (ButtonWidth * 2) + ButtonGap
        Case "101": MinWidth = 20 + (ButtonWidth * 2) + ButtonGap
    End Select
    
    If Me.Width < MinWidth Then Me.Width = MinWidth
        
'   Which Buttons?
'   BUTTON CONSTANTS
'   vbOKOnly = 0   Display OK button only.
'   VbOKCancel = 1   Display OK and Cancel buttons.
'   VbAbortRetryIgnore =2  Display Abort, Retry, and Ignore buttons.
'   VbYesNoCancel = 3   Display Yes, No, and Cancel buttons.
'   VbYesNo = 4   Display Yes and No buttons.
'   VbRetryCancel = 5  Display Retry and Cancel buttons.
        
    With Button1
        .Top = LabelPrompt.Top + LabelPrompt.Height + 12
        .Height = ButtonHeight
        .Visible = False
    End With
    With Button2
        .Top = Button1.Top
        .Height = ButtonHeight
        .Visible = False
    End With
    With Button3
        .Top = Button1.Top
        .Height = ButtonHeight
        .Visible = False
    End With
        
    Select Case Right(BinString, 3) 'digits 8-10 of 10-digit binary string
        Case "000" '0 = OK Only
            With Button1
                .Visible = True
                .caption = "OK"
                .Left = (Me.Width / 2) - (.Width / 2)
            End With
        Case "001" '1 = OK & Cancel
            With Button1
                .Visible = True
                .caption = "OK"
                .Left = (Me.Width / 2) - ((ButtonWidth / 2) * 2)
            End With
            With Button2
                .Visible = True
                .caption = "Cancel"
                .Left = Button1.Left + ButtonWidth + ButtonGap
                .Cancel = True
            End With
        Case "010" '2= Abort, Retry, Ignore
            With Button1
                .Visible = True
                .caption = "Abort"
                .Accelerator = "A"
                .Left = (Me.Width / 2) - ((ButtonWidth / 2) * 3)
            End With
            With Button2
                .Visible = True
                .caption = "Retry"
                .Accelerator = "R"
                .Left = Button1.Left + ButtonWidth + ButtonGap
            End With
            With Button3
                .Visible = True
                .caption = "Ignore"
                .Accelerator = "I"
                .Left = Button2.Left + ButtonWidth + ButtonGap
            End With
        
        Case "011" '3 = Yes, No, Cancel
            With Button1
                .Visible = True
                .caption = "Yes"
                .Accelerator = "Y"
                .Left = (Me.Width / 2) - ((ButtonWidth / 2) * 3)
            End With
            With Button2
                .Visible = True
                .caption = "No"
                .Accelerator = "N"
                .Left = Button1.Left + ButtonWidth + ButtonGap
            End With
            With Button3
                .Visible = True
                .caption = "Cancel"
                .Left = Button2.Left + ButtonWidth + ButtonGap
                .Cancel = True
            End With
        Case "100" '4 = Yes & No
            With Button1
                .Visible = True
                .caption = "Yes"
                .Accelerator = "Y"
                .Left = (Me.Width / 2) - ((ButtonWidth / 2) * 2)
            End With
            With Button2
                .Visible = True
                .caption = "No"
                .Accelerator = "N"
                .Left = Button1.Left + ButtonWidth + ButtonGap
            End With
        Case "101" '5 = Retry & Cancel
            With Button1
                .Visible = True
                .caption = "Retry"
                .Left = (Me.Width / 2) - ((ButtonWidth / 2) * 2)
            End With
            With Button2
                .Visible = True
                .caption = "Cancel"
                .Left = Button1.Left + ButtonWidth + ButtonGap
                .Cancel = True
            End With
        Case Else ' OK Only
            With Button1
                .Visible = True
                .caption = "OK"
                .Left = (Me.Width / 2) - (.Width / 2)
            End With
    End Select
    Me.Height = Button1.Top + 54
        
'   Default Button
'   DEFAULT BUTTON CONSTANTS
'   VbDefaultButton1 = 0   First button is default.
'   VbDefaultButton2 = 256 Second button is default.
'   VbDefaultButton3 = 512 Third button is default.
'   VbDefaultButton4 = 768 Fourth button is default - not implemented here.

    Select Case Left(BinString, 2) 'digits 1-2 of 10-digit binary string
        Case "00":
            Button1.Default = True
            Button1.TabIndex = 0
        Case "01":
            Button2.Default = True
            Button2.TabIndex = 0
        Case "10":
            Button3.Default = True
            Button3.TabIndex = 0
        Case "11": ' Not implemented, use Button1
            Button1.Default = True
            Button1.TabIndex = 0
    End Select
End Sub

'RETURN VALUES CONSTANTS
'vbOK = 1 OK
'vbCancel = 2 Cancel
'vbAbort = 3   Abort
'vbRetry = 4   Retry
'vbIgnore = 5   Ignore
'vbYes = 6 Yes
'vbNo = 7 No

Private Sub Button1_Click()
    Select Case Button1.caption
        Case "OK": UserClick = vbOK
        Case "Cancel": UserClick = vbCancel
        Case "Abort": UserClick = vbAbort
        Case "Retry": UserClick = vbRetry
        Case "Ignore": UserClick = vbIgnore
        Case "Yes": UserClick = vbYes
        Case "No": UserClick = vbNo
    End Select
    Unload Me
End Sub

Private Sub Button2_Click()
    Select Case Button2.caption
        Case "OK": UserClick = vbOK
        Case "Cancel": UserClick = vbCancel
        Case "Abort": UserClick = vbAbort
        Case "Retry": UserClick = vbRetry
        Case "Ignore": UserClick = vbIgnore
        Case "Yes": UserClick = vbYes
        Case "No": UserClick = vbNo
    End Select
    Unload Me
End Sub

Private Sub Button3_Click()
    Select Case Button3.caption
        Case "OK": UserClick = vbOK
        Case "Cancel": UserClick = vbCancel
        Case "Abort": UserClick = vbAbort
        Case "Retry": UserClick = vbRetry
        Case "Ignore": UserClick = vbIgnore
        Case "Yes": UserClick = vbYes
        Case "No": UserClick = vbNo
    End Select
    Unload Me
End Sub


Private Function DECTOBIN(num, digits) As String
'   Converts a number to a binary string
    Dim i As Integer
    DECTOBIN = ""
    For i = digits To 1 Step -1
        If num >= 2 ^ (i - 1) Then
            DECTOBIN = DECTOBIN & "1"
            num = num - (2 ^ (i - 1))
        Else
            DECTOBIN = DECTOBIN & "0"
        End If
    Next i
End Function

