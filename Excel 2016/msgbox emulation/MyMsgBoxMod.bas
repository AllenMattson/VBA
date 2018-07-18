Attribute VB_Name = "MyMsgBoxMod"
Option Explicit

#If VBA7 And Win64 Then
    Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If


Function MyMsgBox(ByVal Prompt As String, _
    Optional ByVal Buttons As Long, _
    Optional ByVal Title As String) As Long
'   Emulates VBA's MsgBox function
'   Does not support the HelpFile or Context arguments
    With UMsgBox
    '   Do the Caption
        If Len(Title) > 0 Then .Caption = Title _
            Else .Caption = Application.Name
        SetImage Buttons
        SetPrompt Prompt
        SetButtons Buttons
        .Height = .cmdLeft.Top + 64
        SetDefaultButton Buttons
        '.StartUpPosition = 0
        '.Left = Application.Left + _
            (0.5 * Application.Width) - (0.5 * .Width)
        '.Top = Application.Top + _
            (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
    MyMsgBox = UMsgBox.UserClick
End Function

Sub TestMyMsgBox()
    Dim Prompt As String
    Dim Buttons As Long
    Dim Title As String
    
'   Tests the MyMsgbox Function
'   Can be deleted from this module
    Prompt = Range("B1").Value
    Buttons = Range("B5").Value
    Title = Range("B6").Value
    Range("B7").Value = MyMsgBox(Prompt, Buttons, Title)
End Sub

Sub TestMsgBox()
    Dim Prompt As String
    Dim Buttons As Integer
    Dim Title As String

'   Tests the standard Msgbox Function
'   Can be deleted from this module
    Prompt = Range("B1").Value
    Buttons = Range("B5").Value
    Title = Range("B6").Value
    Range("B7").Value = MsgBox(Prompt, Buttons, Title)
End Sub

Sub MyMsgBoxTest()
    Dim Prompt As String, Buttons As Long, Title As String, Ans As Long
    Prompt = "You have chosen to save this workbook" & vbCrLf
    Prompt = Prompt & "on a drive that is not available to" & vbCrLf
    Prompt = Prompt & "all employees." & vbCrLf & vbCrLf
    Prompt = Prompt & "OK to continue?"
    Buttons = vbQuestion + vbYesNo
    Title = "We have a problem"
    Ans = MyMsgBox(Prompt, Buttons, Title)
End Sub

Private Sub SetImage(Buttons As Long)
'   Do the icon
'   VbCritical = 16  Display Critical Message icon.
'   VbQuestion = 32  Display Question icon.
'   VbExclamation = 48  Display Exclamation icon.
'   VbInformation   64  Display Information Message icon.
    
    With UMsgBox
        .imgCritical.Visible = False
        .imgExclamation.Visible = False
        .imgInformation.Visible = False
        .imgQuestion.Visible = False
        
        .imgCritical.Top = 8
        .imgExclamation.Top = 8
        .imgInformation.Top = 8
        .imgQuestion.Top = 8
        
        Select Case True
            Case (Buttons And vbInformation) = vbInformation: .imgInformation.Visible = True
            Case (Buttons And vbExclamation) = vbExclamation: .imgExclamation.Visible = True
            Case (Buttons And vbQuestion) = vbQuestion: .imgQuestion.Visible = True
            Case (Buttons And vbCritical) = vbCritical: .imgCritical.Visible = True
        End Select
        
        If Buttons And vbCritical Or Buttons And vbExclamation Or Buttons And vbInformation Or Buttons And vbQuestion Then
            .lblPrompt.Left = 50
        Else
            .lblPrompt.Left = 10
        End If
        
    End With
    
    
End Sub

Private Sub SetPrompt(Prompt As String)
    Const TextTop As Double = 8

'   Do the Prompt
    With UMsgBox.lblPrompt
        .Top = TextTop
        .Font.Name = "Calibri"
        .Font.Size = 12
        .Font.Bold = True
        .Caption = Prompt
        .AutoSize = False
'       Adjust width, based on video resolution
        .Width = GetSystemMetrics(0) * 0.44
        .AutoSize = True
'       Adjust dialog box width
        UMsgBox.Width = .Left + .Width + 30
        If .Height < 16 Then .Height = 16
    End With
End Sub

Private Sub SetButtons(Buttons As Long)
    Dim MinWidth As Double
    Const ButtonWidth As Double = 56
    Const ButtonHeight As Double = 20
    Const ButtonGap As Double = 0

    With UMsgBox
'       Make sure dialog box is wide enough for the buttons
        Select Case True
            Case Buttons And vbOKOnly
                MinWidth = 20 + ButtonWidth
            Case Buttons And vbOKCancel, Buttons And vbYesNo, Buttons And vbRetryCancel
                MinWidth = 20 + (ButtonWidth * 2) + ButtonGap
            Case Buttons And vbAbortRetryIgnore, Buttons And vbYesNoCancel
                MinWidth = 20 + (ButtonWidth * 3) + (ButtonGap * 2)
        End Select
        
        If .Width < MinWidth Then .Width = MinWidth
        
'       Which Buttons?
'       BUTTON CONSTANTS
'       vbOKOnly = 0   Display OK button only.
'       VbOKCancel = 1   Display OK and Cancel buttons.
'       VbAbortRetryIgnore =2  Display Abort, Retry, and Ignore buttons.
'       VbYesNoCancel = 3   Display Yes, No, and Cancel buttons.
'       VbYesNo = 4   Display Yes and No buttons.
'       VbRetryCancel = 5  Display Retry and Cancel buttons.
        
        .cmdLeft.Top = .lblPrompt.Top + .lblPrompt.Height + 12
        .cmdLeft.Height = ButtonHeight
        .cmdLeft.Visible = False
        .cmdMiddle.Top = .lblPrompt.Top + .lblPrompt.Height + 12
        .cmdMiddle.Height = ButtonHeight
        .cmdMiddle.Visible = False
        .cmdRight.Top = .lblPrompt.Top + .lblPrompt.Height + 12
        .cmdRight.Height = ButtonHeight
        .cmdRight.Visible = False
        
        Select Case True
            Case (Buttons And vbRetryCancel) = vbRetryCancel
                .cmdLeft.Visible = True
                .cmdLeft.Caption = "Retry"
                .cmdLeft.Left = (.Width / 2) - ((ButtonWidth / 2) * 2)
                .cmdMiddle.Visible = True
                .cmdMiddle.Caption = "Cancel"
                .cmdMiddle.Left = .cmdLeft.Left + ButtonWidth + ButtonGap
            Case (Buttons And vbYesNo) = vbYesNo
                .cmdLeft.Visible = True
                .cmdLeft.Caption = "Yes"
                .cmdLeft.Left = (.Width / 2) - ((ButtonWidth / 2) * 2)
                .cmdMiddle.Visible = True
                .cmdMiddle.Caption = "No"
                .cmdMiddle.Left = .cmdLeft.Left + ButtonWidth + ButtonGap
            Case (Buttons And vbYesNoCancel) = vbYesNoCancel
                .cmdLeft.Visible = True
                .cmdLeft.Caption = "Yes"
                .cmdLeft.Left = (.Width / 2) - ((ButtonWidth / 2) * 3)
                .cmdMiddle.Visible = True
                .cmdMiddle.Caption = "No"
                .cmdMiddle.Left = .cmdLeft.Left + ButtonWidth + ButtonGap
                .cmdRight.Visible = True
                .cmdRight.Caption = "Cancel"
                .cmdRight.Left = .cmdMiddle.Left + ButtonWidth + ButtonGap
            Case (Buttons And vbAbortRetryIgnore) = vbAbortRetryIgnore
                .cmdLeft.Visible = True
                .cmdLeft.Caption = "Abort"
                .cmdLeft.Left = (.Width / 2) - ((ButtonWidth / 2) * 3)
                .cmdMiddle.Visible = True
                .cmdMiddle.Caption = "Retry"
                .cmdMiddle.Left = .cmdLeft.Left + ButtonWidth + ButtonGap
                .cmdRight.Visible = True
                .cmdRight.Caption = "Ignore"
                .cmdRight.Left = .cmdMiddle.Left + ButtonWidth + ButtonGap
            Case (Buttons And vbOKCancel) = vbOKCancel
                .cmdLeft.Visible = True
                .cmdLeft.Caption = "OK"
                .cmdLeft.Left = (.Width / 2) - ((ButtonWidth / 2) * 2)
                .cmdMiddle.Visible = True
                .cmdMiddle.Caption = "Cancel"
                .cmdMiddle.Left = .cmdLeft.Left + ButtonWidth + ButtonGap
            Case Else
                .cmdLeft.Visible = True
                .cmdLeft.Caption = "OK"
                .cmdLeft.Left = (.Width / 2) - (ButtonWidth / 2)
        End Select
    End With
End Sub

Private Sub SetDefaultButton(Buttons As Long)
'   Default Button
'   DEFAULT BUTTON CONSTANTS
'   VbDefaultButton1 = 0   First button is default.
'   VbDefaultButton2 = 256 Second button is default.
'   VbDefaultButton3 = 512 Third button is default.
'   VbDefaultButton4 = 768 Fourth button is default - not implemented here.
    
    With UMsgBox
        Select Case True
            Case (Buttons And vbDefaultButton4) = vbDefaultButton4
                .cmdLeft.Default = True
                .cmdLeft.TabIndex = 0
            Case (Buttons And vbDefaultButton3) = vbDefaultButton3
                .cmdRight.Default = True
                .cmdRight.TabIndex = 0
            Case (Buttons And vbDefaultButton2) = vbDefaultButton2
                .cmdMiddle.Default = True
                .cmdMiddle.TabIndex = 0
            Case Else
                .cmdLeft.Default = True
                .cmdLeft.TabIndex = 0
        End Select
    End With
End Sub
