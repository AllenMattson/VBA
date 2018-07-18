VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public UserLanguage As Integer

Public Sub UserForm_Initialize()
    MultiPage1.Value = 0
    Select Case Application.International(xlCountryCode)
        Case 34 'Spanish
            UserLanguage = 2
        Case 49 'German
            UserLanguage = 3
        Case Else 'default to English
            UserLanguage = 1 'default
    End Select
    UpdateControls
End Sub
    
Private Sub obEnglish_Click()
    UserLanguage = 1
    ChangeLanguage
End Sub

Private Sub obGerman_Click()
    UserLanguage = 3
    ChangeLanguage
End Sub

Private Sub obSpanish_Click()
    UserLanguage = 2
    ChangeLanguage
End Sub

Sub ChangeLanguage()
    Dim ctl As Control
    Dim Cap As String

    For Each ctl In Me.Controls
        If HasCaption(ctl) Then
            Cap = Translate(ctl.Name, UserLanguage)
            If Cap <> "" Then ctl.Caption = Cap
        End If
    Next ctl

'   Update the caption
    Cap = APPNAME & " - Step " _
      & MultiPage1.Value + 1 & " of " _
      & MultiPage1.Pages.Count
    Me.Caption = Translate(Cap, UserLanguage)
End Sub

Function HasCaption(ctl As Control) As Boolean
    Dim x
    On Error Resume Next
    x = ctl.Caption
    If IsEmpty(x) Or Err <> 0 Then HasCaption = False Else HasCaption = True
End Function
Function Translate(text, language) As String
    Dim txt As String
    On Error Resume Next
    txt = Application.WorksheetFunction.VLookup(text, Sheets("shtLocalization").Range("A1:D32"), language + 1, False)
    If Err <> 0 Then Translate = "" Else Translate = txt
End Function
Sub UpdateControls()
    Dim Cap As String
    Select Case MultiPage1.Value
        Case 0
            BackButton.Enabled = False
            NextButton.Enabled = True
        Case MultiPage1.Pages.Count - 1
            BackButton.Enabled = True
            NextButton.Enabled = False
        Case Else
            BackButton.Enabled = True
            NextButton.Enabled = True
    End Select
    
'   Update the caption
    Cap = APPNAME & " - Step " _
      & MultiPage1.Value + 1 & " of " _
      & MultiPage1.Pages.Count
    Me.Caption = Translate(Cap, UserLanguage)

'   the Name field is required
    If tbName.text = "" Then
        FinishButton.Enabled = False
    Else
        FinishButton.Enabled = True
    End If
End Sub


Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub BackButton_Click()
    MultiPage1.Value = MultiPage1.Value - 1
    UpdateControls
End Sub

Private Sub NextButton_Click()
    MultiPage1.Value = MultiPage1.Value + 1
    UpdateControls
End Sub


Private Sub FinishButton_Click()
    Dim r As Integer
    r = Application.WorksheetFunction. _
      CountA(Range("A:A")) + 1

'   Insert the name
    Cells(r, 1) = tbName.text
    
'   Insert the gender
    Select Case True
        Case obMale: Cells(r, 2) = "Male"
        Case obFemale: Cells(r, 2) = "Female"
        Case obNoAnswer: Cells(r, 2) = "Unknown"
    End Select
    
'   Insert usage
    If cbExcel Then Cells(r, 3) = True Else Cells(r, 3) = False
    If cbWord Then Cells(r, 4) = True Else Cells(r, 4) = False
    If cbAccess Then Cells(r, 5) = True Else Cells(r, 5) = False
    
'   Insert ratings
    If obExcel1 Then Cells(r, 6) = ""
    If obExcel2 Then Cells(r, 6) = 0
    If obExcel3 Then Cells(r, 6) = 1
    If obExcel4 Then Cells(r, 6) = 2
    If obWord1 Then Cells(r, 7) = ""
    If obWord2 Then Cells(r, 7) = 0
    If obWord3 Then Cells(r, 7) = 1
    If obWord4 Then Cells(r, 7) = 2
    If obAccess1 Then Cells(r, 8) = ""
    If obAccess2 Then Cells(r, 8) = 0
    If obAccess3 Then Cells(r, 8) = 1
    If obAccess4 Then Cells(r, 8) = 2
    
    Unload Me
    
End Sub

Private Sub MultiPage1_Change()
    Dim TopPos As Integer, FSpace As Integer
    Dim AtLeastOne As Boolean
'   Set up the Ratings page
    If MultiPage1.Value = 3 Then
        TopPos = Me.lblHeading1.Top + Me.lblHeading1.Height + 4
        FSpace = 8
        AtLeastOne = False

'       Excel user?
        If cbExcel Then
            FrameExcel.Visible = True
            FrameExcel.Top = TopPos
            TopPos = TopPos + FrameExcel.Height + FSpace
            AtLeastOne = True
        Else
            FrameExcel.Visible = False
        End If

'       Word user?
        If cbWord Then
            FrameWord.Visible = True
            FrameWord.Top = TopPos
            TopPos = TopPos + FrameWord.Height + FSpace
            AtLeastOne = True
        Else
            FrameWord.Visible = False
        End If

'       Acess user?
        If cbAccess Then
            FrameAccess.Visible = True
            FrameAccess.Top = TopPos
            TopPos = TopPos + FrameWord.Height + FSpace
            AtLeastOne = True
        Else
            FrameAccess.Visible = False
        End If
        
'       Uses no products?
        If AtLeastOne Then
            lblHeading1.Visible = True
            lblHeading2.Visible = True
            lblHeading3.Visible = True
            lblHeading4.Visible = True
            Label10.Visible = True
            lblFinishMsg.Visible = False
        Else
            lblHeading1.Visible = False
            lblHeading2.Visible = False
            lblHeading3.Visible = False
            lblHeading4.Visible = False
            Label10.Visible = False
            lblFinishMsg.Visible = True
            If tbName = "" Then
                lblFinishMsg.Caption = Translate("AltMessage", UserLanguage)
            Else
                lblFinishMsg.Caption = Translate("FinishMessage", UserLanguage)
             End If
        End If
    End If
End Sub


Private Sub tbName_Change()
    If tbName.text = "" Then FinishButton.Enabled = False Else FinishButton.Enabled = True
End Sub



