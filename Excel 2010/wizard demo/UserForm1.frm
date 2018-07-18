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

Private Sub CancelButton_Click()
    Dim Msg As String
    Dim Ans As Integer
    Msg = "Cancel the wizard?"
    Ans = MsgBox(Msg, vbQuestion + vbYesNo, APPNAME)
    If Ans = vbYes Then Unload Me
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
    Dim r As Long
    
    r = Application.WorksheetFunction. _
      CountA(Range("A:A")) + 1

'   Insert the name
    Cells(r, 1) = tbName.Text
    
'   Insert the gender
    Select Case True
        Case obMale: Cells(r, 2) = "Male"
        Case obFemale: Cells(r, 2) = "Female"
        Case obNoAnswer: Cells(r, 2) = "Unknown"
    End Select
    
'   Insert usage
    Cells(r, 3) = cbExcel
    Cells(r, 4) = cbWord
    Cells(r, 5) = cbAccess
    
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
    Dim TopPos As Long
    Dim FSpace As Long
    Dim AtLeastOne As Boolean
    Dim i As Long
    
'   Set up the Ratings page?
    If MultiPage1.Value = 3 Then
'       Create an array of CheckBox controls
        Dim ProdCB(1 To 3) As MSForms.CheckBox
        Set ProdCB(1) = cbExcel
        Set ProdCB(2) = cbWord
        Set ProdCB(3) = cbAccess
        
'       Create an array of Frame controls
        Dim ProdFrame(1 To 3) As MSForms.Frame
        Set ProdFrame(1) = FrameExcel
        Set ProdFrame(2) = FrameWord
        Set ProdFrame(3) = FrameAccess
        
        TopPos = 22
        FSpace = 8
        AtLeastOne = False

'       Loop through all products
        For i = 1 To 3
            If ProdCB(i) Then
                ProdFrame(i).Visible = True
                ProdFrame(i).Top = TopPos
                TopPos = TopPos + ProdFrame(i).Height + FSpace
                AtLeastOne = True
            Else
                ProdFrame(i).Visible = False
            End If
        Next i
        
'       Uses no products?
        If AtLeastOne Then
            lblHeadings.Visible = True
            Image4.Visible = True
            lblFinishMsg.Visible = False
        Else
            lblHeadings.Visible = False
            Image4.Visible = False
            lblFinishMsg.Visible = True
            If tbName = "" Then
                lblFinishMsg.Caption = _
                  "A name is required in Step 1."
            Else
                lblFinishMsg.Caption = _
                  "Click Finish to exit."
             End If
        End If
    End If
End Sub


Private Sub tbName_Change()
    If tbName.Text = "" Then FinishButton.Enabled = False Else FinishButton.Enabled = True
End Sub

Public Sub UserForm_Initialize()
    MultiPage1.Value = 0
    UpdateControls
End Sub

Sub UpdateControls()
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
    Me.Caption = APPNAME & " Step " _
      & MultiPage1.Value + 1 & " of " _
      & MultiPage1.Pages.Count

'   the Name field is required
    If tbName.Text = "" Then
        FinishButton.Enabled = False
    Else
        FinishButton.Enabled = True
    End If
End Sub
