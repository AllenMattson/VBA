VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UWizard 
   Caption         =   "UserForm1"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6015
   OleObjectBlob   =   "UWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    Dim Msg As String
    Dim Ans As Long
    Msg = "Cancel the wizard?"
    Ans = MsgBox(Msg, vbQuestion + vbYesNo, APPNAME)
    If Ans = vbYes Then Unload Me
End Sub

Private Sub cmdBack_Click()
    Me.mpgWizard.Value = Me.mpgWizard.Value - 1
    UpdateControls
End Sub

Private Sub cmdNext_Click()
    Me.mpgWizard.Value = Me.mpgWizard.Value + 1
    UpdateControls
End Sub


Private Sub cmdFinish_Click()
    Dim r As Long
    
    r = Application.WorksheetFunction. _
      CountA(Range("A:A")) + 1

'   Insert the name
    Cells(r, 1) = Me.tbxName.Text
    
'   Insert the gender
    Select Case True
        Case Me.optMale.Value: Cells(r, 2) = "Male"
        Case Me.optFemale: Cells(r, 2) = "Female"
        Case Me.optNoAnswer: Cells(r, 2) = "Unknown"
    End Select
    
'   Insert usage
    Cells(r, 3) = Me.chkExcel.Value
    Cells(r, 4) = Me.chkWord.Value
    Cells(r, 5) = Me.chkAccess.Value
    
'   Insert ratings
    If Me.optExcelNo.Value Then Cells(r, 6) = ""
    If Me.optExcelPoor.Value Then Cells(r, 6) = 0
    If Me.optExcelGood.Value Then Cells(r, 6) = 1
    If Me.optExcelExc.Value Then Cells(r, 6) = 2
    If Me.optWordNo.Value Then Cells(r, 7) = ""
    If Me.optWordPoor.Value Then Cells(r, 7) = 0
    If Me.optWordGood.Value Then Cells(r, 7) = 1
    If Me.optWordExc.Value Then Cells(r, 7) = 2
    If Me.optAccessNo.Value Then Cells(r, 8) = ""
    If Me.optAccessPoor.Value Then Cells(r, 8) = 0
    If Me.optAccessGood.Value Then Cells(r, 8) = 1
    If Me.optAccessExc.Value Then Cells(r, 8) = 2
    
    Unload Me
End Sub

Private Sub mpgWizard_Change()
    Dim TopPos As Long
    Dim FSpace As Long
    Dim AtLeastOne As Boolean
    Dim i As Long
    
'   Set up the Ratings page?
    If Me.mpgWizard.Value = 3 Then
'       Create an array of CheckBox controls
        Dim ProdCB(1 To 3) As MSForms.CheckBox
        Set ProdCB(1) = Me.chkExcel
        Set ProdCB(2) = Me.chkWord
        Set ProdCB(3) = Me.chkAccess
        
'       Create an array of Frame controls
        Dim ProdFrame(1 To 3) As MSForms.Frame
        Set ProdFrame(1) = Me.frmExcel
        Set ProdFrame(2) = Me.frmWord
        Set ProdFrame(3) = Me.frmAccess
        
        TopPos = 22
        FSpace = 8
        AtLeastOne = False

'       Loop through all products
        For i = 1 To 3
            If ProdCB(i).Value Then
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
            Me.lblHeadings.Visible = True
            Me.imgRating.Visible = True
            Me.lblFinishMsg.Visible = False
        Else
            Me.lblHeadings.Visible = False
            Me.imgRating.Visible = False
            Me.lblFinishMsg.Visible = True
            If Len(Me.tbxName.Text) = 0 Then
                Me.lblFinishMsg.Caption = _
                  "A name is required in Step 1."
            Else
                Me.lblFinishMsg.Caption = _
                  "Click Finish to exit."
             End If
        End If
    End If
End Sub


Private Sub tbxName_Change()
    UpdateControls
End Sub

Public Sub UserForm_Initialize()
    Me.mpgWizard.Value = 0
    UpdateControls
End Sub

Sub UpdateControls()
'   Enable back if not on page 1
    Me.cmdBack.Enabled = Me.mpgWizard.Value > 0
'   Enable next if not on the last page
    Me.cmdNext.Enabled = Me.mpgWizard.Value < Me.mpgWizard.Pages.Count - 1
    
'   Update the caption
    Me.Caption = APPNAME & " Step " _
      & Me.mpgWizard.Value + 1 & " of " _
      & Me.mpgWizard.Pages.Count

'   the Name field is required
    Me.cmdFinish.Enabled = Len(Me.tbxName.Text) > 0
End Sub
