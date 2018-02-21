VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Get Name and Sex"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CloseButton_Click()
    Unload Me
End Sub

Private Sub OKButton_Click()
    Dim NextRow As Long
    
'   Make sure Sheet1 is active
    Sheets("Sheet1").Activate

'   Make sure a name is entered
    If TextName.Text = "" Then
        MsgBox "You must enter a name."
        TextName.SetFocus
        Exit Sub
    End If

'   Determine the next empty row
    NextRow = _
      Application.WorksheetFunction.CountA(Range("A:A")) + 1
'   Transfer the name
    Cells(NextRow, 1) = TextName.Text
    
'   Transfer the sex
    If OptionMale Then Cells(NextRow, 2) = "Male"
    If OptionFemale Then Cells(NextRow, 2) = "Female"
    If OptionUnknown Then Cells(NextRow, 2) = "Unknown"
    
'   Clear the controls for the next entry
    TextName.Text = ""
    OptionUnknown = True
    TextName.SetFocus
End Sub


