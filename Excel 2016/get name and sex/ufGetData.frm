VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufGetData 
   Caption         =   "Get Name and Sex"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   OleObjectBlob   =   "ufGetData.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufGetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim lNextRow As Long
    Dim wf As WorksheetFunction
    
    Set wf = Application.WorksheetFunction
    
'   Make sure a name is entered
    If Len(Me.tbxName.Text) = 0 Then
        MsgBox "You must enter a name."
        Me.tbxName.SetFocus
    Else
    '   Determine the next empty row
        lNextRow = wf.CountA(Sheet1.Range("A:A")) + 1
    '   Transfer the name
        Sheet1.Cells(lNextRow, 1) = Me.tbxName.Text
        
    '   Transfer the sex
        With Sheet1.Cells(lNextRow, 2)
            If Me.optMale.Value Then .Value = "Male"
            If Me.optFemale.Value Then .Value = "Female"
            If Me.optUnknown.Value Then .Value = "Unknown"
        End With
        
    '   Clear the controls for the next entry
        Me.tbxName.Text = vbNullString
        Me.optUnknown.Value = True
        Me.tbxName.SetFocus
    End If
End Sub

