VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSpinEvents 
   Caption         =   "SpinButton / TextBox Demo"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   OleObjectBlob   =   "ufSpinEvents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSpinEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lRow As Long

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdClear_Click()
    Sheet1.Range("A:A").ClearContents
    lRow = 0
End Sub

Private Sub spbDemo_AfterUpdate()
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_AfterUpdate Event"
End Sub

Private Sub spbDemo_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_BeforeUpdate Event"
End Sub

Private Sub spbDemo_Change()
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_Change Event"
End Sub

Private Sub spbDemo_Enter()
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_Enter Event"
End Sub

Private Sub spbDemo_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_Error Event"
End Sub

Private Sub spbDemo_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_Exit Event"
End Sub

Private Sub spbDemo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_KeyDown Event"
End Sub

Private Sub spbDemo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_KeyPress Event"
End Sub

Private Sub spbDemo_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_KeyUp Event"
End Sub

Private Sub spbDemo_SpinDown()
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_SpinDown Event"
End Sub

Private Sub spbDemo_SpinUp()
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "spbDemo_SpinUp Event"
End Sub

Private Sub UserForm_Activate()
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "UserForm_Activate"
End Sub


Private Sub UserForm_Click()
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "UserForm_Click Event"
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "UserForm_DblClick Event"
End Sub

Private Sub UserForm_Deactivate()
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "UserForm_Deactivate Event"
End Sub

Private Sub UserForm_Initialize()
    Range("A:A").ClearContents
    lRow = 1
    Sheet1.Cells(lRow, 1) = "UserForm_Initialize Event"
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "UserForm_KeyDown Event"
End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "UserForm_KeyUp Event"
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "UserForm_MouseDown Event"
End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "UserForm_MouseUp Event"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "UserForm_QueryClose Event"
End Sub

Private Sub UserForm_Terminate()
    lRow = lRow + 1
    Sheet1.Cells(lRow, 1) = "UserForm_Terminate Event"
End Sub
