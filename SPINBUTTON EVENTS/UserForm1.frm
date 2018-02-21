VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "SpinButton / TextBox Demo"
   ClientHeight    =   2370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4980
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Row As Integer

Private Sub OKButton_Click()
    Unload Me
End Sub

Private Sub ClearListButton_Click()
    Range("A:A").ClearContents
    Row = 0
End Sub

Private Sub SpinButton1_AfterUpdate()
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_AfterUpdate Event"
End Sub

Private Sub SpinButton1_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_BeforeUpdate Event"
End Sub

Private Sub SpinButton1_Change()
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_Change Event"
End Sub

Private Sub SpinButton1_Enter()
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_Enter Event"
End Sub

Private Sub SpinButton1_Error(ByVal Number As Integer, ByVal Description As MSForms.ReturnString, ByVal SCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, ByVal CancelDisplay As MSForms.ReturnBoolean)
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_Error Event"
End Sub

Private Sub SpinButton1_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_Exit Event"
End Sub

Private Sub SpinButton1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_KeyDown Event"
End Sub

Private Sub SpinButton1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_KeyPress Event"
End Sub

Private Sub SpinButton1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_KeyUp Event"
End Sub

Private Sub SpinButton1_SpinDown()
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_SpinDown Event"
End Sub

Private Sub SpinButton1_SpinUp()
    Row = Row + 1
    Cells(Row, 1) = "SpinButton1_SpinUp Event"
End Sub

Private Sub UserForm_Activate()
    Row = Row + 1
    Cells(Row, 1) = "UserForm_Activate"
End Sub


Private Sub UserForm_Click()
    Row = Row + 1
    Cells(Row, 1) = "UserForm_Click Event"
End Sub

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Row = Row + 1
    Cells(Row, 1) = "UserForm_DblClick Event"
End Sub

Private Sub UserForm_Deactivate()
    Row = Row + 1
    Cells(Row, 1) = "UserForm_Deactivate Event"
End Sub

Private Sub UserForm_Initialize()
    Range("A:A").ClearContents
    Row = 1
    Cells(Row, 1) = "UserForm_Initialize Event"
End Sub

Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Row = Row + 1
    Cells(Row, 1) = "UserForm_KeyDown Event"
End Sub

Private Sub UserForm_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Row = Row + 1
    Cells(Row, 1) = "UserForm_KeyUp Event"
End Sub

Private Sub UserForm_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Row = Row + 1
    Cells(Row, 1) = "UserForm_MouseDown Event"
End Sub

Private Sub UserForm_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Row = Row + 1
    Cells(Row, 1) = "UserForm_MouseUp Event"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Row = Row + 1
    Cells(Row, 1) = "UserForm_QueryClose Event"
End Sub

Private Sub UserForm_Terminate()
    Row = Row + 1
    Cells(Row, 1) = "UserForm_Terminate Event"
End Sub
