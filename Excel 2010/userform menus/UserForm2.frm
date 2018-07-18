VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm2 
   Caption         =   "ListBox Menu Demo"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   OleObjectBlob   =   "UserForm2.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CancelButton_Click()
    Unload Me
End Sub

Private Sub ExecuteButton_Click()
    Select Case ListBox1.ListIndex
        Case -1
            MsgBox "Select a macro from the list."
            Exit Sub
        Case 0: Me.Hide: Call Macro1
        Case 1: Me.Hide: Call Macro2
        Case 2: Me.Hide: Call Macro3
        Case 3: Me.Hide: Call Macro4
        Case 4: Me.Hide: Call Macro5
        Case 5: Me.Hide: Call Macro6
    End Select
    Unload Me
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Hide
    Select Case ListBox1.ListIndex
        Case 0: Call Macro1
        Case 1: Call Macro2
        Case 2: Call Macro3
        Case 3: Call Macro4
        Case 4: Call Macro5
        Case 5: Call Macro6
    End Select
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    With ListBox1
        .AddItem "Macro1"
        .AddItem "Macro2"
        .AddItem "Macro3"
        .AddItem "Macro4"
        .AddItem "Macro5"
        .AddItem "Macro6"
    End With
End Sub
