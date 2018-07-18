Attribute VB_Name = "Module1"
Option Explicit

Sub MakeForm()
    Dim TempForm  As Object 'VBComponent
    Dim NewButton As Msforms.CommandButton
    Dim Line As Integer
    Dim TheForm

'   Make sure access to the VBProject is allowed
    On Error Resume Next
    Dim x
    Set x = ActiveWorkbook.VBProject
    If Err <> 0 Then
        MsgBox "Your security settings do not allow this macro to run.", vbCritical
        On Error GoTo 0
        Exit Sub
    End If
    
    Application.VBE.MainWindow.Visible = False

'   Create the UserForm
    Set TempForm = ThisWorkbook.VBProject. _
      VBComponents.Add(3) 'vbext_ct_MSForm
    With TempForm
        .Properties("Caption") = "Temporary Form"
        .Properties("Width") = 200
        .Properties("Height") = 100
    End With

'   Add a CommandButton
    Set NewButton = TempForm.Designer.Controls _
      .Add("forms.CommandButton.1")
    With NewButton
        .Caption = "Click Me"
        .Left = 60
        .Top = 40
    End With

'   Add an event-hander sub for the CommandButton
    With TempForm.CodeModule
        Line = .CountOfLines
        .InsertLines Line + 1, "Sub CommandButton1_Click()"
        .InsertLines Line + 2, "MsgBox ""Hello!"""
        .InsertLines Line + 3, "Unload Me"
        .InsertLines Line + 4, "End Sub"
    End With

'   Show the form
    VBA.UserForms.Add(TempForm.Name).Show
'
'   Delete the form
    ThisWorkbook.VBProject.VBComponents.Remove TempForm
End Sub

