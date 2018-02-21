Attribute VB_Name = "modOptionsForm"
Option Explicit

'Passed back to the function from the UserForm
Public GETOPTION_RET_VAL As Variant

Function GetOption(OpArray, Default, Title)
    Dim TempForm As Object 'VBComponent
    Dim NewOptionButton As Msforms.OptionButton
    Dim NewCommandButton1 As Msforms.CommandButton
    Dim NewCommandButton2 As Msforms.CommandButton
    Dim i As Integer, TopPos As Integer
    Dim MaxWidth As Long
    Dim Code As String
    
'   Hide VBE window to prevent screen flashing
    Application.VBE.MainWindow.Visible = False

'   Create the UserForm
    Set TempForm = _
      ThisWorkbook.VBProject.VBComponents.Add(3) 'vbext_ct_MSForm
    TempForm.Properties("Width") = 800
    
'   Add the OptionButtons
    TopPos = 4
    MaxWidth = 0 'Stores width of widest OptionButton
    For i = LBound(OpArray) To UBound(OpArray)
        Set NewOptionButton = TempForm.Designer.Controls. _
          Add("forms.OptionButton.1")
        With NewOptionButton
            .Width = 800
            .Caption = OpArray(i)
            .Height = 15
            .Accelerator = Left(.Caption, 1)
            .Left = 8
            .Top = TopPos
            .Tag = i
            .AutoSize = True
            If Default = i Then .Value = True
            If .Width > MaxWidth Then MaxWidth = .Width
        End With
        TopPos = TopPos + 15
    Next i
    
'   Add the Cancel button
    Set NewCommandButton1 = TempForm.Designer.Controls. _
      Add("forms.CommandButton.1")
    With NewCommandButton1
        .Caption = "Cancel"
        .Cancel = True
        .Height = 18
        .Width = 44
        .Left = MaxWidth + 12
        .Top = 6
    End With

'   Add the OK button
    Set NewCommandButton2 = TempForm.Designer.Controls. _
      Add("forms.CommandButton.1")
    With NewCommandButton2
        .Caption = "OK"
        .Default = True
        .Height = 18
        .Width = 44
        .Left = MaxWidth + 12
        .Top = 28
    End With

'   Add event-hander subs for the CommandButtons
    Code = ""
    Code = Code & "Sub CommandButton1_Click()" & vbCrLf
    Code = Code & "  GETOPTION_RET_VAL=False" & vbCrLf
    Code = Code & "  Unload Me" & vbCrLf
    Code = Code & "End Sub" & vbCrLf
    Code = Code & "Sub CommandButton2_Click()" & vbCrLf
    Code = Code & "  Dim ctl" & vbCrLf
    Code = Code & "  GETOPTION_RET_VAL = False" & vbCrLf
    Code = Code & "  For Each ctl In Me.Controls" & vbCrLf
    Code = Code & "    If TypeName(ctl) = ""OptionButton"" Then" & vbCrLf
    Code = Code & "      If ctl Then GETOPTION_RET_VAL = ctl.Tag" & vbCrLf
    Code = Code & "    End If" & vbCrLf
    Code = Code & "  Next ctl" & vbCrLf
    Code = Code & "  Unload Me" & vbCrLf
    Code = Code & "End Sub"

    With TempForm.CodeModule
        .InsertLines .CountOfLines + 1, Code
    End With
    
'   Adjust the form
    With TempForm
        .Properties("Caption") = Title
        .Properties("Width") = NewCommandButton1.Left + _
           NewCommandButton1.Width + 10
        If .Properties("Width") < 160 Then
            .Properties("Width") = 160
            NewCommandButton1.Left = 106
            NewCommandButton2.Left = 106
        End If
        .Properties("Height") = TopPos + 28
    End With

'   Show the form
    VBA.UserForms.Add(TempForm.Name).Show

'   Delete the form
    ThisWorkbook.VBProject.VBComponents.Remove VBComponent:=TempForm
    
'   Pass the selected option back to the calling procedure
    GetOption = GETOPTION_RET_VAL
End Function

Sub TestGetOption()
    Dim Ops(1 To 5)
    Dim UserOption

'   Make sure access to the VBProject is allowed
    On Error Resume Next
    Dim x
    Set x = ActiveWorkbook.VBProject
    If Err <> 0 Then
        MsgBox "Your security settings do not allow this macro to run.", vbCritical
        On Error GoTo 0
        Exit Sub
    End If
    Ops(1) = "North"
    Ops(2) = "South"
    Ops(3) = "West"
    Ops(4) = "East"
    Ops(5) = "All Regions"
    UserOption = GetOption(Ops, 5, "Select a region")
    MsgBox Ops(UserOption)
End Sub


Sub TestGetOption2()
    Dim Ops()
    Dim UserOption, i, Cnt
    
'   Make sure access to the VBProject is allowed
    On Error Resume Next
    Dim x
    Set x = ActiveWorkbook.VBProject
    If Err <> 0 Then
        MsgBox "Your security settings do not allow this macro to run.", vbCritical
        On Error GoTo 0
        Exit Sub
    End If
    
    Cnt = Application.WorksheetFunction.CountA(Range("A:A"))
    ReDim Ops(1 To Cnt)
    For i = 1 To Cnt
        Ops(i) = Cells(i, 1)
    Next i
    UserOption = GetOption(Ops, 0, "Select a month")
    If UserOption = False Then Exit Sub Else MsgBox Ops(UserOption)
End Sub

Sub TestGetOption3()
    Dim Ops(1 To 9)
    Dim i As Long
    Dim UserOption

'   Make sure access to the VBProject is allowed
    On Error Resume Next
    Dim x
    Set x = ActiveWorkbook.VBProject
    If Err <> 0 Then
        MsgBox "Your security settings do not allow this macro to run.", vbCritical
        On Error GoTo 0
        Exit Sub
    End If
    Ops(1) = "Highway 61 Revisited"
    Ops(2) = "Blood On The Tracks"
    Ops(3) = "Blonde On Blonde"
    Ops(4) = "Time Out Of Mind"
    Ops(5) = "Love And Theft"
    Ops(6) = "Modern Times"
    Ops(7) = "Together Through Life"
    Ops(8) = "My favorite Bob Dylan album is not listed here."
    Ops(9) = "Are you crazy? It's impossible to select a favorite Bob Dylan album!"
    
    UserOption = GetOption(Ops, 0, "Your favorite Dylan album?")
    MsgBox Ops(UserOption)
End Sub

