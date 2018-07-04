Attribute VB_Name = "AddUserForm"
Option Explicit


Sub AddUserForm()
Dim objVBProj As VBProject
Dim objVBComp As VBComponent
Dim objVBFrm As UserForm
Dim objChkBox As Object
Dim x As Integer
Dim sVBA As String
 
Set objVBProj = Application.VBE.ActiveVBProject
Set objVBComp = objVBProj.VBComponents.Add(vbext_ct_MSForm)

With objVBComp
' read form's name and other properties
    Debug.Print "Default Name " & .Name
    Debug.Print "Caption: " & .DesignerWindow.Caption
    Debug.Print "Form is open in the Designer window: " & _
        .HasOpenDesigner
    Debug.Print "Form Name " & .Name
    Debug.Print "Default Width " & .Properties("Width")
    Debug.Print "Default Height " & .Properties("Height")
    
' set form's name, caption and size
    .Name = "ReportSelector"
    .Properties("Caption") = "Request Report"
    .Properties("Width") = 250
    .Properties("Height") = 250
End With
  
Set objVBFrm = objVBComp.Designer
With objVBFrm
    With .Controls.Add("Forms.Label.1", "lbName")
        .Caption = "Department:"
        .AutoSize = True
        .Width = 48
        .Top = 30
        .Left = 20
    End With
    
    With .Controls.Add("Forms.Combobox.1", "cboDept")
        .Width = 110
        .Top = 30
        .Left = 70
    End With

    ' add frame control
    With .Controls.Add("Forms.Frame.1", "frReports")
        .Caption = "Choose Report Type"
        .Top = 60
        .Left = 18
        .Height = 96
    End With
        
    ' add three check boxes
    Set objChkBox = .frReports.Controls.Add("Forms.CheckBox.1")
    With objChkBox
        .Name = "chk1"
        .Caption = "Last Month's Performance Report"
        .WordWrap = False
        .Left = 12
        .Top = 12
        .Height = 20
        .Width = 186
    End With
        
    Set objChkBox = .frReports.Controls.Add("Forms.CheckBox.1")
    With objChkBox
        .Name = "chk2"
        .Caption = "Last Qtr. Performance Report"
        .WordWrap = False
        .Left = 12
        .Top = 32
        .Height = 20
        .Width = 186
    End With
        
    Set objChkBox = .frReports.Controls.Add("Forms.CheckBox.1")
    With objChkBox
         .Name = "chk3"
         .Caption = Year(Now) - 1 & " Performance Report"
         .WordWrap = False
         .Left = 12
         .Top = 54
         .Height = 20
         .Width = 186
    End With

    ' Add and position OK and Cancel buttons
    With .Controls.Add("Forms.CommandButton.1", "cmdOK")
          .Caption = "OK"
          .Default = "True"
          .Height = 20
          .Width = 60
          .Top = objVBFrm.InsideHeight - .Height - 20
          .Left = objVBFrm.InsideWidth - .Width - 10
    End With
        
    With .Controls.Add("Forms.CommandButton.1", "cmdCancel")
        .Caption = "Cancel"
        .Height = 20
        .Width = 60
        .Top = objVBFrm.InsideHeight - .Height - 20
        .Left = objVBFrm.InsideWidth - .Width - 80
    End With
End With
    
'populate the combo box
With objVBComp.CodeModule
    x = .CountOfLines
    .InsertLines x + 1, "Sub UserForm_Initialize()"
    .InsertLines x + 2, vbTab & "With Me.cboDept"
    .InsertLines x + 3, vbTab & vbTab & ".addItem ""Marketing"""
    .InsertLines x + 4, vbTab & vbTab & ".addItem ""Sales"""
    .InsertLines x + 5, vbTab & vbTab & ".addItem ""Finance"""
    .InsertLines x + 6, vbTab & vbTab & _
        ".addItem ""Research & Development"""
    .InsertLines x + 7, vbTab & vbTab & _
        ".addItem ""Human Resources"""
    .InsertLines x + 8, vbTab & "End With"
    .InsertLines x + 9, "End Sub"
        
    ' write a procedure to handle the Cancel button
        
    Dim firstLine As Long
    With objVBComp.CodeModule
         firstLine = .CreateEventProc("Click", "cmdCancel")
        .InsertLines firstLine + 1, "    Unload Me"
    End With
    
    ' write a procedure to handle OK button
    sVBA = "Private Sub cmdOK_Click()" & vbCrLf
    sVBA = sVBA & "  Dim ctrl As Control" & vbCrLf
    sVBA = sVBA & "  Dim chkflag As Integer" & vbCrLf
    sVBA = sVBA & "  Dim strMsg As String" & vbCrLf
    sVBA = sVBA & "  If Me.cboDept.Value = """" Then " & vbCrLf
    sVBA = sVBA & "     MsgBox ""Select the Department.""" & _
                        vbCrLf
    sVBA = sVBA & "     Me.cboDept.SetFocus " & vbCrLf
    sVBA = sVBA & "     Exit Sub" & vbCrLf
    sVBA = sVBA & "  End If" & vbCrLf
    sVBA = sVBA & "  For Each ctrl In Me.Controls " & vbCrLf
    sVBA = sVBA & "     Select Case ctrl.Name" & vbCrLf
    sVBA = sVBA & "       Case ""chk1"", ""chk2"", ""chk3""" _
                            & vbCrLf
    sVBA = sVBA & "         If ctrl.Value = True Then" & vbCrLf
    sVBA = sVBA & "           strMsg = strMsg & vbCrLf & ctrl.Caption " _
                            & Chr(13) & vbCrLf
    sVBA = sVBA & "           chkflag = 1" & vbCrLf
    sVBA = sVBA & "         End If" & vbCrLf
    sVBA = sVBA & "     End Select" & vbCrLf
    sVBA = sVBA & "  Next" & vbCrLf
    sVBA = sVBA & "  If chkflag = 1 Then" & vbCrLf
    sVBA = sVBA & "    MsgBox ""Run the Report(s) for "" " & vbCrLf
    sVBA = sVBA & "    Me.cboDept.Value & "":"""
    sVBA = sVBA & "    & Chr(13) & Chr(13) & strMsg" & vbCrLf

    sVBA = sVBA & "  Else" & vbCrLf
    sVBA = sVBA & "    MsgBox ""Please select Report type.""" _
                        & vbCrLf
    sVBA = sVBA & "  End If" & vbCrLf
    sVBA = sVBA & "End Sub"

    .AddFromString sVBA
End With
Set objVBComp = Nothing
End Sub


Sub ReportGeneratorForm()
        Dim objVBComp As VBComponent

        Set objVBComp = Application.VBE.ActiveVBProject. _
            VBComponents.Add(vbext_ct_MSForm)
        With objVBComp
            .Name = "ReportGenerator"
            .Properties("Caption") = "My Report Form"
        End With
        Set objVBComp = Nothing
End Sub

Sub DeleteReportGenerator()
    Dim objVBComp As VBComponent
    
    Set objVBComp = Application.VBE.ActiveVBProject. _
        VBComponents("ReportGenerator")
    Application.VBE.ActiveVBProject.VBComponents.Remove objVBComp
End Sub
