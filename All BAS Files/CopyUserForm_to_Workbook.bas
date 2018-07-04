Attribute VB_Name = "CopyUserForm_to_Workbook"
Sub UserFormCopy(strFileName As String)

Dim objVBComp As VBComponent
Dim wkb As Workbook

On Error Resume Next
Set wkb = Workbooks(strFileName)

If Err.Number <> 0 Then
    Workbooks.Open ActiveWorkbook.Path & "\" & strFileName
    Set wkb = Workbooks(strFileName)
End If

For Each objVBComp In ThisWorkbook.VBProject.VBComponents
    If objVBComp.Type = 3 Then  ' this is a UserForm
        ' export the UserForm to disk
        objVBComp.Export Filename:=objVBComp.Name
        ' import the UserForm to a specific workbook
        wkb.VBProject.VBComponents.Import Filename:=objVBComp.Name
        ' delete two form files created by the Export method
        Kill objVBComp.Name
        Kill objVBComp.Name & ".frx"
    End If
Next

' add a standard module to the workbook and write code to show the UserForm
Set objVBComp = wkb.VBProject.VBComponents.Add(vbext_ct_StdModule)
objVBComp.CodeModule.AddFromString "Sub ShowReportSelector()" & vbCrLf & "ReportSelector.Show" & vbCrLf & "End Sub" & vbCrLf


' close the Code pane
objVBComp.CodeModule.CodePane.Window.Close

' run the ShowReportSelector procedure to display the form
Application.Run wkb.Name & "!ShowReportSelector"

Set objVBComp = Nothing
Set wkb = Nothing
End Sub

