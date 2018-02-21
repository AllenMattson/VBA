Attribute VB_Name = "Module1"
Option Explicit

Sub ShowForm()
    UserForm1.Show
End Sub

Sub ShowComponents(vbfilename)
    Dim VBP As VBIDE.VBProject
    Dim VBC As VBComponent
    Dim x
    Dim row As Long
    
    Set VBP = Workbooks(vbfilename).VBProject
    
'   Make sure access to the VBProject is allowed
    On Error Resume Next
    Set x = ActiveWorkbook.VBProject
    If Err <> 0 Then
        MsgBox "Your security settings do not allow this macro to run.", vbCritical
        On Error GoTo 0
        Exit Sub
    End If

    Cells.ClearContents

'   Write headers
    Range("A1:C1") = Array("Name", "Type", "Code Lines")
    Range("A1:C1").Font.Bold = True
    row = 1
    For Each VBC In VBP.VBComponents
        row = row + 1
'       Name
        Cells(row, 1) = VBC.Name
        
'       Type
        Select Case VBC.Type
            Case vbext_ct_StdModule
                Cells(row, 2) = "Module"
            Case vbext_ct_ClassModule
                Cells(row, 2) = "Class Module"
            Case vbext_ct_MSForm
                Cells(row, 2) = "UserForm"
            Case vbext_ct_Document
                Cells(row, 2) = "Document Module"
        End Select
'       Lines of code
        Cells(row, 3) = VBC.CodeModule.CountOfLines
    Next VBC
End Sub

