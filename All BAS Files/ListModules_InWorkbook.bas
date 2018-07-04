Attribute VB_Name = "ListModules_InWorkbook"
Sub ModuleList()

Dim objVBComp As VBComponent
Dim listArray()
Dim i As Integer

If ThisWorkbook.VBProject.Protection = vbext_pp_locked Then
    MsgBox "Please unprotect the project to run this " & "procedure."
    Exit Sub
End If

i = 2

For Each objVBComp In ThisWorkbook.VBProject.VBComponents
    ReDim Preserve listArray(1 To 2, 1 To i - 1)
    listArray(1, i - 1) = objVBComp.Name
    listArray(2, i - 1) = GetModuleType(objVBComp)
    i = i + 1
Next
    
With ActiveSheet
    .Cells(1, 1).Resize(1, 2).Value = Array("Module Name", "Module Type")
    .Cells(2, 1).Resize(UBound(listArray, 2), UBound(listArray, 1)).Value = Application.Transpose(listArray)
    .Columns("A:B").AutoFit
End With

Set objVBComp = Nothing
End Sub

Function GetModuleType(comp As VBComponent)
Select Case comp.Type
    Case vbext_ct_StdModule
        GetModuleType = "Standard module"
    Case vbext_ct_ClassModule
        GetModuleType = "Class module"
    Case vbext_ct_MSForm
        GetModuleType = "Microsoft Form"
    Case vbext_ct_ActiveXDesigner
        GetModuleType = "ActiveX Designer"
    Case vbext_ct_Document
        GetModuleType = "Document module"
    Case Else
        GetModuleType = "Unknown"
End Select
End Function
