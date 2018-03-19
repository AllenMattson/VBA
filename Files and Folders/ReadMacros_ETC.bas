Attribute VB_Name = "ReadMacros_ETC"
Sub Macro_read_code()
With ThisWorkbook.VBProject.VBComponents("Macroos").CodeModule
c00 = .Lines(.ProcStartLine("macro3", 0), .ProcCountLines("macro3", 0))
End With
End Sub
Sub Code_in_workbookmodule_lezen()
'3.1.1 Read the complete VBA code
With ThisWorkbook.VBProject.VBComponents(1).CodeModule
MsgBox .Lines(1, .CountOfLines)
End With
End Sub
Sub Modules_copy()
'2.8.5 All macromodules and userforms: copy
For Each cp In ThisWorkbook.VBProject.VBComponents
If cp.Type <> 100 Then
If Workbooks.Count = 1 Then Workbooks.Add
With Workbooks(2).VBProject.VBComponents.Add(vbext_ct_MSForm)
.Name = cp.Name
.CodeModule.AddFromString cp.CodeModule.Lines(1, cp.CodeModule.CountOfLines)
End With
End If
Next
End Sub
Sub Code_in_workbookmodule_lezen()
'3.1.1 Read the complete VBA code
With ThisWorkbook.VBProject.VBComponents(1).CodeModule
MsgBox .Lines(1, .CountOfLines)
End With
End Sub
Sub Modules_namen()
For j = 1 To ThisWorkbook.VBProject.VBComponents.Count
MsgBox ThisWorkbook.VBProject.VBComponents(j).Name
Next
End Sub
