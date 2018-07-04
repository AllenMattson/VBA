Attribute VB_Name = "CopyAllModules_Revised"
Sub CopyAllModules(wkbFrom As String, wkbTo As String)
Dim objVBComp As VBComponent
Dim wkb As Workbook
Dim strFile As String
Set wkb = Workbooks(wkbFrom)
On Error Resume Next
Workbooks(wkbTo).Activate

If Err.Number <> 0 Then Workbooks.Open wkbTo

strFile = wkb.Path & "\vbCode.bas"

If Dir(strFile) <> "" Then Kill strFile

For Each objVBComp In wkb.VBProject.VBComponents
    If objVBComp.Type <> vbext_ct_Document Then
        objVBComp.Export strFile
        Workbooks(wkbTo).VBProject.VBComponents.Import strFile
    End If
Next
        
Set objVBComp = Nothing
Set wkb = Nothing
End Sub
Sub CopyAllModulesRevised(wkbFrom As String, wkbTo As String)
Dim objVBComp As VBComponent
Dim wkb As Workbook
Dim strFile As String
Set wkb = Workbooks(wkbFrom)
On Error Resume Next
Workbooks(wkbTo).Activate
If Err.Number <> 0 Then Workbooks.Open wkbTo
    
strFile = wkb.Path & "\vbCode.bas"
    
If Dir(strFile) <> "" Then Kill strFile

For Each objVBComp In wkb.VBProject.VBComponents
    If objVBComp.Type <> vbext_ct_Document Then
        objVBComp.Export strFile

            With Workbooks(wkbTo)
                If Len(.VBProject.VBComponents(objVBComp.Name).Name) = 0 Then
                    Workbooks(wkbTo).VBProject.VBComponents.Import strFile
                End If
            End With

    End If
Next

Set objVBComp = Nothing
Set wkb = Nothing
End Sub
