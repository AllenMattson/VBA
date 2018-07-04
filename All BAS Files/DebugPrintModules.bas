Attribute VB_Name = "DebugPrintModules"

Public Sub execute()
Dim AppArray() As String

AppArray() = Split(ListAllMacroNames, " ")

For i = 0 To UBound(AppArray)

temp = AppArray(i)

If temp <> "" Then

    If temp <> "execute" And temp <> "ListAllMacroNames" And temp <> "ThisWorkbook" Then
    Application.Run (AppArray(i))

    End If

End If

Next i

End Sub
Function ListAllMacroNames() As String

Dim pj As VBProject
Dim vbcomp As VBComponent
Dim curMacro As String, newMacro As String, MacroName As String, OldPJname As String
Dim x As String
Dim y As String
Dim macros As String

On Error Resume Next
curMacro = ""
'Documents.Add

For Each pj In Application.VBE.VBProjects
If pj.Name = "" Or pj.Name <> oldpj.Name Then
    Debug.Print ("********************************************************************************")
    Debug.Print ("********************************************************************************")
    Debug.Print ("The VB Project Name is: " & pj.Name)
End If
     For Each vbcomp In pj.VBComponents
            If Not vbcomp Is Nothing And vbcomp.Name <> "ThisWorkbook" And Left(vbcomp.Name, 5) <> "Sheet" Then
                If vbcomp.CodeModule <> "" Then ' "Module_name" Then
                Debug.Print ("The VB Module Name is: " & vbcomp.Name)
                    For i = 1 To vbcomp.CodeModule.CountOfLines
                       newMacro = vbcomp.CodeModule.ProcOfLine(Line:=i, _
                          ProcKind:=vbext_pk_Proc)
                        newMacro = MacroName
                       If curMacro <> newMacro Then
                          vbcomp = newMacro
                            If curMacro <> "" And curMacro <> "app_NewDocument" Then
                                macros = curMacro + " " + macros
                            End If
                            
                       End If
                    Next
                End If
            End If
    Next
Next
Debug.Print ("********************************************************************************")
Debug.Print ("********************************************************************************")
End Function


