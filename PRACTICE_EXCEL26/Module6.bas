Attribute VB_Name = "Module6"
Option Explicit

Sub DeleteEmptyModules()
    Dim objVBComp As VBComponent

    Const vbext_ct_StdModule As Long = 1
    Const vbext_ct_ClassModule As Long = 2

    For Each objVBComp In ActiveWorkbook.VBProject.VBComponents
      Select Case objVBComp.Type
        Case vbext_ct_StdModule, vbext_ct_ClassModule
          If objVBComp.CodeModule.CountOfLines < 3 Then
            Debug.Print "(deleted) " & objVBComp.Name & vbTab & _
                "declarations: " & objVBComp.CodeModule. _
            CountOfDeclarationLines & vbTab & _
                "Total code Lines: " & _
                objVBComp.CodeModule.CountOfLines
            ActiveWorkbook.VBProject.VBComponents. _
                Remove objVBComp
          End If
      End Select
    Next
    Set objVBComp = Nothing
End Sub

Function ModuleExists(strModName As String) As Boolean
    Dim objVBProj As VBProject

    Set objVBProj = ThisWorkbook.VBProject

    On Error Resume Next

    ModuleExists = Len(objVBProj.VBComponents(strModName).Name) <> 0
End Function

Function ProcExists(strModName As String, _
                    strProcName As String) As Boolean

    Dim objVBProj As VBProject

    Set objVBProj = ThisWorkbook.VBProject

    On Error Resume Next

    ' first find out if the specified module exists
    If ModuleExists(strModName) = True Then
        ProcExists = objVBProj.VBComponents(strModName) _
            .CodeModule.ProcStartLine(strProcName, vbext_pk_Proc) <> 0
    End If
End Function


