Attribute VB_Name = "ListAllModules_And_Procedures"
Sub ListAll_Modules_and_Procedures()

Sheets.add
Cells(1, 1).value = "Module Name"
Cells(1, 2).value = "Procedure Name"
Cells(1, 3).value = "Number of Lines in Procedure"
Dim objVBProj As VBProject
Dim objVBComp As VBComponent
Dim objVBCode As CodeModule
Dim strCurrent As String
Dim strPrevious As String
Dim x As Integer

Set objVBProj = ThisWorkbook.VBProject

For Each objVBComp In objVBProj.VBComponents
    If InStr(1, "1, 2", objVBComp.Type) Then
        Set objVBCode = objVBComp.CodeModule
        'Debug.Print objVBComp.name
        Range("B900").End(xlUp).Offset(1, -1).value = objVBComp.name
            For x = objVBCode.CountOfDeclarationLines + 1 To objVBCode.CountOfLines
                strCurrent = objVBCode.ProcOfLine(x, vbext_pk_Proc)
                If strCurrent <> strPrevious Then
                    'Debug.Print vbTab & objVBCode.ProcOfLine(x, vbext_pk_Proc)
                    Range("B900").End(xlUp).Offset(1, 0).value = objVBCode.ProcOfLine(x, vbext_pk_Proc)
                    strPrevious = strCurrent
                End If
                'line count of procedure
                Range("B900").End(xlUp).Offset(0, 1).value = x
            Next
    End If
Next

Set objVBCode = Nothing
Set objVBComp = Nothing
Set objVBProj = Nothing
columns.AutoFit
End Sub
