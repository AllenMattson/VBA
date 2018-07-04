Attribute VB_Name = "AddNewProc"
Option Explicit

Sub AddNewProc(strModName As String)
    Dim objVBCode As CodeModule
    Dim objVBProj As VBProject
    Dim strProc As String

    Set objVBProj = ThisWorkbook.VBProject

    Set objVBCode = objVBProj.VBComponents( _
        strModName).CodeModule

    strProc = "Sub CreateWorkBook()" & Chr(13)
    strProc = strProc & Chr(9) & "Workbooks.Add" & Chr(13)
    strProc = strProc & Chr(9) & _
        "ActiveSheet.Name = ""Test"" & Chr (13)"
    strProc = strProc & "End Sub"

    Debug.Print strProc

    With objVBCode
        .InsertLines .CountOfLines + 1, strProc
    End With

    Set objVBCode = Nothing
    Set objVBProj = Nothing
End Sub


