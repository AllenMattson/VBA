Attribute VB_Name = "Module5"
Option Explicit

Sub DeleteModuleCode(strName As String)
    Dim objVBProj As VBProject
    Dim objVBCode As CodeModule
    Dim firstLn As Long
    Dim totLn As Long

    Set objVBProj = ThisWorkbook.VBProject
    Set objVBCode = objVBProj.VBComponents(strName).CodeModule
    With objVBCode
        firstLn = 1
        totLn = .CountOfLines
        .DeleteLines firstLn, totLn
    End With

    Set objVBProj = Nothing
    Set objVBCode = Nothing
End Sub


