Attribute VB_Name = "Module11"
Option Explicit

Sub DeleteProc(strModName As String, strProcName As String)
    Dim objVBProj As VBProject
    Dim objVBCode As CodeModule
    Dim firstLn As Long
    Dim totLn As Long

    Set objVBProj = ThisWorkbook.VBProject
    Set objVBCode = objVBProj.VBComponents(strModName).CodeModule
    With objVBCode
      firstLn = .ProcStartLine(strProcName, vbext_pk_Proc)
      totLn = .ProcCountLines(strProcName, vbext_pk_Proc)
     .DeleteLines firstLn, totLn
    End With

    Set objVBProj = Nothing
    Set objVBCode = Nothing
End Sub


