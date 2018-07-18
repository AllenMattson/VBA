Attribute VB_Name = "Module15"
Option Explicit

Sub ListPrjCompRef()
  Dim objVBPrj As VBIDE.VBProject
  Dim objVBCom As VBIDE.VBComponent
  Dim vbRef As VBIDE.Reference

    ' List VBA projects as well as references and
    ' component names they contain
    For Each objVBPrj In Application.VBE.VBProjects
        Debug.Print objVBPrj.Name
        Debug.Print vbTab & "References"
        For Each vbRef In objVBPrj.References
            With vbRef
               Debug.Print vbTab & vbTab & .Name & "---" & .FullPath
            End With
        Next
        Debug.Print vbTab & "Components"
        For Each objVBCom In objVBPrj.VBComponents
            Debug.Print vbTab & vbTab & objVBCom.Name
        Next
    Next
    Set vbRef = Nothing
    Set objVBCom = Nothing
    Set objVBPrj = Nothing
End Sub


