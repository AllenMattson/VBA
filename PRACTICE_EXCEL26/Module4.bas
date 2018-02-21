Attribute VB_Name = "Module4"
Option Explicit

Sub DeleteModule(strName As String)
    Dim objVBProj As VBProject
    Dim objVBComp As VBComponent

    Set objVBProj = ThisWorkbook.VBProject

    Set objVBComp = objVBProj.VBComponents(strName)

    objVBProj.VBComponents.Remove objVBComp

    Set objVBComp = Nothing
    Set objVBProj = Nothing
End Sub

