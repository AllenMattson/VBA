Attribute VB_Name = "CREATE_ADD_MODULE"
Option Explicit

Sub CreateModule()
    Dim modType As Integer
    Dim strName As String
    Dim strPrompt As String

    strPrompt = "Enter a number representing the type of module:"
    strPrompt = strPrompt & vbCr & "1 (Standard Module)"
    strPrompt = strPrompt & vbCr & "2 (Class Module)"
    modType = Val(InputBox(prompt:=strPrompt, Title:="Insert Module"))
    If modType = 0 Then Exit Sub
    strName = InputBox("Enter the name you want to assign to " & _
              "new module", "Module Name")
    If strName = "" Then Exit Sub
    AddModule modType, strName
End Sub


Sub AddModule(modType As Integer, strName As String)
    Dim objVBProj As VBProject
    Dim objVBComp As VBComponent

    If InStr(1, "1, 2", modType) = 0 Then Exit Sub

    Set objVBProj = ThisWorkbook.VBProject
    Set objVBComp = objVBProj.VBComponents.Add(modType)
    objVBComp.Name = strName

    Application.Visible = True

    Set objVBComp = Nothing
    Set objVBProj = Nothing
End Sub



