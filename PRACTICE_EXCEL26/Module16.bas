Attribute VB_Name = "Module16"
Option Explicit

Sub AddRef()
    Dim objVBProj As VBProject

    Set objVBProj = ThisWorkbook.VBProject

    On Error GoTo ErrorHandle
    objVBProj.References.AddFromFile _
        "C:\Windows\System32\scrrun.dll"
    MsgBox "The reference to the Microsoft Scripting " _
        & "Runtime was set."
    Application.SendKeys "%tr"

ExitHere:
    Set objVBProj = Nothing
    Exit Sub
ErrorHandle:
    MsgBox "The reference to the Microsoft Scripting " & _
        " Runtime already exists."
    GoTo ExitHere
End Sub

Sub AddRef_FromGuid()
    Dim objVBProj As VBProject
    Dim i As Integer
    Dim strName As String
    Dim strGuid As String
    Dim strMajor As Long
    Dim strMinor As Long

    Set objVBProj = ActiveWorkbook.VBProject

    ' Find out what libraries are already installed
    For i = 1 To objVBProj.References.Count
          strName = objVBProj.References(i).Name
          strGuid = objVBProj.References(i).GUID
          strMajor = objVBProj.References(i).Major
          strMinor = objVBProj.References(i).Minor
          Debug.Print strName & " - " & strGuid & _
            ", " & strMajor & ", " & strMinor
    Next i

    ' add a reference to the Microsoft DAO 3.6 Object library
    On Error Resume Next
    ThisWorkbook.VBProject.References.AddFromGuid _
       "{00025E01-0000-0000-C000-000000000046}", 5, 0
End Sub

Sub RemoveRef()
    Dim objVBProj As VBProject
    Dim objRef As Reference
    Dim sRefFile As String

    Set objVBProj = ActiveWorkbook.VBProject

    ' Loop through the references and delete
    ' the reference to DAO library
    For Each objRef In objVBProj.References
        If InStr(1, objRef.Description, "DAO 3.6") > 0 Then
            objVBProj.References.Remove objRef
            Exit For
        End If
    Next objRef
End Sub

Function IsBrokenRef(strRef As String) As Boolean
       'call from Immediate window using the following two lines:
       ' ref = IsBrokenRef("OLE Automation")
       ' Print ref

Dim objVBProj As VBProject
        Dim objRef As Reference

        Set objVBProj = ThisWorkbook.VBProject

        For Each objRef In objVBProj.References

        If strRef = objRef.Name And objRef.IsBroken Then
            IsBrokenRef = True
            Exit Function
        End If
        Next

        IsBrokenRef = False
    End Function



