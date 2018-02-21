Attribute VB_Name = "Module7"
Option Explicit

Sub CopyAModule(wkbFrom As String, _
                wkbTo As String, _
                strFromMod As String)
    Dim wkb As Workbook
    Dim strFile As String

    Set wkb = Workbooks(wkbFrom)

    strFile = wkb.Path & "\vbCode.bas"
    wkb.VBProject.VBComponents(strFromMod).Export strFile

    On Error Resume Next
    Set wkb = Workbooks(wkbTo)
    If Err.Number <> 0 Then
        Workbooks.Open wkbTo
        Set wkb = Workbooks(wkbTo)
    End If

    wkb.VBProject.VBComponents.Import strFile
    wkb.Save

    Set wkb = Nothing
End Sub



