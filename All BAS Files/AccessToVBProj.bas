Attribute VB_Name = "AccessToVBProj"
Option Explicit

Sub AccessToVBProj()
    Dim objVBProject As VBProject
    Dim strMsg1 As String
    Dim strMsg2 As String
    Dim response As Integer

    On Error Resume Next

    If Application.Version >= "15.0" Then
        Set objVBProject = ActiveWorkbook.VBProject

        strMsg2 = "The access to the VBA "
        strMsg2 = strMsg2 + " project must be trusted for this "
        strMsg2 = strMsg2 + "procedure to work."
        strMsg2 = strMsg2 + vbCrLf + vbCrLf
        strMsg2 = strMsg2 + " Click 'OK' to view instructions,"
        strMsg2 = strMsg2 + "  or click 'Cancel' to exit."

        If Err.Number <> 0 Then
            strMsg1 = "Please change the security settings to "
            strMsg1 = strMsg1 & "allow access to the VBA project:"
            strMsg1 = strMsg1 & Chr(10) & "1. "
            strMsg1 = strMsg1 & "Choose Developer | Macro Security."
            strMsg1 = strMsg1 & Chr(10) & "2. "
            strMsg1 = strMsg1 & "Check the 'Trust access to the " _
                 & " VBA project object model'. "
            strMsg1 = strMsg1 & Chr(10) & "3. Click OK."

            response = MsgBox(strMsg2, vbCritical + vbOKCancel, _
                        "Access to VB Project is not trusted")

                If response = 1 Then
                    Workbooks.Add
                    With ActiveSheet
                      .Shapes.AddTextbox _
                      (msoTextOrientationHorizontal, _
                         Left:=0, Top:=0, Width:=300, _
                         Height:=100).Select
                      Selection.Characters.Text = strMsg1
                      .Shapes(1).Fill.PresetTextured _
                           PresetTexture:=msoTextureBlueTissuePaper
                      .Shapes(1).Shadow.Type = msoShadow6
                    End With
                End If
            Exit Sub
        End If

        MsgBox "There are " & objVBProject.References.Count & _
            " project references in " & objVBProject.Name & "."
    End If
End Sub

Function IsProjProtected() As Boolean
    Dim objVBProj As VBProject

    Set objVBProj = ActiveWorkbook.VBProject

    If objVBProj.Protection = vbext_pp_locked Then
        IsProjProtected = True
    Else
        IsProjProtected = False
    End If
End Function



