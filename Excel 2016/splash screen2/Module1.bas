Attribute VB_Name = "Module1"
Sub TestSplashScreen()
    With UserForm1
        .StartUpPosition = 0
        .Left = Application.Left + (0.5 * Application.Width) - (0.5 * .Width)
        .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
        .Show
    End With
End Sub

Private Sub KillTheForm()
    Unload UserForm1
End Sub


