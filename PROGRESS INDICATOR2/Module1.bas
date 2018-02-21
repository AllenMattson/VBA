Attribute VB_Name = "Module1"
Sub ShowUserForm()
    With UserForm1
        'Use a color from the workbook's theme
        .LabelProgress.BackColor = ActiveWorkbook.Theme. _
            ThemeColorScheme.Colors(msoThemeAccent1)
        .LabelProgress.Width = 0
        .Show
    End With
End Sub


