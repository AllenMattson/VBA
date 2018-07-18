Attribute VB_Name = "Module1"
Option Explicit

Sub ShowUserFormMp()
    With UProgressMp
        'Use a color from the workbook's theme
        .lblProgress.BackColor = ActiveWorkbook.Theme. _
            ThemeColorScheme.Colors(msoThemeAccent1)
        .lblProgress.Width = 0
        .Show
    End With
End Sub


Sub ShowUserFormHidden()
    With UProgressH
        'Use a color from the workbook's theme
        .lblProgress.BackColor = ActiveWorkbook.Theme. _
            ThemeColorScheme.Colors(msoThemeAccent1)
        .lblProgress.Width = 0
        .Show
    End With
End Sub

