Attribute VB_Name = "Module1"
Option Explicit

Sub ShowThemeColors()
' Puts the theme colors in a range of cells
  Dim r As Long, c As Long
  For r = 1 To 6
    For c = 1 To 10
        With Cells(r, c).Interior
        .ThemeColor = c
        Select Case c
            Case 1 'Text/Background 1
            Select Case r
                Case 1: .TintAndShade = 0
                Case 2: .TintAndShade = -0.05
                Case 3: .TintAndShade = -0.15
                Case 4: .TintAndShade = -0.25
                Case 5: .TintAndShade = -0.35
                Case 6: .TintAndShade = -0.5
            End Select
        Case 2 'Text/Background 2
            Select Case r
                Case 1: .TintAndShade = 0
                Case 2: .TintAndShade = 0.5
                Case 3: .TintAndShade = 0.35
                Case 4: .TintAndShade = 0.25
                Case 5: .TintAndShade = 0.15
                Case 6: .TintAndShade = 0.05
            End Select
        Case 3 'Text/Background 3
            Select Case r
                Case 1: .TintAndShade = 0
                Case 2: .TintAndShade = -0.1
                Case 3: .TintAndShade = -0.25
                Case 4: .TintAndShade = -0.5
                Case 5: .TintAndShade = -0.75
                Case 6: .TintAndShade = -0.9
            End Select
        Case Else  'Text/Background 4, and Accent 1-6
            Select Case r
                Case 1: .TintAndShade = 0
                Case 2: .TintAndShade = 0.8
                Case 3: .TintAndShade = 0.6
                Case 4: .TintAndShade = 0.4
                Case 5: .TintAndShade = -0.25
                Case 6: .TintAndShade = -0.5
            End Select
        End Select
        Cells(r, c) = .TintAndShade
        End With
    Next c
  Next r
End Sub

