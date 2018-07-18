Attribute VB_Name = "SpellingOptionsObject"
Option Explicit

Sub SpellCheck()
    ' set spelling options
    With Application.SpellingOptions
        .UserDict = "Custom.dic"
        .IgnoreCaps = True
        .IgnoreMixedDigits = True
        .SuggestMainOnly = False
        .IgnoreFileNames = True
    End With

    ' run a spell check
    Cells.CheckSpelling
End Sub


