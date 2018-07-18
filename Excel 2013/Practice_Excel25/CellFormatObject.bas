Attribute VB_Name = "CellFormatObject"
Option Explicit

Sub Reformat()
    ' Set search criteria
    With Application.FindFormat.Font
        .Name = "Arial"
        .FontStyle = "Regular"
        .Size = 10
    End With

    ' Set replacement criteria
    With Application.ReplaceFormat.Font
        .Name = "Tahoma"
        .FontStyle = "Bold"
        .Size = 11
    End With

    With Application.ReplaceFormat.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
    End With

    ' Perform the replace
    Sheets(1).UsedRange.Replace _
        what:="", _
        Replacement:="", _
        SearchFormat:=True, _
        ReplaceFormat:=True

    ' Reset the Find and Replace formats
    Application.FindFormat.Clear
    Application.ReplaceFormat.Clear
End Sub



