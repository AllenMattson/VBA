Attribute VB_Name = "CharactersObject"
Option Explicit

Sub Format1stLetters()
    Dim myChr As Characters
    Dim cell As Variant
    Dim i As Integer

    For Each cell In Sheets(1).UsedRange
        If Not IsNumeric(cell) Then
            Set myChr = cell.Characters(1, 1)
            myChr.Font.Color = RGB(128, 0, 255)
            For i = 1 To Len(cell.Text)
                If Asc(Mid(cell, i, 1)) = 32 Then
                    Set myChr = cell.Characters(i + 1, 1)
                    myChr.Font.Color = RGB(255, 0, 0)
                End If
            Next
        End If
    Next
End Sub


