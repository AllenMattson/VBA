Function FindLastRow( _
    ByVal Col As Long) As Long
    
    'Gives you the last cell with data in the specified row
    '  Will not work correctly if the last row is hidden

    With ActiveSheet
        FindLastRow = .Cells(.Rows.Count, Col).End(xlUp).Row
    End With
    
End Function



Sub SampleUsage()
    
    'Sample usage for FindLastRow()
    
    Dim LastRow As Long
    Dim ColNum As Long
    
    ColNum = 3
    LastRow = FindLastRow(ColNum)
    
    MsgBox "The last row in column number " & ColNum & " is " & LastRow

End Sub
