Function FindLastCol( _
    ByVal Row As Long) As Long
    
    'Gives you the last cell with data in the specified row
    '  Will not work correctly if the last row is hidden

    With ActiveSheet
        FindLastCol = .Cells(Row, .Columns.Count).End(xlToLeft).Column
    End With
    
End Function



Sub SampleUsage()
    
    'Sample usage for FindLastCol()
    
    Dim LastCol As Long
    Dim RowNum As Long
    
    RowNum = 7
    LastCol = FindLastCol(RowNum)
    
    MsgBox "The last column in row number " & RowNum & " is " & LastCol

End Sub
