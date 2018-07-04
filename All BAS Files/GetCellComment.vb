Function GetCellComment(cell As Range) As String
    On Error Resume Next

    GetCellComment = cell.Comment.Text

    If Err <> 0 Then CellComment = ""

End Function
