Function GetFilePathFromString(FileName As String) As String

    GetFilePathFromString = Left(FileName, InStrRev(Path, "\"))

End Function
