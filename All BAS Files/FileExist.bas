Function FileExist( _
    ByVal FilePath As String, _
    ByVal FileName As String, _
    Optional ByVal FileType As String = "NotGiven") As Boolean

    Dim fName As String

    FileExist = False
    
    FilePath = IIf(Right(FilePath, 1) <> "\", FilePath & "\", FilePath)
    
    If FileType <> "NotGiven" Then

        If Right(FileName, 1) = "." Then
            FileName = Left(FileName, Len(FileName) - 1)
        End If

        If Left(FileType, 1) <> "." Then
            FileType = "." & FileType
        End If
        
        fName = FilePath & FileName & FileType
        
    Else
        fName = FilePath & FileName
        
    End If
      
    If Dir(fName) <> "" Then
        FileExist = True
    End If
        
End Function



Sub SampleUsage()

    'Sample usage for FileExist function
    '  Note: file extension is included in the filename
    

    Dim fPath As String     'Directory where files should be
    Dim fName As String     'File name (pulled from spreadsheet)

    fPath = "C:\"
    fName = "Test.xls"
    
    If FileExist(fPath, fName) Then
        MsgBox "File Found"
    Else
        MsgBox "File Not Found"
    End If

End Sub

