Attribute VB_Name = "FillFileNamesFromList"
Sub FillFileNamesFromList()

    'Takes a list of file names without extensions, verifies they exist, then
    '  pastes the file name w/ extension in the result column.
    
    'NOTE: If there are file with the same name but different extension,
    '  the last extension listed in the array is passed back to the result
    '  column. In the code below, PDF will over-right (trump) tif
    
    'Requires the following additiona functions
    '  FindLastRow
    '  FileExist

    Dim FirstRow As Long
    Dim LastRow As Long
    Dim iRow As Long
    Dim iFileType As Long
    Dim FileTypeCnt As Long
    Dim SearchCol As Long
    Dim LinkCol As Long
    Dim FileNameCol As Long
    Dim FileTypes(1 To 5) As String
    Dim SearchFolder As String
    Dim SearchName As String
    Dim strLink As String
    Dim IncludeHyperLink As Boolean
    Dim HyperLinkCol As Long

    'Settings: These must be defined
    FirstRow = 2                         'First row with data to search
    SearchCol = 11                       'Column where search names exist
    FileNameCol = 5                      'Column where the file name will be placed
    IncludeHyperLink = True              'Option to include a hyperlink to the file
    HyperLinkCol = 10                     'Column where hyperlinks will be placed (Must be set if above is True)
    
    'Find the last row in the NameCol
    LastRow = FindLastRow(SearchCol)
    
    'Folder where the files are located. Be sure to add the trailing "\"
    'SearchFolder = "C:\Users\Allen\Documents\Books\106 visual basic ebooks" \ ""

    'Defines the file type to search for. Add more as needed.
    FileTypes(1) = "tif"
    FileTypes(2) = "pdf"
    FileTypes(3) = "txt"
    FileTypes(4) = "doc"
    FileTypes(5) = ""


    For iRow = FirstRow To LastRow
        SearchName = Cells(iRow, SearchCol)
         SearchFolder = Cells(iRow, 1) & "\"
        For iFileType = LBound(FileTypes) To UBound(FileTypes)
            If FileTypes(iFileType) <> "" Then
                
                If FileExist(SearchFolder, SearchName, FileTypes(iFileType)) Then
                    
                    Cells(iRow, FileNameCol) = SearchName & "." & FileTypes(iFileType)
                        
                    'Includes hyperlink to file if set to true
                    If IncludeHyperLink Then
                        strLink = "=HyperLink(" & Chr(34) & SearchFolder & _
                            SearchName & "." & FileTypes(iFileType) & Chr(34) & "," & _
                            Chr(34) & "Link" & Chr(34) & ")"
                        
                        Cells(iRow, HyperLinkCol).Formula = strLink
                    End If
                End If
            End If
        Next
    Next

End Sub
Function FindLastRow( _
    ByVal Col As Long) As Long
    
    'Gives you the last cell with data in the specified row
    '  Will not work correctly if the last row is hidden

    With ActiveSheet
        FindLastRow = .Cells(.Rows.Count, Col).End(xlUp).Row
    End With
    
End Function
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

