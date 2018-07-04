Sub ListFilesInFolder( _
    ByVal SourceFolderName As String, _
    Optional ByVal IncludeSubfolders As Boolean)

    'Originally created by Leith Ross
    'Retreived from http://www.excelforum.com/excel-programming/645683-list-files-in-folder.html
    
    'Lists information about the files in SourceFolder
    'Example: ListFilesInFolder "C:\FolderName\", True

    On Error GoTo ExitSub

    Dim FSO As Object
    Dim SourceFolder As Object
    Dim SubFolder As Object
    Dim FileItem As Object
    Dim r As Long

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set SourceFolder = FSO.GetFolder(SourceFolderName)
    r = Range("A65536").End(xlUp).Row + 1

    For Each FileItem In SourceFolder.Files
        'display file properties
        Cells(r, 1).Formula = FileItem.Name
        
        '***** Remove the ' character in lines below to get information *****
        'Cells(r, 2).Formula = FileItem.Path
        'Cells(r, 3).Formula = FileItem.Size
        'Cells(r, 4).Formula = FileItem.DateCreated
        'Cells(r, 5).Formula = FileItem.DateLastModified
        'Cells(r, 6).Formula = GetFileOwner(SourceFolder.Path, FileItem.Name)
        
        r = r + 1 ' next row number
        'X = SourceFolder.Path
    
    Next FileItem

    If IncludeSubfolders Then
        For Each SubFolder In SourceFolder.SubFolders
            ListFilesInFolder SubFolder.Path, True
        Next SubFolder
    End If
    
    '***** Remove the single ' character in the below lines to adjust the column windths
    'Columns("A:G").ColumnWidth = 4
    'Columns("H:I").AutoFit
    'Columns("J:L").ColumnWidth = 12
    'Columns("M:P").ColumnWidth = 8

ExitSub:
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing

End Sub


Function GetFileOwner( _
    ByVal FilePath As String, _
    ByVal FileName As String)

    'Originally created by Leith Ross
    'Retreived from http://www.excelforum.com/excel-programming/645683-list-files-in-folder.html

    On Error GoTo ExitSub
    
    Dim objFolder As Object
    Dim objFolderItem As Object
    Dim objShell As Object

    FileName = StrConv(FileName, vbUnicode)
    FilePath = StrConv(FilePath, vbUnicode)

    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.Namespace(StrConv(FilePath, vbFromUnicode))

    If Not objFolder Is Nothing Then
        Set objFolderItem = objFolder.ParseName(StrConv(FileName, vbFromUnicode))
    End If

    If Not objFolderItem Is Nothing Then
        GetFileOwner = objFolder.GetDetailsOf(objFolderItem, 8)
    Else
        GetFileOwner = ""
    End If

ExitSub:
    Set objShell = Nothing
    Set objFolder = Nothing
    Set objFolderItem = Nothing

End Function


Sub SampleUsage()

    Dim fPath As String
    
    fPath = "V:\Corporate\Tax\Private\Indirect Tax\Certs"

    Call ListFilesInFolder(fPath, True)
    
End Sub
