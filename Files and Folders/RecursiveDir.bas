Attribute VB_Name = "RecursiveDir"
Option Explicit

Sub GetAllFiles()
    Dim Directory As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Select a location containing the files you want to list."
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        Else
            Directory = .SelectedItems(1) & "\"
        End If
    End With
    
    Cells.ClearContents
    Call RecursiveDir(Directory)
End Sub

Public Sub RecursiveDir(ByVal CurrDir As String)
    Dim Dirs() As String
    Dim NumDirs As Long
    Dim FileName As String
    Dim PathAndName As String
    Dim i As Long
    Dim Filesize As Double

'   Make sure path ends in backslash
    If Right(CurrDir, 1) <> "\" Then CurrDir = CurrDir & "\"

'   Put column headings on active sheet
    Cells(1, 1) = "Path"
    Cells(1, 2) = "Filename"
    Cells(1, 3) = "Size"
    Cells(1, 4) = "Date/Time"
    Range("A1:D1").Font.Bold = True
    
'   Get files
    On Error Resume Next
    FileName = Dir(CurrDir & "*.*", vbDirectory)
    Do While Len(FileName) <> 0
      If Left(FileName, 1) <> "." Then 'Current dir
        PathAndName = CurrDir & FileName
        If (GetAttr(PathAndName) And vbDirectory) = vbDirectory Then
          'store found directories
           ReDim Preserve Dirs(0 To NumDirs) As String
           Dirs(NumDirs) = PathAndName
           NumDirs = NumDirs + 1
        Else
          'Write the path and file to the sheet
          Cells(WorksheetFunction.CountA(Range("A:A")) + 1, 1) = CurrDir
          Cells(WorksheetFunction.CountA(Range("B:B")) + 1, 2) = FileName
          'adjust for filesize > 2 gigabytes
          Filesize = FileLen(PathAndName)
          If Filesize < 0 Then Filesize = Filesize + 4294967296#
          Cells(WorksheetFunction.CountA(Range("C:C")) + 1, 3) = Filesize
          Cells(WorksheetFunction.CountA(Range("D:D")) + 1, 4) = FileDateTime(PathAndName)
        End If
    End If
        FileName = Dir()
    Loop
    ' Process the found directories, recursively
    For i = 0 To NumDirs - 1
        RecursiveDir Dirs(i)
    Next i
End Sub

