Attribute VB_Name = "Module1"
Option Explicit

Sub FileInfo()
    Dim c As Long, r As Long, i As Long
    Dim Directory As String
    Dim FileName As Object 'FolderItem2
    Dim objShell As Object 'IShellDispatch5
    Dim objFolder As Object 'Folder3
    Dim d
    
    r = 1
'   Create the object
    Set objShell = CreateObject("Shell.Application")
    d = GetDirectory
    If d = False Then Exit Sub 'canceled
    Set objFolder = objShell.Namespace(d)

'   Insert headers on active sheet
    Worksheets.Add
    c = 0
    For i = 0 To 40
        c = c + 1
        Cells(1, c) = objFolder.GetDetailsOf(objFolder.Items, i)
    Next i

'   Loop through the files
    r = 1
    For Each FileName In objFolder.Items
        c = 0
        r = r + 1
        For i = 0 To 40
            c = c + 1
            Cells(r, c) = objFolder.GetDetailsOf(FileName, i)
        Next i
    Next FileName
'   Make it a table
    ActiveSheet.ListObjects.Add xlSrcRange, _
      Range("A1").CurrentRegion
End Sub

Function GetDirectory()
'   Prompt for the folder
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Select a location containing the files you want to list."
        .Show
        If .SelectedItems.Count = 0 Then
            GetDirectory = False
        Else
            GetDirectory = .SelectedItems(1)
        End If
    End With
End Function
