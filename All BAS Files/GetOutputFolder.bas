Function GetOutputFolder(Optional startFolder As Variant = -1) As Variant
    
    Dim fldr As FileDialog
    Dim vItem As Variant
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a Folder"
        .AllowMultiSelect = False
        If startFolder = -1 Then
            .InitialFileName = Application.DefaultFilePath
        Else
            If Right(startFolder, 1) <> "\" Then
                .InitialFileName = startFolder & "\"
            Else
                .InitialFileName = startFolder
            End If
        End If
        If .Show <> -1 Then GoTo NextCode
        vItem = .SelectedItems(1)
    End With
NextCode:
    GetOutputFolder = vItem
    Set fldr = Nothing
End Function
