Attribute VB_Name = "Module1"
'Option Explicit

Sub GetAFolder()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Select a location for the backup."
        .Show
        If .SelectedItems.Count = 0 Then
            MsgBox "Canceled"
        Else
            MsgBox .SelectedItems(1)
        End If
    End With
End Sub


