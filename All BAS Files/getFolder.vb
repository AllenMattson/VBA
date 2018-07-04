Function getFolder(Optional dialogTitle As String = vbNullString, _
                   Optional dialogButtonName As String = vbNullString, _
                   Optional dialogStartFolder As String = vbNullString, _
                   Optional dialogView As MsoFileDialogView = msoFileDialogViewList) As String
    
  '****************************************************************************
  ' Description:  Returns the selected folder path as a string value
  '
  ' Author:       taxbender
  ' Contributors:
  ' Sources:
  ' Last Updated: 12/30/2015
  ' Dependencies: Var - cEnableErrorHandling
  ' Known Issues: None
  '****************************************************************************
  
  If cEnableErrorHandling Then On Error Resume Next
 
  Dim folderSelection As Variant
  
  With Application.FileDialog(msoFileDialogFolderPicker)
    
    If dialogTitle <> vbNullString Then: .Title = dialogTitle
    If dialogButtonName <> vbNullString Then: .ButtonName = dialogButtonName
      
    .InitialView = dialogView
    .AllowMultiSelect = False
    
    If dialogStartFolder <> vbNullString Then
      If Dir(dialogStartFolder, vbDirectory) <> vbNullString Then
        If Right(dialogStartFolder, 1) <> "\" Then
          dialogStartFolder = dialogStartFolder & "\"
          
          '*** Set initial directory to input value if it exists
          .InitialFileName = dialogStartFolder
        
        End If
      Else
        
        '*** Set initial diretory to same directory as file
        .InitialFileName = CurDir
      
      End If
    End If
    
    .Show
    
    Err.Clear
  
    '*** Set to selected item; if cancel will cause error
    folderSelection = .SelectedItems(1)
    
    If Err.Number <> 0 Then: folderSelection = vbNullString
  
  End With
  
  '*** Set function to string value of folderSelection
  getFolder = CStr(folderSelection)
  
End Function
