Public Function checkFolder(strFolderPath As String) As Boolean

  '******************************************************************************
  ' Description:  Check if folder path exists
  '
  ' Author:       taxbender
  ' Contributors:
  ' Sources:
  ' Last Updated: 12/30/2015
  ' Dependencies: Runtime Reference:  Microsoft Scripting Runtime
  '               Var - cEnableErrorHandling
  '               Sub - errMessage
  ' Known Issues: None
  '******************************************************************************

  If cEnableErrorHandling Then: On Error GoTo errHandler
  
  Dim fso As Scripting.FileSystemObject
  
  Set fso = New FileSystemObject
      
  If fso.FolderExists(strFolderPath) <> 0 Then: checkFolder = True
  
exitMe:
  Set fso = Nothing
  Exit Function
  
errHandler:
  errMessage "checkFolder", Err.Number, Err.Description
  Resume exitMe

End Function
