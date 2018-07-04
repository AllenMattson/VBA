Public Function getMyDocsPath() As String
  
  '******************************************************************************
  ' Description:  Retreives current user's My Documents path without trailing "/"
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

  If cEnableErrorHandling Then On Error GoTo errHandler
 
  Dim oShell As Object
  
  Set oShell = CreateObject("WScript.Shell")
  
  getMyDocsPath = oShell.SpecialFolders("mydocuments")
  
exitMe:
  Set oShell = Nothing
  Exit Function

errHandler:
  errMessage "getMyDocsPath", Err.Number, Err.Description
  Resume exitMe
  
End Function
